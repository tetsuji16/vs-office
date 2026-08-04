'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM } = require('jsdom');

// Minimal webview HTML shell that viewer.js expects to find by id.
const SHELL = `<!doctype html><html><head></head><body>
  <header id="toolbar">
    <div><strong id="filename">VS Office</strong><span id="status" class="badge"></span></div>
    <div class="actions">
      <button id="discard">元に戻す</button>
      <button id="saveAs">名前を付けて保存</button>
      <button id="save" class="primary">保存</button>
    </div>
  </header>
  <div id="notice"></div>
  <nav id="tabs"><button data-tab="preview" class="active">レイアウト表示</button><button data-tab="outline">安全編集</button></nav>
  <main>
    <section id="preview" class="tab-panel"></section>
    <section id="outline" class="tab-panel" hidden></section>
  </main>
  <div id="toast" role="status"></div>
</body></html>`;

function setupViewer(model) {
  const dom = new JSDOM(SHELL, { runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  const posted = [];
  const vscodeApi = {
    postMessage: (msg) => posted.push(msg),
    posted,
    getState: () => undefined,
    setState: () => undefined,
  };
  // viewer.js calls the global acquireVsCodeApi(); expose it on the Node global.
  global.acquireVsCodeApi = () => vscodeApi;
  window.acquireVsCodeApi = () => vscodeApi;
  // Expose jsdom globals that viewer.js touches.
  global.window = window;
  global.document = window.document;
  // viewer.js is an IIFE that runs once on load; clear the require cache so each
  // test gets a fresh module bound to this test's jsdom globals.
  delete require.cache[require.resolve('../media/viewer.js')];
  // Load viewer.js (IIFE runs immediately, querying the DOM).
  require('../media/viewer.js');
  // Dispatch the model message the way the webview receives it.
  window.dispatchEvent(new window.MessageEvent('message', { data: { type: 'model', model } }));
  return { window, vscode: vscodeApi };
}

test('viewer renders an XLSX grid and marks edited cells', () => {
  const model = {
    kind: 'xlsx',
    filename: 'book.xlsx',
    canEdit: true,
    editCount: 1,
    sheets: [{
      name: 'Data',
      part: 'xl/worksheets/sheet1.xml',
      cells: [
        { id: 'xlsx:0:0', ref: 'A1', type: 'n', formula: false, formulaText: '', value: 'Hello', text: 'Hello' },
        { id: 'xlsx:0:1', ref: 'B1', type: 'n', formula: true, formulaText: 'SUM(A1:A3)', value: '=SUM(A1:A3)', text: '=SUM(A1:A3)' },
        { id: 'xlsx:0:2', ref: 'C1', type: 's', formula: false, formulaText: '', value: 'Changed', text: 'Original' },
      ],
    }],
  };
  const { window } = setupViewer(model);
  const doc = window.document;
  // Grid rendered with the cell values.
  assert.ok(doc.querySelector('table'), 'xlsx grid not rendered');
  const inputs = Array.from(doc.querySelectorAll('td input'));
  assert.ok(inputs.some((i) => i.value === 'Hello'), 'cell value not rendered');
  // Edited cell (C1: value !== text) shows the "編集済み" badge.
  const badges = Array.from(doc.querySelectorAll('.edited-badge')).map((b) => b.textContent);
  assert.ok(badges.includes('編集済み'), 'edited badge not shown for changed cell');
  // Formula cell is no longer read-only: input is enabled and not .formula-cell.
  const formulaInput = Array.from(doc.querySelectorAll('td input')).find((i) => i.value === '=SUM(A1:A3)');
  assert.ok(formulaInput, 'formula value not rendered');
  assert.equal(formulaInput.disabled, false, 'formula cell should be editable now');
});

test('viewer renders a DOCX paragraph outline with edited badge', () => {
  const model = {
    kind: 'docx',
    filename: 'doc.docx',
    canEdit: true,
    editCount: 1,
    paragraphs: [{
      id: 'docx:p0',
      part: 'word/document.xml',
      ordinal: 0,
      text: 'Original paragraph',
      value: 'Edited paragraph',
      tIds: ['docx:0', 'docx:1'],
      tLengths: [8, 10],
    }],
    summary: '1 段落',
  };
  const { window } = setupViewer(model);
  const doc = window.document;
  // The paragraph value is rendered into an editable input.
  const input = doc.querySelector('input.edited-input') || Array.from(doc.querySelectorAll('#outline input')).find((i) => i.value === 'Edited paragraph');
  assert.ok(input, 'paragraph value not rendered into an input');
  assert.equal(input.value, 'Edited paragraph');
  const badges = Array.from(doc.querySelectorAll('.edited-badge')).map((b) => b.textContent);
  assert.ok(badges.includes('編集済み'), 'docx edited badge missing');
});

test('viewer reflects read-only mode when canEdit is false', () => {
  const model = {
    kind: 'xlsx',
    filename: 'signed.xlsx',
    canEdit: false,
    editBlockReason: '署名付き',
    editCount: 0,
    sheets: [{
      name: 'Sheet1',
      part: 'xl/worksheets/sheet1.xml',
      cells: [{ id: 'xlsx:0:0', ref: 'A1', type: 'n', formula: false, formulaText: '', value: 'x', text: 'x' }],
    }],
  };
  const { window } = setupViewer(model);
  const doc = window.document;
  const input = doc.querySelector('td input');
  assert.equal(input.disabled, true, 'cell should be disabled in read-only mode');
  assert.ok(doc.body.textContent.includes('署名付き'), 'block reason not shown');
});

test('viewer renders PPTX shapes as editable cards with colour/delete controls', () => {
  const model = {
    kind: 'pptx',
    filename: 'deck.pptx',
    canEdit: true,
    editBlockReason: '',
    editCount: 0,
    slides: [{
      number: 1,
      imageCount: 0,
      items: [{ id: 'pptx:0:0', value: 'Title', text: 'Title' }],
      shapes: [{
        id: 'pptx:0:0',
        type: 'shape',
        text: 'Title',
        value: 'Title',
        fillColor: 'ff0000',
        lineColor: '00ff00',
        fillValue: '#ff0000',
        lineValue: '#00ff00',
        deleted: false,
      }],
    }],
  };
  const { window, vscode } = setupViewer(model);
  const doc = window.document;
  const rows = doc.querySelectorAll('.shape-row');
  assert.equal(rows.length, 1, 'one shape card expected');
  const textInput = rows[0].querySelector('input:not([type="color"])');
  assert.equal(textInput.value, 'Title', 'shape text input value');
  const colors = rows[0].querySelectorAll('input[type="color"]');
  assert.equal(colors.length, 2, 'fill + line colour inputs');
  assert.equal(colors[0].value, '#ff0000');
  assert.equal(colors[1].value, '#00ff00');
  const del = rows[0].querySelector('button');
  assert.equal(del.textContent, '削除');
  // Changing a colour should post a stage-edit with the :fill id.
  colors[0].value = '#123456';
  colors[0].dispatchEvent(new window.Event('change'));
  assert.ok(vscode.posted.some((m) => m.type === 'stage-edit' && m.edit.id === 'pptx:0:0:fill' && m.edit.value === '#123456'), 'fill edit not posted');
  // Clicking delete should post a delete stage-edit.
  del.dispatchEvent(new window.Event('click'));
  assert.ok(vscode.posted.some((m) => m.type === 'stage-edit' && m.edit.id === 'pptx:0:0:delete'), 'delete edit not posted');
});
