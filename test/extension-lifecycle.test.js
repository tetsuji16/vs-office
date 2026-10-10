'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');
const path = require('node:path');
const JSZip = require('jszip');

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

async function makeDocx() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('word/document.xml', '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p></w:body></w:document>');
  zip.file('word/styles.xml', '<w:styles xmlns:w="w"><w:style w:styleId="Normal"/></w:styles>');
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function harness(initialFiles = {}) {
  const files = new Map(Object.entries(initialFiles).map(([name, bytes]) => [path.resolve(name), Buffer.from(bytes)]));
  const dirs = new Set();
  const uri = (filePath) => ({ fsPath: path.resolve(filePath), toString() { return this.fsPath; } });
  const vscode = {
    Uri: {
      file: uri,
      parse: uri,
      joinPath(base, ...parts) { return uri(path.join(base.fsPath, ...parts)); },
    },
    EventEmitter: class { constructor() { this.event = () => ({ dispose() {} }); } fire() {} },
    workspace: {
      fs: {
        async stat(target) { const b = files.get(target.fsPath); if (!b) throw new Error('ENOENT'); return { size: b.length }; },
        async readFile(target) { const b = files.get(target.fsPath); if (!b) throw new Error(`ENOENT ${target.fsPath}`); return Buffer.from(b); },
        async writeFile(target, bytes) { files.set(target.fsPath, Buffer.from(bytes)); },
        async createDirectory(target) { dirs.add(target.fsPath); },
        async copy(from, to) { const b = files.get(from.fsPath); if (!b) throw new Error('ENOENT'); files.set(to.fsPath, Buffer.from(b)); },
        async rename(from, to) { const b = files.get(from.fsPath); if (!b) throw new Error('ENOENT'); files.set(to.fsPath, b); files.delete(from.fsPath); },
        async delete(target) { files.delete(target.fsPath); },
      },
      getConfiguration() { return { get(_key, fallback) { return fallback; } }; },
    },
    window: {
      registerCustomEditorProvider(_id, provider) { captured.provider = provider; return { dispose() {} }; },
      showInformationMessage() {},
      async showOpenDialog() { return undefined; },
    },
    commands: { registerCommand() { return { dispose() {} }; }, async executeCommand() {} },
  };
  const captured = {};
  const originalLoad = Module._load;
  Module._load = function(request, parent, isMain) {
    if (request === 'vscode') return vscode;
    return originalLoad.call(this, request, parent, isMain);
  };
  const extensionPath = require.resolve('../extension');
  delete require.cache[extensionPath];
  const extension = require('../extension');
  Module._load = originalLoad;
  extension.activate({ extensionUri: uri('/extension'), subscriptions: [] });
  return { provider: captured.provider, files, uri };
}

test('backup recovery retains disk baseline and pending edits, then saves recovered edits', async () => {
  const original = await makeDocx();
  const file = path.resolve('/docs/report.docx');
  const backupFile = path.resolve('/docs/report.backup');
  const h = harness({ [file]: original });
  const document = await h.provider.openCustomDocument(h.uri(file));
  const model = await document.officePackage.createViewModel();
  document.officePackage.stageEdit({ id: model.items[0].id, value: 'Recovered text' });

  await h.provider.backupCustomDocument(document, { destination: h.uri(backupFile) });
  const restored = await h.provider.openCustomDocument(h.uri(file), { backupId: backupFile });
  assert.deepEqual(restored.officePackage.originalBytes, original, 'restored package keeps original disk bytes as its baseline');
  assert.equal(restored.officePackage.edits.get(model.items[0].id), 'Recovered text');

  await h.provider.saveCustomDocument(restored);
  const saved = await JSZip.loadAsync(h.files.get(file));
  assert.match(await saved.file('word/document.xml').async('string'), /Recovered text/);
  assert.equal(restored.officePackage.hasEdits(), false);
});

test('no-edit Save As copies identical bytes and rejects a destination with the wrong extension', async () => {
  const original = await makeDocx();
  const source = path.resolve('/docs/source.docx');
  const destination = path.resolve('/docs/copy.docx');
  const h = harness({ [source]: original });
  const document = await h.provider.openCustomDocument(h.uri(source));

  await h.provider.saveCustomDocumentAs(document, h.uri(destination));
  assert.deepEqual(h.files.get(destination), original);
  await assert.rejects(h.provider.saveCustomDocumentAs(document, h.uri('/docs/copy.pptx')), /拡張子は \.docx/);
  assert.equal(h.files.has(path.resolve('/docs/copy.pptx')), false);
});

test('sendDocument refreshes webview DOCX bytes from the package baseline', async () => {
  const original = await makeDocx();
  const source = path.resolve('/docs/preview.docx');
  const h = harness({ [source]: original });
  const document = await h.provider.openCustomDocument(h.uri(source));
  const messages = [];
  document.webviews.add({ async postMessage(message) { messages.push(message); } });

  await h.provider.sendDocument(document);
  const payload = messages.find((message) => message.type === 'docx-bytes');
  assert.ok(payload, 'DOCX refresh includes the source bytes');
  assert.deepEqual(Buffer.from(payload.bytes), original);
  assert.ok(messages.some((message) => message.type === 'model' && message.model.kind === 'docx'));
});
