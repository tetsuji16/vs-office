'use strict';

const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { OfficePackage } = require(path.join(__dirname, '..', 'src', 'office-package'));

const ROOT = path.join(__dirname, '..');
const token = 'test-nonce';

async function buildModel(file) {
  const bytes = fs.readFileSync(path.join(ROOT, 'samples', file));
  const pkg = await OfficePackage.open(file, bytes);
  const model = await pkg.createViewModel();
  return { model, bytes: Array.from(bytes) };
}

function htmlFor(kind) {
  const css = fs.readFileSync(path.join(ROOT, 'media', 'viewer.css'), 'utf8');
  const viewerJs = fs.readFileSync(path.join(ROOT, 'media', 'viewer.js'), 'utf8');
  const jszipSrc = fs.readFileSync(path.join(ROOT, 'node_modules', 'jszip', 'dist', 'jszip.min.js'), 'utf8');
  const docxPreviewSrc = fs.readFileSync(path.join(ROOT, 'node_modules', 'docx-preview', 'dist', 'docx-preview.js'), 'utf8');
  return `<!doctype html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <style>${css}</style>
  <title>VS Office ${kind}</title>
</head>
<body>
  <header id="toolbar">
    <div><strong id="filename">VS Office</strong><span id="status" class="badge">読み込み中</span></div>
    <div class="actions">
      <button id="discard" disabled>元に戻す</button>
      <button id="saveAs" disabled>名前を付けて保存</button>
      <button id="save" class="primary" disabled>保存 (Ctrl+S)</button>
    </div>
  </header>
  <div id="notice"></div>
  <nav id="tabs" hidden>
    <button data-tab="preview" class="active">プレビュー</button>
    <button data-tab="outline">安全編集</button>
  </nav>
  <main>
    <section id="preview" class="tab-panel"></section>
    <section id="outline" class="tab-panel" hidden></section>
  </main>
  <div id="toast" role="status"></div>
  <div id="toasts" role="status"></div>
  <script nonce="${token}">${jszipSrc}</script>
  <script nonce="${token}">${docxPreviewSrc}</script>
  <script nonce="${token}">
    window.__posted = [];
    window.acquireVsCodeApi = function () {
      return { postMessage: function (message) { window.__posted.push(message); }, getState: function () { return undefined; }, setState: function () {} };
    };
  </script>
  <script nonce="${token}">${viewerJs}</script>
</body>
</html>`;
}

(async () => {
  const files = { docx: 'sample.docx', pptx: 'sample.pptx', xlsx: 'sample.xlsx' };
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 1600 } });
  const errors = [];
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', (e) => errors.push('PAGEERROR: ' + e.message));

  for (const [kind, file] of Object.entries(files)) {
    const { model, bytes } = await buildModel(file);
    const payload = {
      type: 'model',
      model: { ...model, canEdit: true, editBlockReason: '', filename: file, editCount: 0 },
    };
    const html = htmlFor(kind);
    await page.setContent(html, { waitUntil: 'networkidle' });
    await page.evaluate((p) => {
      window.dispatchEvent(new MessageEvent('message', { data: p }));
    }, payload);
    if (kind === 'docx') {
      await page.evaluate((b) => {
        window.dispatchEvent(new MessageEvent('message', { data: { type: 'docx-bytes', bytes: b } }));
      }, bytes);
    }
    await page.waitForTimeout(1500);
    const statusText = await page.locator('#status').textContent();
    if (statusText !== '編集可能') throw new Error(`${kind}: unexpected editable status: ${statusText}`);
    const tabsVisible = await page.locator('#tabs').isVisible();
    if (kind === 'xlsx' ? tabsVisible : !tabsVisible) {
      throw new Error(`${kind}: tab visibility does not match the active file type`);
    }
    const out = path.join(ROOT, 'samples', `render-${kind}.png`);
    await page.screenshot({ path: out, fullPage: true });
    console.log('captured', out);

    // Switch to the "安全編集" (edit) tab and capture it too (DOCX/PPTX only).
    if (kind !== 'xlsx') {
      await page.evaluate(() => {
        const btn = document.querySelector('#tabs button[data-tab="outline"]');
        if (btn) btn.click();
      });
      await page.waitForTimeout(500);
      const outEdit = path.join(ROOT, 'samples', `render-${kind}-edit.png`);
      await page.screenshot({ path: outEdit, fullPage: true });
      console.log('captured', outEdit);

      if (kind === 'pptx') {
        const labels = await page.locator('.color-field span').allTextContents();
        if (!labels.includes('塗り') || !labels.includes('枠線')) throw new Error('pptx: visible colour labels missing');
        const deleteButton = page.locator('.shape-row button').first();
        await deleteButton.click();
        if ((await deleteButton.textContent()) !== '再クリック') throw new Error('pptx: delete confirmation was not armed');
        const stagedAfterFirstClick = await page.evaluate(() => window.__posted.some((message) => message?.edit?.id?.endsWith(':delete')));
        if (stagedAfterFirstClick) throw new Error('pptx: first delete click staged a destructive edit');
        const outConfirm = path.join(ROOT, 'samples', 'render-pptx-delete-confirm.png');
        await page.screenshot({ path: outConfirm, fullPage: true });
        console.log('captured', outConfirm);
      }
    }

    // DOM assertions for key correctness signals.
    if (kind === 'xlsx') {
      const formulaBadge = await page.evaluate(() => {
        const badges = Array.from(document.querySelectorAll('td .formula, span.formula'));
        return badges.map((b) => b.textContent);
      });
      console.log('xlsx formula badges:', JSON.stringify(formulaBadge));
      const cellValues = await page.evaluate(() =>
        Array.from(document.querySelectorAll('td input')).map((i) => i.value));
      console.log('xlsx cell values:', JSON.stringify(cellValues));
      const badgeContained = await page.evaluate(() => {
        const badge = document.querySelector('td .cell-badge.formula');
        if (!badge) return false;
        const cell = badge.closest('td');
        const badgeRect = badge.getBoundingClientRect();
        const cellRect = cell.getBoundingClientRect();
        return badgeRect.left >= cellRect.left && badgeRect.right <= cellRect.right
          && badgeRect.top >= cellRect.top && badgeRect.bottom <= cellRect.bottom;
      });
      if (!badgeContained) throw new Error('xlsx: formula badge is clipped or outside its cell');
    }

    await page.setViewportSize({ width: 560, height: 900 });
    const pageOverflow = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth);
    if (pageOverflow) throw new Error(`${kind}: narrow viewport causes page-level horizontal overflow`);
    const narrowOut = path.join(ROOT, 'samples', `render-${kind}-narrow.png`);
    await page.screenshot({ path: narrowOut, fullPage: true });
    console.log('captured', narrowOut);
    await page.setViewportSize({ width: 1280, height: 1600 });
  }
  if (errors.length) {
    console.log('--- console/page errors ---');
    [...new Set(errors)].forEach((e) => console.log(e));
  } else {
    console.log('no console/page errors');
  }
  await browser.close();
})();
