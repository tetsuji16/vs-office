(function () {
  'use strict';
  const vscode = acquireVsCodeApi();
  let model = null;
  let docxBytes = null;

  const $ = (selector) => document.querySelector(selector);
  const filename = $('#filename');
  const status = $('#status');
  const notice = $('#notice');
  const preview = $('#preview');
  const outline = $('#outline');
  const tabs = $('#tabs');
  const saveButton = $('#save');
  const saveAsButton = $('#saveAs');
  const discardButton = $('#discard');
  const toast = $('#toast');

  function text(tag, value, className) {
    const el = document.createElement(tag);
    el.textContent = value;
    if (className) el.className = className;
    return el;
  }

  function showToast(message, error = false) {
    toast.textContent = message;
    toast.className = error ? 'show error' : 'show';
    setTimeout(() => { toast.className = ''; }, 5000);
  }

  function editRow(item, label) {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.appendChild(text('span', label, 'row-label'));
    const input = document.createElement('input');
    input.value = item.value;
    input.disabled = !model.canEdit || item.formula;
    input.title = item.formula ? '数式セルは閲覧のみです' : 'Enterで変更をステージします';
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && input.value !== item.value) {
        vscode.postMessage({ type: 'stage-edit', edit: { id: item.id, value: input.value } });
      }
    });
    input.addEventListener('change', () => {
      if (input.value !== item.value) vscode.postMessage({ type: 'stage-edit', edit: { id: item.id, value: input.value } });
    });
    row.appendChild(input);
    if (item.formula) row.appendChild(text('span', 'ƒx', 'formula'));
    return row;
  }

  function renderCommon() {
    filename.textContent = model.filename;
    status.textContent = model.editCount ? `未保存 ${model.editCount}件` : '閲覧モード';
    status.className = model.editCount ? 'badge changed' : 'badge';
    saveButton.disabled = !model.editCount || !model.canEdit;
    saveAsButton.disabled = !model.editCount || !model.canEdit;
    discardButton.disabled = !model.editCount;
    notice.innerHTML = '';
    const message = model.canEdit
      ? 'Ctrl+Sで元ファイルへ保存できます。一時ファイルを検証してから原子的に置換し、未編集パートのSHA-256一致とZIP CRCを確認します。'
      : model.editBlockReason;
    notice.appendChild(text('div', message, model.canEdit ? 'notice safe' : 'notice warning'));
  }

  function renderDocx() {
    tabs.hidden = false;
    outline.innerHTML = '';
    outline.appendChild(text('h2', '段落の安全編集'));
    outline.appendChild(text('p', '書式（太字・斜体等）は各 run のまま保持し、テキストのみを段落単位で置換します。文字数に応じて run へ自動配分され、改行による通常の再配置は起こり得ます。', 'muted'));
    const list = document.createElement('div');
    list.className = 'edit-list';
    model.paragraphs.forEach((para, i) => list.appendChild(editParagraph(para, `#${i + 1}`)));
    outline.appendChild(list);
    renderDocxPreview();
  }

  function editParagraph(para, label) {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.appendChild(text('span', label, 'row-label'));
    const input = document.createElement('input');
    input.value = para.value;
    input.disabled = !model.canEdit;
    input.title = 'Enterで変更をステージします';
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && input.value !== para.value) {
        vscode.postMessage({ type: 'stage-edit', edit: { id: para.id, value: input.value } });
      }
    });
    input.addEventListener('change', () => {
      if (input.value !== para.value) vscode.postMessage({ type: 'stage-edit', edit: { id: para.id, value: input.value } });
    });
    row.appendChild(input);
    return row;
  }

  async function renderDocxPreview() {
    preview.innerHTML = '';
    if (!docxBytes || !window.docx?.renderAsync) {
      preview.appendChild(text('p', '高忠実度プレビューを準備中です…', 'muted'));
      return;
    }
    const container = document.createElement('div');
    container.className = 'docx-container';
    preview.appendChild(container);
    try {
      await window.docx.renderAsync(new Uint8Array(docxBytes).buffer, container, null, {
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreLastRenderedPageBreak: false,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        breakPages: true,
      });
    } catch (error) {
      container.replaceChildren(text('p', `プレビュー失敗: ${error.message}`, 'error-text'));
    }
  }

  function renderPptx() {
    tabs.hidden = true;
    preview.innerHTML = '';
    outline.innerHTML = '';
    const intro = text('p', '軽量構造ビューです。座標・テーマ・マスター・画像は変更せず、既存 a:t テキストだけを編集します。上段がスライド内テキストの視覚的プレビュー、下段が編集フィールドです。', 'muted');
    preview.appendChild(intro);
    const deck = document.createElement('div');
    deck.className = 'slides';
    model.slides.forEach((slide) => {
      const card = document.createElement('article');
      card.className = 'slide';
      const head = document.createElement('div');
      head.className = 'slide-head';
      head.appendChild(text('strong', `Slide ${slide.number}`));
      head.appendChild(text('span', `画像 ${slide.imageCount}`, 'muted'));
      card.appendChild(head);
      const previewBox = document.createElement('div');
      previewBox.className = 'slide-preview';
      if (!slide.items.length) previewBox.appendChild(text('p', '（テキストなし）', 'muted'));
      slide.items.forEach((item) => {
        const line = document.createElement('p');
        line.className = 'slide-line';
        line.textContent = item.value || ' ';
        previewBox.appendChild(line);
      });
      card.appendChild(previewBox);
      const list = document.createElement('div');
      list.className = 'edit-list';
      slide.items.forEach((item, i) => list.appendChild(editRow(item, `${i + 1}`)));
      card.appendChild(list);
      deck.appendChild(card);
    });
    preview.appendChild(deck);
  }

  function colNumber(ref) {
    const letters = ref.replace(/[0-9]/g, '');
    let n = 0;
    for (const char of letters) n = n * 26 + char.charCodeAt(0) - 64;
    return n;
  }

  function renderXlsx() {
    tabs.hidden = true;
    preview.innerHTML = '';
    outline.innerHTML = '';
    const switcher = document.createElement('div');
    switcher.className = 'sheet-tabs';
    const host = document.createElement('div');
    preview.append(switcher, host);

    function showSheet(sheet, activeButton) {
      switcher.querySelectorAll('button').forEach((b) => b.classList.toggle('active', b === activeButton));
      host.innerHTML = '';
      const maxRow = Math.min(200, Math.max(1, ...sheet.cells.map((c) => Number(c.ref.match(/[0-9]+/)[0]))));
      const maxCol = Math.min(50, Math.max(1, ...sheet.cells.map((c) => colNumber(c.ref))));
      const map = new Map(sheet.cells.map((c) => [c.ref, c]));
      const tableWrap = document.createElement('div');
      tableWrap.className = 'grid-wrap';
      const table = document.createElement('table');
      const header = document.createElement('tr');
      header.appendChild(document.createElement('th'));
      for (let col = 1; col <= maxCol; col += 1) {
        let value = '';
        let n = col;
        while (n) { n -= 1; value = String.fromCharCode(65 + (n % 26)) + value; n = Math.floor(n / 26); }
        header.appendChild(text('th', value));
      }
      table.appendChild(header);
      for (let row = 1; row <= maxRow; row += 1) {
        const tr = document.createElement('tr');
        tr.appendChild(text('th', String(row)));
        for (let col = 1; col <= maxCol; col += 1) {
          let letters = '';
          let n = col;
          while (n) { n -= 1; letters = String.fromCharCode(65 + (n % 26)) + letters; n = Math.floor(n / 26); }
          const item = map.get(`${letters}${row}`);
          const td = document.createElement('td');
          if (item) {
            if (item.formula) {
              const input = document.createElement('input');
              input.value = item.value;
              input.disabled = true;
              input.readOnly = true;
              input.className = 'formula-cell';
              input.title = item.formulaText ? `数式 (閲覧のみ): =${item.formulaText}` : '数式セルは閲覧のみです';
              td.appendChild(input);
              td.appendChild(text('span', 'ƒx', 'formula'));
            } else {
              const input = document.createElement('input');
              input.value = item.value;
              input.disabled = !model.canEdit;
              input.addEventListener('change', () => {
                if (input.value !== item.value) vscode.postMessage({ type: 'stage-edit', edit: { id: item.id, value: input.value } });
              });
              td.appendChild(input);
            }
          }
          tr.appendChild(td);
        }
        table.appendChild(tr);
      }
      tableWrap.appendChild(table);
      host.appendChild(tableWrap);
      if (sheet.cells.length > 20000 || maxRow >= 200 || maxCol >= 50) {
        host.appendChild(text('p', '軽量表示のため、最大200行×50列まで表示しています。', 'muted'));
      }
    }

    model.sheets.forEach((sheet, i) => {
      const button = text('button', sheet.name);
      button.addEventListener('click', () => showSheet(sheet, button));
      switcher.appendChild(button);
      if (i === 0) showSheet(sheet, button);
    });
  }

  function render() {
    if (!model) return;
    renderCommon();
    if (model.kind === 'docx') renderDocx();
    else if (model.kind === 'pptx') renderPptx();
    else renderXlsx();
  }

  tabs.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-tab]');
    if (!button) return;
    tabs.querySelectorAll('button').forEach((b) => b.classList.toggle('active', b === button));
    preview.hidden = button.dataset.tab !== 'preview';
    outline.hidden = button.dataset.tab !== 'outline';
  });
  saveButton.addEventListener('click', () => vscode.postMessage({ type: 'save' }));
  saveAsButton.addEventListener('click', () => vscode.postMessage({ type: 'save-as' }));
  discardButton.addEventListener('click', () => vscode.postMessage({ type: 'discard' }));

  window.addEventListener('message', (event) => {
    const message = event.data;
    if (message.type === 'model') {
      model = message.model;
      render();
    } else if (message.type === 'docx-bytes') {
      docxBytes = message.bytes;
      renderDocxPreview();
    } else if (message.type === 'error') {
      showToast(message.message, true);
    } else if (message.type === 'saved') {
      const backup = message.backupPath ? ' · バックアップ作成済み' : '';
      showToast(`保存・検証完了: 未編集 ${message.report.untouchedPartsVerified} パート一致${backup}`);
    }
  });

  vscode.postMessage({ type: 'ready' });
}());
