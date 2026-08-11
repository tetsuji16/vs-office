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
    const stack = document.getElementById('toasts') || toast.parentElement;
    const el = document.createElement('div');
    el.className = error ? 'toast-item error' : 'toast-item';
    el.textContent = message;
    stack.appendChild(el);
    setTimeout(() => { el.classList.add('hide'); }, 5000);
    setTimeout(() => { el.remove(); }, 5500);
  }

  function editRow(item, label) {
    const row = document.createElement('div');
    row.className = 'edit-row';
    row.appendChild(text('span', label, 'row-label'));
    const input = document.createElement('input');
    input.value = item.value;
    input.setAttribute('aria-label', `${label} ${item.formula ? '数式' : '値'}`);
    input.disabled = !model.canEdit || item.formula;
    const edited = !item.formula && item.value !== item.text;
    if (edited) {
      row.classList.add('edited');
      input.classList.add('edited-input');
    }
    input.title = item.formula
      ? '数式セルは閲覧のみです'
      : edited
        ? `変更前: ${item.text}\n変更後: ${item.value}`
        : 'Enterで変更をステージします';
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
    else if (edited) row.appendChild(text('span', '編集済み', 'edited-badge'));
    return row;
  }

  function renderCommon() {
    filename.textContent = model.filename;
    status.textContent = model.editCount
      ? `未保存 ${model.editCount}件`
      : model.canEdit
        ? '編集可能'
        : '閲覧のみ';
    status.className = model.editCount ? 'badge changed' : 'badge';
    saveButton.disabled = !model.editCount || !model.canEdit;
    saveAsButton.disabled = !model.editCount || !model.canEdit;
    discardButton.disabled = !model.editCount;
    notice.innerHTML = '';
    const message = model.canEdit
      ? 'Ctrl+Sで元ファイルへ保存できます。一時ファイルを検証してから原子的に置換し、未編集パーツのSHA-256一致とZIP CRCを確認します。'
      : model.editBlockReason;
    notice.appendChild(text('div', message, model.canEdit ? 'notice safe' : 'notice warning'));
    // Tabs are shown for DOCX (preview/edit) and PPTX (preview/edit). XLSX edits
    // inline, so its single grid view needs no tab switch.
    tabs.hidden = model.kind === 'xlsx';
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
    input.setAttribute('aria-label', `${label} 段落`);
    const edited = para.value !== para.text;
    if (edited) {
      row.classList.add('edited');
      input.classList.add('edited-input');
    }
    input.disabled = !model.canEdit;
    input.title = edited
      ? `変更前: ${para.text}\n変更後: ${para.value}`
      : 'Enterで変更をステージします';
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && input.value !== para.value) {
        vscode.postMessage({ type: 'stage-edit', edit: { id: para.id, value: input.value } });
      }
    });
    input.addEventListener('change', () => {
      if (input.value !== para.text) vscode.postMessage({ type: 'stage-edit', edit: { id: para.id, value: input.value } });
    });
    row.appendChild(input);
    if (edited) row.appendChild(text('span', '編集済み', 'edited-badge'));
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

  function editShape(shape) {
    const card = document.createElement('div');
    card.className = 'edit-row shape-row';
    const textInput = document.createElement('input');
    textInput.value = shape.value;
    textInput.disabled = !model.canEdit;
    textInput.placeholder = '図形テキスト';
    textInput.title = '図形のテキスト';
    textInput.setAttribute('aria-label', '図形テキスト');
    textInput.addEventListener('change', () => {
      if (textInput.value !== shape.text) vscode.postMessage({ type: 'stage-edit', edit: { id: `${shape.id}:text`, value: textInput.value } });
    });
    card.appendChild(textInput);
    const fill = document.createElement('input');
    fill.type = 'color';
    fill.value = shape.fillValue || '#ffffff';
    fill.disabled = !model.canEdit;
    fill.title = `塗りつぶし色 (元: ${shape.fillColor || 'なし'})`;
    fill.addEventListener('change', () => vscode.postMessage({ type: 'stage-edit', edit: { id: `${shape.id}:fill`, value: fill.value } }));
    fill.setAttribute('aria-label', '塗りつぶし色');
    const fillField = document.createElement('label');
    fillField.className = 'color-field';
    fillField.append(text('span', '塗り'), fill);
    card.appendChild(fillField);
    const line = document.createElement('input');
    line.type = 'color';
    line.value = shape.lineValue || '#000000';
    line.disabled = !model.canEdit;
    line.title = `枠線色 (元: ${shape.lineColor || 'なし'})`;
    line.addEventListener('change', () => vscode.postMessage({ type: 'stage-edit', edit: { id: `${shape.id}:line`, value: line.value } }));
    line.setAttribute('aria-label', '枠線色');
    const lineField = document.createElement('label');
    lineField.className = 'color-field';
    lineField.append(text('span', '枠線'), line);
    card.appendChild(lineField);
    const del = document.createElement('button');
    del.textContent = '削除';
    del.disabled = !model.canEdit;
    del.title = '誤操作防止のため2回クリックで削除をステージします';
    del.setAttribute('aria-label', '図形を削除（確認が必要）');
    let deleteArmed = false;
    const resetDelete = () => {
      deleteArmed = false;
      del.textContent = '削除';
      del.classList.remove('armed');
      del.setAttribute('aria-label', '図形を削除（確認が必要）');
    };
    del.addEventListener('blur', resetDelete);
    del.addEventListener('click', () => {
      if (!deleteArmed) {
        deleteArmed = true;
        del.textContent = '再クリック';
        del.classList.add('armed');
        del.setAttribute('aria-label', '図形の削除を確定');
        return;
      }
      vscode.postMessage({ type: 'stage-edit', edit: { id: `${shape.id}:delete`, value: '' } });
      resetDelete();
    });
    card.appendChild(del);
    return card;
  }

  function renderPptx() {
    preview.innerHTML = '';
    outline.innerHTML = '';
    const intro = text('p', 'スライド内テキストの構造プレビューです。実際の位置・フォント・画像はPowerPointと異なる場合があります。「安全編集」では図形テキスト、塗りつぶし色、枠線色を変更できます。テーマ・マスター・アニメーション・画像は変更しません。', 'muted');
    preview.appendChild(intro);
    const deck = document.createElement('div');
    deck.className = 'slides';
    model.slides.forEach((slide) => {
      const card = document.createElement('article');
      card.className = 'slide';
      const head = document.createElement('div');
      head.className = 'slide-head';
      head.appendChild(text('strong', `Slide ${slide.number}`));
      head.appendChild(text('span', `画像 ${slide.imageCount} · 図形 ${slide.shapes.length}`, 'muted'));
      card.appendChild(head);
      const previewBox = document.createElement('div');
      previewBox.className = 'slide-preview';
      if (!slide.items.length) previewBox.appendChild(text('p', '（テキストなし）', 'muted'));
      slide.items.forEach((item) => {
        const line = document.createElement('p');
        line.className = 'slide-line';
        line.textContent = item.value || ' ';
        previewBox.appendChild(line);
      });
      card.appendChild(previewBox);
      deck.appendChild(card);
    });
    preview.appendChild(deck);

    const editIntro = text('p', '図形ごとの編集フィールド。テキスト・塗りつぶし色・枠線色を直接編集し、図形ごと削除も可能です。', 'muted');
    outline.appendChild(editIntro);
    model.slides.forEach((slide, si) => {
      const secHead = text('h3', `Slide ${slide.number}`, 'slide-edit-head');
      outline.appendChild(secHead);
      const list = document.createElement('div');
      list.className = 'edit-list';
      if (!slide.shapes.length) list.appendChild(text('p', '（図形なし）', 'muted'));
      slide.shapes.forEach((shape, i) => list.appendChild(editShape(shape, i)));
      outline.appendChild(list);
    });
  }

  function colNumber(ref) {
    const letters = ref.replace(/[0-9]/g, '');
    let n = 0;
    for (const char of letters) n = n * 26 + char.charCodeAt(0) - 64;
    return n;
  }

  function renderXlsx() {
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
            const input = document.createElement('input');
            input.value = item.value;
            input.setAttribute('aria-label', `${item.ref} ${item.formula ? '数式' : 'セル値'}`);
            const edited = !item.formula && item.value !== item.text;
            if (edited) {
              input.classList.add('edited-input');
            }
            input.disabled = !model.canEdit;
            input.title = item.formula
              ? `数式セル（編集可）: ${item.formulaText ? '=' + item.formulaText : ''}`
              : edited
                ? `変更前: ${item.text}\n変更後: ${item.value}`
                : 'Enterで変更をステージします';
            input.addEventListener('change', () => {
              if (input.value !== item.value) vscode.postMessage({ type: 'stage-edit', edit: { id: item.id, value: input.value } });
            });
            td.appendChild(input);
            if (item.formula) td.appendChild(text('span', 'ƒx', 'formula cell-badge'));
            else if (edited) td.appendChild(text('span', '変更', 'edited-badge cell-badge'));
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
