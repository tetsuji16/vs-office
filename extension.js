'use strict';

const vscode = require('vscode');
const path = require('path');
const crypto = require('crypto');
const {
  OfficePackage,
  OfficePackageError,
  MAX_FILE_BYTES,
} = require('./src/office-package');

function nonce() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let value = '';
  for (let i = 0; i < 32; i += 1) value += chars[Math.floor(Math.random() * chars.length)];
  return value;
}

class VsOfficeDocument {
  constructor(uri, officePackage) {
    this.uri = uri;
    this.officePackage = officePackage;
    this.disposed = false;
    this.webviews = new Set();
  }

  dispose() {
    this.disposed = true;
  }
}

class VsOfficeProvider {
  constructor(context) {
    this.context = context;
    this.changeEmitter = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this.changeEmitter.event;
  }

  async openCustomDocument(uri) {
    const stat = await vscode.workspace.fs.stat(uri);
    if (stat.size > MAX_FILE_BYTES) {
      throw new OfficePackageError(`安全のため ${Math.round(MAX_FILE_BYTES / 1024 / 1024)} MB を超えるファイルは開きません。`);
    }
    const bytes = await vscode.workspace.fs.readFile(uri);
    const pkg = await OfficePackage.open(uri.fsPath, bytes);
    return new VsOfficeDocument(uri, pkg);
  }

  async resolveCustomEditor(document, webviewPanel) {
    const webview = webviewPanel.webview;
    document.webviews.add(webview);
    webviewPanel.onDidDispose(() => document.webviews.delete(webview));
    webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'media'),
        vscode.Uri.joinPath(this.context.extensionUri, 'node_modules', 'docx-preview', 'dist'),
        vscode.Uri.joinPath(this.context.extensionUri, 'node_modules', 'jszip', 'dist'),
      ],
    };

    const scriptUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'media', 'viewer.js'));
    const styleUri = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'media', 'viewer.css'));
    const docxPreviewUri = webview.asWebviewUri(vscode.Uri.joinPath(
      this.context.extensionUri,
      'node_modules',
      'docx-preview',
      'dist',
      'docx-preview.min.js',
    ));
    const jszipUri = webview.asWebviewUri(vscode.Uri.joinPath(
      this.context.extensionUri,
      'node_modules',
      'jszip',
      'dist',
      'jszip.min.js',
    ));
    const token = nonce();
    webview.html = this.html(webview, scriptUri, styleUri, jszipUri, docxPreviewUri, token);

    const sendModel = async () => {
      const model = await document.officePackage.createViewModel();
      await webview.postMessage({ type: 'model', model });
      if (model.kind === 'docx') {
        const source = document.officePackage.originalBytes;
        await webview.postMessage({
          type: 'docx-bytes',
          bytes: source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
        });
      }
    };

    webview.onDidReceiveMessage(async (message) => {
      try {
        if (message.type === 'ready') {
          await sendModel();
        } else if (message.type === 'stage-edit') {
          const hadPrevious = document.officePackage.edits.has(message.edit.id);
          const previous = document.officePackage.edits.get(message.edit.id);
          document.officePackage.stageEdit(message.edit);
          this.changeEmitter.fire({
            document,
            label: 'Edit Office content',
            undo: async () => {
              if (hadPrevious) document.officePackage.stageEdit({ id: message.edit.id, value: previous });
              else document.officePackage.unstageEdit(message.edit.id);
              await this.sendDocument(document);
            },
            redo: async () => {
              document.officePackage.stageEdit(message.edit);
              await this.sendDocument(document);
            },
          });
          await sendModel();
        } else if (message.type === 'discard') {
          await vscode.commands.executeCommand('workbench.action.files.revert');
        } else if (message.type === 'save') {
          await vscode.commands.executeCommand('workbench.action.files.save');
        } else if (message.type === 'save-as') {
          await vscode.commands.executeCommand('workbench.action.files.saveAs');
        }
      } catch (error) {
        const text = error instanceof Error ? error.message : String(error);
        await webview.postMessage({ type: 'error', message: text });
      }
    });
  }

  async saveCustomDocument(document) {
    await this.writeValidated(document, document.uri, true);
  }

  async saveCustomDocumentAs(document, destination) {
    await this.writeValidated(document, destination, false);
  }

  async revertCustomDocument(document) {
    const bytes = await vscode.workspace.fs.readFile(document.uri);
    document.officePackage = await OfficePackage.open(document.uri.fsPath, bytes);
    await this.sendDocument(document);
  }

  async backupCustomDocument(document, context) {
    const result = document.officePackage.hasEdits()
      ? await document.officePackage.exportValidatedCopy()
      : { bytes: document.officePackage.originalBytes };
    await vscode.workspace.fs.writeFile(context.destination, result.bytes);
    return {
      id: context.destination.toString(),
      delete: () => vscode.workspace.fs.delete(context.destination).then(undefined, () => undefined),
    };
  }

  async sendDocument(document) {
    const model = await document.officePackage.createViewModel();
    await Promise.all(Array.from(document.webviews, (webview) => webview.postMessage({ type: 'model', model })));
  }

  async writeValidated(document, target, overwriteOriginal) {
    const pkg = document.officePackage;
    if (!pkg.hasEdits()) {
      vscode.window.showInformationMessage('変更はありません。');
      return;
    }
    if (!pkg.canEdit) {
      throw new OfficePackageError(pkg.editBlockReason || 'このファイルは安全編集できません。');
    }

    const result = await pkg.exportValidatedCopy();
    const parsed = path.parse(target.fsPath);
    const temp = vscode.Uri.file(path.join(parsed.dir, `.${parsed.base}.${crypto.randomUUID()}.tmp`));
    let backup;
    try {
      if (overwriteOriginal) {
        const current = await vscode.workspace.fs.readFile(target);
        const openedHash = crypto.createHash('sha256').update(pkg.originalBytes).digest('hex');
        const currentHash = crypto.createHash('sha256').update(current).digest('hex');
        if (openedHash !== currentHash) {
          throw new OfficePackageError('ファイルが外部で変更されています。再読み込みするか「名前を付けて保存」を選んでください。');
        }
      }
      if (overwriteOriginal && vscode.workspace.getConfiguration('vsOffice').get('createBackupOnSave', true)) {
        const backupDir = vscode.Uri.file(path.join(parsed.dir, '.vs-office-backups'));
        await vscode.workspace.fs.createDirectory(backupDir);
        const stamp = new Date().toISOString().replace(/[:.]/g, '-');
        backup = vscode.Uri.joinPath(backupDir, `${parsed.name}.${stamp}${parsed.ext}`);
        await vscode.workspace.fs.copy(target, backup, { overwrite: false });
      }
      await vscode.workspace.fs.writeFile(temp, result.bytes);
      const written = await vscode.workspace.fs.readFile(temp);
      await OfficePackage.open(target.fsPath, written);
      await vscode.workspace.fs.rename(temp, target, { overwrite: true });
    } catch (error) {
      await vscode.workspace.fs.delete(temp).then(undefined, () => undefined);
      throw error;
    }

    if (overwriteOriginal) {
      document.officePackage = await OfficePackage.open(document.uri.fsPath, result.bytes);
      await this.sendDocument(document);
    }
    await Promise.all(Array.from(document.webviews, (webview) => webview.postMessage({
      type: 'saved',
      report: result.report,
      path: target.fsPath,
      backupPath: backup?.fsPath,
    })));
    vscode.window.showInformationMessage(`保存しました: ${path.basename(target.fsPath)}`);
  }

  html(webview, scriptUri, styleUri, jszipUri, docxPreviewUri, token) {
    const csp = [
      "default-src 'none'",
      `img-src ${webview.cspSource} data: blob:`,
      `font-src ${webview.cspSource} data:`,
      `style-src ${webview.cspSource} 'unsafe-inline'`,
      `script-src 'nonce-${token}'`,
    ].join('; ');
    return `<!doctype html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy" content="${csp}">
  <link rel="stylesheet" href="${styleUri}">
  <title>VS Office</title>
</head>
<body>
  <header id="toolbar">
    <div>
      <strong id="filename">VS Office</strong>
      <span id="status" class="badge">読み込み中</span>
    </div>
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
  <div id="toast" role="status" aria-live="polite"></div>
  <div id="toasts" role="status" aria-live="polite"></div>
  <script nonce="${token}" src="${jszipUri}"></script>
  <script nonce="${token}" src="${docxPreviewUri}"></script>
  <script nonce="${token}" src="${scriptUri}"></script>
</body>
</html>`;
  }
}

function activate(context) {
  const provider = new VsOfficeProvider(context);
  context.subscriptions.push(vscode.window.registerCustomEditorProvider(
    'vsOffice.editor',
    provider,
    { supportsMultipleEditorsPerDocument: false, webviewOptions: { retainContextWhenHidden: true } },
  ));
  context.subscriptions.push(vscode.commands.registerCommand('vsOffice.openOfficeFile', async () => {
    const files = await vscode.window.showOpenDialog({
      canSelectMany: false,
      filters: { 'Office Open XML': ['docx', 'pptx', 'xlsx'] },
    });
    if (files?.[0]) await vscode.commands.executeCommand('vscode.openWith', files[0], 'vsOffice.editor');
  }));
}

function deactivate() {}

module.exports = { activate, deactivate };
