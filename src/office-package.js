'use strict';

const JSZip = require('jszip');
const crypto = require('crypto');
const path = require('path');

const MAX_FILE_BYTES = 100 * 1024 * 1024;
const MAX_UNCOMPRESSED_BYTES = 500 * 1024 * 1024;
const MAX_ENTRIES = 10000;
const MAX_TEXT_NODES = 50000;

class OfficePackageError extends Error {}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXml(value) {
  return String(value)
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&#([0-9]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 10)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function sha256(data) {
  return crypto.createHash('sha256').update(data).digest('hex');
}

function extKind(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.docx') return 'docx';
  if (ext === '.pptx') return 'pptx';
  if (ext === '.xlsx') return 'xlsx';
  throw new OfficePackageError('DOCX / PPTX / XLSX のみ対応しています。');
}

function replaceNthTextNode(xml, tagName, ordinal, value) {
  const pattern = new RegExp(`(<${tagName}\\b[^>]*>)([\\s\\S]*?)(<\\/${tagName}>)`, 'g');
  let current = -1;
  let found = false;
  const replaced = xml.replace(pattern, (all, open, old, close) => {
    current += 1;
    if (current !== ordinal) return all;
    found = true;
    const preserve = /^\s|\s$/.test(value) && !/xml:space=/.test(open)
      ? open.replace(/>$/, ' xml:space="preserve">')
      : open;
    return `${preserve}${escapeXml(value)}${close}`;
  });
  if (!found) throw new OfficePackageError('編集対象が見つかりません。ファイルが変更された可能性があります。');
  return replaced;
}

function collectTextNodes(xml, tagName, part, prefix) {
  const nodes = [];
  const pattern = new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)<\\/${tagName}>`, 'g');
  let match;
  let ordinal = 0;
  while ((match = pattern.exec(xml)) && nodes.length < MAX_TEXT_NODES) {
    const text = decodeXml(match[1].replace(/<[^>]+>/g, ''));
    nodes.push({ id: `${prefix}:${ordinal}`, part, ordinal, text });
    ordinal += 1;
  }
  return nodes;
}

/**
 * Distribute `text` across `lengths` runs, proportional to each run's original
 * length, so that formatting on every run is preserved. The final run absorbs
 * any rounding remainder. Empty runs stay empty.
 */
function splitProportional(text, lengths) {
  if (lengths.length === 0) return [];
  const total = lengths.reduce((a, b) => a + b, 0);
  if (total === 0) return lengths.map(() => '');
  const n = text.length;
  let allocated = 0;
  return lengths.map((len, i) => {
    if (i === lengths.length - 1) return text.slice(allocated);
    const share = Math.max(0, Math.round((n * len) / total));
    const slice = text.slice(allocated, allocated + share);
    allocated += slice.length;
    return slice;
  });
}

/**
 * Group `w:t` runs into their enclosing `w:p` paragraphs. Each paragraph keeps
 * the ids and original lengths of the runs it spans so an edit can be re-fanned
 * back out without disturbing `w:rPr` (run formatting).
 */
function collectParagraphs(xml, part, prefix) {
  const paragraphs = [];
  const paraPattern = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  let paraMatch;
  let ordinal = 0;
  while ((paraMatch = paraPattern.exec(xml)) && paragraphs.length < MAX_TEXT_NODES) {
    const inner = paraMatch[1];
    const runs = collectTextNodes(inner, 'w:t', part, `${prefix}:p${ordinal}`);
    if (runs.length === 0) {
      paragraphs.push({ id: `${prefix}:p${ordinal}`, part, ordinal, text: '', tIds: [], tLengths: [] });
      ordinal += 1;
      continue;
    }
    paragraphs.push({
      id: `${prefix}:p${ordinal}`,
      part,
      ordinal,
      text: runs.map((r) => r.text).join(''),
      tIds: runs.map((r) => r.id),
      tLengths: runs.map((r) => r.text.length),
    });
    ordinal += 1;
  }
  return paragraphs;
}

/**
 * Replace the text inside every `w:t` of a single paragraph block with the
 * corresponding entry of `runValues`, preserving each run's open tag (and thus
 * its `w:rPr` formatting). Used to apply a paragraph-level edit that was
 * re-fanned across runs by `splitProportional`.
 */
function replaceTextNodesInBlock(block, runValues) {
  let i = -1;
  return block.replace(/(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g, (all, open, old, close) => {
    i += 1;
    if (i >= runValues.length) return all;
    const value = runValues[i];
    const preserve = (/^\s|\s$/.test(value) && !/xml:space=/.test(open))
      ? open.replace(/>$/, ' xml:space="preserve">')
      : open;
    return `${preserve}${escapeXml(value)}${close}`;
  });
}

function resolveTarget(basePart, target) {
  if (target.startsWith('/')) return target.slice(1);
  return path.posix.normalize(path.posix.join(path.posix.dirname(basePart), target));
}

function attributes(fragment) {
  const result = {};
  for (const match of fragment.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*("([^"]*)"|'([^']*)')/g)) {
    result[match[1]] = decodeXml(match[3] ?? match[4] ?? '');
  }
  return result;
}

function relationshipMap(xml, basePart) {
  const result = new Map();
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const attrs = attributes(match[1]);
    if (attrs.Id && attrs.Target && attrs.TargetMode !== 'External') {
      result.set(attrs.Id, resolveTarget(basePart, attrs.Target));
    }
  }
  return result;
}

class OfficePackage {
  constructor(filePath, originalBytes, zip, kind) {
    this.filePath = filePath;
    this.originalBytes = Buffer.from(originalBytes);
    this.zip = zip;
    this.kind = kind;
    this.edits = new Map();
    this.partCache = new Map();
    this.entryHashes = new Map();
    this.canEdit = true;
    this.editBlockReason = '';
  }

  static async open(filePath, bytes) {
    const kind = extKind(filePath);
    let zip;
    try {
      zip = await JSZip.loadAsync(bytes, { checkCRC32: true, createFolders: false });
    } catch (error) {
      throw new OfficePackageError(`OfficeファイルのZIP整合性検査に失敗しました: ${error.message}`);
    }
    const pkg = new OfficePackage(filePath, bytes, zip, kind);
    await pkg.audit();
    return pkg;
  }

  async audit() {
    const names = Object.keys(this.zip.files);
    if (names.length > MAX_ENTRIES) throw new OfficePackageError('ZIP内のエントリ数が安全上限を超えています。');
    let total = 0;
    for (const name of names) {
      if (name.includes('..') || name.startsWith('/') || /^[A-Za-z]:/.test(name)) {
        throw new OfficePackageError(`危険なZIPパスを検出しました: ${name}`);
      }
      const entry = this.zip.file(name);
      if (!entry) continue;
      const data = await entry.async('nodebuffer');
      total += data.length;
      if (total > MAX_UNCOMPRESSED_BYTES) throw new OfficePackageError('展開後サイズが安全上限を超えています。');
      this.entryHashes.set(name, sha256(data));
    }
    if (!this.zip.file('[Content_Types].xml')) throw new OfficePackageError('[Content_Types].xml がありません。');
    const required = { docx: 'word/document.xml', pptx: 'ppt/presentation.xml', xlsx: 'xl/workbook.xml' }[this.kind];
    if (!this.zip.file(required)) throw new OfficePackageError(`必須パート ${required} がありません。`);

    const signed = names.some((name) => name.toLowerCase().startsWith('_xmlsignatures/'));
    const encrypted = names.some((name) => /encryptedpackage|encryptioninfo/i.test(name));
    if (signed || encrypted) {
      this.canEdit = false;
      this.editBlockReason = signed
        ? 'デジタル署名付きファイルは、署名を壊すため編集できません。閲覧のみ可能です。'
        : '暗号化されたOfficeファイルは編集できません。';
    }
  }

  async readText(part) {
    if (this.partCache.has(part)) return this.partCache.get(part);
    const entry = this.zip.file(part);
    if (!entry) throw new OfficePackageError(`パートが見つかりません: ${part}`);
    const value = await entry.async('string');
    this.partCache.set(part, value);
    return value;
  }

  stageEdit(edit) {
    if (!this.canEdit) throw new OfficePackageError(this.editBlockReason);
    if (!edit || typeof edit.id !== 'string' || typeof edit.value !== 'string') {
      throw new OfficePackageError('編集データが不正です。');
    }
    if (edit.value.length > 100000) throw new OfficePackageError('1項目の編集上限を超えています。');
    this.edits.set(edit.id, edit.value);
  }

  discardEdits() {
    this.edits.clear();
  }

  unstageEdit(id) {
    this.edits.delete(id);
  }

  hasEdits() {
    return this.edits.size > 0;
  }

  async createViewModel() {
    const base = {
      kind: this.kind,
      filename: path.basename(this.filePath),
      canEdit: this.canEdit,
      editBlockReason: this.editBlockReason,
      editCount: this.edits.size,
      originalSha256: sha256(this.originalBytes),
    };
    if (this.kind === 'docx') return { ...base, ...(await this.docxModel()) };
    if (this.kind === 'pptx') return { ...base, ...(await this.pptxModel()) };
    return { ...base, ...(await this.xlsxModel()) };
  }

  async docxModel() {
    const part = 'word/document.xml';
    const xml = await this.readText(part);
    const nodes = collectTextNodes(xml, 'w:t', part, 'docx');
    const paragraphs = collectParagraphs(xml, part, 'docx');
    return {
      items: nodes.map((node) => ({ ...node, value: this.edits.get(node.id) ?? node.text })),
      paragraphs: paragraphs.map((p) => ({
        ...p,
        value: this.edits.get(p.id) ?? p.text,
        // Per-run edited values, re-fanned from the paragraph edit when present.
        runValues: this.edits.has(p.id)
          ? splitProportional(this.edits.get(p.id), p.tLengths)
          : p.tIds.map((id) => this.edits.get(id) ?? (nodes.find((n) => n.id === id)?.text ?? '')),
      })),
      summary: `${nodes.length} 個のテキスト断片 / ${paragraphs.length} 段落`,
      fidelityPreview: true,
    };
  }

  async pptxModel() {
    const presentationPart = 'ppt/presentation.xml';
    const presentation = await this.readText(presentationPart);
    const rels = await this.readText('ppt/_rels/presentation.xml.rels');
    const relMap = relationshipMap(rels, presentationPart);
    const slideParts = [];
    for (const m of presentation.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*\/?\s*>/g)) {
      const target = relMap.get(m[1]);
      if (target && this.zip.file(target)) slideParts.push(target);
    }
    const slides = [];
    for (let i = 0; i < slideParts.length; i += 1) {
      const part = slideParts[i];
      const xml = await this.readText(part);
      const nodes = collectTextNodes(xml, 'a:t', part, `pptx:${i}`);
      slides.push({
        number: i + 1,
        part,
        items: nodes.map((node) => ({ ...node, value: this.edits.get(node.id) ?? node.text })),
        imageCount: (xml.match(/<a:blip\b/g) || []).length,
      });
    }
    return { slides, summary: `${slides.length} スライド` };
  }

  async xlsxModel() {
    const workbookPart = 'xl/workbook.xml';
    const workbook = await this.readText(workbookPart);
    const rels = await this.readText('xl/_rels/workbook.xml.rels');
    const relMap = relationshipMap(rels, workbookPart);
    let shared = [];
    if (this.zip.file('xl/sharedStrings.xml')) {
      const sharedXml = await this.readText('xl/sharedStrings.xml');
      shared = Array.from(sharedXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g), (m) =>
        decodeXml(Array.from(m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g), (t) => t[1]).join('')),
      );
    }
    const sheets = [];
    let sheetIndex = 0;
    for (const m of workbook.matchAll(/<sheet\b([^>]*)\/?\s*>/g)) {
      const sheetAttrs = attributes(m[1]);
      const part = relMap.get(sheetAttrs['r:id']);
      if (!part || !this.zip.file(part)) continue;
      const xml = await this.readText(part);
      const cells = [];
      let cellOrdinal = 0;
      for (const cellMatch of xml.matchAll(/<c\b([^>]*)\br="([A-Z]+[0-9]+)"([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attrs = `${cellMatch[1]} ${cellMatch[3]}`;
        const body = cellMatch[4];
        const type = /\bt="([^"]+)"/.exec(attrs)?.[1] || 'n';
        const formula = /<f\b/.test(body);
        const formulaText = (/<f\b[^>]*>([\s\S]*?)<\/f>/.exec(body)?.[1] ?? '').trim();
        const raw = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? '';
        const inline = Array.from(body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g), (x) => x[1]).join('');
        let value = type === 's' ? (shared[Number(raw)] ?? '') : type === 'inlineStr' ? decodeXml(inline) : decodeXml(raw);
        const id = `xlsx:${sheetIndex}:${cellOrdinal}`;
        value = this.edits.get(id) ?? value;
        cells.push({ id, part, ordinal: cellOrdinal, ref: cellMatch[2], type, formula, formulaText, value });
        cellOrdinal += 1;
        if (cells.length >= 20000) break;
      }
      sheets.push({ name: sheetAttrs.name || `Sheet ${sheetIndex + 1}`, part, cells });
      sheetIndex += 1;
    }
    return { sheets, summary: `${sheets.length} シート` };
  }

  async buildChangedParts() {
    const changed = new Map();
    const model = await this.createViewModel();
    if (this.kind === 'docx') {
      let xml = await this.readText('word/document.xml');
      // Paragraph-level edits are re-fanned across their runs, preserving each
      // run's formatting (w:rPr). Apply these before individual w:t edits so the
      // two edit granularities can coexist.
      for (const para of model.paragraphs) {
        if (!this.edits.has(para.id)) continue;
        let pIdx = -1;
        xml = xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (block) => {
          pIdx += 1;
          if (pIdx !== para.ordinal) return block;
          return replaceTextNodesInBlock(block, para.runValues);
        });
      }
      for (const item of model.items) {
        if (!this.edits.has(item.id)) continue;
        xml = replaceNthTextNode(xml, 'w:t', item.ordinal, this.edits.get(item.id));
      }
      changed.set('word/document.xml', xml);
    } else if (this.kind === 'pptx') {
      for (const slide of model.slides) {
        let xml = await this.readText(slide.part);
        let touched = false;
        for (const item of slide.items) {
          if (!this.edits.has(item.id)) continue;
          xml = replaceNthTextNode(xml, 'a:t', item.ordinal, this.edits.get(item.id));
          touched = true;
        }
        if (touched) changed.set(slide.part, xml);
      }
    } else {
      for (const sheet of model.sheets) {
        let xml = await this.readText(sheet.part);
        let current = -1;
        let touched = false;
        xml = xml.replace(/<c\b([^>]*)\br="([A-Z]+[0-9]+)"([^>]*)>([\s\S]*?)<\/c>/g, (all, left, ref, right, body) => {
          current += 1;
          const id = sheet.cells.find((c) => c.ordinal === current)?.id;
          if (!id || !this.edits.has(id)) return all;
          const cell = sheet.cells.find((c) => c.id === id);
          if (cell.formula) throw new OfficePackageError(`数式セル ${ref} は編集できません。`);
          touched = true;
          const value = this.edits.get(id);
          const attrs = `${left} r="${ref}"${right}`
            .replace(/\s+t="[^"]*"/g, '')
            .trim();
          if (cell.type === 'n' && /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(value.trim())) {
            return `<c ${attrs}><v>${escapeXml(value.trim())}</v></c>`;
          }
          return `<c ${attrs} t="inlineStr"><is><t${/^\s|\s$/.test(value) ? ' xml:space="preserve"' : ''}>${escapeXml(value)}</t></is></c>`;
        });
        if (touched) changed.set(sheet.part, xml);
      }
    }
    return changed;
  }

  async exportValidatedCopy() {
    if (!this.hasEdits()) throw new OfficePackageError('変更がありません。');
    const changed = await this.buildChangedParts();
    if (changed.size === 0) throw new OfficePackageError('有効な変更がありません。');

    const outputZip = new JSZip();
    for (const name of Object.keys(this.zip.files)) {
      const entry = this.zip.files[name];
      if (entry.dir) {
        outputZip.folder(name);
        continue;
      }
      const data = changed.has(name) ? changed.get(name) : await entry.async('nodebuffer');
      outputZip.file(name, data, { date: entry.date, createFolders: false });
    }
    const bytes = await outputZip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      platform: 'DOS',
    });
    const reopened = await OfficePackage.open(this.filePath, bytes);

    const untouchedMismatches = [];
    for (const [name, before] of this.entryHashes) {
      if (changed.has(name)) continue;
      if (reopened.entryHashes.get(name) !== before) untouchedMismatches.push(name);
    }
    if (untouchedMismatches.length) {
      throw new OfficePackageError(`未編集パートが変化したため保存を中止しました: ${untouchedMismatches.join(', ')}`);
    }

    await reopened.createViewModel();
    return {
      bytes: Uint8Array.from(bytes),
      report: {
        changedParts: Array.from(changed.keys()),
        untouchedPartsVerified: this.entryHashes.size - changed.size,
        outputSha256: sha256(bytes),
        zipCrcVerified: true,
        originalUntouched: true,
      },
    };
  }
}

module.exports = {
  OfficePackage,
  OfficePackageError,
  MAX_FILE_BYTES,
  escapeXml,
  decodeXml,
  splitProportional,
  collectParagraphs,
};
