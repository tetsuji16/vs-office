'use strict';

const JSZip = require('jszip');
const crypto = require('crypto');
const path = require('path');

const MAX_FILE_BYTES = 100 * 1024 * 1024;
const MAX_UNCOMPRESSED_BYTES = 500 * 1024 * 1024;
const MAX_ENTRY_BYTES = 200 * 1024 * 1024;
const MAX_ENTRIES = 10000;
const MAX_TEXT_NODES = 50000;

// Match both populated cells (`<c>...</c>`) and styled/blank self-closing cells
// (`<c .../>`). Treating only the paired form as a cell lets a regex that
// starts on a self-closing cell consume the body of the next populated cell,
// shifting every following reference/value pair in real-world workbooks.
const XLSX_CELL_PATTERN_SOURCE = '<((?:[A-Za-z_][\\w.-]*:)?c)\\b([^>]*?)(?:\\/\\s*>|>([\\s\\S]*?)<\\/\\1>)';

function xlsxCellPattern() {
  return new RegExp(XLSX_CELL_PATTERN_SOURCE, 'g');
}

class OfficePackageError extends Error {}

function escapeXml(value) {
  // XML 1.0 forbids control characters 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, and the
  // lone-surrogate range 0xD800-0xDFFF. A forbidden char in saved output would
  // make Word/Excel refuse to open the file, so strip it before escaping.
  const stripped = String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '') // XML 1.0 forbidden control chars
    .replace(/\p{Surrogate}/gu, ''); // lone surrogates are never valid XML
  return stripped
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Normalize a user-supplied colour to a 6-digit lowercase hex string, or null if invalid. */
function normalizeHex(value) {
  if (typeof value !== 'string') return null;
  const m = /^#?([0-9a-fA-F]{6})$/.exec(value.trim());
  return m ? m[1].toLowerCase() : null;
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
 * Replace the text inside every `a:t` of a single shape block with the
 * corresponding entry of `runValues`, preserving each run's open tag (and thus
 * its `a:rPr` formatting). Used to apply a shape-level text edit that was
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

/**
 * Extract every drawing shape (`p:sp`) from a slide, capturing its combined
 * text and the solid-fill / line colours so the UI can edit them directly
 * without discarding the rest of the OOXML shape markup.
 */
function collectShapes(xml, part, prefix) {
  const shapes = [];
  const spPattern = /<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g;
  let spMatch;
  let ordinal = 0;
  while ((spMatch = spPattern.exec(xml)) && shapes.length < MAX_TEXT_NODES) {
    const inner = spMatch[1];
    const spPr = /<p:spPr\b[^>]*>([\s\S]*?)<\/p:spPr>/.exec(inner)?.[1] || '';
    const lineXml = /<a:ln\b[^>]*>([\s\S]*?)<\/a:ln>/.exec(spPr)?.[1] || '';
    // Fill properties precede line/effect properties in the spPr schema. Stop
    // there so a nested line/effect colour is never reported as the shape fill.
    const fillBoundary = spPr.search(/<a:(?:ln|effectLst|effectDag|scene3d|sp3d|extLst)\b/);
    const shapeFillXml = fillBoundary < 0 ? spPr : spPr.slice(0, fillBoundary);
    const fill = /<a:solidFill>\s*<a:srgbClr val="([0-9a-fA-F]{6})"/.exec(shapeFillXml);
    const line = /<a:solidFill>\s*<a:srgbClr val="([0-9a-fA-F]{6})"/.exec(lineXml);
    const textNodes = collectTextNodes(inner, 'a:t', part, `${prefix}:${ordinal}`);
    const text = textNodes.map((n) => n.text).join('');
    shapes.push({
      id: `${prefix}:${ordinal}`,
      part,
      ordinal,
      type: 'shape',
      text,
      fillColor: (fill?.[1] || '').toLowerCase(),
      lineColor: (line?.[1] || '').toLowerCase(),
      tIds: textNodes.map((n) => n.id),
      tLengths: textNodes.map((n) => n.text.length),
    });
    ordinal += 1;
  }
  return shapes;
}

const FILL_CHOICE = /<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b[^>]*(?:\/>|>[\s\S]*?<\/a:\1>)/;

function replaceFillChoice(xml, replacement) {
  return FILL_CHOICE.test(xml) ? xml.replace(FILL_CHOICE, replacement) : `${xml}${replacement}`;
}

function updateShapeProperties(block, updater) {
  if (/<p:spPr\b[^>]*\/>/.test(block)) {
    return block.replace(/<p:spPr\b([^>]*)\/>/, (_, attrs) => `<p:spPr${attrs}>${updater('')}</p:spPr>`);
  }
  if (/<p:spPr\b[^>]*>[\s\S]*?<\/p:spPr>/.test(block)) {
    return block.replace(/(<p:spPr\b[^>]*>)([\s\S]*?)(<\/p:spPr>)/, (_, open, inner, close) =>
      `${open}${updater(inner)}${close}`);
  }
  // A valid p:sp normally has spPr, but keep the edit well-formed for minimal
  // or producer-specific input that omitted it.
  return block.replace(/(<p:sp\b[^>]*>)/, `$1<p:spPr>${updater('')}</p:spPr>`);
}

/** Replace the text of the n-th `p:sp` in `xml` with `runValues` re-fanned across its `a:t` runs. */
function replaceShapeText(xml, ordinal, runValues) {
  let idx = -1;
  return xml.replace(/<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g, (block) => {
    idx += 1;
    if (idx !== ordinal) return block;
    let j = -1;
    return block.replace(/(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/g, (all, open, old, close) => {
      j += 1;
      if (j >= runValues.length) return all;
      const value = runValues[j];
      const preserve = (/^\s|\s$/.test(value) && !/xml:space=/.test(open))
        ? open.replace(/>$/, ' xml:space="preserve">')
        : open;
      return `${preserve}${escapeXml(value)}${close}`;
    });
  });
}

/** Set/replace the solid-fill colour of the n-th `p:sp` in `xml`. */
function setShapeSolidFill(xml, ordinal, color) {
  let idx = -1;
  return xml.replace(/<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g, (block) => {
    idx += 1;
    if (idx !== ordinal) return block;
    const replacement = `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
    return updateShapeProperties(block, (inner) => {
      const boundary = inner.search(/<a:(?:ln|effectLst|effectDag|scene3d|sp3d|extLst)\b/);
      const before = boundary < 0 ? inner : inner.slice(0, boundary);
      const after = boundary < 0 ? '' : inner.slice(boundary);
      return `${replaceFillChoice(before, replacement)}${after}`;
    });
  });
}

/** Set/replace the line colour of the n-th `p:sp` in `xml`. */
function setShapeLineColor(xml, ordinal, color) {
  let idx = -1;
  return xml.replace(/<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g, (block) => {
    idx += 1;
    if (idx !== ordinal) return block;
    const fill = `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
    return updateShapeProperties(block, (inner) => {
      if (/<a:ln\b[^>]*\/>/.test(inner)) {
        return inner.replace(/<a:ln\b([^>]*)\/>/, (_, attrs) => `<a:ln${attrs}>${fill}</a:ln>`);
      }
      if (/<a:ln\b[^>]*>[\s\S]*?<\/a:ln>/.test(inner)) {
        return inner.replace(/(<a:ln\b[^>]*>)([\s\S]*?)(<\/a:ln>)/, (_, open, body, close) =>
          `${open}${FILL_CHOICE.test(body) ? body.replace(FILL_CHOICE, fill) : `${fill}${body}`}${close}`);
      }
      const boundary = inner.search(/<a:(?:effectLst|effectDag|scene3d|sp3d|extLst)\b/);
      const line = `<a:ln>${fill}</a:ln>`;
      return boundary < 0 ? `${inner}${line}` : `${inner.slice(0, boundary)}${line}${inner.slice(boundary)}`;
    });
  });
}

/** Remove the n-th `p:sp` from `xml`. */
function removeNthShape(xml, ordinal) {
  let idx = -1;
  return xml.replace(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g, (block) => {
    idx += 1;
    return idx === ordinal ? '' : block;
  });
}

function resolveTarget(basePart, target) {
  // OPC permits package-absolute part names such as "/xl/worksheets/sheet1.xml".
  // Accept those by resolving from the package root, while continuing to reject
  // traversal (including percent-encoded variants) and Windows-style paths.
  let decoded;
  try { decoded = decodeURIComponent(target); } catch { return null; }
  if (decoded.includes('..') || decoded.includes('\\') || decoded.includes('\0')) return null;
  const candidate = decoded.startsWith('/')
    ? decoded.replace(/^\/+/, '')
    : path.posix.join(path.posix.dirname(basePart), decoded);
  const normalized = path.posix.normalize(candidate);
  if (!normalized || normalized === '..' || normalized.startsWith('../') || normalized.startsWith('/')) return null;
  return normalized;
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
    const resolved = attrs.Target ? resolveTarget(basePart, attrs.Target) : null;
    if (attrs.Id && resolved && attrs.TargetMode !== 'External') {
      result.set(attrs.Id, resolved);
    }
  }
  return result;
}

function appendSharedStringEntries(xml, values) {
  if (!values.length) return xml;
  const rootMatch = /<((?:[A-Za-z_][\w.-]*:)?sst)\b([^>]*)>/.exec(xml);
  if (!rootMatch) throw new OfficePackageError('共有文字列テーブルの形式が不正です。');

  const rootTag = rootMatch[1];
  const prefix = rootTag.includes(':') ? `${rootTag.split(':')[0]}:` : '';
  const selfClosing = /\/\s*>$/.test(rootMatch[0]);
  let rootAttrs = rootMatch[2].replace(/\/\s*$/, '');
  const existingEntries = Array.from(xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?si\b/g)).length;
  const uniqueCount = existingEntries + values.length;
  if (/\buniqueCount\s*=/.test(rootAttrs)) {
    rootAttrs = rootAttrs.replace(/(\buniqueCount\s*=\s*)(["'])\d+\2/, `$1"${uniqueCount}"`);
  } else {
    rootAttrs += ` uniqueCount="${uniqueCount}"`;
  }
  const openTag = `<${rootTag}${rootAttrs}>`;
  const entries = values.map((value) =>
    `<${prefix}si><${prefix}t${/^\s|\s$/.test(value) ? ' xml:space="preserve"' : ''}>${escapeXml(value)}</${prefix}t></${prefix}si>`,
  ).join('');

  if (selfClosing) {
    return xml.replace(rootMatch[0], `${openTag}${entries}</${rootTag}>`);
  }
  let updated = xml.replace(rootMatch[0], openTag);
  const closeTag = `</${rootTag}>`;
  const closeIndex = updated.lastIndexOf(closeTag);
  if (closeIndex < 0) throw new OfficePackageError('共有文字列テーブルの終了タグがありません。');
  // extLst, when present, must remain the final child of the SST root.
  const extMatch = /<(?:[A-Za-z_][\w.-]*:)?extLst\b/.exec(updated.slice(rootMatch.index, closeIndex));
  const insertIndex = extMatch ? rootMatch.index + extMatch.index : closeIndex;
  updated = `${updated.slice(0, insertIndex)}${entries}${updated.slice(insertIndex)}`;
  return updated;
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
      if (data.length > MAX_ENTRY_BYTES) {
        throw new OfficePackageError(`個別パーツの展開サイズが安全上限（${MAX_ENTRY_BYTES} バイト）を超えています。`);
      }
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
    // Reject ids that could never match a real node id. Real ids are
    // `<kind>:<ordinal>`, `<kind>:p<ordinal>` (paragraphs), `<kind>:<sheet>:<cell>`
    // (xlsx cells), or `<kind>:<slide>:<shape>:(text|fill|line|delete)` (pptx shapes),
    // e.g. "docx:0", "pptx:2:p1", "xlsx:0:3", "pptx:0:1:fill". This blocks crafted
    // ids with path separators or other junk from webview input.
    if (!/^[a-z]+:p?\d+(:\d+)*(:(text|fill|line|delete))?$/.test(edit.id)) {
      throw new OfficePackageError('編集データが不正です。');
    }
    // Colour edits must be a valid 6-digit hex (with or without leading '#').
    if (/(:fill|:line)$/.test(edit.id) && !/^#?[0-9a-fA-F]{6}$/.test(edit.value.trim())) {
      throw new OfficePackageError('色は #RRGGBB 形式で指定してください。');
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
      const shapes = collectShapes(xml, part, `pptx:${i}`).map((shape) => ({
        ...shape,
        value: this.edits.get(`${shape.id}:text`) ?? shape.text,
        fillValue: this.edits.get(`${shape.id}:fill`) ?? (shape.fillColor ? `#${shape.fillColor}` : ''),
        lineValue: this.edits.get(`${shape.id}:line`) ?? (shape.lineColor ? `#${shape.lineColor}` : ''),
        deleted: this.edits.has(`${shape.id}:delete`),
      }));
      slides.push({
        number: i + 1,
        part,
        items: nodes.map((node) => ({ ...node, value: this.edits.get(node.id) ?? node.text })),
        shapes,
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
      shared = Array.from(sharedXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?si\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?si>/g), (m) =>
        decodeXml(Array.from(m[1].matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g), (t) => t[1]).join('')),
      );
    }
    const sheets = [];
    let sheetIndex = 0;
    for (const m of workbook.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b([^>]*)\/?\s*>/g)) {
      const sheetAttrs = attributes(m[1]);
      const part = relMap.get(sheetAttrs['r:id']);
      if (!part || !this.zip.file(part)) continue;
      const xml = await this.readText(part);
      const cells = [];
      let cellOrdinal = 0;
      for (const cellMatch of xml.matchAll(xlsxCellPattern())) {
        const attrsText = cellMatch[2];
        const cellAttrs = attributes(attrsText);
        const ref = cellAttrs.r;
        if (!ref || !/^[A-Za-z]+[0-9]+$/.test(ref)) continue;
        const body = cellMatch[3] || '';
        const type = cellAttrs.t || 'n';
        const formula = /<(?:[A-Za-z_][\w.-]*:)?f\b/.test(body);
        const formulaText = decodeXml((/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?f>/.exec(body)?.[1] ?? '').trim());
        const raw = /<(?:[A-Za-z_][\w.-]*:)?v\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?v>/.exec(body)?.[1] ?? '';
        const inline = Array.from(body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g), (x) => x[1]).join('');
        const original = type === 's' ? (shared[Number(raw)] ?? '') : type === 'inlineStr' ? decodeXml(inline) : decodeXml(raw);
        // `text` is the original, un-edited cell content (used by the UI to
        // detect edits); `value` is the editable copy that may be overridden.
        let value = original;
        if (formula && formulaText) value = `=${formulaText}`;
        const id = `xlsx:${sheetIndex}:${cellOrdinal}`;
        value = this.edits.get(id) ?? value;
        cells.push({ id, part, ordinal: cellOrdinal, ref: ref.toUpperCase(), type, formula, formulaText, value, text: original });
        cellOrdinal += 1;
        if (cells.length >= 20000) break;
      }
      sheets.push({ name: sheetAttrs.name || `Sheet ${sheetIndex + 1}`, part, cells });
      sheetIndex += 1;
    }
    return { sheets, sharedStrings: shared, summary: `${sheets.length} シート` };
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
        // Text-node ordinals also refer to the original XML. Apply legacy
        // per-node edits before any shape can be deleted and shift them.
        for (const item of slide.items) {
          if (!this.edits.has(item.id)) continue;
          xml = replaceNthTextNode(xml, 'a:t', item.ordinal, this.edits.get(item.id));
          touched = true;
        }
        // Shape ordinals are calculated from the original XML.  Deleting a
        // lower-numbered shape shifts every following ordinal, so always
        // apply shape operations from the end of the slide towards the start.
        // This also makes a mixed delete/edit batch address the intended
        // original shapes rather than whichever shape happened to move into
        // an earlier slot.
        for (const shape of [...(slide.shapes || [])].sort((a, b) => b.ordinal - a.ordinal)) {
          const baseId = shape.id;
          if (this.edits.has(`${baseId}:delete`)) {
            xml = removeNthShape(xml, shape.ordinal);
            touched = true;
            continue;
          }
          if (this.edits.has(`${baseId}:text`)) {
            const value = this.edits.get(`${baseId}:text`);
            xml = replaceShapeText(xml, shape.ordinal, splitProportional(value, shape.tLengths));
            touched = true;
          }
          if (this.edits.has(`${baseId}:fill`)) {
            const color = normalizeHex(this.edits.get(`${baseId}:fill`));
            if (color) { xml = setShapeSolidFill(xml, shape.ordinal, color); touched = true; }
          }
          if (this.edits.has(`${baseId}:line`)) {
            const color = normalizeHex(this.edits.get(`${baseId}:line`));
            if (color) { xml = setShapeLineColor(xml, shape.ordinal, color); touched = true; }
          }
        }
        if (touched) changed.set(slide.part, xml);
      }
    } else {
      const sharedStrings = model.sharedStrings ? model.sharedStrings.slice() : [];
      let sharedDirty = false;
      for (const sheet of model.sheets) {
        let xml = await this.readText(sheet.part);
        let current = -1;
        let touched = false;
        xml = xml.replace(xlsxCellPattern(), (all, cellTag, attrsText, body = '') => {
          const ref = attributes(attrsText).r;
          if (!ref || !/^[A-Za-z]+[0-9]+$/.test(ref)) return all;
          current += 1;
          const id = sheet.cells.find((c) => c.ordinal === current)?.id;
          if (!id || !this.edits.has(id)) return all;
          const cell = sheet.cells.find((c) => c.id === id);
          touched = true;
          const value = this.edits.get(id);
          const childPrefix = cellTag.includes(':') ? `${cellTag.split(':')[0]}:` : '';
          const attrs = attrsText
            .replace(/\s+t="[^"]*"/g, '')
            .replace(/\s+t='[^']*'/g, '')
            .trim();
          if (/^=/.test(value.trim())) {
            // User supplied a formula: keep the cell as a formula cell, replace
            // the <f> content, and drop the cached <v> so the app recalculates.
            const fText = escapeXml(value.trim().slice(1));
            let inner;
            if (/<(?:[A-Za-z_][\w.-]*:)?f\b/.test(body)) {
              inner = body
                .replace(/<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f>/, `<${childPrefix}f>${fText}</${childPrefix}f>`)
                .replace(/<(?:[A-Za-z_][\w.-]*:)?v\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?v>/g, '');
            } else {
              inner = `<${childPrefix}f>${fText}</${childPrefix}f>`;
            }
            // Formula result types are meaningful in OOXML. Preserve an
            // existing non-numeric result type (for example `str` or `b`),
            // but leave numeric/new formulas untyped so recalculation yields
            // numeric values instead of forcing them into strings.
            const resultType = cell.formula && ['b', 'd', 'e', 'str'].includes(cell.type)
              ? ` t="${cell.type}"`
              : '';
            return `<${cellTag} ${attrs}${resultType}>${inner}</${cellTag}>`;
          }
          if (cell.type === 'n' && /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(value.trim())) {
            return `<${cellTag} ${attrs}><${childPrefix}v>${escapeXml(value.trim())}</${childPrefix}v></${cellTag}>`;
          }
          if (cell.type === 's') {
            // Preserve the shared-string structure: append (or reuse) a shared
            // string entry and keep the cell as t="s" rather than inlining it.
            const idx = sharedStrings.indexOf(value);
            const useIdx = idx >= 0 ? idx : (sharedStrings.push(value), sharedStrings.length - 1);
            sharedDirty = true;
            return `<${cellTag} ${attrs} t="s"><${childPrefix}v>${useIdx}</${childPrefix}v></${cellTag}>`;
          }
          return `<${cellTag} ${attrs} t="inlineStr"><${childPrefix}is><${childPrefix}t${/^\s|\s$/.test(value) ? ' xml:space="preserve"' : ''}>${escapeXml(value)}</${childPrefix}t></${childPrefix}is></${cellTag}>`;
        });
        if (touched) changed.set(sheet.part, xml);
      }
      if (sharedDirty && this.zip.file('xl/sharedStrings.xml')) {
        const originalSharedXml = await this.readText('xl/sharedStrings.xml');
        const appended = sharedStrings.slice(model.sharedStrings.length);
        if (appended.length) {
          changed.set('xl/sharedStrings.xml', appendSharedStringEntries(originalSharedXml, appended));
        }
      }
    }
    return changed;
  }

  async exportValidatedCopy() {
    if (!this.hasEdits()) throw new OfficePackageError('変更がありません。');
    const model = await this.createViewModel();
    // Every staged edit must reference a real node/paragraph id. A crafted or
    // stale id (e.g. from a webview that drifted out of sync) must never be
    // silently ignored as a "successful no-op" save.
    const validIds = new Set();
    if (model.items) for (const item of model.items) validIds.add(item.id);
    if (model.paragraphs) for (const p of model.paragraphs) validIds.add(p.id);
    if (model.slides) {
      for (const s of model.slides) {
        for (const item of s.items) validIds.add(item.id);
        for (const shape of (s.shapes || [])) {
          validIds.add(`${shape.id}:text`);
          validIds.add(`${shape.id}:fill`);
          validIds.add(`${shape.id}:line`);
          validIds.add(`${shape.id}:delete`);
        }
      }
    }
    if (model.sheets) for (const s of model.sheets) for (const c of s.cells) validIds.add(c.id);
    for (const id of this.edits.keys()) {
      if (!validIds.has(id)) {
        throw new OfficePackageError('編集対象が見つかりません。ファイルが変更された可能性があります。');
      }
    }
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
