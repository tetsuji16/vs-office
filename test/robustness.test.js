'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { OfficePackage, escapeXml, decodeXml } = require('../src/office-package');

async function makeZip(parts) {
  const zip = new JSZip();
  Object.entries(parts).forEach(([name, value]) => zip.file(name, value));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

test('escapeXml round-trips through decodeXml for every XML special', () => {
  const payload = '& < > " \' <tag> "quoted" </tag> &amp;';
  const escaped = escapeXml(payload);
  // The five XML predefined entities must all be present.
  for (const ent of ['&amp;', '&lt;', '&gt;', '&quot;', '&apos;']) {
    assert.ok(escaped.includes(ent), `expected ${ent} in escaped output`);
  }
  assert.equal(decodeXml(escaped), payload);
});

test('DOCX edit escapes XML special characters and the result re-parses', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  const dangerous = '<script> & "quotes" \'apos\' </script>';
  pkg.stageEdit({ id: model.items[0].id, value: dangerous });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  // Raw angle brackets must NOT survive unescaped.
  assert.ok(!doc.includes('<script>'), 'unescaped tag leaked into package');
  assert.ok(doc.includes('&lt;script&gt;'), 'expected escaped script tag');
  // And the model decodes it back to the original dangerous string.
  const reModel = await OfficePackage.open('sample.docx', result.bytes);
  const reItems = (await reModel.createViewModel()).items;
  assert.equal(reItems[0].value, dangerous);
});

test('stageEdit rejects non-string id and non-string value', () => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('word/document.xml', '<w:document xmlns:w="w"><w:body/></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' }).then(async (bytes) => {
    const pkg = await OfficePackage.open('x.docx', bytes);
    assert.throws(() => pkg.stageEdit({ id: 5, value: 'v' }), /編集データが不正/);
    assert.throws(() => pkg.stageEdit({ id: 'd', value: 9 }), /編集データが不正/);
    assert.throws(() => pkg.stageEdit({ id: 'd', value: null }), /編集データが不正/);
    assert.throws(() => pkg.stageEdit({}), /編集データが不正/);
  });
});

test('stageEdit enforces the per-edit size ceiling', () => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('word/document.xml', '<w:document xmlns:w="w"><w:body/></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' }).then(async (bytes) => {
    const pkg = await OfficePackage.open('x.docx', bytes);
    const huge = 'x'.repeat(100001);
    assert.throws(() => pkg.stageEdit({ id: 'docx:0', value: huge }), /編集上限/);
    // Exactly at the limit is allowed.
    const atLimit = 'x'.repeat(100000);
    pkg.stageEdit({ id: 'docx:0', value: atLimit });
    assert.equal(pkg.edits.get('docx:0'), atLimit);
  });
});

test('staging an edit for a non-existent node is ignored without corrupting the package', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>One</w:t></w:r></w:p></w:body></w:document>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  // Stage an edit whose id claims ordinal 99, which does not exist.
  pkg.stageEdit({ id: 'docx:99', value: 'ghost' });
  // The unknown id is skipped during build; the real node stays untouched and
  // export still succeeds (defence-in-depth: a bad id never corrupts output).
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  assert.ok(doc.includes('>One<'), 'the real text node was altered by a phantom edit');
});

test('empty-string edit deletes text while preserving the run element', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t xml:space="preserve">keep me</w:t></w:r></w:p></w:body></w:document>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: model.items[0].id, value: '' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  // The run survives; its text is now empty.
  assert.ok(/<w:r>/.test(doc), 'run element was lost');
  assert.ok(!doc.includes('keep me'), 'old text still present');
});

test('exportValidatedCopy refuses when nothing changed', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  await assert.rejects(() => pkg.exportValidatedCopy(), /変更がありません/);
});

test('conservative save leaves untouched parts byte-identical by SHA-256', async () => {
  const image = Buffer.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]);
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>',
    'word/media/image1.png': image,
    'word/styles.xml': '<w:styles xmlns:w="w"><w:style w:styleId="Normal"/></w:styles>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: model.items[0].id, value: 'Alpha' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  assert.deepEqual(await reopened.file('word/media/image1.png').async('nodebuffer'), image);
  // Four parts total; only word/document.xml changed, so three stay untouched.
  assert.equal(result.report.untouchedPartsVerified, 3);
  assert.equal(result.report.changedParts.length, 1);
});

test('XLSX numeric edit keeps a shared-string neighbour untouched', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': '<workbook xmlns:r="r"><sheets><sheet r:id="rId1" name="Data"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Target="worksheets/sheet1.xml" Id="rId1"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>SharedA</t></si><si><t>SharedB</t></si></sst>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData></worksheet>',
    'xl/styles.xml': '<styleSheet/>',
  });
  const pkg = await OfficePackage.open('sample.xlsx', original);
  const model = await pkg.createViewModel();
  // Edit the second shared-string cell only.
  pkg.stageEdit({ id: model.sheets[0].cells[1].id, value: 'LocalB' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const shared = await reopened.file('xl/sharedStrings.xml').async('string');
  // The untouched shared string must remain; only the edited cell becomes inlineStr.
  assert.ok(shared.includes('SharedA'), 'untouched shared string was dropped');
  const sheet = await reopened.file('xl/worksheets/sheet1.xml').async('string');
  assert.ok(/B1[^>]*t="inlineStr"/.test(sheet), 'edited cell not converted to inlineStr');
});

test('XLSX formula cells expose the formula text for read-only display', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': '<workbook xmlns:r="r"><sheets><sheet r:id="rId1" name="Data"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Target="worksheets/sheet1.xml" Id="rId1"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1"><f>SUM(B1:B3)</f><v>6</v></c></row></sheetData></worksheet>',
    'xl/styles.xml': '<styleSheet/>',
  });
  const pkg = await OfficePackage.open('sample.xlsx', original);
  const model = await pkg.createViewModel();
  const cell = model.sheets[0].cells[0];
  assert.equal(cell.formula, true);
  assert.equal(cell.formulaText, 'SUM(B1:B3)');
  assert.equal(cell.value, '6');
  // Formula cells must remain read-only (editing is refused by the save pipeline).
  pkg.stageEdit({ id: cell.id, value: '99' });
  await assert.rejects(() => pkg.exportValidatedCopy(), /数式セル/);
});

module.exports = { makeZip, contentTypes };
