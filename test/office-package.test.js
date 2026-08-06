'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { OfficePackage } = require('../src/office-package');

async function makeZip(parts) {
  const zip = new JSZip();
  Object.entries(parts).forEach(([name, value]) => zip.file(name, value));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

test('DOCX edits one text node and preserves every untouched entry', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p></w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"><w:style w:styleId="Normal"/></w:styles>',
    'word/media/image1.png': Buffer.from([1, 2, 3, 4]),
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.items.length, 2);
  pkg.stageEdit({ id: model.items[0].id, value: 'こんにちは' });
  const result = await pkg.exportValidatedCopy();
  assert.deepEqual(result.report.changedParts, ['word/document.xml']);
  assert.equal(result.report.untouchedPartsVerified, 3);
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  assert.match(await reopened.file('word/document.xml').async('string'), /<w:t>こんにちは<\/w:t>/);
  assert.deepEqual(await reopened.file('word/media/image1.png').async('nodebuffer'), Buffer.from([1, 2, 3, 4]));
});

test('PPTX follows presentation relationship order and edits slide text only', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId2"/><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Target="slides/slide1.xml" Id="rId1"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>',
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="p" xmlns:a="a"><a:t>One</a:t></p:sld>',
    'ppt/slides/slide2.xml': '<p:sld xmlns:p="p" xmlns:a="a"><a:t>Two</a:t></p:sld>',
    'ppt/theme/theme1.xml': '<a:theme xmlns:a="a"/>',
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.slides[0].items[0].text, 'Two');
  pkg.stageEdit({ id: model.slides[0].items[0].id, value: 'Second' });
  const result = await pkg.exportValidatedCopy();
  assert.deepEqual(result.report.changedParts, ['ppt/slides/slide2.xml']);
});

test('XLSX converts a shared-string cell locally and refuses formula changes', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': '<workbook xmlns:r="r"><sheets><sheet r:id="rId1" name="Data"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Target="worksheets/sheet1.xml" Id="rId1"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><t>Shared</t></si></sst>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1" s="2" t="s"><v>0</v></c><c r="B1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>',
    'xl/styles.xml': '<styleSheet/>',
  });
  const pkg = await OfficePackage.open('sample.xlsx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.sheets[0].cells[0].value, 'Shared');
  pkg.stageEdit({ id: model.sheets[0].cells[0].id, value: 'Local' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const sheet = await reopened.file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sheet, /<c r="A1" s="2" t="s"><v>\d+<\/v><\/c>/);
  const sst = await reopened.file('xl/sharedStrings.xml').async('string');
  assert.ok(sst.includes('<t>Local</t>'), 'shared string table gains the edited value');
  assert.ok(sst.includes('<t>Shared</t>'), 'original shared value retained');

  const pkg2 = await OfficePackage.open('sample.xlsx', original);
  const model2 = await pkg2.createViewModel();
  // Formula cells are now editable: editing B1 with a new formula rewrites <f>.
  assert.equal(model2.sheets[0].cells[1].formula, true);
  assert.equal(model2.sheets[0].cells[1].value, '=1+1');
  pkg2.stageEdit({ id: model2.sheets[0].cells[1].id, value: '=2*3' });
  const result2 = await pkg2.exportValidatedCopy();
  const reopened2 = await JSZip.loadAsync(result2.bytes, { checkCRC32: true });
  const sheet2 = await reopened2.file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sheet2, /<c r="B1"[^>]*><f>2\*3<\/f><\/c>/);
});

test('digitally signed packages are view-only', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:t>Signed</w:t></w:document>',
    '_xmlsignatures/sig1.xml': '<Signature/>',
  });
  const pkg = await OfficePackage.open('signed.docx', original);
  assert.equal(pkg.canEdit, false);
  assert.throws(() => pkg.stageEdit({ id: 'docx:0', value: 'x' }), /署名/);
});
