'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { OfficePackage, escapeXml } = require('../src/office-package');

async function makeZip(parts) {
  const zip = new JSZip();
  Object.entries(parts).forEach(([name, value]) => zip.file(name, value));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

function hasForbiddenControl(s) {
  for (const ch of s) {
    const c = ch.charCodeAt(0);
    if (c <= 0x08 || c === 0x0b || c === 0x0c || (c >= 0x0e && c <= 0x1f)) return true;
  }
  return false;
}

test('escapeXml strips XML 1.0 forbidden control characters', () => {
  // Build a string containing NUL (0x00) and BEL (0x07) via code points.
  const input = 'hi' + String.fromCharCode(0x00) + ' there' + String.fromCharCode(0x07) + 'end';
  const out = escapeXml(input);
  assert.ok(!hasForbiddenControl(out), 'control char survived escaping');
  // Normal content still escapes the five XML predefined entities.
  assert.ok(escapeXml('<a>').includes('&lt;a&gt;'), 'angle brackets not escaped');
});

test('saving a value with control characters yields valid (control-free) XML, not corruption', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  const ctrl = 'A' + String.fromCharCode(0x00) + 'B' + String.fromCharCode(0x07) + 'C';
  pkg.stageEdit({ id: model.items[0].id, value: ctrl });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  assert.ok(!hasForbiddenControl(doc), 'control char leaked into produced XML');
  assert.ok(doc.includes('ABC'), 'edited text present');
});

test('stageEdit rejects an id with path separators or non-numeric suffix', () => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('word/document.xml', '<w:document xmlns:w="w"><w:body/></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' }).then(async (bytes) => {
    const pkg = await OfficePackage.open('x.docx', bytes);
    assert.throws(() => pkg.stageEdit({ id: 'docx:0/../../x', value: 'x' }), /編集データが不正/);
    assert.throws(() => pkg.stageEdit({ id: 'docx', value: 'x' }), /編集データが不正/);
    // A well-formed id is still accepted.
    pkg.stageEdit({ id: 'docx:0', value: 'ok' });
    assert.equal(pkg.edits.get('docx:0'), 'ok');
  });
});

test('exportValidatedCopy rejects an edit whose id matches no real node (orphan edit)', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  pkg.stageEdit({ id: 'docx:99999', value: 'ghost' });
  await assert.rejects(() => pkg.exportValidatedCopy(), /編集対象が見つかりません|有効な変更がありません/);
});

test('relationship targets with ".." are ignored, never resolved outside the package', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/><p:sldId r:id="rId2"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="../outside/slide1.xml" Type="t"/><Relationship Id="rId2" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="p" xmlns:a="a"><a:t>One</a:t></p:sld>',
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  // The "../" target is skipped; only the in-package slide (rId2) is loaded.
  assert.equal(model.slides.length, 1);
  assert.equal(model.slides[0].items[0].text, 'One');
});
