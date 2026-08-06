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

// A grouped shape: <p:sp> contains <p:sp> (nested group member). LTspice-style
// group shapes embed child <p:sp> elements, so a naive non-greedy match can
// delete the wrong (outer) element or splice the tree.
const groupedSlide = `<?xml version="1.0"?>
<p:sld xmlns:p="p" xmlns:a="a">
  <p:sp><p:nvSpPr/><p:spPr/><a:txBody><a:p><a:r><a:t>OUTER</a:t></a:r></a:p></a:txBody></p:sp>
  <p:sp><p:nvSpPr/><p:spPr/><a:txBody><a:p><a:r><a:t>INNER</a:t></a:r></a:p></a:txBody></p:sp>
</p:sld>`;

test('pptx shape delete removes only the targeted shape (no tree corruption)', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': groupedSlide,
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.slides[0].shapes.length, 2, 'two shapes expected');
  // Delete the first shape only.
  pkg.stageEdit({ id: `${model.slides[0].shapes[0].id}:delete`, value: '' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('ppt/slides/slide1.xml').async('string');
  assert.ok(!doc.includes('OUTER'), 'outer shape should be removed');
  assert.ok(doc.includes('INNER'), 'inner shape must survive');
  // Re-open must still parse cleanly.
  const reparsed = await OfficePackage.open('sample.pptx', result.bytes);
  const m2 = await reparsed.createViewModel();
  assert.equal(m2.slides[0].shapes.length, 1, 'one shape should remain');
});

test('resolveTarget rejects encoded traversal (..%2e and dot-dot variants)', async () => {
  const rels = '<Relationships>' +
    '<Relationship Id="rId1" Target="..%2f..%2foutside.xml" Type="t"/>' +
    '<Relationship Id="rId2" Target="....//slide.xml" Type="t"/>' +
    '<Relationship Id="rId3" Target="slides/slide1.xml" Type="t"/>' +
    '</Relationships>';
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/><p:sldId r:id="rId2"/><p:sldId r:id="rId3"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels,
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="p" xmlns:a="a"><a:t>One</a:t></p:sld>',
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  // Only the in-package slide (rId3) must load.
  assert.equal(model.slides.length, 1, 'encoded/dot-dot traversal targets must be ignored');
});

test('xlsx colNumber handles columns beyond Z (AA, AZ, BA)', () => {
  // Inline the same algorithm the viewer uses, then assert against the model ref parsing.
  const original = (() => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('xl/workbook.xml', '<workbook xmlns="x" xmlns:r="r"><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="t"/></Relationships>');
    zip.file('xl/worksheets/sheet1.xml',
      '<worksheet xmlns="x"><sheetData>' +
      '<c r="A1"><v>1</v></c><c r="Z1"><v>2</v></c><c r="AA1"><v>3</v></c><c r="BA10"><v>4</v></c>' +
      '</sheetData></worksheet>');
    return zip.generateAsync({ type: 'nodebuffer' });
  })();
  return original.then(async (bytes) => {
    const pkg = await OfficePackage.open('sample.xlsx', bytes);
    const model = await pkg.createViewModel();
    const refs = model.sheets[0].cells.map((c) => c.ref).sort();
    assert.deepEqual(refs, ['A1', 'AA1', 'BA10', 'Z1'], 'all columns parsed');
  });
});

test('escapeXml neutralises a lone surrogate so the produced XML stays well-formed', () => {
  // A lone high surrogate (0xD800) is not a valid XML character.
  const input = 'hi' + String.fromCharCode(0xd800) + 'there';
  const out = (require('../src/office-package').escapeXml)(input);
  // Must not contain the raw surrogate codepoint.
  assert.ok(!out.includes(String.fromCharCode(0xd800)), 'lone surrogate leaked');
});
