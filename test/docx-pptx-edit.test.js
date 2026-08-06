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

// A shape with a gradient/pattern fill (NOT a solidFill) plus a line with no solid colour.
// Editing the shape fill must inject a solidFill without destroying the existing markup.
const shapeDoc = `<?xml version="1.0"?>
<p:sld xmlns:p="p" xmlns:a="a">
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>
    </p:spPr>
    <a:txBody><a:p><a:r><a:t>Label</a:t></a:r></a:p></a:txBody>
  </p:sp>
</p:sld>`;

test('pptx: set fill on a shape that had a gradient fill injects solidFill, keeps geometry', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': shapeDoc,
  });
  const pkg = await OfficePackage.open('s.pptx', original);
  const model = await pkg.createViewModel();
  const shape = model.slides[0].shapes[0];
  pkg.stageEdit({ id: `${shape.id}:fill`, value: '#112233' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await rz.file('ppt/slides/slide1.xml').async('string');
  assert.ok(doc.includes('<a:solidFill><a:srgbClr val="112233"/>'), 'solidFill injected');
  // Geometry must survive.
  assert.ok(doc.includes('prst="rect"'), 'preset geometry preserved');
  assert.ok(doc.includes('<a:t>Label</a:t>'), 'text preserved');
  // Re-open must still parse.
  const reparsed = await OfficePackage.open('s.pptx', result.bytes);
  const m2 = await reparsed.createViewModel();
  assert.equal(m2.slides[0].shapes[0].fillColor, '112233', 'new fill colour read back');
});

test('pptx: set a line colour on a shape with no line injects a line', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': shapeDoc,
  });
  const pkg = await OfficePackage.open('s.pptx', original);
  const model = await pkg.createViewModel();
  const shape = model.slides[0].shapes[0];
  pkg.stageEdit({ id: `${shape.id}:line`, value: '#abcdef' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await rz.file('ppt/slides/slide1.xml').async('string');
  assert.ok(doc.includes('<a:ln>') && doc.includes('srgbClr val="abcdef"'), 'line with colour injected');
});

test('pptx: delete a shape keeps other slides/parts intact and re-opens cleanly', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': shapeDoc,
  });
  const pkg = await OfficePackage.open('s.pptx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: `${model.slides[0].shapes[0].id}:delete`, value: '' });
  const result = await pkg.exportValidatedCopy();
  const reparsed = await OfficePackage.open('s.pptx', result.bytes);
  const m2 = await reparsed.createViewModel();
  assert.equal(m2.slides[0].shapes.length, 0, 'shape removed');
});

// DOCX: edit a single w:t text node (not a paragraph) must not disturb siblings.
const docxTwoRuns = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>First</w:t></w:r><w:r><w:t>Second</w:t></w:r></w:p>
  </w:body>
</w:document>`;

test('docx: editing one w:t node leaves the other w:t nodes intact', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': docxTwoRuns,
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('d.docx', original);
  const model = await pkg.createViewModel();
  const first = model.items[0];
  assert.equal(first.text, 'First');
  pkg.stageEdit({ id: first.id, value: 'CHANGED' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await rz.file('word/document.xml').async('string');
  const texts = Array.from(doc.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g), (m) => m[1]);
  assert.deepEqual(texts, ['CHANGED', 'Second'], 'only the targeted run changed');
});
