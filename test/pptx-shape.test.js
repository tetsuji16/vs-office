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

test('pptx model extracts shapes with text and fill/line colours', async () => {
  const slide = `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:nvSpPr/><p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></p:spPr><p:txBody><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.slides[0].shapes.length, 1);
  const shape = model.slides[0].shapes[0];
  assert.equal(shape.text, 'Title');
  assert.equal(shape.fillColor, 'ff0000');
  assert.equal(shape.lineColor, '00ff00');
});

test('pptx shape text edit re-fans across runs and preserves markup', async () => {
  const slide = `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>Hello</a:t></a:r><a:r><a:t> world</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  const shapeId = model.slides[0].shapes[0].id;
  pkg.stageEdit({ id: `${shapeId}:text`, value: 'Changed text' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const xml = await reopened.file('ppt/slides/slide1.xml').async('string');
  // Text is re-fanned across runs, so verify both run pieces are present.
  assert.ok(xml.includes('Chang') && xml.includes('ed text'), 'edited shape text not present');
  // Untouched parts remain byte-identical.
  assert.equal(result.report.untouchedPartsVerified, 3);
});

test('pptx shape fill and line colour edits rewrite srgbClr', async () => {
  const slide = `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:nvSpPr/><p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:txBody><a:p><a:r><a:t>X</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  const shapeId = model.slides[0].shapes[0].id;
  pkg.stageEdit({ id: `${shapeId}:fill`, value: '#00FF00' });
  pkg.stageEdit({ id: `${shapeId}:line`, value: '#0000FF' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const xml = await reopened.file('ppt/slides/slide1.xml').async('string');
  assert.ok(/<a:srgbClr val="00ff00"/.test(xml), 'fill colour not updated');
  assert.ok(/<a:ln\b[^>]*><a:solidFill><a:srgbClr val="0000ff"/.test(xml), 'line colour not injected');
});

test('pptx shape delete removes the sp element', async () => {
  const slide = `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>Keep</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>Drop</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  // Delete the second shape (ordinal 1).
  const shapeId = model.slides[0].shapes[1].id;
  pkg.stageEdit({ id: `${shapeId}:delete`, value: '' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const xml = await reopened.file('ppt/slides/slide1.xml').async('string');
  assert.ok(xml.includes('Keep'), 'remaining shape lost');
  assert.ok(!xml.includes('Drop'), 'deleted shape still present');
});

test('pptx applies a batch of shape edits against their original ordinals', async () => {
  const slide = `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>Third</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
  });
  const pkg = await OfficePackage.open('batch.pptx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: `${model.slides[0].shapes[0].id}:delete`, value: '' });
  pkg.stageEdit({ id: `${model.slides[0].shapes[1].id}:delete`, value: '' });
  pkg.stageEdit({ id: `${model.slides[0].shapes[2].id}:fill`, value: '#123456' });

  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const xml = await reopened.file('ppt/slides/slide1.xml').async('string');
  assert.ok(!xml.includes('First') && !xml.includes('Second'), 'both selected shapes must be removed');
  assert.ok(xml.includes('Third'), 'an unselected later shape must remain');
  assert.match(xml, /<a:srgbClr val="123456"/);
});

test('invalid colour value is rejected at save time', async () => {
  const slide = `<p:sld xmlns:p="p" xmlns:a="a"><p:sp><p:nvSpPr/><p:spPr/><p:txBody><a:p><a:r><a:t>X</a:t></a:r></a:p></p:txBody></p:sp></p:sld>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'ppt/presentation.xml': '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId r:id="rId1"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml" Type="t"/></Relationships>',
    'ppt/slides/slide1.xml': slide,
  });
  const pkg = await OfficePackage.open('sample.pptx', original);
  const model = await pkg.createViewModel();
  const shapeId = model.slides[0].shapes[0].id;
  assert.throws(() => pkg.stageEdit({ id: `${shapeId}:fill`, value: 'not-a-colour' }), /色は #RRGGBB/);
});
