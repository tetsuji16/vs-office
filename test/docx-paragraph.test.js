'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const {
  OfficePackage,
  splitProportional,
} = require('../src/office-package');

async function makeZip(parts) {
  const zip = new JSZip();
  Object.entries(parts).forEach(([name, value]) => zip.file(name, value));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

test('splitProportional distributes text across runs by original length', () => {
  assert.deepEqual(splitProportional('HELLO', [3, 2]), ['HEL', 'LO']);
  assert.deepEqual(splitProportional('Hi', [3, 2]), ['H', 'i']);
  assert.deepEqual(splitProportional('', [3, 2]), ['', '']);
  assert.deepEqual(splitProportional('xyz', [0, 0]), ['', '']);
  // a single run keeps every character
  assert.deepEqual(splitProportional('entire paragraph text', [21]), ['entire paragraph text']);
  // long text overflowing the ratio rounds sensibly and the last run absorbs the remainder
  const out = splitProportional('abcdefghij', [1, 1, 1]);
  assert.equal(out.join(''), 'abcdefghij');
  assert.ok(out.every((s) => s.length >= 2 && s.length <= 4));
});

test('DOCX model exposes paragraphs grouped from w:t runs with format preservation', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml':
      '<w:document xmlns:w="w"><w:body>'
      + '<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>'
      + '<w:p><w:r><w:t>Second</w:t></w:r></w:p>'
      + '</w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.paragraphs.length, 2);
  assert.equal(model.paragraphs[0].text, 'Hello world');
  assert.equal(model.paragraphs[1].text, 'Second');
  // Each paragraph records the w:t ids it spans so edits can be re-fanned out.
  assert.equal(model.paragraphs[0].tIds.length, 2);
});

test('paragraph edit re-fans text across runs without touching formatting', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml':
      '<w:document xmlns:w="w"><w:body>'
      + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>world</w:t></w:r></w:p>'
      + '</w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  const para = model.paragraphs[0];
  pkg.stageEdit({ id: para.id, value: 'Bonjour le monde' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  // The two runs and their distinct formatting must survive.
  assert.ok(doc.includes('<w:rPr><w:b/></w:rPr>'), 'bold formatting lost');
  assert.ok(doc.includes('<w:rPr><w:i/></w:rPr>'), 'italic formatting lost');
  // And the joined text equals the edited value.
  const reModel = await OfficePackage.open('sample.docx', result.bytes);
  assert.equal((await reModel.createViewModel()).paragraphs[0].text, 'Bonjour le monde');
});

test('empty paragraph edit clears every run but keeps the paragraph structure', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml':
      '<w:document xmlns:w="w"><w:body>'
      + '<w:p><w:r><w:t>Keep</w:t></w:r><w:r><w:t>Structure</w:t></w:r></w:p>'
      + '</w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: model.paragraphs[0].id, value: '' });
  const result = await pkg.exportValidatedCopy();
  const doc = await (await JSZip.loadAsync(result.bytes, { checkCRC32: true }))
    .file('word/document.xml').async('string');
  assert.ok(/<w:p>/.test(doc), 'paragraph element removed');
  assert.ok(!doc.includes('Keep') && !doc.includes('Structure'), 'text survived');
});

test('paragraph edit and individual w:t edit can coexist on different paragraphs', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml':
      '<w:document xmlns:w="w"><w:body>'
      + '<w:p><w:r><w:t>First</w:t></w:r></w:p>'
      + '<w:p><w:r><w:t>Second</w:t></w:r></w:p>'
      + '</w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: model.paragraphs[0].id, value: 'Premier' });
  pkg.stageEdit({ id: model.items[1].id, value: 'Deuxième' });
  const result = await pkg.exportValidatedCopy();
  const reModel = await (await OfficePackage.open('sample.docx', result.bytes)).createViewModel();
  assert.equal(reModel.paragraphs[0].text, 'Premier');
  assert.equal(reModel.paragraphs[1].text, 'Deuxième');
});
