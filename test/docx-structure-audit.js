'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { OfficePackage, splitProportional } = require('../src/office-package');

async function makeZip(parts) {
  const zip = new JSZip();
  Object.entries(parts).forEach(([name, value]) => zip.file(name, value));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

// Two runs: run0 length 1 ("A"), run1 length 2 ("BC").
// Editing the paragraph to the SHORTER text "AC" must not drop the "B".
const twoRunDoc = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>BC</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

test('splitProportional must not drop characters when new text is shorter than runs', () => {
  // run lengths 1 and 2; new text "AC" (2 chars) is shorter than 3.
  const out = splitProportional('AC', [1, 2]);
  assert.equal(out.join(''), 'AC', 'characters were dropped during fan-out');
  assert.equal(out.length, 2, 'run count must stay constant');
});

test('docx paragraph edit keeps every character when replacement is shorter (regression)', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': twoRunDoc,
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  // Replace the whole paragraph "ABC" with the shorter "AC".
  pkg.stageEdit({ id: model.paragraphs[0].id, value: 'AC' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  // Concatenate all w:t text in document order.
  const texts = Array.from(doc.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g), (m) => m[1]);
  assert.equal(texts.join(''), 'AC', `expected "AC" but got "${texts.join('')}"`);
  // Bold and italic formatting on the two runs must survive.
  assert.ok(/<w:rPr><w:b\/><\/w:rPr>/.test(doc), 'bold rPr lost');
  assert.ok(/<w:rPr><w:i\/><\/w:rPr>/.test(doc), 'italic rPr lost');
});

test('docx paragraph edit longer than runs still preserves every character', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': twoRunDoc,
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
  });
  const pkg = await OfficePackage.open('sample.docx', original);
  const model = await pkg.createViewModel();
  pkg.stageEdit({ id: model.paragraphs[0].id, value: 'XYZZY' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const doc = await reopened.file('word/document.xml').async('string');
  const texts = Array.from(doc.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g), (m) => m[1]);
  assert.equal(texts.join(''), 'XYZZY', `expected "XYZZY" but got "${texts.join('')}"`);
});
