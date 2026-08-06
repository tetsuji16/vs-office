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

const workbook = '<workbook xmlns="x" xmlns:r="r"><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>';
const rels = '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="t"/></Relationships>';
const shared = '<?xml version="1.0"?><sst xmlns="x"><si><t>SharedA</t></si><si><t>SharedB</t></si></sst>';

// A sheet with: a shared-string cell (A1), an inline-string cell (B1), a formula
// cell (C1), and a number cell (D1). Covers every branch of xlsxModel + buildChangedParts.
const sheet = `<?xml version="1.0"?>
<worksheet xmlns="x">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="inlineStr"><is><t>InlineHere</t></is></c>
      <c r="C1" t="str"><f>A1&amp;B1</f><v>SharedAInlineHere</v></c>
      <c r="D1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>`;

test('xlsx: edit a shared-string cell preserves t="s" and updates the shared table', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/sharedStrings.xml': shared,
    'xl/worksheets/sheet1.xml': sheet,
  });
  const pkg = await OfficePackage.open('s.xlsx', original);
  const model = await pkg.createViewModel();
  const a1 = model.sheets[0].cells.find((c) => c.ref === 'A1');
  assert.equal(a1.text, 'SharedA');
  pkg.stageEdit({ id: a1.id, value: 'ChangedShared' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const out = await rz.file('xl/worksheets/sheet1.xml').async('string');
  // Cell stays a shared-string reference (type s), pointing at a new shared entry.
  assert.ok(/<c[^>]*r="A1"[^>]*t="s"/.test(out), 'cell type preserved as shared-string');
  const sst = await rz.file('xl/sharedStrings.xml').async('string');
  assert.ok(sst.includes('ChangedShared'), 'shared string table updated with new value');
  // Original untouched string must remain for other cells that reference it.
  assert.ok(sst.includes('SharedA'), 'original shared value retained');
  // Re-open must read the edited value back.
  const reparsed = await OfficePackage.open('s.xlsx', result.bytes);
  const m2 = await reparsed.createViewModel();
  const a1b = m2.sheets[0].cells.find((c) => c.ref === 'A1');
  assert.equal(a1b.text, 'ChangedShared', 'edited shared value read back');
});

test('xlsx: edit an inline-string cell updates the inline text', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': sheet,
  });
  const pkg = await OfficePackage.open('s.xlsx', original);
  const model = await pkg.createViewModel();
  const b1 = model.sheets[0].cells.find((c) => c.ref === 'B1');
  assert.equal(b1.text, 'InlineHere');
  pkg.stageEdit({ id: b1.id, value: 'NewInline' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const out = await rz.file('xl/worksheets/sheet1.xml').async('string');
  assert.ok(out.includes('<is><t>NewInline</t></is>'), 'inline string updated');
});

test('xlsx: edit a formula cell replaces the formula and drops cached value', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': sheet,
  });
  const pkg = await OfficePackage.open('s.xlsx', original);
  const model = await pkg.createViewModel();
  const c1 = model.sheets[0].cells.find((c) => c.ref === 'C1');
  assert.equal(c1.formula, true);
  assert.equal(c1.formulaText, 'A1&B1');
  pkg.stageEdit({ id: c1.id, value: '=A1+B1' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const out = await rz.file('xl/worksheets/sheet1.xml').async('string');
  assert.ok(out.includes('<f>A1+B1</f>'), 'formula replaced (XML-escaped ampersand is correct)');
  assert.ok(!/<v>/.test(out.split('r="C1"')[1].split('</c>')[0]), 'cached <v> dropped');
});

test('xlsx: edit a number cell keeps numeric type and value', async () => {
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': sheet,
  });
  const pkg = await OfficePackage.open('s.xlsx', original);
  const model = await pkg.createViewModel();
  const d1 = model.sheets[0].cells.find((c) => c.ref === 'D1');
  assert.equal(d1.type, 'n');
  pkg.stageEdit({ id: d1.id, value: '99' });
  const result = await pkg.exportValidatedCopy();
  const rz = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const out = await rz.file('xl/worksheets/sheet1.xml').async('string');
  assert.ok(/<c[^>]*r="D1"[^>]*><v>99<\/v><\/c>/.test(out), 'numeric value preserved as <v>99</v>');
});
