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
  assert.ok(/<c[^>]*r="C1"[^>]*t="str"/.test(out), 'existing string-result formula type is preserved');
  assert.ok(!/<v>/.test(out.split('r="C1"')[1].split('</c>')[0]), 'cached <v> dropped');
});

test('xlsx: editing a numeric formula does not force its result type to string', async () => {
  const numericFormulaSheet = `<?xml version="1.0"?>
<worksheet xmlns="x"><sheetData><row r="1">
  <c r="A1" t="n"><f>1+1</f><v>2</v></c>
</row></sheetData></worksheet>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': numericFormulaSheet,
  });
  const pkg = await OfficePackage.open('numeric-formula.xlsx', original);
  const model = await pkg.createViewModel();
  const a1 = model.sheets[0].cells.find((cell) => cell.ref === 'A1');
  pkg.stageEdit({ id: a1.id, value: '=2*3' });
  const result = await pkg.exportValidatedCopy();
  const out = await (await JSZip.loadAsync(result.bytes, { checkCRC32: true }))
    .file('xl/worksheets/sheet1.xml').async('string');
  const cellXml = out.match(/<c\b[^>]*r="A1"[^>]*>[\s\S]*?<\/c>/)[0];
  assert.ok(cellXml.includes('<f>2*3</f>'));
  assert.ok(!/\bt="str"/.test(cellXml), 'numeric formula must not be serialized as a string result');
  assert.ok(!/<v>/.test(cellXml), 'stale cached numeric value is removed');
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

test('xlsx: reads and edits namespace-prefixed sheets with package-absolute relationship targets', async () => {
  const prefixedWorkbook = '<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets><x:sheet name="編集検証" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>';
  const absoluteRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="/xl/worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>';
  const prefixedSheet = '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" t="str"><x:v>Sample</x:v></x:c><x:c r="B1" t="n"><x:v>42</x:v></x:c><x:c r="C1" t="n"><x:f>B1*2</x:f><x:v>84</x:v></x:c></x:row></x:sheetData></x:worksheet>';
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': prefixedWorkbook,
    'xl/_rels/workbook.xml.rels': absoluteRels,
    'xl/worksheets/sheet1.xml': prefixedSheet,
  });
  const pkg = await OfficePackage.open('prefixed.xlsx', original);
  const model = await pkg.createViewModel();
  assert.equal(model.sheets.length, 1);
  assert.equal(model.sheets[0].name, '編集検証');
  assert.equal(model.sheets[0].cells.find((cell) => cell.ref === 'B1').text, '42');
  assert.equal(model.sheets[0].cells.find((cell) => cell.ref === 'C1').formulaText, 'B1*2');
  const target = model.sheets[0].cells.find((cell) => cell.ref === 'B1');
  const formulaTarget = model.sheets[0].cells.find((cell) => cell.ref === 'C1');
  pkg.stageEdit({ id: target.id, value: '99' });
  pkg.stageEdit({ id: formulaTarget.id, value: '=B1*3' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await OfficePackage.open('prefixed.xlsx', result.bytes);
  const reopenedModel = await reopened.createViewModel();
  assert.equal(reopenedModel.sheets[0].cells.find((cell) => cell.ref === 'B1').text, '99');
  assert.equal(reopenedModel.sheets[0].cells.find((cell) => cell.ref === 'C1').formulaText, 'B1*3');
  const outputSheet = await (await JSZip.loadAsync(result.bytes, { checkCRC32: true })).file('xl/worksheets/sheet1.xml').async('string');
  assert.ok(outputSheet.includes('<x:f>B1*3</x:f>'), 'formula edit preserves the source namespace prefix');
});

test('xlsx: self-closing blank cells do not steal the next populated cell value', async () => {
  const realWorldSheet = `<?xml version="1.0"?>
<x:worksheet xmlns:x="x">
  <x:sheetData>
    <x:row r="1">
      <x:c r="A1" t="str"><x:v>Title</x:v></x:c>
      <x:c r="B1" s="3" />
      <x:c r="C1" s="3" />
    </x:row>
    <x:row r="2">
      <x:c r="A2" t="str"><x:v>Sales total</x:v></x:c>
      <x:c r="B2"><x:v>15310</x:v></x:c>
    </x:row>
  </x:sheetData>
</x:worksheet>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/worksheets/sheet1.xml': realWorldSheet,
  });
  const pkg = await OfficePackage.open('self-closing.xlsx', original);
  const model = await pkg.createViewModel();
  const byRef = Object.fromEntries(model.sheets[0].cells.map((cell) => [cell.ref, cell]));

  assert.equal(byRef.A1.text, 'Title');
  assert.equal(byRef.B1.text, '');
  assert.equal(byRef.C1.text, '');
  assert.equal(byRef.A2.text, 'Sales total');
  assert.equal(byRef.B2.text, '15310');

  pkg.stageEdit({ id: byRef.B1.id, value: 'new value' });
  const result = await pkg.exportValidatedCopy();
  const reopened = await OfficePackage.open('self-closing.xlsx', result.bytes);
  const edited = (await reopened.createViewModel()).sheets[0].cells.find((cell) => cell.ref === 'B1');
  assert.equal(edited.text, 'new value', 'a self-closing blank cell can be edited and read back');
});

test('xlsx: editing one shared string preserves unrelated rich text and extensions', async () => {
  const richSharedStrings = `<?xml version="1.0"?>
<x:sst xmlns:x="x" count="2" uniqueCount="2">
  <x:si><x:r><x:rPr><x:b/></x:rPr><x:t>Rich</x:t></x:r></x:si>
  <x:si><x:t>Plain</x:t></x:si>
  <x:extLst><x:ext uri="{must-survive}"/></x:extLst>
</x:sst>`;
  const sharedSheet = `<?xml version="1.0"?>
<x:worksheet xmlns:x="x"><x:sheetData><x:row r="1">
  <x:c r="A1" t="s"><x:v>0</x:v></x:c>
  <x:c r="B1" t="s"><x:v>1</x:v></x:c>
</x:row></x:sheetData></x:worksheet>`;
  const original = await makeZip({
    '[Content_Types].xml': contentTypes,
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels,
    'xl/sharedStrings.xml': richSharedStrings,
    'xl/worksheets/sheet1.xml': sharedSheet,
  });
  const pkg = await OfficePackage.open('rich-shared.xlsx', original);
  const model = await pkg.createViewModel();
  const b1 = model.sheets[0].cells.find((cell) => cell.ref === 'B1');
  pkg.stageEdit({ id: b1.id, value: 'Changed' });
  const result = await pkg.exportValidatedCopy();
  const zip = await JSZip.loadAsync(result.bytes, { checkCRC32: true });
  const sst = await zip.file('xl/sharedStrings.xml').async('string');

  assert.ok(sst.includes('<x:rPr><x:b/></x:rPr>'), 'unrelated rich-text formatting survives byte-for-byte');
  assert.ok(sst.includes('<x:ext uri="{must-survive}"/>'), 'unrelated extension metadata survives');
  assert.match(sst, /count="2"/, 'cell usage count remains unchanged');
  assert.match(sst, /uniqueCount="3"/, 'one appended unique string is counted');
  assert.ok(sst.includes('<x:si><x:t>Changed</x:t></x:si>'), 'new entry uses the existing namespace prefix');
  assert.ok(sst.indexOf('<x:si><x:t>Changed</x:t></x:si>') < sst.indexOf('<x:extLst>'), 'extLst remains the final SST child');

  const reopened = await OfficePackage.open('rich-shared.xlsx', result.bytes);
  const cells = Object.fromEntries((await reopened.createViewModel()).sheets[0].cells.map((cell) => [cell.ref, cell.text]));
  assert.equal(cells.A1, 'Rich');
  assert.equal(cells.B1, 'Changed');
});
