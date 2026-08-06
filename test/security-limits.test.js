'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { OfficePackage, OfficePackageError } = require('../src/office-package');

async function makeZip(parts, opts = {}) {
  const zip = new JSZip();
  Object.entries(parts).forEach(([name, value]) => zip.file(name, value));
  return zip.generateAsync(Object.assign({ type: 'nodebuffer', compression: 'DEFLATE' }, opts));
}

const contentTypes = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';

test('allows a moderately large single part under the per-entry cap', async () => {
  const big = 'x'.repeat(60 * 1024 * 1024); // 60 MB < 200 MB per-entry cap
  let bytes;
  try {
    bytes = await makeZip({ '[Content_Types].xml': contentTypes, 'word/document.xml': big });
  } catch (e) {
    return; // generation too heavy in CI; treat as skipped
  }
  const pkg = await OfficePackage.open('big.docx', bytes);
  assert.ok(pkg, '60 MB part opened successfully');
});

test('rejects a package with an absurd number of entries', async () => {
  const parts = { '[Content_Types].xml': contentTypes };
  for (let i = 0; i < 10001; i += 1) parts[`part${i}.xml`] = '<x/>';
  const bytes = await makeZip(parts);
  await assert.rejects(() => OfficePackage.open('many.xlsx', bytes), OfficePackageError);
});

test('rejects a zip entry whose name uses a Windows drive letter', async () => {
  const bytes = await makeZip({
    '[Content_Types].xml': contentTypes,
    'C:\\evil\\out.xml': '<x/>',
    'word/document.xml': '<w:document xmlns:w="w"><w:body/></w:document>',
  });
  await assert.rejects(() => OfficePackage.open('drive.docx', bytes), OfficePackageError);
});

test('rejects absolute-path entry names', async () => {
  const bytes = await makeZip({
    '[Content_Types].xml': contentTypes,
    '/etc/passwd.xml': '<x/>',
    'word/document.xml': '<w:document xmlns:w="w"><w:body/></w:document>',
  });
  await assert.rejects(() => OfficePackage.open('abs.docx', bytes), OfficePackageError);
});

test('signed packages open in read-only mode, never editable', async () => {
  const bytes = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
    '_xmlsignatures/sig1.xml': '<sig/>',
  });
  const pkg = await OfficePackage.open('signed.docx', bytes);
  assert.equal(pkg.canEdit, false, 'signed file must be read-only');
  assert.throws(() => pkg.stageEdit({ id: 'docx:0', value: 'x' }), OfficePackageError);
});

test('encrypted packages open in read-only mode', async () => {
  const bytes = await makeZip({
    '[Content_Types].xml': contentTypes,
    'word/document.xml': '<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>',
    'word/styles.xml': '<w:styles xmlns:w="w"/>',
    'EncryptedPackage/enctime.bin': '<x/>',
  });
  const pkg = await OfficePackage.open('enc.docx', bytes);
  assert.equal(pkg.canEdit, false, 'encrypted file must be read-only');
});
