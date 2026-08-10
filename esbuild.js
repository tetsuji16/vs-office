'use strict';

const esbuild = require('esbuild');
const fs = require('fs/promises');
const path = require('path');

const root = __dirname;
const production = process.argv.includes('--production');

async function copy(source, destination) {
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await fs.copyFile(source, destination);
}

async function main() {
  const distDir = path.join(root, 'dist');
  const vendorDir = path.join(root, 'media', 'vendor');
  await Promise.all([
    fs.rm(distDir, { recursive: true, force: true }),
    fs.rm(vendorDir, { recursive: true, force: true }),
  ]);

  await esbuild.build({
    entryPoints: [path.join(root, 'extension.js')],
    bundle: true,
    external: ['vscode'],
    format: 'cjs',
    legalComments: 'none',
    minify: production,
    outfile: path.join(distDir, 'extension.js'),
    platform: 'node',
    sourcemap: production ? false : true,
    sourcesContent: false,
    target: 'node18',
  });

  await Promise.all([
    copy(
      path.join(root, 'node_modules', 'docx-preview', 'dist', 'docx-preview.min.js'),
      path.join(vendorDir, 'docx-preview.min.js'),
    ),
    copy(
      path.join(root, 'node_modules', 'jszip', 'dist', 'jszip.min.js'),
      path.join(vendorDir, 'jszip.min.js'),
    ),
    copy(
      path.join(root, 'node_modules', 'docx-preview', 'LICENSE'),
      path.join(distDir, 'licenses', 'docx-preview-LICENSE.txt'),
    ),
    copy(
      path.join(root, 'node_modules', 'jszip', 'LICENSE.markdown'),
      path.join(distDir, 'licenses', 'jszip-LICENSE.md'),
    ),
  ]);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
