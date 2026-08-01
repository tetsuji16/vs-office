# VS Office - Lightweight Editor

<p align="center"><img src="assets/vs-office-icon.png" width="160" alt="VS Office icon"></p>

VS Office is a lightweight Visual Studio Code extension for viewing `.docx`, `.pptx`, and `.xlsx` files and making carefully limited edits without rebuilding the entire Office document.

## Safety principles

- Save changes to the original file with standard `Ctrl+S`, or choose Save As when you need a separate copy.
- Write to a temporary file in the same directory and replace the original atomically only after validation succeeds.
- Create a recoverable copy in `.vs-office-backups` before overwriting by default. This can be disabled in Settings.
- Modify only the targeted OOXML text or cell data. Themes, images, masters, styles, formulas, embedded files, and other untouched package parts are not regenerated.
- Recheck ZIP CRC values after saving and verify that every untouched OOXML part remains byte-for-byte identical using SHA-256 hashes. Saving fails if validation detects a mismatch.
- Open digitally signed or encrypted files in read-only mode.
- Keep formula cells read-only.

## Supported features

| Format | Viewing | Conservative editing |
| --- | --- | --- |
| DOCX | Page preview powered by `docx-preview`, plus a text list | Existing `w:t` text runs |
| PPTX | Structural view of slide order, text, and image counts | Existing `a:t` text runs |
| XLSX | Lightweight grid of up to 200 rows by 50 columns | Non-formula cells; shared-string cells are converted individually to `inlineStr` |

## Development

```powershell
npm install
npm test
```

Open this folder in Visual Studio Code and press `F5` to start an Extension Development Host.

## Important limitations

Browser HTML and Microsoft Office use different text layout engines, so pixel-perfect preview parity cannot be guaranteed. Editing text can also cause normal line wrapping or page and slide reflow when the document is reopened in Word or PowerPoint. VS Office is designed to preserve all unrelated layout information and package content despite those unavoidable text-layout changes. Always verify final deliverables in Microsoft Office.

## Open-source components

- [docxjs / docx-preview](https://github.com/VolodymyrBaydalka/docxjs) (Apache-2.0) renders DOCX content as HTML.
- [JSZip](https://github.com/Stuk/jszip) (MIT / GPLv3) reads and writes OOXML ZIP containers and validates CRC values.

[SheetJS Community Edition](https://git.sheetjs.com/sheetjs/sheetjs) was evaluated for XLSX support. It is not used for saving because regenerating an entire workbook can discard unknown features or layout information. PPTXjs-based renderers were also evaluated, but the current PPTX implementation intentionally uses a structural view because those projects could not provide the required fidelity guarantees.
