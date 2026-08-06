# Change Log

## 0.6.0

Adversarial review hardening and UI consistency (手抜きなし敵対的レビュー対応).

### Robustness / security fixes
- `escapeXml` now also strips lone surrogate code points (U+D800–U+DFFF) in addition to XML 1.0 control characters, so a saved file never contains a character that would make Word/Excel refuse to open it.
- `xlsxModel` decodes XML entities in formula text (`&amp;` → `&`), so formula cells display and round-trip the real formula text instead of the escaped form.
- Added a per-entry uncompressed-size cap (`MAX_ENTRY_BYTES = 200 MB`) on top of the existing total cap, so a single oversized part is rejected during audit rather than after a full inflate.
- Shared-string cell edits now preserve the `t="s"` structure: the value is written into the shared string table (reusing an existing entry when possible) and the cell keeps its shared-string type, instead of being silently rewritten as an inline string. This keeps the OOXML structure intact (保守的方針).

### UI consistency
- The webview now shows Preview / Edit tabs for DOCX and PPTX. PPTX moves shape-editing cards to the Edit tab while the Preview tab keeps the visual slide preview, matching the DOCX layout.
- Error/save toasts are now stacked (`#toasts` container) instead of overwriting each other, so multiple messages are all visible.

### Tests
- Added `test/docx-structure-audit.js`, `test/adversarial-deep.test.js`, `test/xlsx-edit.test.js`, `test/docx-pptx-edit.test.js`, `test/security-limits.test.js` covering shared-string preservation, formula escaping, lone-surrogate stripping, path traversal variants, signed/encrypted read-only mode, and per-entry size caps. Full suite: 54 tests, all passing.

## 0.5.0

- PPTX: 本格的な図形編集を追加。各図形（p:sp）のテキスト・塗りつぶし色（solidFill/srgbClr）・枠線色を直接編集し、図形ごと削除も可能に。テーマ・マスター・アニメーション・画像は変更せず、編集した図形の内部マークアップのみ書き換えます。
- PPTX: 編集 id 体系を拡張（`pptx:<slide>:<shape>:(text|fill|line|delete)`）
- XLSX: 数式セルを編集可能に（`<f>` を書き換え、計算値はアプリ再計算に委ねる）
- DOCX/XLSX/PPTX: 編集済みセル・段落に「編集済み」バッジと変更前後ツールチップを表示
- Webview 単体テスト（jsdom）を追加
- .vsix を再ビルド（v0.5.0）
