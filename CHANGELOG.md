# Change Log

## 0.6.3

- 拡張ホスト側のコードを esbuild で単一ファイルにバンドルし、VSIX から開発用ソースと `node_modules` を除外。
- Webview が必要とするブラウザ用ライブラリと第三者ライセンスだけを公開物へ収録。
- Marketplace 公開ワークフローを Node.js 24 および Node 24 対応版の公式Actionsへ更新。

## 0.6.2

- XLSX の名前空間接頭辞付き要素（`x:sheet`、`x:c`、`x:f` など）を読み書きできるように改善。
- OPCで有効なパッケージ絶対Relationship Target（例: `/xl/worksheets/sheet1.xml`）を安全に解決し、実ファイルが0シートとして表示される問題を修正。
- 実運用に近いXLSXサンプルと回帰テストで、数値セルの編集・再保存・再読込を検証。

## 0.6.1

- PPTX の複数編集で、図形削除後にテキスト断片の序数がずれて別の文字列を書き換える問題を修正。
- 図形の塗り色と枠線色を独立して検出・更新し、自己終了型の `p:spPr` でも有効な OOXML 内に色要素を配置。
- Marketplace への同一バージョン再公開を安全にスキップし、VSIX から GitHub ワークフローを除外。
- XML 制御文字の正規表現を可視エスケープへ変更し、ソースを通常のテキスト差分としてレビュー可能に改善。

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
