# Change Log

## 0.3.0

- DOCX: 段落単位の安全編集を追加。段落内の複数 run へ文字数比率で自動配分し、太字・斜体などの書式（w:rPr）を保持したままテキストだけを置換します。
- DOCX: 編集 UI をテキスト断片（w:t）単位から段落単位に刷新。
- PPTX: スライド内テキストの視覚的プレビュー（16:9 スライド風）を構造ビューに追加。
- XLSX: 数式セルを読み取り専用表示にし、数式文字列をツールチップで確認可能に（編集は引き続きブロック）。
- テストを 4 件から 19 件に拡充。XML エスケープ完全性、攻撃的入力・境界値、段落編集の書式保持、数式セル保護をカバー。
- 保守的保存パイプライン（未編集パーツの SHA-256 一致＋ZIP CRC 検証）を維持。

## 0.2.2

- Add a new product hero image for GitHub and the Visual Studio Marketplace.
- Rewrite the product introduction around focused editing, format support, and save protection.
- Add a dark Marketplace gallery banner that matches the product artwork.

## 0.2.1

- Publish the Marketplace and GitHub introduction in English.

## 0.2.0

- Add lightweight DOCX page preview and conservative text editing.
- Add PPTX slide structure preview and conservative text editing.
- Add XLSX grid preview and non-formula cell editing.
- Support standard `Ctrl+S`, Save As, undo, redo and VS Code dirty state.
- Validate ZIP CRC and untouched OOXML part hashes before saving.
- Save through a validated temporary file and atomic replacement.
- Detect external file changes and block unsafe overwrites.
- Add optional automatic backups and a transparent Marketplace icon.
- Standardize the product name and public identifiers as VS Office.
