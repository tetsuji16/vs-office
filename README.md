# VS Office

<p align="center"><img src="assets/vs-office-icon.png" width="160" alt="VS Office icon"></p>

VS Code 内で `.docx`、`.pptx`、`.xlsx` を軽量に閲覧し、安全範囲に限定して編集する拡張機能です。

## 設計原則

- 通常の `Ctrl+S` で元ファイルへ保存できます。名前を付けて保存も選べます。
- 保存時は同じフォルダーの一時ファイルへ書き、検証合格後に原子的に置換します。
- 既定では上書き前のファイルを `.vs-office-backups` に退避します。設定で無効化できます。
- OOXML パッケージ中の対象テキスト／セルだけを変更し、テーマ、画像、マスター、スタイル、数式、埋め込みなどは再生成しません。
- 保存後にZIP CRCを再検査し、未編集パートが元データとSHA-256で一致することを確認します。不一致なら保存処理を失敗させます。
- デジタル署名付き・暗号化ファイルは閲覧専用です。
- 数式セルは閲覧専用です。

## 対応機能

| 形式 | 表示 | 保守的編集 |
| --- | --- | --- |
| DOCX | `docx-preview` によるページ表示＋テキスト一覧 | 既存の `w:t` テキスト断片 |
| PPTX | スライド順・テキスト・画像数の構造表示 | 既存の `a:t` テキスト断片 |
| XLSX | 最大200行×50列の軽量グリッド | 非数式セル。共有文字列は対象セルだけ `inlineStr` 化 |

## 開発

```powershell
npm install
npm test
```

VS Code でこのフォルダーを開き、`F5` で Extension Development Host を起動してください。

## 重要な制限

ブラウザHTMLとMicrosoft Officeでは文字組みエンジンが異なるため、プレビューの完全一致は保証できません。また、文字を変更すればWord/PowerPoint側で通常の改行・ページ送り変化が起こり得ます。この拡張は、その不可避な変化以外のレイアウト情報を変更しないことを重視しています。最終成果物はMicrosoft Officeで確認してください。

## 採用したOSS

- [docxjs / docx-preview](https://github.com/VolodymyrBaydalka/docxjs) (Apache-2.0): DOCXのHTMLレンダリング。
- [JSZip](https://github.com/Stuk/jszip) (MIT / GPLv3): OOXML ZIPコンテナの読み書きとCRC検証。

XLSXの調査では [SheetJS Community Edition](https://git.sheetjs.com/sheetjs/sheetjs) も検討しましたが、ワークブック全体を再生成すると未知機能やレイアウトを落とし得るため、保存処理には採用していません。PPTXjs系は表示実装が古く、忠実度を保証できないため、PPTXは構造ビューに限定しています。
