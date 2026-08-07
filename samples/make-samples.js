'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const OUT = path.join(__dirname);

async function zip(name, parts) {
  const z = new JSZip();
  for (const [k, v] of Object.entries(parts)) z.file(k, v);
  const buf = await z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(path.join(OUT, name), buf);
  console.log('wrote', name);
}

const ctDocx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const docxRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const docxDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>VS Office サンプル文書</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第1章 概要</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>太字の強調</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>斜体の語句</w:t></w:r></w:p>
<w:p><w:r><w:t>段落編集テスト用のテキストです。編集しても書式が保たれます。</w:t></w:r></w:p>
</w:body>
</w:document>`;

const docxStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`;

const ctPptx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

const pptxRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const pptxPres = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
</p:presentation>`;

const pptxPresRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`;

const pptxSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="6000000" cy="1000000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:t>タイトル図形</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="2500000"/><a:ext cx="6000000" cy="1500000"/></a:xfrm><a:solidFill><a:srgbClr val="FFD966"/></a:solidFill></p:spPr>
<p:txBody><a:bodyPr/><a:p><a:r><a:t>塗りつぶし付き図形（黄色）</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:spTree></p:cSld>
</p:sld>`;

const ctXlsx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;

const xlsxRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const xlsxWb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const xlsxWbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

const xlsxSst = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
<si><t>名前</t></si><si><t>スコア</t></si><si><t>太郎</t></si><si><t>花子</t></si>
</sst>`;

const xlsxSheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>95</v></c></row>
<row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3"><f>SUM(B2:B2)</f><v>95</v></c></row>
</sheetData>
</worksheet>`;

(async () => {
  await zip('sample.docx', {
    '[Content_Types].xml': ctDocx,
    '_rels/.rels': docxRels,
    'word/document.xml': docxDoc,
    'word/styles.xml': docxStyles,
  });
  await zip('sample.pptx', {
    '[Content_Types].xml': ctPptx,
    '_rels/.rels': pptxRels,
    'ppt/presentation.xml': pptxPres,
    'ppt/_rels/presentation.xml.rels': pptxPresRels,
    'ppt/slides/slide1.xml': pptxSlide,
  });
  await zip('sample.xlsx', {
    '[Content_Types].xml': ctXlsx,
    '_rels/.rels': xlsxRels,
    'xl/workbook.xml': xlsxWb,
    'xl/_rels/workbook.xml.rels': xlsxWbRels,
    'xl/sharedStrings.xml': xlsxSst,
    'xl/worksheets/sheet1.xml': xlsxSheet,
  });
  console.log('done');
})();
