/**
 * Create a synthetic DOCX fixture with straight lines and a horizontal rule
 * drawn as DrawingML shapes (`wps:wsp`), for upstream #972.
 *
 * Every drawing is a bare `w:drawing` in its run (no `mc:AlternateContent`
 * wrapper) and carries a `wp:docPr descr` so the spec can find its painted
 * image by alt text:
 *
 *   - "Vertical line": inline `prst="line"`, zero width (cx="0"), 1.5pt red
 *   - "Horizontal line": inline `prst="line"`, zero height (cy="0"), 1pt navy
 *   - "Diagonal line": inline `prst="line"` with flipV, 2pt green
 *   - "Full-width rule": anchored `wp:wrapTopAndBottom` `prst="rect"` as wide
 *     as the text column, zero height, `a:noFill`, 1.5pt grey outline
 *   - "Margin connector": anchored `wp:wrapNone` `prst="straightConnector1"`
 *     (`wps:cNvCnPr`), zero width, dashed, in the left page margin
 *
 * All text and package metadata are synthetic.
 *
 * Run: bun scripts/create-vector-lines-fixture.mjs
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'e2e/fixtures/vector-lines-and-rules.docx');
const FIXTURE_DATE = new Date('2026-01-01T00:00:00Z');

const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
/** Letter page with 1in margins: the text column is 6.5in wide. */
const COLUMN_WIDTH_EMU = 5943600;

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Synthetic Lines and Rules</dc:title>
  <dc:creator>docx-editor fixture generator</dc:creator>
  <cp:lastModifiedBy>docx-editor fixture generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>docx-editor fixture generator</Application>
</Properties>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="0" w:after="160" w:line="276" w:lineRule="auto"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
</w:styles>`;

function run(text) {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function paragraph(...runs) {
  return `<w:p>${runs.join('')}</w:p>`;
}

/**
 * A `wps:wsp` with a preset geometry and an `a:ln` outline; `extra` goes
 * inside `a:ln` (a dash), `fill` before it (`a:noFill`).
 */
function shape({ prst, cx, cy, lineWidth, color, flips = '', fill = '', extra = '', connector }) {
  const nonVisual = connector ? '<wps:cNvCnPr/>' : '<wps:cNvSpPr/>';
  return `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="${WPS_URI}">
        <wps:wsp>
          ${nonVisual}
          <wps:spPr>
            <a:xfrm${flips}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
            <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
            ${fill}
            <a:ln w="${lineWidth}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${extra}</a:ln>
          </wps:spPr>
          <wps:bodyPr/>
        </wps:wsp>
      </a:graphicData>
    </a:graphic>`;
}

function inlineDrawing({ id, descr, cx, cy, graphic }) {
  return `<w:r><w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="${id}" name="Shape ${id}" descr="${descr}"/>
      <wp:cNvGraphicFramePr/>
      ${graphic}
    </wp:inline>
  </w:drawing></w:r>`;
}

function anchoredDrawing({ id, descr, cx, cy, positionH, positionV, wrap, graphic }) {
  return `<w:r><w:drawing>
    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
      relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      ${positionH}
      ${positionV}
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      ${wrap}
      <wp:docPr id="${id}" name="Shape ${id}" descr="${descr}"/>
      <wp:cNvGraphicFramePr/>
      ${graphic}
    </wp:anchor>
  </w:drawing></w:r>`;
}

const VERTICAL_LINE = inlineDrawing({
  id: 1,
  descr: 'Vertical line',
  cx: 0,
  cy: 914400,
  graphic: shape({
    prst: 'line',
    cx: 0,
    cy: 914400,
    lineWidth: 19050,
    color: 'C00000',
    connector: true,
  }),
});

const HORIZONTAL_LINE = inlineDrawing({
  id: 2,
  descr: 'Horizontal line',
  cx: 1828800,
  cy: 0,
  graphic: shape({
    prst: 'line',
    cx: 1828800,
    cy: 0,
    lineWidth: 12700,
    color: '1F3864',
    connector: true,
  }),
});

const DIAGONAL_LINE = inlineDrawing({
  id: 3,
  descr: 'Diagonal line',
  cx: 914400,
  cy: 457200,
  graphic: shape({
    prst: 'line',
    cx: 914400,
    cy: 457200,
    lineWidth: 25400,
    color: '00B050',
    flips: ' flipV="1"',
    connector: true,
  }),
});

const FULL_WIDTH_RULE = anchoredDrawing({
  id: 4,
  descr: 'Full-width rule',
  cx: COLUMN_WIDTH_EMU,
  cy: 0,
  positionH: '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>',
  positionV: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
  wrap: '<wp:wrapTopAndBottom/>',
  graphic: shape({
    prst: 'rect',
    cx: COLUMN_WIDTH_EMU,
    cy: 0,
    lineWidth: 19050,
    color: '404040',
    fill: '<a:noFill/>',
  }),
});

const MARGIN_CONNECTOR = anchoredDrawing({
  id: 5,
  descr: 'Margin connector',
  cx: 0,
  cy: 1371600,
  positionH: '<wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>',
  positionV: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
  wrap: '<wp:wrapNone/>',
  graphic: shape({
    prst: 'straightConnector1',
    cx: 0,
    cy: 1371600,
    lineWidth: 12700,
    color: '7030A0',
    extra: '<a:prstDash val="dash"/>',
    connector: true,
  }),
});

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    ${paragraph(run('Plain paragraph before the drawings.'))}
    ${paragraph(run('A vertical line '), VERTICAL_LINE, run(' stands between these words.'))}
    ${paragraph(run('A horizontal line '), HORIZONTAL_LINE, run(' sits on this line of text.'))}
    ${paragraph(run('A diagonal line '), DIAGONAL_LINE, run(' rises to the right.'))}
    ${paragraph(FULL_WIDTH_RULE, run('The rule above spans the text column.'))}
    ${paragraph(MARGIN_CONNECTOR, run('A dashed connector runs down the left margin beside this paragraph.'))}
    ${paragraph(run('Plain paragraph after the drawings.'))}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const zip = new JSZip();
const zipOptions = { date: FIXTURE_DATE, createFolders: false };
zip.file('[Content_Types].xml', CONTENT_TYPES_XML, zipOptions);
zip.file('_rels/.rels', RELS_XML, zipOptions);
zip.file('docProps/core.xml', CORE_XML, zipOptions);
zip.file('docProps/app.xml', APP_XML, zipOptions);
zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML, zipOptions);
zip.file('word/styles.xml', STYLES_XML, zipOptions);
zip.file('word/document.xml', DOCUMENT_XML, zipOptions);

const buffer = await zip.generateAsync({ type: 'nodebuffer' });
fs.writeFileSync(OUT, buffer);
console.log(`Created ${OUT}`);
