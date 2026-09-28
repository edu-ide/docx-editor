/**
 * Word writes lines, shapes and connectors inside `mc:AlternateContent` (a
 * `wps:wsp` in mc:Choice, a VML copy in mc:Fallback); older files carry VML
 * shapes in `w:pict`, and charts sit in a `w:drawing` without a picture. The
 * fork does not model these. It used to drop them on read, so the editor's
 * default save — which rewrites an edited paragraph from the model — deleted
 * them without a warning. They are now kept as their XML, written back as
 * parsed, and a line among them paints through its preview.
 */
import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { parseDocumentBody } from '../documentParser';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';
import { attemptSelectiveSave } from '../selectiveSave';
import { serializeDocument } from '../serializer/documentSerializer';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import type { ImageRun, ParagraphBlock } from '../../layout-engine/types';
import type { Document, Paragraph, PreservedXmlContent, RunContent } from '../../types/document';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS =
  `xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ` +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  `xmlns:wps="${WPS}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
  'mc:Ignorable="w14 wps"';

/** A Word-saved horizontal line, as Word writes it (`a:` declared on the graphic). */
const WORD_LINE =
  '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="0"/>' +
  '<wp:docPr id="7" name="Straight Connector 7"/>' +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvCnPr/><wps:spPr>` +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm>' +
  '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
  '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
  '</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>' +
  '</mc:Choice><mc:Fallback><w:pict><v:line id="Straight Connector 7" from="0,0" to="72pt,0" ' +
  'strokeweight="1pt"/></w:pict></mc:Fallback></mc:AlternateContent>';

const WORD_TEXT_BOX =
  '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="457200"/>' +
  '<wp:docPr id="9" name="Text Box 9"/>' +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>` +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Stamp box</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
  '<mc:Fallback><w:pict><v:shape id="Text Box 9" type="#_x0000_t202" style="width:144pt;height:36pt">' +
  '<v:textbox><w:txbxContent><w:p><w:r><w:t>Stamp box</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
  '</v:shape></w:pict></mc:Fallback></mc:AlternateContent>';

const VML_LINE = '<w:pict><v:line from="0,0" to="72pt,0" strokeweight="1pt"/></w:pict>';

const VML_WATERMARK =
  '<w:pict><v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" ' +
  'style="position:absolute;width:400pt;height:100pt"><v:textpath string="DRAFT"/></v:shape></w:pict>';

const CHART =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4572000" cy="2743200"/>' +
  '<wp:docPr id="3" name="Chart 3"/>' +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing>';

/** Word ink: `wpi` is named only by Requires and declared on Word's document root. */
const INK =
  '<mc:AlternateContent><mc:Choice Requires="wpi"><w:drawing>' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  '<wp:docPr id="5" name="Ink 5"/>' +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk">` +
  '<w14:contentPart bwMode="auto" r:id="rId7"/></a:graphicData></a:graphic></wp:inline></w:drawing>' +
  '</mc:Choice><mc:Fallback><w:pict><v:shape id="Ink 5" style="width:72pt;height:36pt"/></w:pict>' +
  '</mc:Fallback></mc:AlternateContent>';

/** A shape using a prefix nobody declares: it cannot be written back self-contained. */
const UNDECLARED =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="4" name="Shape 4"/>' +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="${WPS}"><wps:wsp><zz:mystery/><wps:spPr>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing>';

/** A paragraph with text around one run holding `element`. */
function holder(paraId: string, element: string): string {
  return (
    `<w:p w14:paraId="${paraId}" w14:textId="7F000001"><w:r><w:t xml:space="preserve">Before </w:t></w:r>` +
    `<w:r>${element}</w:r><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`
  );
}

const OTHER =
  '<w:p w14:paraId="2A000009" w14:textId="7F000009"><w:r><w:t>Other paragraph</w:t></w:r></w:p>';

function documentXml(body: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>' +
    '</w:sectPr></w:body></w:document>'
  );
}

function modelOf(body: string): Document {
  return { package: { document: parseDocumentBody(documentXml(body)) } };
}

function runContent(doc: Document, index = 0): RunContent[] {
  const paragraph = doc.package.document.content[index] as Paragraph;
  return paragraph.content.flatMap((item) => (item.type === 'run' ? item.content : []));
}

function kept(doc: Document, index = 0): PreservedXmlContent[] {
  return runContent(doc, index).filter((c): c is PreservedXmlContent => c.type === 'preservedXml');
}

async function packageOf(body: string): Promise<ArrayBuffer> {
  const zip = new JSZip();
  const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}">` +
      `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  zip.file('word/document.xml', documentXml(body));
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}"></Relationships>`
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function documentXmlOf(buffer: ArrayBuffer): Promise<string> {
  return (await JSZip.loadAsync(buffer)).file('word/document.xml')!.async('text');
}

function editFirstText(doc: Document, index: number, text: string): void {
  for (const item of runContent(doc, index)) {
    if (item.type === 'text') {
      item.text = text;
      return;
    }
  }
  throw new Error('no text to edit');
}

describe('run content the editor does not model is kept as parsed', () => {
  test('a Word-saved line keeps its wrapper and fallback, and previews as a line', () => {
    const [line] = kept(modelOf(holder('1A000001', WORD_LINE)));
    expect(line?.xml).toBe(WORD_LINE);
    expect(line?.preview?.image.vectorShape?.shapeType).toBe('line');
  });

  test('a VML line and a chart are kept as parsed', () => {
    expect(kept(modelOf(holder('1A000001', VML_LINE)))[0]?.xml).toBe(VML_LINE);
    expect(kept(modelOf(holder('1A000001', CHART)))[0]?.xml).toBe(CHART);
  });

  test('text boxes and watermarks keep their own paths', () => {
    const textBox = modelOf(holder('1A000001', WORD_TEXT_BOX));
    expect(kept(textBox)).toEqual([]);
    expect(runContent(textBox).some((c) => c.type === 'shape')).toBe(true);
    expect(kept(modelOf(holder('1A000001', VML_WATERMARK)))).toEqual([]);
  });

  test("a prefix only Word's root declares is declared on the kept element", () => {
    const [ink] = kept(modelOf(holder('1A000001', INK)));
    expect(ink?.xml).toBe(
      INK.replace(
        '<mc:AlternateContent>',
        '<mc:AlternateContent xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk">'
      )
    );
  });

  test('a shape whose prefix cannot be declared keeps its old reading', () => {
    const content = runContent(modelOf(holder('1A000001', UNDECLARED)));
    expect(content.some((c) => c.type === 'preservedXml')).toBe(false);
    expect(content.some((c) => c.type === 'drawing')).toBe(true);
  });

  test('the editor path through ProseMirror writes it back as parsed', () => {
    const doc = modelOf(holder('1A000001', WORD_LINE) + holder('1A000002', VML_LINE));
    const xml = serializeDocument(fromProseDoc(toProseDoc(doc), doc));
    expect(xml).toContain(WORD_LINE);
    expect(xml).toContain(VML_LINE);
  });

  test('a Word-saved line paints as its stroke', () => {
    const doc = modelOf(holder('1A000001', WORD_LINE));
    const paragraph = toFlowBlocks(toProseDoc(doc)).find(
      (b) => b.kind === 'paragraph'
    ) as ParagraphBlock;
    const run = paragraph.runs.find((r): r is ImageRun => r.kind === 'image');
    expect(run?.src.startsWith('data:image/svg+xml,')).toBe(true);
    expect(run?.width).toBe(96);
  });
});

describe('saving a paragraph that holds a Word shape keeps the shape', () => {
  const tags = (xml: string) =>
    ['mc:AlternateContent', 'wps:wsp', 'w:pict'].map(
      (tag) => (xml.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length
    );

  test('selective save of the edited paragraph writes the shape back', async () => {
    const buffer = await packageOf(holder('1A000001', WORD_LINE) + OTHER);
    const doc = await parseDocx(buffer.slice(0), { preloadFonts: false });
    editFirstText(doc, 0, 'Edited ');
    const saved = await attemptSelectiveSave(doc, buffer.slice(0), {
      changedParaIds: new Set(['1A000001']),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(saved).not.toBeNull();
    const xml = await documentXmlOf(saved!);
    expect(tags(xml)).toEqual([1, 1, 1]);
    expect(xml).toContain(WORD_LINE);
    expect(xml).toContain('Edited ');
  });

  test('a full repack writes the shape back instead of refusing the save', async () => {
    const buffer = await packageOf(holder('1A000001', WORD_LINE) + holder('1A000002', VML_LINE));
    const doc = await parseDocx(buffer.slice(0), { preloadFonts: false });
    const xml = await documentXmlOf(await repackDocx(doc, { updateModifiedDate: false }));
    expect(tags(xml)).toEqual([1, 1, 2]);
    expect(xml).toContain(WORD_LINE);
    expect(xml).toContain(VML_LINE);
  });
});
