/**
 * Line and rule drawings are kept as parsed, so saving writes them back as they
 * were: the fork used to save a drawing without a picture as an empty picture
 * frame, losing the `wps:wsp`. The headless path (model -> XML), the editor path
 * (model -> PM -> model -> XML) and a full repack must all write the drawing
 * itself; the painted line (the vector shape of its preview) survives the PM
 * round-trip.
 */
import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { parseDocumentBody } from '../documentParser';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';
import { serializeDocument, serializeDocumentBody } from '../serializer/documentSerializer';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import type { Document, Image, Paragraph } from '../../types/document';

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(' ');
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const STROKE = '<a:ln w="12700"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill></a:ln>';
const FIXTURE_PATH = join(
  __dirname,
  '..',
  '..',
  '..',
  '..',
  '..',
  'e2e',
  'fixtures',
  'vector-lines-and-rules.docx'
);

function shape(
  prst: string,
  cx: number,
  cy: number,
  spPr: string,
  xfrm = '',
  nv = '<wps:cNvCnPr/>'
) {
  return (
    `<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp>${nv}<wps:spPr>` +
    `<a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${spPr}</wps:spPr>` +
    '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>'
  );
}

function inline(id: number, cx: number, cy: number, graphic: string) {
  return (
    '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Line ${id}"/>${graphic}` +
    '</wp:inline></w:drawing>'
  );
}

function anchor(id: number, cx: number, cy: number, wrap: string, graphic: string) {
  return (
    '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
    'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}` +
    `<wp:docPr id="${id}" name="Shape ${id}"/>${graphic}</wp:anchor></w:drawing>`
  );
}

const DRAWINGS = [
  inline(11, 0, 914400, shape('line', 0, 914400, STROKE)),
  inline(12, 914400, 0, shape('line', 914400, 0, STROKE)),
  inline(13, 914400, 457200, shape('straightConnector1', 914400, 457200, STROKE, ' flipH="1"')),
  anchor(
    14,
    5943600,
    0,
    '<wp:wrapTopAndBottom/>',
    shape('rect', 5943600, 0, `<a:noFill/>${STROKE}`, '', '<wps:cNvSpPr/>')
  ),
  anchor(15, 0, 1828800, '<wp:wrapNone/>', shape('line', 0, 1828800, STROKE, ' flipV="1"')),
];

function parse(): Document {
  const body = parseDocumentBody(
    `<w:document ${NS}><w:body>${DRAWINGS.map(
      (drawing) => `<w:p><w:r>${drawing}</w:r><w:r><w:t>x</w:t></w:r></w:p>`
    ).join('')}</w:body></w:document>`
  );
  return { package: { document: body } };
}

function savedDrawings(xml: string): string[] {
  return xml.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/g) ?? [];
}

/** The painted image of each paragraph's drawing, kept as parsed with a preview. */
function previews(doc: Document): Image[] {
  return doc.package.document.content.map((block) => {
    const run = (block as Paragraph).content[0];
    const kept = run?.type === 'run' ? run.content[0] : undefined;
    if (kept?.type !== 'preservedXml' || !kept.preview) {
      throw new Error('expected a drawing kept as parsed, with a preview');
    }
    return kept.preview.image;
  });
}

/**
 * The drawings as a save writes them: as parsed, plus the `a:` declaration the
 * test document gives on its root, which a repacked root does not carry.
 */
const WRITTEN = DRAWINGS.map((drawing) =>
  drawing.replace('<w:drawing>', `<w:drawing xmlns:a="${A_NS}">`)
);

describe('saving writes line and rule drawings back as parsed', () => {
  test('the headless model -> XML path writes each drawing as parsed', () => {
    expect(savedDrawings(serializeDocumentBody(parse().package.document))).toEqual(WRITTEN);
  });

  test('the editor path through ProseMirror writes each drawing as parsed', () => {
    const doc = parse();
    expect(savedDrawings(serializeDocument(fromProseDoc(toProseDoc(doc), doc)))).toEqual(WRITTEN);
  });

  test('the painted line survives the ProseMirror round-trip on the model', () => {
    const doc = parse();
    const before = previews(doc).map((image) => image.vectorShape);
    const after = previews(fromProseDoc(toProseDoc(doc), doc)).map((image) => image.vectorShape);
    expect(after).toEqual(before);
    expect(before.map((vectorShape) => vectorShape?.shapeType)).toEqual([
      'line',
      'line',
      'straightConnector1',
      'rect',
      'line',
    ]);
  });

  test('a full repack of the e2e fixture writes its drawings as parsed', async () => {
    const buffer = readFileSync(FIXTURE_PATH);
    const original = await (await JSZip.loadAsync(buffer)).file('word/document.xml')!.async('text');
    const doc = await parseDocx(
      buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
      { preloadFonts: false }
    );
    const zip = await JSZip.loadAsync(await repackDocx(doc, { updateModifiedDate: false }));
    const xml = await zip.file('word/document.xml')!.async('text');
    // The fixture breaks some start tags across lines; XML does not keep the
    // whitespace between attributes, so tags compare with it collapsed.
    const tagSpaces = (drawing: string) =>
      drawing.replace(/<[^>]+>/g, (tag) => tag.replace(/\s+/g, ' '));
    expect(savedDrawings(original)).toHaveLength(5);
    expect(savedDrawings(xml)).toEqual(savedDrawings(original).map(tagSpaces));
  });
});
