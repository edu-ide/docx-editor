/**
 * Painting `wps:wsp` lines and rules (upstream #972) is render-only: saving
 * must write these drawings exactly as before. The fork saves a drawing
 * without a picture as a picture frame (the `wps:wsp` is not written back);
 * the vector shape must not add a stroke, flips or a new size to it, on the
 * headless path (model -> XML) or the editor path (model -> PM -> model -> XML).
 */
import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
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

/** The picture frame the fork writes for a drawing it read without a picture. */
function savedFrame(
  id: number,
  cx: number,
  cy: number,
  anchorWrap?: '<wp:wrapTopAndBottom/>' | '<wp:wrapNone/>'
): string {
  const extent = `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`;
  const graphic =
    `<wp:docPr id="${id}" name="Picture ${id}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    '<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>';
  if (!anchorWrap) {
    return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${extent}${graphic}</wp:inline></w:drawing>`;
  }
  return (
    '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
    'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `${extent}${anchorWrap}${graphic}</wp:anchor></w:drawing>`
  );
}

function parse(): Document {
  const body = parseDocumentBody(
    `<w:document ${NS}><w:body>${DRAWINGS.map(
      (drawing) => `<w:p><w:r>${drawing}</w:r><w:r><w:t>x</w:t></w:r></w:p>`
    ).join('')}</w:body></w:document>`
  );
  return { package: { document: body } };
}

function savedDrawings(xml: string): string[] {
  return xml.match(/<w:drawing>.*?<\/w:drawing>/g) ?? [];
}

function images(doc: Document): Image[] {
  return doc.package.document.content.map((block) => {
    const run = (block as Paragraph).content[0];
    if (run?.type !== 'run' || run.content[0]?.type !== 'drawing') {
      throw new Error('expected a drawing run');
    }
    return run.content[0].image;
  });
}

describe('saving line and rule drawings is unchanged by vector painting', () => {
  test('the headless model -> XML path writes the same picture frames', () => {
    expect(savedDrawings(serializeDocumentBody(parse().package.document))).toEqual([
      savedFrame(11, 0, 914400),
      savedFrame(12, 914400, 0),
      savedFrame(13, 914400, 457200),
      savedFrame(14, 5943600, 0, '<wp:wrapTopAndBottom/>'),
      savedFrame(15, 0, 1828800, '<wp:wrapNone/>'),
    ]);
  });

  test('the editor path through ProseMirror writes the same picture frames', () => {
    const doc = parse();
    // fromProseDoc drops the docPr id, so the serializer numbers the frames itself.
    expect(savedDrawings(serializeDocument(fromProseDoc(toProseDoc(doc), doc)))).toEqual([
      savedFrame(100000, 0, 914400),
      savedFrame(100001, 914400, 0),
      savedFrame(100002, 914400, 457200),
      savedFrame(100003, 5943600, 0, '<wp:wrapTopAndBottom/>'),
      savedFrame(100004, 0, 1828800, '<wp:wrapNone/>'),
    ]);
  });

  test('the vector shape survives the ProseMirror round-trip on the model', () => {
    const doc = parse();
    const before = images(doc).map((image) => image.vectorShape);
    const after = images(fromProseDoc(toProseDoc(doc), doc)).map((image) => image.vectorShape);
    expect(after).toEqual(before);
    expect(before.map((shape) => shape?.shapeType)).toEqual([
      'line',
      'line',
      'straightConnector1',
      'rect',
      'line',
    ]);
  });
});
