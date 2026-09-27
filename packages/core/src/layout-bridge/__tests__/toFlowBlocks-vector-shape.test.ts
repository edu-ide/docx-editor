/**
 * A `wps:wsp` line or rule stays a PM `image` node and an `image` run, but the
 * run paints its stroke instead of an empty picture (upstream #972): the source
 * is an SVG of the line, and an axis with no extent takes the stroke width
 * instead of a phantom 100px.
 */
import { describe, expect, test } from 'bun:test';
import type { Node as PMNode } from 'prosemirror-model';
import { parseDocumentBody } from '../../docx/documentParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../toFlowBlocks';
import type { ImageRun, ParagraphBlock } from '../../layout-engine/types';
import type { Document } from '../../types/document';

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(' ');
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const SVG_PREFIX = 'data:image/svg+xml,';
/** `a:ln w="12700"` (1pt) is 4/3 px. */
const STROKE_1PT = 12700 / 9525;

function stroke(widthEmu: number, rgb: string, extra = ''): string {
  return `<a:ln w="${widthEmu}"><a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>${extra}</a:ln>`;
}

function shape(prst: string, cx: number, cy: number, spPr: string, xfrm = ''): string {
  return (
    `<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvCnPr/><wps:spPr>` +
    `<a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${spPr}</wps:spPr>` +
    '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>'
  );
}

function inlineRun(cx: number, cy: number, graphic: string): string {
  return (
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="6" name="Line 6"/>${graphic}` +
    '</wp:inline></w:drawing></w:r>'
  );
}

function topAndBottomRun(cx: number, cy: number, graphic: string): string {
  return (
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
    'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapTopAndBottom/><wp:docPr id="7" name="Rule 7"/>${graphic}</wp:anchor></w:drawing></w:r>`
  );
}

function convert(runXml: string): { pmImage: PMNode; run: ImageRun } {
  const body = parseDocumentBody(
    `<w:document ${NS}><w:body><w:p>${runXml}<w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>`
  );
  const document: Document = { package: { document: body } };
  const pmDoc = toProseDoc(document);
  const pmImages: PMNode[] = [];
  pmDoc.descendants((node) => {
    if (node.type.name === 'image') pmImages.push(node);
    return true;
  });
  const pmImage = pmImages[0];
  if (!pmImage) throw new Error('expected a PM image node');
  const paragraph = toFlowBlocks(pmDoc).find((b) => b.kind === 'paragraph') as ParagraphBlock;
  const run = paragraph.runs.find((r) => r.kind === 'image');
  if (!run || run.kind !== 'image') throw new Error('expected an image run');
  return { pmImage, run };
}

/** The SVG markup an image run's `data:` source carries. */
function svgOf(run: ImageRun): string {
  expect(run.src.startsWith(SVG_PREFIX)).toBe(true);
  return decodeURIComponent(run.src.slice(SVG_PREFIX.length));
}

/** Numeric attribute of the first `<line>` in an SVG string. */
function lineAttr(svg: string, name: string): number {
  const match = new RegExp(`<line[^>]*\\s${name}="([^"]+)"`).exec(svg);
  if (!match) throw new Error(`no ${name} on the line in ${svg}`);
  return Number(match[1]);
}

describe('line drawings through PM and flow', () => {
  test('a vertical line stays an image node and run, stroke-wide instead of 100px', () => {
    const { pmImage, run } = convert(
      inlineRun(0, 914400, shape('line', 0, 914400, stroke(12700, '1F3864')))
    );

    // PM keeps the drawing's own size (no width for a zero-width line) and the shape.
    expect(pmImage.attrs.width).toBeNull();
    expect(pmImage.attrs.height).toBe(96);
    expect(pmImage.attrs.src).toBe('');
    expect(pmImage.attrs.vectorShape).toEqual({
      shapeType: 'line',
      outline: { width: 12700, color: { rgb: '1F3864' } },
    });

    expect(run.kind).toBe('image');
    expect(run.width).toBeCloseTo(STROKE_1PT, 6);
    expect(run.height).toBe(96);

    const svg = svgOf(run);
    expect(svg).toContain('stroke="#1F3864"');
    expect(lineAttr(svg, 'stroke-width')).toBeCloseTo(STROKE_1PT, 3);
    // Centred on the zero axis, running the full length.
    expect(lineAttr(svg, 'x1')).toBeCloseTo(STROKE_1PT / 2, 3);
    expect(lineAttr(svg, 'x2')).toBeCloseTo(STROKE_1PT / 2, 3);
    expect(lineAttr(svg, 'y1')).toBe(0);
    expect(lineAttr(svg, 'y2')).toBe(96);
  });

  test('a horizontal line takes its stroke width as its height', () => {
    const { run } = convert(
      inlineRun(914400, 0, shape('line', 914400, 0, stroke(12700, '000000')))
    );
    expect(run.width).toBe(96);
    expect(run.height).toBeCloseTo(STROKE_1PT, 6);
    const svg = svgOf(run);
    expect(lineAttr(svg, 'y1')).toBeCloseTo(STROKE_1PT / 2, 3);
    expect(lineAttr(svg, 'y2')).toBeCloseTo(STROKE_1PT / 2, 3);
    expect(lineAttr(svg, 'x2')).toBe(96);
  });

  test('a diagonal line keeps its extent and flips pick its corners', () => {
    const plain = convert(
      inlineRun(914400, 457200, shape('line', 914400, 457200, stroke(12700, '000000')))
    );
    expect([plain.run.width, plain.run.height]).toEqual([96, 48]);
    const down = svgOf(plain.run);
    expect([
      lineAttr(down, 'x1'),
      lineAttr(down, 'y1'),
      lineAttr(down, 'x2'),
      lineAttr(down, 'y2'),
    ]).toEqual([0, 0, 96, 48]);

    for (const flips of [' flipH="1"', ' flipV="1"']) {
      const flipped = svgOf(
        convert(
          inlineRun(914400, 457200, shape('line', 914400, 457200, stroke(12700, '000000'), flips))
        ).run
      );
      expect([
        lineAttr(flipped, 'x1'),
        lineAttr(flipped, 'y1'),
        lineAttr(flipped, 'x2'),
        lineAttr(flipped, 'y2'),
      ]).toEqual([0, 48, 96, 0]);
    }

    const both = svgOf(
      convert(
        inlineRun(
          914400,
          457200,
          shape('line', 914400, 457200, stroke(12700, '000000'), ' flipH="1" flipV="1"')
        )
      ).run
    );
    expect([
      lineAttr(both, 'x1'),
      lineAttr(both, 'y1'),
      lineAttr(both, 'x2'),
      lineAttr(both, 'y2'),
    ]).toEqual([0, 0, 96, 48]);
  });

  test('a preset dash becomes a dash array in stroke widths', () => {
    const { run } = convert(
      inlineRun(
        914400,
        0,
        shape('straightConnector1', 914400, 0, stroke(19050, '000000', '<a:prstDash val="dash"/>'))
      )
    );
    // `dash` is 4 on, 3 off, in multiples of the 2px line width.
    expect(svgOf(run)).toContain('stroke-dasharray="8 6"');
  });

  test('a dash name that is no preset draws an unbroken line', () => {
    const { run } = convert(
      inlineRun(
        914400,
        0,
        shape('line', 914400, 0, stroke(19050, '000000', '<a:prstDash val="__proto__"/>'))
      )
    );
    expect(svgOf(run)).not.toContain('stroke-dasharray');
  });

  test('a stroke with no color of its own paints black, and a hairline paints 1px', () => {
    const { run } = convert(inlineRun(914400, 0, shape('line', 914400, 0, '<a:ln w="0"/>')));
    const svg = svgOf(run);
    expect(svg).toContain('stroke="#000000"');
    expect(lineAttr(svg, 'stroke-width')).toBe(1);
    expect(run.height).toBe(1);
  });

  test('a full-width rule is one band across the column, not a 100px box', () => {
    const rule = shape('rect', 5943600, 0, `<a:noFill/>${stroke(19050, '404040')}`);
    const { pmImage, run } = convert(topAndBottomRun(5943600, 0, rule));
    expect(pmImage.attrs.vectorShape.shapeType).toBe('rect');
    expect(run.wrapType).toBe('topAndBottom');
    expect(run.width).toBe(624);
    expect(run.height).toBe(2);
    const svg = svgOf(run);
    expect(lineAttr(svg, 'x1')).toBe(0);
    expect(lineAttr(svg, 'x2')).toBe(624);
    expect(lineAttr(svg, 'y1')).toBe(1);
    expect(lineAttr(svg, 'stroke-width')).toBe(2);
  });

  test('a rule with height draws both outline edges as one band', () => {
    // 19050 EMU tall with a 19050 EMU outline: edges at 0 and 2px, each 2px wide.
    const rule = shape('rect', 5943600, 19050, `<a:noFill/>${stroke(19050, '404040')}`);
    const { run } = convert(inlineRun(5943600, 19050, rule));
    expect(run.height).toBe(4);
    const svg = svgOf(run);
    expect(lineAttr(svg, 'y1')).toBe(2);
    expect(lineAttr(svg, 'stroke-width')).toBe(4);
  });

  test('a picture run is unchanged', () => {
    const picture =
      '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId9"/>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const { pmImage, run } = convert(picture);
    expect(pmImage.attrs.vectorShape).toBeNull();
    expect(run.src).toBe('');
    expect([run.width, run.height]).toEqual([96, 48]);
  });
});
