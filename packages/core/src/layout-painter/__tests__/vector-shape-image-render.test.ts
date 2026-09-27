/**
 * Lines and rules drawn as `wps:wsp` shapes paint their stroke through every
 * image path of a page (upstream #972): inline in a line, as a topAndBottom
 * block, in the page's floating layer and in the header's floating images.
 * None of them paints an empty picture box or reserves a phantom 100px.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../../docx/documentParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import { measureParagraph } from '../../layout-bridge/measuring/measureParagraph';
import { layoutDocument } from '../../layout-engine';
import type { ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types';
import { renderPage } from '../renderPage';

let restoreCanvas: (() => void) | undefined;

beforeAll(() => {
  GlobalRegistrator.register();
  const prototype = HTMLCanvasElement.prototype;
  const { getContext } = prototype;
  // Deterministic text widths for measuring: 7px per character.
  Object.assign(prototype, {
    getContext(type: string) {
      return type === '2d'
        ? ({
            font: '',
            measureText: (value: string) => ({ width: value.length * 7 }),
          } as unknown as CanvasRenderingContext2D)
        : null;
    },
  });
  restoreCanvas = () => Object.assign(prototype, { getContext });
});

afterAll(() => {
  restoreCanvas?.();
  GlobalRegistrator.unregister();
});

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(' ');
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const SVG_PREFIX = 'data:image/svg+xml,';
const CONTENT_WIDTH = 624;
/** `a:ln w="12700"` (1pt) is 4/3 px. */
const STROKE_1PT = 12700 / 9525;

function shape(
  prst: string,
  cx: number,
  cy: number,
  stroke: { widthEmu: number; rgb: string },
  fill = ''
): string {
  return (
    `<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvCnPr/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fill}` +
    `<a:ln w="${stroke.widthEmu}"><a:solidFill><a:srgbClr val="${stroke.rgb}"/></a:solidFill></a:ln>` +
    '</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>'
  );
}

function inlineDrawing(id: number, cx: number, cy: number, graphic: string): string {
  return (
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Line ${id}"/>${graphic}` +
    '</wp:inline></w:drawing></w:r>'
  );
}

function anchoredDrawing(
  id: number,
  cx: number,
  cy: number,
  placement: { h: string; v: string; wrap: string },
  graphic: string
): string {
  return (
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
    'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    `<wp:simplePos x="0" y="0"/>${placement.h}${placement.v}` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${placement.wrap}` +
    `<wp:docPr id="${id}" name="Shape ${id}"/>${graphic}</wp:anchor></w:drawing></w:r>`
  );
}

function words(value: string): string {
  return `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
}

function flowOf(paragraphsXml: string): {
  blocks: ParagraphBlock[];
  measures: ParagraphMeasure[];
} {
  const body = parseDocumentBody(
    `<w:document ${NS}><w:body>${paragraphsXml}</w:body></w:document>`
  );
  const blocks = toFlowBlocks(toProseDoc({ package: { document: body } })).filter(
    (block): block is ParagraphBlock => block.kind === 'paragraph'
  );
  return { blocks, measures: blocks.map((block) => measureParagraph(block, CONTENT_WIDTH)) };
}

const NAVY = { widthEmu: 12700, rgb: '1F3864' };
const BODY = [
  `<w:p>${words('Before ')}${inlineDrawing(1, 0, 914400, shape('line', 0, 914400, NAVY))}${words(' after')}</w:p>`,
  `<w:p>${words('Text ')}${inlineDrawing(2, 914400, 0, shape('line', 914400, 0, NAVY))}${words(' more')}</w:p>`,
  `<w:p>${words('Plain paragraph')}</w:p>`,
  `<w:p>${anchoredDrawing(
    3,
    5943600,
    0,
    {
      h: '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>',
      v: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
      wrap: '<wp:wrapTopAndBottom/>',
    },
    shape('rect', 5943600, 0, { widthEmu: 19050, rgb: '404040' }, '<a:noFill/>')
  )}${words('Below the rule')}</w:p>`,
  `<w:p>${anchoredDrawing(
    4,
    0,
    1828800,
    {
      h: '<wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>',
      v: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
      wrap: '<wp:wrapNone/>',
    },
    shape('line', 0, 1828800, { widthEmu: 12700, rgb: 'C00000' })
  )}${words('Beside a margin line')}</w:p>`,
].join('');
// A letterhead fold mark: a 0.25pt horizontal line anchored to the page in the header.
const HEADER = `<w:p>${anchoredDrawing(
  5,
  107950,
  0,
  {
    h: '<wp:positionH relativeFrom="page"><wp:posOffset>180340</wp:posOffset></wp:positionH>',
    v: '<wp:positionV relativeFrom="page"><wp:posOffset>3780790</wp:posOffset></wp:positionV>',
    wrap: '<wp:wrapNone/>',
  },
  shape('line', 107950, 0, { widthEmu: 3175, rgb: '00445E' })
)}</w:p>`;

function renderFixturePage(): HTMLElement {
  const body = flowOf(BODY);
  const header = flowOf(HEADER);
  const layout = layoutDocument(body.blocks, body.measures, {
    pageSize: { w: 816, h: 1056 },
    margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
  });
  const blockLookup = new Map(
    body.blocks.map((block, index) => [String(block.id), { block, measure: body.measures[index]! }])
  );
  return renderPage(
    layout.pages[0]!,
    { pageNumber: 1, totalPages: 1, section: 'body' },
    {
      document,
      blockLookup,
      headerContent: {
        blocks: header.blocks,
        measures: header.measures,
        height: header.measures.reduce((sum, measure) => sum + measure.totalHeight, 0),
      },
    }
  );
}

/** The `<line>` of the SVG an `<img>` paints. */
function paintedLine(img: HTMLImageElement): Element {
  const src = img.getAttribute('src') ?? '';
  expect(src.startsWith(SVG_PREFIX)).toBe(true);
  const markup = decodeURIComponent(src.slice(SVG_PREFIX.length));
  const line = new DOMParser().parseFromString(markup, 'image/svg+xml').querySelector('line');
  if (!line) throw new Error(`no <line> in ${markup}`);
  return line;
}

function px(value: string): number {
  return parseFloat(value);
}

describe('vector shape images on a painted page', () => {
  test('an inline vertical line paints its stroke at the stroke width', () => {
    const img = renderFixturePage().querySelector<HTMLImageElement>('img.layout-run-image');
    expect(img).toBeTruthy();
    expect(px(img!.style.width)).toBeCloseTo(STROKE_1PT, 6);
    const line = paintedLine(img!);
    expect(Number(line.getAttribute('stroke-width'))).toBeCloseTo(STROKE_1PT, 3);
    expect(line.getAttribute('stroke')).toBe('#1F3864');
    expect(Number(line.getAttribute('y2'))).toBe(96);
  });

  test('an inline horizontal line does not push its line down', () => {
    const { measures } = flowOf(BODY);
    const [verticalLine, horizontalLine, plain] = measures.map((m) => m.lines[0]!.lineHeight);
    expect(horizontalLine).toBeCloseTo(plain, 6);
    expect(verticalLine).toBeGreaterThanOrEqual(96);
    expect(verticalLine).toBeLessThan(96 + 50);

    const lineElements = renderFixturePage().querySelectorAll<HTMLElement>('.layout-line');
    expect(px(lineElements[1].style.height)).toBeCloseTo(plain, 6);
  });

  test('a full-width topAndBottom rule paints as a 2px band', () => {
    const img = renderFixturePage().querySelector<HTMLImageElement>('.layout-block-image img');
    expect(img).toBeTruthy();
    expect(img!.width).toBe(CONTENT_WIDTH);
    expect(img!.height).toBe(2);
    const line = paintedLine(img!);
    expect(line.getAttribute('stroke')).toBe('#404040');
    expect(Number(line.getAttribute('x2'))).toBe(CONTENT_WIDTH);
  });

  test('an anchored wrapNone line paints in the page floating layer', () => {
    const img = renderFixturePage().querySelector<HTMLImageElement>(
      '.layout-page-floating-image img'
    );
    expect(img).toBeTruthy();
    expect(px(img!.style.width)).toBeCloseTo(STROKE_1PT, 6);
    expect(px(img!.style.height)).toBe(192);
    expect(paintedLine(img!).getAttribute('stroke')).toBe('#C00000');
  });

  test('a header fold mark paints a 1px line instead of an 11x100 box', () => {
    const img = renderFixturePage().querySelector<HTMLImageElement>('.layout-page-header img');
    expect(img).toBeTruthy();
    expect(px(img!.style.width)).toBe(11);
    // 0.25pt is a third of a pixel; a hairline still paints one pixel.
    expect(px(img!.style.height)).toBe(1);
    expect(paintedLine(img!).getAttribute('stroke')).toBe('#00445E');
  });

  test('no painted image is an empty picture box', () => {
    const images = [...renderFixturePage().querySelectorAll<HTMLImageElement>('img')];
    expect(images).toHaveLength(5);
    for (const img of images) {
      expect(img.getAttribute('src')).not.toBe('');
      expect(img.width).not.toBe(100);
      expect(img.height).not.toBe(100);
    }
  });
});
