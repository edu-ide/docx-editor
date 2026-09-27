/**
 * Straight lines and horizontal rules drawn as `wps:wsp` shapes (upstream
 * #972). Such a drawing has no picture: the parser reads its preset geometry,
 * stroke and flips into `Image.vectorShape` so layout can paint it as a line.
 * Every other drawing keeps `vectorShape` unset.
 */
import { describe, expect, test } from 'bun:test';
import { parseDrawing } from '../imageParser';
import { parseXml } from '../xmlParser';
import type { XmlElement } from '../xmlParser';
import type { Image } from '../../types/document';

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
].join(' ');
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const BLACK_LINE = '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';

function parseDrawingFromXml(innerXml: string): Image | null {
  const doc = parseXml(`<w:drawing ${NS}>${innerXml}</w:drawing>`);
  return parseDrawing((doc.elements as XmlElement[])[0], undefined, undefined);
}

/** A `wps:wsp` shape as Insert > Shapes writes it, with `spPr` children after the geometry. */
function shape(
  prst: string,
  cx: number,
  cy: number,
  options: { spPr?: string; xfrm?: string; nv?: string; style?: string } = {}
): string {
  return (
    `<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp>${options.nv ?? '<wps:cNvCnPr/>'}` +
    `<wps:spPr><a:xfrm${options.xfrm ?? ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${options.spPr ?? BLACK_LINE}</wps:spPr>` +
    `${options.style ?? ''}<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>`
  );
}

function inline(cx: number, cy: number, graphic: string): Image | null {
  return parseDrawingFromXml(
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="6" name="Line 6"/>${graphic}</wp:inline>`
  );
}

function anchored(cx: number, cy: number, wrap: string, graphic: string): Image | null {
  return parseDrawingFromXml(
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
      'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="page"><wp:posOffset>180340</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="page"><wp:posOffset>3780790</wp:posOffset></wp:positionV>' +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}` +
      `<wp:docPr id="13" name="Shape 13"/>${graphic}</wp:anchor>`
  );
}

describe('wps line drawings parse to a vector shape', () => {
  test('a zero-width vertical line keeps its extent and reads its stroke', () => {
    const image = inline(0, 914400, shape('line', 0, 914400));
    expect(image?.size).toEqual({ width: 0, height: 914400 });
    expect(image?.src).toBeUndefined();
    expect(image?.vectorShape).toEqual({
      shapeType: 'line',
      outline: { width: 9525, color: { rgb: '000000' } },
    });
  });

  test('a zero-height horizontal line and a diagonal line parse the same way', () => {
    expect(inline(914400, 0, shape('line', 914400, 0))?.vectorShape?.shapeType).toBe('line');
    expect(inline(914400, 457200, shape('line', 914400, 457200))?.vectorShape).toEqual({
      shapeType: 'line',
      outline: { width: 9525, color: { rgb: '000000' } },
    });
  });

  test('flips pick the diagonal a line runs along', () => {
    const flipH = inline(914400, 457200, shape('line', 914400, 457200, { xfrm: ' flipH="1"' }));
    expect(flipH?.vectorShape?.flipH).toBe(true);
    expect(flipH?.vectorShape?.flipV).toBeUndefined();
    const both = inline(
      914400,
      457200,
      shape('line', 914400, 457200, { xfrm: ' flipH="1" flipV="1"' })
    );
    expect(both?.vectorShape).toMatchObject({ flipH: true, flipV: true });
    // The flips stay on the vector shape; the picture transform is not set.
    expect(both?.transform).toBeUndefined();
  });

  test('a straight connector (wps:cNvCnPr) keeps its preset', () => {
    const connector = inline(
      914400,
      0,
      shape('straightConnector1', 914400, 0, {
        spPr:
          '<a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>' +
          '<a:tailEnd type="triangle"/></a:ln>',
      })
    );
    expect(connector?.vectorShape?.shapeType).toBe('straightConnector1');
    expect(connector?.vectorShape?.outline).toMatchObject({
      width: 19050,
      color: { rgb: 'C00000' },
    });
  });

  test('a preset dash is kept on the outline', () => {
    const dashed = inline(
      914400,
      0,
      shape('line', 914400, 0, {
        spPr:
          '<a:ln w="12700"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>' +
          '<a:prstDash val="sysDash"/></a:ln>',
      })
    );
    expect(dashed?.vectorShape?.outline.style).toBe('sysDash');
  });

  test('a line stroked only through its style keeps the style line', () => {
    const styled = inline(
      914400,
      0,
      shape('line', 914400, 0, {
        spPr: '',
        style:
          '<wps:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>' +
          '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></wps:style>',
      })
    );
    expect(styled?.vectorShape).toEqual({
      shapeType: 'line',
      outline: { width: 9525, color: { themeColor: 'accent1' } },
    });
  });

  test('an anchored line keeps its anchor and gets the vector shape', () => {
    const image = anchored(107950, 0, '<wp:wrapNone/>', shape('line', 107950, 0));
    expect(image?.wrap.type).toBe('inFront');
    expect(image?.position?.horizontal).toEqual({ relativeTo: 'page', posOffset: 180340 });
    expect(image?.vectorShape?.shapeType).toBe('line');
  });
});

describe('an outlined rectangle acting as a horizontal rule parses to a vector shape', () => {
  const RULE_LINE =
    '<a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="404040"/></a:solidFill></a:ln>';

  test('a full-width zero-height rectangle with an outline and no fill is a rule', () => {
    const rule = anchored(
      5943600,
      0,
      '<wp:wrapTopAndBottom/>',
      shape('rect', 5943600, 0, { nv: '<wps:cNvSpPr/>', spPr: RULE_LINE })
    );
    expect(rule?.vectorShape).toEqual({
      shapeType: 'rect',
      outline: { width: 19050, color: { rgb: '404040' } },
    });
  });

  test('a rectangle no taller than its outline still reads as one rule', () => {
    const rule = inline(
      5943600,
      19050,
      shape('rect', 5943600, 19050, { nv: '<wps:cNvSpPr/>', spPr: RULE_LINE })
    );
    expect(rule?.vectorShape?.shapeType).toBe('rect');
  });

  test('a filled, a hollow-box or an unstroked rectangle is not a rule', () => {
    const filled = inline(
      5943600,
      0,
      shape('rect', 5943600, 0, {
        nv: '<wps:cNvSpPr/>',
        spPr: `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>${BLACK_LINE}`,
      })
    );
    expect(filled?.vectorShape).toBeUndefined();

    const styleFilled = inline(
      5943600,
      0,
      shape('rect', 5943600, 0, {
        nv: '<wps:cNvSpPr/>',
        style: '<wps:style><a:fillRef idx="1"><a:srgbClr val="4472C4"/></a:fillRef></wps:style>',
      })
    );
    expect(styleFilled?.vectorShape).toBeUndefined();

    const box = inline(
      1828800,
      914400,
      shape('rect', 1828800, 914400, { nv: '<wps:cNvSpPr/>', spPr: RULE_LINE })
    );
    expect(box?.vectorShape).toBeUndefined();

    const unstroked = inline(
      5943600,
      0,
      shape('rect', 5943600, 0, {
        nv: '<wps:cNvSpPr/>',
        spPr: '<a:noFill/><a:ln><a:noFill/></a:ln>',
      })
    );
    expect(unstroked?.vectorShape).toBeUndefined();
  });
});

describe('other drawings keep vectorShape unset', () => {
  test('a picture', () => {
    const picture = parseDrawingFromXml(
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="img"/><pic:cNvPicPr/></pic:nvPicPr>' +
        '<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
        '</a:graphicData></a:graphic></wp:inline>'
    );
    expect(picture?.rId).toBe('rId1');
    expect(picture?.vectorShape).toBeUndefined();
  });

  test('other presets, unstroked, rotated and zero-size lines', () => {
    expect(inline(914400, 914400, shape('ellipse', 914400, 914400))?.vectorShape).toBeUndefined();
    expect(
      inline(914400, 457200, shape('rightArrow', 914400, 457200))?.vectorShape
    ).toBeUndefined();
    const unstroked = inline(
      914400,
      0,
      shape('line', 914400, 0, { spPr: '<a:ln><a:noFill/></a:ln>' })
    );
    expect(unstroked?.vectorShape).toBeUndefined();
    const rotated = inline(914400, 0, shape('line', 914400, 0, { xfrm: ' rot="5400000"' }));
    expect(rotated?.vectorShape).toBeUndefined();
    const point = inline(0, 0, shape('line', 0, 0));
    expect(point).not.toBeNull();
    expect(point?.vectorShape).toBeUndefined();
  });

  test('a negative extent or an outline width outside ST_LineWidth', () => {
    expect(inline(-5, 914400, shape('line', -5, 914400))?.vectorShape).toBeUndefined();
    for (const width of ['-9525', 'wide', '20116801']) {
      const line = inline(
        914400,
        0,
        shape('line', 914400, 0, {
          spPr: `<a:ln w="${width}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`,
        })
      );
      expect(line?.vectorShape).toBeUndefined();
    }
  });

  test('a group of lines', () => {
    const group = inline(
      914400,
      914400,
      `<a:graphic><a:graphicData uri="${WPG_URI}"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>` +
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>' +
        '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></wpg:grpSpPr>' +
        '<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="914400"/></a:xfrm>' +
        `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${BLACK_LINE}</wps:spPr><wps:bodyPr/></wps:wsp>` +
        '</wpg:wgp></a:graphicData></a:graphic>'
    );
    expect(group?.vectorShape).toBeUndefined();
  });

  test('a text box is still left to the text box parser', () => {
    const textBox = parseDrawingFromXml(
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="5943600" cy="0"/><wp:docPr id="2" name="Text Box 2"/>' +
        `<a:graphic><a:graphicData uri="${WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>` +
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="5943600" cy="0"/></a:xfrm>' +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${BLACK_LINE}</wps:spPr>` +
        '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Label</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
        '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline>'
    );
    expect(textBox).toBeNull();
  });
});
