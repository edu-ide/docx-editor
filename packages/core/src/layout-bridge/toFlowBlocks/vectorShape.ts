/**
 * Vector Shape Pictures
 *
 * Paints a DrawingML line or horizontal rule (`ImageVectorShape`) as an SVG
 * picture, so an image run whose drawing has no picture draws its stroke
 * through the same `<img>` paths a picture takes: inline in a line, as a
 * topAndBottom block, and in the page, cell and header/footer floating layers
 * (upstream #972).
 *
 * The SVG only lives in the layout run's `src`. The document model and
 * ProseMirror keep the drawing's empty `src`, so saving never mistakes it for
 * a new image to embed.
 */

import type { ImageVectorShape } from '../../types/document';
import { sanitizeRgbHex } from '../../docx/drawingUtils';

const EMU_PER_PIXEL = 9525;

/**
 * Stroke color when the outline names no sRGB color. Theme colors are not
 * resolved for drawing outlines; a text box outline paints black the same way.
 */
const DEFAULT_STROKE_COLOR = '#000000';

/**
 * ECMA-376 §20.1.10.48 preset dashes (`a:prstDash`): dash and gap lengths in
 * multiples of the line width. `solid`, and any other value, is unbroken. A
 * Map, so a name read from the file never reaches an object prototype key.
 */
const PRESET_DASH_PATTERNS: ReadonlyMap<string, readonly number[]> = new Map([
  ['dot', [1, 3]],
  ['dash', [4, 3]],
  ['lgDash', [8, 3]],
  ['dashDot', [4, 3, 1, 3]],
  ['lgDashDot', [8, 3, 1, 3]],
  ['lgDashDotDot', [8, 3, 1, 3, 1, 3]],
  ['sysDot', [1, 1]],
  ['sysDash', [3, 1]],
  ['sysDashDot', [3, 1, 1, 1]],
  ['sysDashDotDot', [3, 1, 1, 1, 1, 1]],
]);

/** The picture an image run paints for a line or rule, in layout pixels. */
export interface VectorShapePicture {
  /** `data:image/svg+xml` source that draws the stroke. */
  src: string;
  width: number;
  height: number;
}

interface Stroke {
  color: string;
  width: number;
  dashArray?: readonly number[];
}

/** Three decimals are far below a device pixel and keep the markup short. */
function svgNumber(value: number): string {
  return String(Math.round(value * 1000) / 1000);
}

/** An SVG of one straight line from (x1, y1) to (x2, y2) in a `width` x `height` box. */
function linePicture(
  width: number,
  height: number,
  [x1, y1, x2, y2]: readonly [number, number, number, number],
  stroke: Stroke
): VectorShapePicture {
  const dashArray = stroke.dashArray
    ? ` stroke-dasharray="${stroke.dashArray.map(svgNumber).join(' ')}"`
    : '';
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${svgNumber(width)}" ` +
    `height="${svgNumber(height)}" viewBox="0 0 ${svgNumber(width)} ${svgNumber(height)}">` +
    `<line x1="${svgNumber(x1)}" y1="${svgNumber(y1)}" x2="${svgNumber(x2)}" y2="${svgNumber(y2)}" ` +
    `stroke="${stroke.color}" stroke-width="${svgNumber(stroke.width)}"${dashArray}/></svg>`;
  return { src: `data:image/svg+xml,${encodeURIComponent(svg)}`, width, height };
}

/**
 * The picture for a line or rule whose drawing extent is `extent` (pixels; 0
 * on the axis a straight vertical or horizontal line has no size on).
 *
 * - A line runs corner to corner of its extent, between the other two corners
 *   when it is flipped on one axis. An axis with no extent takes the stroke
 *   width instead, with the line through its middle.
 * - A rule's top and bottom outline edges overlap into one band as tall as the
 *   rectangle plus its outline width, drawn as one line that wide.
 *
 * The stroke is at least one pixel: Word draws a hairline (`w="0"`), and any
 * thinner outline, one device pixel wide.
 */
export function vectorShapePicture(
  shape: ImageVectorShape,
  extent: { width: number; height: number }
): VectorShapePicture {
  const strokeWidth = Math.max(1, (shape.outline.width ?? 0) / EMU_PER_PIXEL);
  const rgb = sanitizeRgbHex(shape.outline.color?.rgb);
  const dashes = shape.outline.style ? PRESET_DASH_PATTERNS.get(shape.outline.style) : undefined;
  const stroke: Stroke = {
    color: rgb ? `#${rgb}` : DEFAULT_STROKE_COLOR,
    width: strokeWidth,
    dashArray: dashes?.map((length) => length * strokeWidth),
  };

  if (shape.shapeType === 'rect') {
    const band = extent.height + strokeWidth;
    return linePicture(extent.width, band, [0, band / 2, extent.width, band / 2], {
      ...stroke,
      width: band,
    });
  }

  const width = extent.width > 0 ? extent.width : strokeWidth;
  const height = extent.height > 0 ? extent.height : strokeWidth;
  const [x1, x2] = extent.width > 0 ? [0, width] : [width / 2, width / 2];
  const [top, bottom] = extent.height > 0 ? [0, height] : [height / 2, height / 2];
  const flipped = Boolean(shape.flipH) !== Boolean(shape.flipV);
  return linePicture(
    width,
    height,
    [x1, flipped ? bottom : top, x2, flipped ? top : bottom],
    stroke
  );
}
