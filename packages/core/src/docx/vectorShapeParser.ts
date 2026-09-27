/**
 * Vector Shape Parser - lines and rules drawn as `wps:wsp` shapes
 *
 * A `w:drawing` whose graphic is a `wps:wsp` shape has no picture. When that
 * shape is a straight line or connector, or an unfilled outlined rectangle
 * flat enough to read as a horizontal rule, it paints as a stroke instead
 * (upstream #972). This reads that stroke for `Image.vectorShape`; every
 * other drawing keeps the picture path.
 *
 * OOXML Structure:
 * wp:inline or wp:anchor
 *   └── a:graphic
 *       └── a:graphicData
 *           └── wps:wsp
 *               ├── wps:spPr
 *               │   ├── a:xfrm (flipH, flipV, rot)
 *               │   ├── a:prstGeom (prst: line, straightConnector1, rect)
 *               │   ├── a:noFill / a:solidFill (a rule has no fill)
 *               │   └── a:ln (stroke width, color, dash)
 *               └── wps:style (a:lnRef / a:fillRef fallbacks)
 */

import type { ImageSize, ImageVectorShape } from '../types/document';
import { findByFullName, getAttribute, type XmlElement } from './xmlParser';
import { hasFill, hasTextContent, parseShape } from './shapeParser';

/** ECMA-376 `ST_LineWidth` upper bound (1584pt). */
const MAX_LINE_WIDTH_EMU = 20116800;

/**
 * Read the stroke of a `wps:wsp` drawing that paints as a straight line or a
 * horizontal rule rather than a picture.
 *
 * - A line is a `line` or `straightConnector1` preset with a visible outline.
 * - A rule is an unfilled `rect` with no text whose outline is its only ink
 *   and whose height is no more than the outline width, so its top and
 *   bottom edges paint as one horizontal band. The text column is not known
 *   while parsing, so a rule is recognized by that shape, not by spanning it.
 *
 * Everything else keeps the picture path and gets undefined: pictures, groups,
 * other presets and custom geometry, a rotated shape (rotation is not
 * painted), a shape with no visible outline or with an outline width outside
 * `ST_LineWidth`, and a negative or zero-size extent.
 *
 * @param container - The wp:inline or wp:anchor element
 * @param size - The drawing's extent (wp:extent) in EMUs
 */
export function parseVectorShape(
  container: XmlElement,
  size: ImageSize
): ImageVectorShape | undefined {
  const graphicData = findByFullName(findByFullName(container, 'a:graphic'), 'a:graphicData');
  const wsp = findByFullName(graphicData, 'wps:wsp');
  if (!wsp || size.width < 0 || size.height < 0 || (size.width === 0 && size.height === 0)) {
    return undefined;
  }

  // Read the preset itself: parseShape reports custom geometry as `rect` too.
  const prstGeom = findByFullName(findByFullName(wsp, 'wps:spPr'), 'a:prstGeom');
  const preset = getAttribute(prstGeom, null, 'prst');
  if (preset !== 'line' && preset !== 'straightConnector1' && preset !== 'rect') return undefined;

  const shape = parseShape(wsp);
  const { outline } = shape;
  if (!outline || (shape.transform?.rotation ?? 0) % 360 !== 0) return undefined;
  const strokeWidth = outline.width;
  if (strokeWidth !== undefined && !(strokeWidth >= 0 && strokeWidth <= MAX_LINE_WIDTH_EMU)) {
    return undefined;
  }
  if (
    preset === 'rect' &&
    (hasFill(shape) ||
      hasTextContent(shape) ||
      size.width === 0 ||
      size.height > (strokeWidth ?? 0))
  ) {
    return undefined;
  }

  const vectorShape: ImageVectorShape = { shapeType: preset, outline };
  if (shape.transform?.flipH) vectorShape.flipH = true;
  if (shape.transform?.flipV) vectorShape.flipV = true;
  return vectorShape;
}
