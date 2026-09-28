/**
 * Pictures, and the run content kept as parsed that paints like one
 * (a Word line or rule, or the frame that keeps a shape's place).
 */

import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../../schema';
import type { Image, PreservedXmlContent } from '../../../types/document';
import { emuToPixels } from '../../../docx/imageParser';
import { isWrapNone } from '../../../docx/wrapTypes';

/**
 * Convert an Image to a ProseMirror image node
 *
 * DOCX images have size in EMUs (English Metric Units), which must be
 * converted to pixels for proper HTML rendering.
 * 914400 EMU = 1 inch = 96 CSS pixels
 *
 * Image types in DOCX:
 * 1. Inline (wp:inline) - flows with text like a character
 * 2. Floating/Anchored (wp:anchor) with wrap types:
 *    - Square/Tight/Through: text wraps around image
 *      - wrapText='left' → text on LEFT, image floats RIGHT
 *      - wrapText='right' → text on RIGHT, image floats LEFT
 *      - wrapText='bothSides' → depends on horizontal alignment
 *    - TopAndBottom: image on its own line, text above/below only
 *    - None/Behind/InFront: positioned image, no text wrap
 */
export function convertImage(image: Image): PMNode {
  // Convert EMU to pixels for proper sizing
  const widthPx = image.size?.width ? emuToPixels(image.size.width) : undefined;
  const heightPx = image.size?.height ? emuToPixels(image.size.height) : undefined;

  // Determine wrap type and float direction
  const wrapType = image.wrap.type;
  const wrapText = image.wrap.wrapText;
  const hAlign = image.position?.horizontal?.alignment;

  // Determine CSS float based on wrap settings
  // In DOCX: wrapText='left' means "text flows on the left" → image is on right → float: right
  //          wrapText='right' means "text flows on the right" → image is on left → float: left
  let cssFloat: 'left' | 'right' | 'none' | undefined;

  if (wrapType === 'inline') {
    cssFloat = 'none'; // Inline images don't float
  } else if (wrapType === 'topAndBottom') {
    cssFloat = 'none'; // Block images don't float
  } else if (wrapType === 'square' || wrapType === 'tight' || wrapType === 'through') {
    // These wrap types support text wrapping around the image
    if (wrapText === 'left') {
      cssFloat = 'right'; // Text on left → image floats right
    } else if (wrapText === 'right') {
      cssFloat = 'left'; // Text on right → image floats left
    } else if (wrapText === 'bothSides' || wrapText === 'largest') {
      // Use horizontal alignment to determine float
      if (hAlign === 'left') {
        cssFloat = 'left';
      } else if (hAlign === 'right') {
        cssFloat = 'right';
      } else {
        cssFloat = 'none'; // Center or no alignment → block
      }
    } else {
      // Default: use horizontal alignment
      if (hAlign === 'left') {
        cssFloat = 'left';
      } else if (hAlign === 'right') {
        cssFloat = 'right';
      } else {
        cssFloat = 'none';
      }
    }
  } else {
    // Behind, inFront, etc. - positioned images, no float
    cssFloat = 'none';
  }

  // Determine display mode for CSS
  let displayMode: 'inline' | 'block' | 'float' = 'inline';
  if (wrapType === 'inline') {
    displayMode = 'inline';
  } else if (wrapType === 'topAndBottom') {
    displayMode = 'block';
  } else if (isWrapNone(wrapType)) {
    // wrapNone (behind / inFront): positioned float, painted out of paragraph flow.
    displayMode = 'float';
  } else if (cssFloat && cssFloat !== 'none') {
    displayMode = 'float';
  } else {
    // Centered square/tight/through images without a wrapping side fall back to block.
    displayMode = 'block';
  }

  // Build transform string if needed (rotation, flip)
  let transform: string | undefined;
  if (image.transform) {
    const transforms: string[] = [];
    if (image.transform.rotation) {
      transforms.push(`rotate(${image.transform.rotation}deg)`);
    }
    if (image.transform.flipH) {
      transforms.push('scaleX(-1)');
    }
    if (image.transform.flipV) {
      transforms.push('scaleY(-1)');
    }
    if (transforms.length > 0) {
      transform = transforms.join(' ');
    }
  }

  // Convert wrap distances from EMU to pixels for margins. Nullish, not truthy:
  // an explicit `w:distL="0"` (image butted flush against the wrapped text) is a
  // meaningful 0 that must survive — collapsing it to `undefined` lets the
  // float-zone fall back to its non-zero default (12px L/R), opening a phantom
  // gap. Only an ABSENT distance should fall back. Same falsy-zero class as the
  // page-margin/header fixes (#740).
  const distTop = image.wrap.distT != null ? emuToPixels(image.wrap.distT) : undefined;
  const distBottom = image.wrap.distB != null ? emuToPixels(image.wrap.distB) : undefined;
  const distLeft = image.wrap.distL != null ? emuToPixels(image.wrap.distL) : undefined;
  const distRight = image.wrap.distR != null ? emuToPixels(image.wrap.distR) : undefined;

  // Build position data for floating images
  let position:
    | {
        horizontal?: { relativeTo?: string; posOffset?: number; align?: string };
        vertical?: { relativeTo?: string; posOffset?: number; align?: string };
      }
    | undefined;
  if (image.position) {
    position = {
      horizontal: image.position.horizontal
        ? {
            relativeTo: image.position.horizontal.relativeTo,
            posOffset: image.position.horizontal.posOffset,
            align: image.position.horizontal.alignment,
          }
        : undefined,
      vertical: image.position.vertical
        ? {
            relativeTo: image.position.vertical.relativeTo,
            posOffset: image.position.vertical.posOffset,
            align: image.position.vertical.alignment,
          }
        : undefined,
    };
  }

  // Convert outline to border attrs
  let borderWidth: number | undefined;
  let borderColor: string | undefined;
  let borderStyle: string | undefined;
  if (image.outline && image.outline.width) {
    // Convert EMU to pixels (1 EMU = 1/914400 inch, 1 inch = 96 px)
    borderWidth = Math.round((image.outline.width / 914400) * 96 * 100) / 100;
    if (image.outline.color?.rgb) {
      borderColor = `#${image.outline.color.rgb}`;
    }
    // Map OOXML dash styles to CSS border styles
    const styleMap: Record<string, string> = {
      solid: 'solid',
      dot: 'dotted',
      dash: 'dashed',
      lgDash: 'dashed',
      dashDot: 'dashed',
      lgDashDot: 'dashed',
      lgDashDotDot: 'dashed',
      sysDot: 'dotted',
      sysDash: 'dashed',
      sysDashDot: 'dashed',
      sysDashDotDot: 'dashed',
    };
    borderStyle = image.outline.style ? styleMap[image.outline.style] || 'solid' : 'solid';
  }

  // Effect extent (shadow/glow padding) is parsed in EMU; convert to px so
  // the renderer can apply it as outer margin.
  const effectExtentTop = image.padding?.top ? emuToPixels(image.padding.top) : undefined;
  const effectExtentBottom = image.padding?.bottom ? emuToPixels(image.padding.bottom) : undefined;
  const effectExtentLeft = image.padding?.left ? emuToPixels(image.padding.left) : undefined;
  const effectExtentRight = image.padding?.right ? emuToPixels(image.padding.right) : undefined;

  return schema.node('image', {
    src: image.src || '',
    alt: image.alt,
    title: image.title,
    width: widthPx,
    height: heightPx,
    rId: image.rId,
    wrapType: wrapType,
    displayMode: displayMode,
    cssFloat: cssFloat,
    transform: transform,
    distTop: distTop,
    distBottom: distBottom,
    distLeft: distLeft,
    distRight: distRight,
    position: position,
    borderWidth: borderWidth,
    borderColor: borderColor,
    borderStyle: borderStyle,
    wrapText: wrapText,
    hlinkHref: image.hlinkHref,
    cropTop: image.crop?.top,
    cropRight: image.crop?.right,
    cropBottom: image.crop?.bottom,
    cropLeft: image.crop?.left,
    opacity: image.opacity,
    effectExtentTop,
    effectExtentBottom,
    effectExtentLeft,
    effectExtentRight,
    layoutInCell: image.layoutInCell,
    allowOverlap: image.allowOverlap,
    // A line or rule's stroke, kept in EMUs for layout to paint and for fromProseDoc.
    vectorShape: image.vectorShape,
  });
}

/**
 * Convert run content kept as parsed to its atom; the preview carries the
 * attributes of the image it paints as, so layout draws it like one.
 */
export function convertPreservedXml(content: PreservedXmlContent): PMNode {
  return schema.node('preservedXml', {
    xml: content.xml,
    preview: content.preview ? convertImage(content.preview.image).attrs : null,
  });
}
