/**
 * Document Run/RunContent/Hyperlink/Image/Shape/Field/Math → PM nodes
 * (Document → ProseMirror direction).
 *
 * Each Document content type has a factory that produces the matching PM
 * node(s). `convertRun` merges the run's `formatting` with the paragraph's
 * style cascade before projecting via `textFormattingToMarks` from ./marks.ts.
 *
 * `convertInlineSdt` lives in ./paragraph.ts (not here) — same cycle-break
 * decision as fromProseDoc, since it recurses through ./paragraph.ts's
 * `convertParagraph` content walker.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../../schema';
import type {
  Run,
  RunContent,
  Hyperlink,
  Shape,
  SimpleField,
  ComplexField,
  MathEquation,
  TextFormatting,
} from '../../../types/document';
import { emuToPixels } from '../../../docx/imageParser';
import { convertImage, convertPreservedXml } from './image';
import { mergeTextFormatting } from '../../../utils/textFormattingMerge';
import type { StyleResolver } from '../../styles';
import { textFormattingToMarks } from './marks';

/**
 * Convert a SimpleField or ComplexField to a ProseMirror field node.
 * Preserves run formatting (bold, fontSize, color, etc.) as PM marks.
 * Accepts styleFormatting so fields inherit paragraph-level formatting
 * (same as convertRun does for regular text runs).
 */
export function convertField(
  field: SimpleField | ComplexField,
  styleFormatting?: TextFormatting
): PMNode | null {
  // Extract display text and formatting from field content/result
  let displayText = '';
  let fieldFormatting: TextFormatting | undefined;
  const runs = field.type === 'simpleField' ? field.content : field.fieldResult;
  if (runs) {
    for (const r of runs) {
      if (r.type === 'run') {
        for (const c of r.content) {
          if (c.type === 'text') displayText += c.text;
        }
        // Use formatting from the first run that has it
        if (!fieldFormatting && r.formatting) {
          fieldFormatting = r.formatting;
        }
      }
    }
  }

  // Fall back to the field's structural-run formatting (the run holding the
  // fldChars) when no result run carried any — e.g. a PAGE field collapsed into
  // a single run, where the w:rPr lives on the field run, not a result run.
  // Without this the field renders at the default size/color.
  const resolvedFieldFormatting =
    fieldFormatting ?? (field.type === 'complexField' ? field.formatting : undefined);

  // Merge style formatting with field run formatting (inline takes precedence)
  const mergedFormatting = mergeTextFormatting(styleFormatting, resolvedFieldFormatting);
  const marks = textFormattingToMarks(mergedFormatting);

  return schema.node(
    'field',
    {
      fieldType: field.fieldType,
      instruction: field.instruction,
      displayText,
      fieldKind: field.type === 'simpleField' ? 'simple' : 'complex',
      fldLock: field.fldLock ?? false,
      dirty: field.dirty ?? false,
      fieldCode: nestedFieldCodeAttr(field),
    },
    undefined,
    marks
  );
}

/**
 * The field code of a complex field whose instruction holds nested fields, as the PM
 * `fieldCode` attr (upstream 390c1772): the instruction string cannot carry a field, and
 * save writes these runs back in place. Other fields keep `null` and save from `instruction`.
 */
function nestedFieldCodeAttr(field: SimpleField | ComplexField): string | null {
  if (field.type !== 'complexField') return null;
  const nested = field.fieldCode.some((run) =>
    run.content.some((content) => content.type === 'fieldChar')
  );
  return nested ? JSON.stringify({ instruction: field.instruction, runs: field.fieldCode }) : null;
}

/**
 * Convert a MathEquation to a ProseMirror math node.
 */
export function convertMathEquation(math: MathEquation): PMNode | null {
  return schema.node('math', {
    display: math.display,
    ommlXml: math.ommlXml,
    plainText: math.plainText || '',
  });
}

/**
 * Convert a Run to ProseMirror text nodes with marks
 *
 * @param run - The run to convert
 * @param styleFormatting - Text formatting from the paragraph's style (e.g., Heading1's font size/color)
 */
export function convertRun(
  run: Run,
  styleFormatting?: TextFormatting,
  styleResolver?: StyleResolver | null
): PMNode[] {
  const nodes: PMNode[] = [];

  // Merge style formatting with run's inline formatting
  // Inline formatting takes precedence over style formatting
  //
  // Use getRunStyleOwnProperties (not resolveRunStyle) to avoid docDefaults
  // from the character style overriding paragraph style properties.
  // The styleFormatting parameter already includes docDefaults from paragraph
  // style resolution, so we only need the character style's own properties.
  const runStyleFormatting = run.formatting?.styleId
    ? styleResolver?.getRunStyleOwnProperties(run.formatting.styleId)
    : undefined;
  const mergedFormatting = mergeTextFormatting(
    mergeTextFormatting(styleFormatting, runStyleFormatting),
    run.formatting
  );
  const marks = textFormattingToMarks(mergedFormatting);

  for (const content of run.content) {
    const contentNodes = convertRunContent(content, marks);
    nodes.push(...contentNodes);
  }

  return nodes;
}

/**
 * Convert RunContent to ProseMirror nodes
 */
function convertRunContent(content: RunContent, marks: ReturnType<typeof schema.mark>[]): PMNode[] {
  switch (content.type) {
    case 'text':
      if (content.text) {
        return [schema.text(content.text, marks)];
      }
      return [];

    case 'break':
      if (content.breakType === 'textWrapping' || !content.breakType) {
        // Carry marks (including any enclosing hyperlink) so the break is
        // recognized as inside the hyperlink on the way back out.
        return [schema.node('hardBreak', null, undefined, marks)];
      }
      // Page breaks not supported in inline content
      return [];

    case 'tab':
      // Carry marks (including any enclosing hyperlink) so round-trip keeps
      // the tab inside the hyperlink — TOC entries depend on this.
      return [schema.node('tab', null, undefined, marks)];

    case 'drawing':
      if (content.image) {
        return [convertImage(content.image)];
      }
      return [];

    case 'preservedXml':
      return [convertPreservedXml(content)];

    case 'shape': {
      // Shapes with text body are handled as text boxes at block level
      // Other shapes render as inline SVG
      const shp = content.shape;
      if (shp.textBody && shp.textBody.content.length > 0) {
        // Skip - handled by extractTextBoxesFromParagraph
        return [];
      }
      return [convertShape(shp)];
    }

    case 'footnoteRef':
      // Footnote reference - render as superscript number with footnoteRef mark
      const footnoteMark = schema.mark('footnoteRef', {
        id: content.id.toString(),
        noteType: 'footnote',
      });
      return [schema.text(content.id.toString(), [...marks, footnoteMark])];

    case 'endnoteRef':
      // Endnote reference - render as superscript number with footnoteRef mark
      const endnoteMark = schema.mark('footnoteRef', {
        id: content.id.toString(),
        noteType: 'endnote',
      });
      return [schema.text(content.id.toString(), [...marks, endnoteMark])];

    default:
      return [];
  }
}

/**
 * Convert a Hyperlink to ProseMirror nodes with link mark
 *
 * @param hyperlink - The hyperlink to convert
 * @param styleFormatting - Text formatting from the paragraph's style
 */
export function convertHyperlink(
  hyperlink: Hyperlink,
  styleFormatting?: TextFormatting,
  styleResolver?: StyleResolver | null
): PMNode[] {
  const nodes: PMNode[] = [];

  // Create link mark — internal anchors use #bookmarkName format
  const href = hyperlink.href || (hyperlink.anchor ? `#${hyperlink.anchor}` : '');
  const linkMark = schema.mark('hyperlink', {
    href,
    tooltip: hyperlink.tooltip,
    rId: hyperlink.rId,
  });

  for (const child of hyperlink.children) {
    if (child.type === 'run') {
      // Merge style formatting with run's inline formatting. Mirror convertRun
      // and pull *only* the character style's own properties — resolveRunStyle
      // walks all the way up to docDefaults, and merging that on top of the
      // already-resolved paragraph style re-introduces docDefaults' rPr (e.g.
      // sz=24 in the parity doc) and clobbers the paragraph style's run size.
      // TOC3 entries are wrapped in <w:hyperlink> so this path is the one
      // that decides their font size; using resolveRunStyle here made TOC3
      // render at the doc default 12pt instead of the TOC3 style's 10pt.
      const runStyleFormatting = child.formatting?.styleId
        ? styleResolver?.getRunStyleOwnProperties(child.formatting.styleId)
        : undefined;
      const mergedFormatting = mergeTextFormatting(
        mergeTextFormatting(styleFormatting, runStyleFormatting),
        child.formatting
      );
      const runMarks = textFormattingToMarks(mergedFormatting);
      // Add link mark to run marks
      const allMarks = [...runMarks, linkMark];

      // Delegate to convertRunContent so tabs, breaks, fields, footnote refs
      // etc. inside a hyperlink round-trip. (Drawings and shapes inside a
      // hyperlink don't carry the hyperlink mark through convertImage /
      // convertShape today — linked-image round-trip is a separate gap.)
      for (const content of child.content) {
        nodes.push(...convertRunContent(content, allMarks));
      }
    }
  }

  return nodes;
}

/**
 * Convert a Shape to a ProseMirror shape node (inline SVG)
 */
function convertShape(shape: Shape): PMNode {
  const widthPx = shape.size?.width ? emuToPixels(shape.size.width) : 100;
  const heightPx = shape.size?.height ? emuToPixels(shape.size.height) : 80;

  let fillColor: string | undefined;
  let fillType: string = 'solid';
  let gradientType: string | undefined;
  let gradientAngle: number | undefined;
  let gradientStops: string | undefined;
  if (shape.fill) {
    fillType = shape.fill.type;
    if (shape.fill.color?.rgb) {
      fillColor = `#${shape.fill.color.rgb}`;
    }
    // Extract gradient data
    if (shape.fill.type === 'gradient' && shape.fill.gradient) {
      const g = shape.fill.gradient;
      gradientType = g.type;
      gradientAngle = g.angle;
      // Convert stops to serializable format with CSS colors
      gradientStops = JSON.stringify(
        g.stops.map((s) => ({
          position: s.position,
          color: s.color.rgb ? `#${s.color.rgb}` : '#000000',
        }))
      );
    }
  }

  let outlineWidth: number | undefined;
  let outlineColor: string | undefined;
  let outlineStyle: string | undefined;
  if (shape.outline) {
    if (shape.outline.width) {
      outlineWidth = Math.round((shape.outline.width / 914400) * 96 * 100) / 100;
    }
    if (shape.outline.color?.rgb) {
      outlineColor = `#${shape.outline.color.rgb}`;
    }
    outlineStyle = shape.outline.style || 'solid';
  }

  let transform: string | undefined;
  if (shape.transform) {
    const transforms: string[] = [];
    if (shape.transform.rotation) {
      transforms.push(`rotate(${shape.transform.rotation}deg)`);
    }
    if (shape.transform.flipH) {
      transforms.push('scaleX(-1)');
    }
    if (shape.transform.flipV) {
      transforms.push('scaleY(-1)');
    }
    if (transforms.length > 0) {
      transform = transforms.join(' ');
    }
  }

  return schema.node('shape', {
    shapeType: shape.shapeType || 'rect',
    shapeId: shape.id,
    width: widthPx,
    height: heightPx,
    fillColor,
    fillType,
    gradientType,
    gradientAngle,
    gradientStops,
    outlineWidth,
    outlineColor,
    outlineStyle,
    transform,
  });
}
