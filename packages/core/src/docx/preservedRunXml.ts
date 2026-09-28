/**
 * Run children the editor cannot write back from its model are kept as their
 * XML, so a save writes them unchanged.
 *
 * The editor models pictures (`pic:pic`) as images and text boxes as shapes.
 * Everything else a run can draw — a Word shape, connector or group, a chart,
 * SmartArt, the `mc:AlternateContent` Word wraps them in, a VML shape or an
 * embedded object — used to be dropped on read or turned into an empty picture
 * frame. The editor's default save rewrites an edited paragraph from the model,
 * so editing the paragraph that held one deleted it without a warning.
 */

import type {
  DrawingContent,
  MediaFile,
  PreservedXmlContent,
  RelationshipMap,
  RunContent,
} from '../types/document';
import { parseImage } from './imageParser';
import { WORD_NAMESPACES } from './serializer/wordNamespaces';
import { parseVmlImageContent } from './vmlImageParser';
import { isWatermarkShape } from './vmlWatermarkParser';
import {
  NAMESPACES,
  elementToXml,
  findAllDeep,
  findDeep,
  getAttribute,
  getChildElements,
  getLocalName,
  type XmlElement,
} from './xmlParser';

/** Prefixes the document and header/footer serializers declare on their root element. */
const ROOT_PREFIXES = new Set([
  'wpc',
  'mc',
  'o',
  'r',
  'm',
  'v',
  'wp14',
  'wp',
  'w10',
  'w',
  'w14',
  'w15',
  'w16se',
  'w16cid',
  'w16',
  'w16cex',
  'w16sdtdh',
  'wne',
  'wpg',
  'wps',
]);

/**
 * Parse a `w:drawing` into drawing content: a picture, or a picture-less image
 * (a Word shape, whose `vectorShape` is set when it is a line or rule).
 */
export function parseDrawingContent(
  element: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null
): DrawingContent | null {
  const image = parseImage(element, rels ?? undefined, media ?? undefined);
  return image ? { type: 'drawing', image } : null;
}

/** A drawing whose graphic is a picture — the only kind modelled as an editable image. */
function isPictureDrawing(drawing: XmlElement): boolean {
  const graphicData = findDeep(drawing, 'a', 'graphicData');
  return getAttribute(graphicData, null, 'uri') === NAMESPACES.pic;
}

/** Text boxes are modelled by enrichParagraphTextBoxes, which reads them from the paragraph XML. */
function holdsTextBox(element: XmlElement): boolean {
  return findDeep(element, 'w', 'txbxContent') !== null;
}

function isWatermark(element: XmlElement): boolean {
  return findAllDeep(element, 'v', 'shape').some(
    (shape) =>
      isWatermarkShape(shape, (getAttribute(shape, null, 'id') ?? '').toLowerCase()) ||
      findDeep(shape, 'v', 'textpath') !== null
  );
}

/**
 * The element with a declaration for each namespace prefix it uses that neither
 * it nor the serializers' root declares — Word declares `wpi:` (ink) or `cx:`
 * (chartex) on its root, a non-Word writer may declare `a:` or `pic:` there —
 * so a full repack, which writes its own root, still writes valid XML. Null when
 * such a prefix has no known namespace.
 */
function selfContained(element: XmlElement): XmlElement | null {
  const used = new Set<string>();
  const declared = new Set<string>();
  const visit = (el: XmlElement): void => {
    const name = el.name ?? '';
    if (name.includes(':')) used.add(name.slice(0, name.indexOf(':')));
    for (const [attr, value] of Object.entries(el.attributes ?? {})) {
      if (attr.startsWith('xmlns:')) declared.add(attr.slice('xmlns:'.length));
      else if (attr.includes(':')) used.add(attr.slice(0, attr.indexOf(':')));
      // mc:Choice Requires="wps" names a prefix that must be in scope.
      if (attr === 'Requires') {
        for (const prefix of String(value).split(/\s+/).filter(Boolean)) used.add(prefix);
      }
    }
    for (const child of el.elements ?? []) {
      if (child.type === 'element') visit(child);
    }
  };
  visit(element);

  const known: Record<string, string> = { ...WORD_NAMESPACES, ...NAMESPACES };
  const declarations: Record<string, string> = {};
  for (const prefix of used) {
    if (prefix === 'xml' || declared.has(prefix) || ROOT_PREFIXES.has(prefix)) continue;
    if (!known[prefix]) return null;
    declarations[`xmlns:${prefix}`] = known[prefix];
  }
  if (Object.keys(declarations).length === 0) return element;
  return { ...element, attributes: { ...element.attributes, ...declarations } };
}

function preserve(element: XmlElement, preview: DrawingContent | null): PreservedXmlContent | null {
  const whole = selfContained(element);
  if (!whole) return null;
  return { type: 'preservedXml', xml: elementToXml(whole), ...(preview ? { preview } : {}) };
}

/**
 * `w:drawing` in a run. A picture stays an image and a text box is left to the
 * text box pass. Anything else is kept as parsed and shows its picture-less
 * image, which paints a line or rule and keeps any other shape's place.
 */
export function parseDrawingRunContent(
  element: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null
): RunContent[] {
  if (holdsTextBox(element)) return [];
  const drawing = parseDrawingContent(element, rels, media);
  if (isPictureDrawing(element)) return drawing ? [drawing] : [];
  const kept = preserve(element, drawing);
  if (kept) return [kept];
  return drawing ? [drawing] : [];
}

/**
 * `mc:AlternateContent` in a run. Pictures in the chosen branch stay images and
 * a text box is left to the text box pass. Anything else is kept whole, fallback
 * included, and painted when its drawing is a line or rule.
 */
export function parseAlternateContentRunContent(
  element: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null
): RunContent[] {
  const branches = getChildElements(element);
  const chosen =
    branches.find((el) => getLocalName(el.name ?? '') === 'Choice') ??
    branches.find((el) => getLocalName(el.name ?? '') === 'Fallback');
  if (!chosen || holdsTextBox(chosen)) return [];

  const drawings = getChildElements(chosen).filter(
    (el) => getLocalName(el.name ?? '') === 'drawing'
  );
  if (drawings.length > 0 && drawings.every(isPictureDrawing)) {
    return drawings
      .map((drawing) => parseDrawingContent(drawing, rels, media))
      .filter((drawing): drawing is DrawingContent => Boolean(drawing?.image.src));
  }
  const preview =
    drawings
      .map((drawing) => parseDrawingContent(drawing, rels, media))
      .find((drawing) => drawing?.image.vectorShape) ?? null;
  const kept = preserve(element, preview);
  return kept ? [kept] : [];
}

/**
 * VML `w:pict` / `w:object` in a run. A picture stays an image and a watermark
 * is modelled from the header by extractWatermark. An embedded object keeps its
 * data, its picture only previewing it; any other VML is kept as parsed.
 */
export function parseVmlRunContent(
  element: XmlElement,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null
): RunContent[] {
  const picture = parseVmlImageContent(element, rels, media);
  if (getLocalName(element.name ?? '') === 'object') {
    const kept = preserve(element, picture);
    if (kept) return [kept];
    return picture ? [picture] : [];
  }
  if (picture) return [picture];
  if (isWatermark(element)) return [];
  const kept = preserve(element, null);
  return kept ? [kept] : [];
}
