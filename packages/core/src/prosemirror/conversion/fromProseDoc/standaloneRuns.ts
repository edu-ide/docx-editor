/**
 * Inline nodes that each stand in a run of their own: pictures, shapes, and
 * content kept as parsed (a Word shape, chart or VML), which goes back to the
 * model with its XML and the picture-less image it paints as.
 */

import type { Node as PMNode } from 'prosemirror-model';
import type { PreservedXmlContent, Run } from '../../../types/document';
import { createImageRun, createShapeRun } from './runs';

/**
 * Create the run for content kept as parsed. Saving writes `xml` back as it
 * was; the preview is rebuilt through the image node it was converted from.
 */
export function createPreservedXmlRun(node: PMNode): Run {
  const previewAttrs = node.attrs.preview as Record<string, unknown> | null;
  const preview = previewAttrs
    ? createImageRun(node.type.schema.nodes.image.create(previewAttrs)).content[0]
    : undefined;
  const content: PreservedXmlContent = {
    type: 'preservedXml',
    xml: node.attrs.xml as string,
    ...(preview?.type === 'drawing' ? { preview } : {}),
  };
  return { type: 'run', content: [content] };
}

/** The run each standalone inline node becomes, by node type name. */
export const STANDALONE_RUNS: ReadonlyMap<string, (node: PMNode) => Run> = new Map([
  ['image', createImageRun],
  ['shape', createShapeRun],
  ['preservedXml', createPreservedXmlRun],
]);
