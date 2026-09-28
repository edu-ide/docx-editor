/**
 * The image run for an image node's attributes — a picture, or the preview of
 * content kept as parsed (a Word line or rule, or a frame keeping a shape's
 * place), which layout paints the same way.
 */

import type { ImageRun } from '../../layout-engine/types';
import type { ImageVectorShape } from '../../types/document';
import { constrainImageToPage } from './shared';
import { vectorShapePicture } from './vectorShape';

/** The revision an image run takes from its tracked-change marks. */
export type ImageRunChange = Pick<
  ImageRun,
  'isInsertion' | 'isDeletion' | 'changeAuthor' | 'changeDate' | 'changeRevisionId'
>;

export function imageRunFor(
  attrs: Record<string, unknown>,
  change: ImageRunChange,
  pageContentHeight: number | undefined,
  pmStart: number,
  pmEnd: number
): ImageRun {
  // A wps:wsp line or rule has no picture: it paints its stroke as an SVG,
  // and an axis it has no extent on takes the stroke width, not 100px.
  const vectorShape = attrs.vectorShape as ImageVectorShape | null;
  const picture = vectorShape
    ? vectorShapePicture(
        vectorShape,
        constrainImageToPage(
          (attrs.width as number) || 0,
          (attrs.height as number) || 0,
          pageContentHeight
        )
      )
    : {
        src: attrs.src as string,
        ...constrainImageToPage(
          (attrs.width as number) || 100,
          (attrs.height as number) || 100,
          pageContentHeight
        ),
      };
  // The tracked-change marks let an inserted/deleted picture paint in the
  // revision color and resolve with the rest of the change.
  return {
    kind: 'image',
    src: picture.src,
    width: picture.width,
    height: picture.height,
    alt: attrs.alt as string | undefined,
    transform: attrs.transform as string | undefined,
    wrapType: attrs.wrapType as string | undefined,
    displayMode: attrs.displayMode as 'inline' | 'block' | 'float' | undefined,
    cssFloat: attrs.cssFloat as 'left' | 'right' | 'none' | undefined,
    distTop: attrs.distTop as number | undefined,
    distBottom: attrs.distBottom as number | undefined,
    distLeft: attrs.distLeft as number | undefined,
    distRight: attrs.distRight as number | undefined,
    position: attrs.position as ImageRun['position'] | undefined,
    cropTop: attrs.cropTop as number | undefined,
    cropRight: attrs.cropRight as number | undefined,
    cropBottom: attrs.cropBottom as number | undefined,
    cropLeft: attrs.cropLeft as number | undefined,
    opacity: attrs.opacity as number | undefined,
    isInsertion: change.isInsertion,
    isDeletion: change.isDeletion,
    changeAuthor: change.changeAuthor,
    changeDate: change.changeDate,
    changeRevisionId: change.changeRevisionId,
    pmStart,
    pmEnd,
  };
}
