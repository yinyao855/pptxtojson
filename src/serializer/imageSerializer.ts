/**
 * Serializes PicNodeData to pptxtojson Image, Video, or Audio element.
 * Resolves blip embed to base64 src; crop to rect; border from line.
 */

import type { PicNodeData } from '../model/nodes/PicNode';
import type { RenderContext } from './RenderContext';
import { resolveRelTarget } from '../parser/RelParser';
import { encodeMediaForWebDisplay } from '../utils/mediaWebConvert';
import { lineStyleToBorder } from './borderMapper';
import type { Image as ImageElement, Video, Audio } from '../adapter/types';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

function dataUrlToRawBase64(dataUrl: string): string {
  const m = /^data:[^;]+;base64,(.+)$/.exec(dataUrl);
  return m ? m[1] : '';
}

/**
 * Serialize picture node to Image, Video, or Audio element.
 */
export function pictureToElement(
  node: PicNodeData,
  ctx: RenderContext,
  order: number,
): ImageElement | Video | Audio {
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);
  const embedId = node.blipEmbed ?? node.mediaRId;
  let src = '';
  let blob = '';
  if (embedId) {
    const rel = ctx.slide.rels.get(embedId);
    if (rel) {
      const basePath = ctx.slide.slidePath.replace(/\/[^/]+$/, '');
      const mediaPath = resolveRelTarget(basePath, rel.target);
      const data = ctx.presentation.media.get(mediaPath);
      if (data) {
        src = encodeMediaForWebDisplay(mediaPath, data);
        blob = dataUrlToRawBase64(src);
      }
    }
  }
  if (node.isVideo) {
    return { type: 'video', left, top, width, height, blob: blob || undefined, src: src || undefined, order };
  }
  if (node.isAudio) {
    return { type: 'audio', left, top, width, height, blob: blob || '', order };
  }
  const spPr = node.source.child('spPr');
  const ln = spPr.exists() ? spPr.child('ln') : node.source.child('__none__');
  const borderResult = ln.exists() ? lineStyleToBorder(ln, ctx) : {
    border: { borderColor: '#000000', borderWidth: 0, borderType: 'solid' as const },
    borderStrokeDasharray: '',
  };
  let rect: ImageElement['rect'] | undefined;
  if (node.crop && (node.crop.top !== 0 || node.crop.bottom !== 0 || node.crop.left !== 0 || node.crop.right !== 0)) {
    rect = {
      t: node.crop.top,
      b: node.crop.bottom,
      l: node.crop.left,
      r: node.crop.right,
    };
  }
  return {
    type: 'image',
    left,
    top,
    width,
    height,
    src,
    rotate: node.rotation,
    isFlipH: node.flipH,
    isFlipV: node.flipV,
    order,
    rect,
    geom: 'rect',
    borderColor: borderResult.border.borderColor,
    borderWidth: borderResult.border.borderWidth,
    borderType: borderResult.border.borderType,
    borderStrokeDasharray: borderResult.borderStrokeDasharray || '',
  };
}
