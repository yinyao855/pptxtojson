/**
 * Serializes PicNodeData to pptxtojson Image, Video, or Audio element.
 * Resolves blip embed to base64 src; crop to rect; border from line.
 */

import type { PicNodeData } from '../model/nodes/PicNode';
import type { RenderContext } from '../resolve/RenderContext';
import { resolveRelTarget } from '../parser/RelParser';
import { getMimeType, toDataUrl } from '../utils/media';
import { lineStyleToBorder } from './borderMapper';
import type { Image as ImageElement, Video, Audio } from '../adapter/types';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

function arrayBufferToBase64(data: Uint8Array): string {
  let binary = '';
  const len = data.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(data[i]);
  }
  if (typeof btoa !== 'undefined') return btoa(binary);
  const NodeBuffer = (typeof globalThis !== 'undefined' && (globalThis as unknown as { Buffer?: { from(a: Uint8Array): { toString(e: string): string } } }).Buffer);
  if (NodeBuffer) return NodeBuffer.from(data).toString('base64');
  return btoa(binary);
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
        const b64 = arrayBufferToBase64(data);
        const mime = getMimeType(mediaPath);
        src = toDataUrl(b64, mime);
        blob = b64;
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
