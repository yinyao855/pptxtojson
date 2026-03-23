/**
 * Image serializer — converts PicNodeData into positioned HTML image/video/audio elements.
 */

import type { PicNodeData } from '../model/nodes/PicNode';
import type { RenderContext } from './RenderContext';
import { SafeXmlNode } from '../parser/XmlParser';
import { encodeMediaForWebDisplay } from '../utils/mediaWebConvert';
import { lineStyleToBorder } from './borderMapper';
import type { Image, Video, Audio } from '../adapter/types';
import { getOrCreateBlobUrl, resolveMediaPath } from '../utils/media';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

/**
 * Check if a file extension is an unsupported legacy format (WMF only now; EMF is handled).
 */
function isUnsupportedFormat(path: string): boolean {
  const ext = path.split('.').pop()?.toLowerCase() || '';
  return ext === 'wmf';
}

/**
 * Check if a file path is an EMF image.
 */
function isEmfFormat(path: string): boolean {
  const ext = path.split('.').pop()?.toLowerCase() || '';
  return ext === 'emf';
}

/** OOXML fixed-point scale (100000 = 100%). */
const OOXML_100K = 100000;

/**
 * Build `filters` for PPTist: `sharpen`, `colorTemperature`, `saturation`, `brightness`, `contrast`.
 *
 * - **ISO / DrawingML**: `<a:lum bright contrast>` on `<a:blip>`.
 * - **Office 2010+** (same as legacy `src1/fill.js` `getPicFilters`): `a:extLst` → `ext` →
 *   `a14:imgProps` / `a14:imgLayer` / `a14:imgEffect` → `a14:saturation`, `a14:brightnessContrast`,
 *   `a14:sharpenSoften`, `a14:colorTemperature`.
 *
 * Extension effects are applied after `lum` and may override brightness/contrast when both exist.
 */
function buildImageFilters(node: PicNodeData): Image['filters'] | undefined {
  const blipFill = node.source.child('blipFill');
  if (!blipFill.exists()) return undefined;
  const blip = blipFill.child('blip');
  if (!blip.exists()) return undefined;

  const out: NonNullable<Image['filters']> = {};

  applyLumToFilters(blip, out);
  applyExtLstImageEffectsToFilters(blip, out);

  return Object.keys(out).length > 0 ? out : undefined;
}

/** `<a:lum>` — brightness / contrast (values typically −100000…100000, scale 100000 = 100%). */
function applyLumToFilters(blip: SafeXmlNode, out: NonNullable<Image['filters']>): void {
  const lum = blip.child('lum');
  if (!lum.exists()) return;
  const bright = lum.numAttr('bright');
  const contrast = lum.numAttr('contrast');
  if (bright !== undefined && bright !== 0) {
    out.brightness = bright / OOXML_100K;
  }
  if (contrast !== undefined && contrast !== 0) {
    out.contrast = contrast / OOXML_100K;
  }
}

/**
 * `a:extLst` / `a14:img*` image adjustments (namespace-agnostic `localName` from DOM).
 */
function applyExtLstImageEffectsToFilters(blip: SafeXmlNode, out: NonNullable<Image['filters']>): void {
  const extLst = blip.child('extLst');
  if (!extLst.exists()) return;

  for (const ext of extLst.children()) {
    if (ext.localName !== 'ext') continue;
    const imgProps = ext.child('imgProps');
    if (!imgProps.exists()) continue;
    const imgLayer = imgProps.child('imgLayer');
    if (!imgLayer.exists()) continue;

    for (const imgEffect of imgLayer.children()) {
      if (imgEffect.localName !== 'imgEffect') continue;
      for (const el of imgEffect.allChildren()) {
        switch (el.localName) {
          case 'saturation': {
            const sat = el.numAttr('sat');
            if (sat !== undefined) {
              out.saturation = sat / OOXML_100K;
            }
            break;
          }
          case 'brightnessContrast': {
            const bright = el.numAttr('bright');
            const contrast = el.numAttr('contrast');
            if (bright !== undefined && bright !== 0) {
              out.brightness = bright / OOXML_100K;
            }
            if (contrast !== undefined && contrast !== 0) {
              out.contrast = contrast / OOXML_100K;
            }
            break;
          }
          case 'sharpenSoften': {
            const amount = el.numAttr('amount');
            if (amount !== undefined && amount !== 0) {
              // Positive = sharpen, negative = soften (PPTist only has `sharpen`; use signed value).
              out.sharpen = amount / OOXML_100K;
            }
            break;
          }
          case 'colorTemperature': {
            const ct = el.numAttr('colorTemp');
            if (ct !== undefined) {
              out.colorTemperature = ct;
            }
            break;
          }
          default:
            break;
        }
      }
    }
  }
}

/**
 * Resolve a media URL from a relationship ID.
 */
function resolveMediaUrl(rId: string | undefined, ctx: RenderContext): string | undefined {
  if (!rId) return undefined;

  const rel = ctx.slide.rels.get(rId);
  if (!rel) return undefined;

  // Check if target is an external URL
  if (rel.target.startsWith('http://') || rel.target.startsWith('https://')) {
    return rel.target;
  }

  // Resolve from embedded media
  const mediaPath = resolveMediaPath(rel.target);
  const data = ctx.presentation.media.get(mediaPath);
  if (!data) return undefined;

  return getOrCreateBlobUrl(mediaPath, data, ctx.mediaUrlCache);
}

/**
 * Render a video element.
 */
function renderVideo(
  node: PicNodeData,
  ctx: RenderContext,
  order: number,
  box: { left: number; top: number; width: number; height: number },
): Video {
  // Try to get video URL from mediaRId
  const videoUrl = resolveMediaUrl(node.mediaRId, ctx);

  // Also try to show poster image from blipEmbed
  let posterUrl: string | undefined;
  if (node.blipEmbed) {
    const rel = ctx.slide.rels.get(node.blipEmbed);
    if (rel) {
      const mediaPath = resolveMediaPath(rel.target);
      const data = ctx.presentation.media.get(mediaPath);
      if (data && !isUnsupportedFormat(mediaPath)) {
        posterUrl = getOrCreateBlobUrl(mediaPath, data, ctx.mediaUrlCache);
      }
    }
  }

  const blob = videoUrl || undefined;
  const src = posterUrl ?? videoUrl ?? undefined;

  return {
    type: 'video',
    ...box,
    blob,
    src,
    order,
  };
}

/**
 * Render an audio element.
 */
function renderAudio(
  node: PicNodeData,
  ctx: RenderContext,
  order: number,
  box: { left: number; top: number; width: number; height: number },
): Audio {
  const audioUrl = resolveMediaUrl(node.mediaRId, ctx);
  const blob = audioUrl || '';
  // TODO: optional cover image from blipEmbed

  return {
    type: 'audio',
    ...box,
    blob,
    order,
  };
}

/**
 * Render an image element.
 */
function renderImage(
  node: PicNodeData,
  ctx: RenderContext,
  order: number,
  box: { left: number; top: number; width: number; height: number },
): Image {
  const embedId = node.blipEmbed;
  let src = '';

  if (!embedId) {
    return buildImage(node, ctx, order, box, src, undefined);
  }

  const rel = ctx.slide.rels.get(embedId);
  if (!rel) {
    return buildImage(node, ctx, order, box, src, undefined);
  }

  const mediaPath = resolveMediaPath(rel.target);

  const data = ctx.presentation.media.get(mediaPath);
  if (!data) {
    return buildImage(node, ctx, order, box, src, undefined);
  }

  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

  src = encodeMediaForWebDisplay(mediaPath, bytes);
  return buildImage(node, ctx, order, box, src, buildImageFilters(node));
}

function buildImage(
  node: PicNodeData,
  ctx: RenderContext,
  order: number,
  box: { left: number; top: number; width: number; height: number },
  src: string,
  filters: Image['filters'] | undefined,
): Image {
  const spPr = node.source.child('spPr');
  const ln = spPr.exists() ? spPr.child('ln') : node.source.child('__none__');
  const borderResult = ln.exists()
    ? lineStyleToBorder(ln, ctx)
    : {
        border: { borderColor: '#000000', borderWidth: 0, borderType: 'solid' as const },
        borderStrokeDasharray: '',
      };

  let rect: Image['rect'] | undefined;
  if (
    node.crop &&
    (node.crop.top !== 0 || node.crop.bottom !== 0 || node.crop.left !== 0 || node.crop.right !== 0)
  ) {
    // OOXML srcRect
    rect = {
      t: node.crop.top,
      b: node.crop.bottom,
      l: node.crop.left,
      r: node.crop.right,
    };
  }

  return {
    type: 'image',
    ...box,
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
    ...(filters && Object.keys(filters).length > 0 ? { filters } : {}),
  };
}

/**
 * Serialize picture node to Image, Video, or Audio element.
 *
 * Handles:
 * - Standard images (png, jpg, gif, svg, bmp)
 * - Unsupported formats (wmf) with placeholder
 * - Video elements with controls
 * - Audio elements with controls
 * - Crop via `rect` (fractions)
 * - Rotation and flip on Image
 */
export function pictureToElement(
  node: PicNodeData,
  ctx: RenderContext,
  order: number,
): Image | Video | Audio {
  const box = {
    left: pxToPt(node.position.x),
    top: pxToPt(node.position.y),
    width: pxToPt(node.size.w),
    height: pxToPt(node.size.h),
  };

  if (node.isVideo) {
    return renderVideo(node, ctx, order, box);
  }

  if (node.isAudio) {
    return renderAudio(node, ctx, order, box);
  }

  return renderImage(node, ctx, order, box);
}
