/**
 * Background serializer — resolves and applies slide/layout/master backgrounds.
 */

import type { SafeXmlNode } from '../parser/XmlParser';
import type { RenderContext } from './RenderContext';
import { resolveColor, resolveGradientFill } from './StyleResolver';
import { hexToRgb, rgbToHex } from '../utils/color';
import type { RelEntry } from '../parser/RelParser';
import { encodeMediaForWebDisplay } from '../utils/mediaWebConvert';
import { gradientFillDataToValue } from './fillMapper';
import type { Fill } from '../adapter/types';
import { resolveMediaPath } from '../utils/media';

function defaultFill(): Fill {
  return { type: 'color', value: '#ffffff' };
}

/** Same formula as BackgroundRenderer.compositeOnWhite → opaque hex for JSON. */
function compositeOnWhiteToHex(r: number, g: number, b: number, a: number): string {
  const cr = Math.round(r * a + 255 * (1 - a));
  const cg = Math.round(g * a + 255 * (1 - a));
  const cb = Math.round(b * a + 255 * (1 - a));
  return rgbToHex(cr, cg, cb);
}

/**
 * Resolve slide fill
 * 
 * Background priority: slide.background -> layout.background -> master.background.
 * The first found background is used.
 */
export function resolveSlideFill(ctx: RenderContext): Fill {
  // Find the first available background in the inheritance chain,
  // and track which rels map to use for resolving image references
  let bgNode: SafeXmlNode | undefined;
  let bgRels: Map<string, RelEntry> = ctx.slide.rels;

  if (ctx.slide.background?.exists()) {
    bgNode = ctx.slide.background;
    bgRels = ctx.slide.rels;
  } else if (ctx.layout.background?.exists()) {
    bgNode = ctx.layout.background;
    bgRels = ctx.layout.rels;
  } else if (ctx.master.background?.exists()) {
    bgNode = ctx.master.background;
    bgRels = ctx.master.rels;
  }

  if (!bgNode?.exists()) return defaultFill();

  // Parse p:bg > p:bgPr
  const bgPr = bgNode.child('bgPr');
  if (bgPr.exists()) {
    return renderBgPr(bgPr, ctx, bgRels);
  }

  // Parse p:bg > p:bgRef (theme reference)
  const bgRef = bgNode.child('bgRef');
  if (bgRef.exists()) {
    return renderBgRef(bgRef, ctx);
  }

  return defaultFill();
}

/**
 * Render background from bgPr (background properties).
 * Contains direct fill definitions: solidFill, gradFill, blipFill, etc.
 */
export function renderBgPr(bgPr: SafeXmlNode, ctx: RenderContext, rels?: Map<string, RelEntry>): Fill {
  // solidFill
  const solidFill = bgPr.child('solidFill');
  if (solidFill.exists()) {
    const { color, alpha } = resolveColor(solidFill, ctx);
    const hex = color.startsWith('#') ? color : `#${color}`;
    if (alpha < 1) {
      const { r, g, b } = hexToRgb(hex);
      return { type: 'color', value: compositeOnWhiteToHex(r, g, b, alpha) };
    } else {
      return { type: 'color', value: hex };
    }
  }

  // gradFill
  const gradFill = bgPr.child('gradFill');
  if (gradFill.exists()) {
    const data = resolveGradientFill(bgPr, ctx);
    if (data) {
      return { type: 'gradient', value: gradientFillDataToValue(data) };
    }
    return defaultFill();
  }

  // blipFill (image background)
  const blipFill = bgPr.child('blipFill');
  if (blipFill.exists()) {
    const blip = blipFill.child('blip');
    const embedId = blip.attr('embed') ?? blip.attr('r:embed');
    if (embedId) {
      const relsMap = rels ?? ctx.slide.rels;
      const rel = relsMap.get(embedId);
      if (rel) {
        const mediaPath = resolveMediaPath(rel.target);
        const data = ctx.presentation.media.get(mediaPath);
        if (data) {
          const picBase64 = encodeMediaForWebDisplay(mediaPath, data);
          // TODO: get opacity from alphaModFix
          return {
            type: 'image',
            value: { picBase64, opacity: 1 },
          };
        }
      }
    }
    return defaultFill();
  }

  const noFill = bgPr.child('noFill');
  if (noFill.exists()) {
    return defaultFill();
  }

  return defaultFill();
}

/**
 * Render background from bgRef (theme format scheme reference).
 * Simplified: just resolve the color from the reference.
 */
export function renderBgRef(bgRef: SafeXmlNode, ctx: RenderContext): Fill {
  // bgRef may contain a color child (schemeClr, srgbClr, etc.)
  const { color, alpha } = resolveColor(bgRef, ctx);
  if (color && color !== '#000000') {
    const hex = color.startsWith('#') ? color : `#${color}`;
    if (alpha < 1) {
      const { r, g, b } = hexToRgb(hex);
      return { type: 'color', value: compositeOnWhiteToHex(r, g, b, alpha) };
    }
    return { type: 'color', value: hex };
  }
  return defaultFill();
}
