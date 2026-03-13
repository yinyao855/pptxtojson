/**
 * Resolves slide background to types.Fill (same order as BackgroundRenderer:
 * slide.background → layout.background → master.background).
 */

import type { SafeXmlNode } from '../parser/XmlParser';
import type { RenderContext } from '../resolve/RenderContext';
import { spPrToFill, bgRefToFill } from './fillMapper';
import type { Fill } from '../adapter/types';

function defaultFill(): Fill {
  return { type: 'color', value: '#ffffff' };
}

/**
 * Resolve slide fill for the given render context.
 * Uses the first available background in: slide → layout → master.
 * For bgPr (direct fill) uses spPrToFill with the correct rels/basePath for blip.
 */
export function resolveSlideFill(ctx: RenderContext): Fill {
  let bgNode: SafeXmlNode | undefined;
  let rels = ctx.slide.rels;
  let basePath = ctx.slide.slidePath.replace(/\/[^/]+$/, '');

  if (ctx.slide.background?.exists()) {
    bgNode = ctx.slide.background;
    rels = ctx.slide.rels;
    basePath = ctx.slide.slidePath.replace(/\/[^/]+$/, '');
  } else if (ctx.layout.background?.exists()) {
    bgNode = ctx.layout.background;
    rels = ctx.layout.rels;
    const layoutPath = ctx.presentation.slideToLayout.get(ctx.slide.index) || '';
    basePath = layoutPath.replace(/\/[^/]+$/, '');
  } else if (ctx.master.background?.exists()) {
    bgNode = ctx.master.background;
    rels = ctx.master.rels;
    const layoutPath = ctx.presentation.slideToLayout.get(ctx.slide.index) || '';
    const masterPath = ctx.presentation.layoutToMaster.get(layoutPath) || '';
    basePath = masterPath.replace(/\/[^/]+$/, '');
  }

  if (!bgNode?.exists()) return defaultFill();

  const bgPr = bgNode.child('bgPr');
  if (bgPr.exists()) {
    const fill = spPrToFill(bgPr, ctx, { rels, basePath });
    if (fill.type === 'color' && fill.value === 'transparent') return defaultFill();
    return fill;
  }

  const bgRef = bgNode.child('bgRef');
  if (bgRef.exists()) {
    const fill = bgRefToFill(bgRef, ctx);
    if (fill.type === 'color' && fill.value === 'transparent') return defaultFill();
    return fill;
  }

  return defaultFill();
}
