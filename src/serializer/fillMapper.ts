/**
 * Maps StyleResolver fill output to pptxtojson Fill types (ColorFill, ImageFill, GradientFill, PatternFill).
 */

import type { SafeXmlNode } from '../parser/XmlParser';
import type { RenderContext } from './RenderContext';
import { resolveColor, resolveGradientFill, type GradientFillData } from './StyleResolver';
import { resolveRelTarget } from '../parser/RelParser';
import type { RelEntry } from '../parser/RelParser';
import { encodeMediaForWebDisplay } from '../utils/mediaWebConvert';
import type { Fill, ColorFill, ImageFill, GradientFill, PatternFill } from '../adapter/types';

export interface SpPrToFillOptions {
  rels: Map<string, RelEntry>;
  basePath: string;
}

const PX_TO_PT = 0.75;

function ensureHex(color: string): string {
  const s = color.trim();
  if (s.startsWith('#')) return s;
  return `#${s}`;
}

/** Convert GradientFillData to types.GradientFill.value (path, rot, colors). */
export function gradientFillDataToValue(data: GradientFillData): GradientFill['value'] {
  const path: GradientFill['value']['path'] =
    data.type === 'linear'
      ? 'line'
      : data.pathType === 'rect'
        ? 'rect'
        : data.pathType === 'circle' || data.pathType === 'shape'
          ? data.pathType
          : 'circle';
  const rot = data.type === 'linear' ? data.angle : 0;
  const colors = data.stops.map((s) => ({
    pos: `${s.position.toFixed(1)}%`,
    color: ensureHex(s.color),
  }));
  return { path, rot, colors };
}

/**
 * Resolve fill from shape properties (`spPr`) to types.Fill.
 * Slide backgrounds use `backgroundSerializer.bgPrToFill` / `bgRefToFill` instead (aligned with BackgroundRenderer).
 * When resolving blip from a shape on layout/master slide, pass options with that part's rels and basePath.
 */
export function spPrToFill(
  spPr: SafeXmlNode,
  ctx: RenderContext,
  options?: SpPrToFillOptions,
): Fill {
  const solidFill = spPr.child('solidFill');
  if (solidFill.exists()) {
    const { color } = resolveColor(solidFill, ctx);
    const value = ensureHex(color);
    return { type: 'color', value };
  }

  const gradFill = spPr.child('gradFill');
  if (gradFill.exists()) {
    const data = resolveGradientFill(spPr, ctx);
    if (data) {
      return { type: 'gradient', value: gradientFillDataToValue(data) };
    }
  }

  const blipFill = spPr.child('blipFill');
  if (blipFill.exists()) {
    const blip = blipFill.child('blip');
    const embedId = blip.attr('embed') ?? blip.attr('r:embed');
    if (embedId) {
      const rels = options?.rels ?? ctx.slide.rels;
      const basePath = options?.basePath ?? ctx.slide.slidePath.replace(/\/[^/]+$/, '');
      const rel = rels.get(embedId);
      if (rel) {
        const mediaPath = resolveRelTarget(basePath, rel.target);
        const data = ctx.presentation.media.get(mediaPath);
        if (data) {
          const picBase64 = encodeMediaForWebDisplay(mediaPath, data);
          return {
            type: 'image',
            value: { picBase64, opacity: 1 },
          };
        }
      }
    }
  }

  const pattFill = spPr.child('pattFill');
  if (pattFill.exists()) {
    const preset = pattFill.attr('prst') ?? 'solid';
    let foregroundColor = '#000000';
    let backgroundColor = '#ffffff';
    const fgClr = pattFill.child('fgClr');
    if (fgClr.exists()) {
      const { color } = resolveColor(fgClr, ctx);
      foregroundColor = ensureHex(color);
    }
    const bgClr = pattFill.child('bgClr');
    if (bgClr.exists()) {
      const { color } = resolveColor(bgClr, ctx);
      backgroundColor = ensureHex(color);
    }
    return {
      type: 'pattern',
      value: { type: preset, foregroundColor, backgroundColor },
    };
  }

  const grpFill = spPr.child('grpFill');
  if (grpFill.exists() && ctx.groupFillNode) {
    return spPrToFill(ctx.groupFillNode, ctx);
  }

  const noFill = spPr.child('noFill');
  if (noFill.exists()) {
    return { type: 'color', value: 'transparent' };
  }

  return { type: 'color', value: '#ffffff' };
}
