/**
 * Maps StyleResolver fill output to pptxtojson Fill types (ColorFill, ImageFill, GradientFill, PatternFill).
 */

import type { SafeXmlNode } from '../parser/XmlParser';
import type { PlaceholderInfo } from '../model/nodes/BaseNode';
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

  // No explicit fill in OOXML — do not assume white (PPT treats as no fill / see-through).
  return { type: 'color', value: 'transparent' };
}

function dirnamePackagePath(path: string): string {
  if (!path) return '';
  const i = path.lastIndexOf('/');
  return i >= 0 ? path.slice(0, i) : '';
}

/** True when spPr defines how the interior is painted (including explicit no fill). */
export function spPrHasExplicitFill(spPr: SafeXmlNode): boolean {
  if (!spPr.exists()) return false;
  return (
    spPr.child('solidFill').exists() ||
    spPr.child('gradFill').exists() ||
    spPr.child('blipFill').exists() ||
    spPr.child('pattFill').exists() ||
    spPr.child('grpFill').exists() ||
    spPr.child('noFill').exists()
  );
}

function findPlaceholderNode(
  placeholders: SafeXmlNode[],
  info: PlaceholderInfo,
): SafeXmlNode | undefined {
  for (const ph of placeholders) {
    let phEl: SafeXmlNode | undefined;
    const nvSpPr = ph.child('nvSpPr');
    if (nvSpPr.exists()) {
      phEl = nvSpPr.child('nvPr').child('ph');
    }
    if (!phEl || !phEl.exists()) {
      const nvPicPr = ph.child('nvPicPr');
      if (nvPicPr.exists()) {
        phEl = nvPicPr.child('nvPr').child('ph');
      }
    }
    if (!phEl || !phEl.exists()) continue;

    const phType = phEl.attr('type');
    const phIdx = phEl.numAttr('idx');

    if (info.idx !== undefined && phIdx === info.idx) return ph;
    if (info.type && phType === info.type) return ph;
  }
  return undefined;
}

function isTransparentOnlyFill(fill: Fill): boolean {
  return fill.type === 'color' && fill.value === 'transparent';
}

/**
 * Resolve shape `spPr` fill for serialization. When the slide shape omits fill (only xfrm/geom etc.),
 * PowerPoint inherits from the matching layout then master placeholder — same order as text lstStyle.
 */
export function resolveShapeFill(
  spPr: SafeXmlNode,
  ctx: RenderContext,
  placeholder?: PlaceholderInfo,
): Fill {
  const direct: Fill = spPr.exists()
    ? spPrToFill(spPr, ctx)
    : { type: 'color', value: 'transparent' };

  if (spPr.exists() && spPrHasExplicitFill(spPr)) {
    return direct;
  }

  if (placeholder) {
    const layoutPh = findPlaceholderNode(
      ctx.layout.placeholders.map((e) => e.node),
      placeholder,
    );
    if (layoutPh) {
      const phSpPr = layoutPh.child('spPr');
      if (phSpPr.exists() && spPrHasExplicitFill(phSpPr)) {
        const fill = spPrToFill(phSpPr, ctx, {
          rels: ctx.layout.rels,
          basePath: dirnamePackagePath(ctx.layoutPath),
        });
        if (!isTransparentOnlyFill(fill)) return fill;
      }
    }

    const masterPh = findPlaceholderNode(ctx.master.placeholders, placeholder);
    if (masterPh) {
      const phSpPr = masterPh.child('spPr');
      if (phSpPr.exists() && spPrHasExplicitFill(phSpPr)) {
        const fill = spPrToFill(phSpPr, ctx, {
          rels: ctx.master.rels,
          basePath: dirnamePackagePath(ctx.masterPath),
        });
        if (!isTransparentOnlyFill(fill)) return fill;
      }
    }
  }

  return direct;
}
