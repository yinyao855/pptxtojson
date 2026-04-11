/**
 * Shape serializer — mirrors `pptx-renderer-main/src/renderer/ShapeRenderer.renderShape` control flow
 * and naming, but emits pptxtojson `Shape` / `Text` objects (`adapter/types.ts`) instead of DOM.
 */

import type { ShapeNodeData, LineEndInfo, TextBody } from '../model/nodes/ShapeNode';
import type { PlaceholderInfo } from '../model/nodes/BaseNode';
import type { RenderContext } from './RenderContext';
import {
  resolveFill,
  resolveLineStyle,
  resolveGradientStroke,
  resolveGradientFill,
  resolveColorToCss,
  resolveColor,
  resolveThemeFillReference,
  type GradientFillData,
} from './StyleResolver';
import { renderTextBody, findPlaceholderNode, type RenderTextBodyOptions } from './textSerializer';
import { renderCustomGeometry } from '../shapes/customGeometry';
import { getPresetShapePath, getMultiPathPreset, type PresetSubPath } from '../shapes/presets';
import { emuToPt } from '../parser/units';
import { hexToRgb, rgbToHex } from '../utils/color';
import { SafeXmlNode } from '../parser/XmlParser';
import { resolveRelTarget } from '../parser/RelParser';
import { resolveMediaToUrl } from '../utils/mediaWebConvert';
import { isAllowedExternalUrl } from '../utils/urlSafety';
import { lineStyleToBorder, type BorderResult } from './borderMapper';
import type { AutoFit, Fill, GradientFill, ImageFill, Shadow, Shape, Text } from '../adapter/types';

// ---------------------------------------------------------------------------
// Units (shape positions/sizes are in px in node; JSON uses pt)
// ---------------------------------------------------------------------------

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return Number((px * PX_TO_PT).toFixed(4));
}

// ---------------------------------------------------------------------------
// Shape blipFill (image fill) — resolve to base64 for JSON (renderer uses blob URL)
// ---------------------------------------------------------------------------

/** Resolve blip opacity from alphaModFix / alphaMod / alphaOff modifiers. */
function resolveBlipOpacity(blip: SafeXmlNode): number {
  let alpha = 1;
  const alphaModFix = blip.child('alphaModFix');
  if (alphaModFix.exists()) {
    alpha *= (alphaModFix.numAttr('amt') ?? 100000) / 100000;
  }
  const alphaMod = blip.child('alphaMod');
  if (alphaMod.exists()) {
    alpha *= (alphaMod.numAttr('val') ?? 100000) / 100000;
  }
  const alphaOff = blip.child('alphaOff');
  if (alphaOff.exists()) {
    alpha += (alphaOff.numAttr('val') ?? 0) / 100000;
  }
  return Math.max(0, Math.min(1, alpha));
}

/** Resolve shape blipFill to embedded image data + opacity for JSON. */
async function resolveShapeBlipFill(
  blipFill: SafeXmlNode,
  ctx: RenderContext,
): Promise<{ url: string; opacity: number } | null> {
  const blip = blipFill.child('blip');
  const embedId = blip.attr('embed') ?? blip.attr('r:embed');
  if (!embedId) return null;
  const rel = ctx.slide.rels.get(embedId);
  if (!rel) return null;
  const basePath = ctx.slide.slidePath.replace(/\/[^/]+$/, '');
  const mediaPath = resolveRelTarget(basePath, rel.target);
  const data = ctx.presentation.media.get(mediaPath);
  if (!data) return null;
  const url = await resolveMediaToUrl(mediaPath, data, ctx.mediaMode, ctx.mediaUrlCache);
  if (!url) return null;
  const opacity = resolveBlipOpacity(blip);
  return { url, opacity };
}

// ---------------------------------------------------------------------------
// Line End Marker (Arrowhead) Helpers — same as ShapeRenderer (JSON does not emit SVG markers)
// ---------------------------------------------------------------------------

/** True if the text body has at least one non-empty run (avoids covering shapes with empty placeholder text). */
function hasVisibleText(textBody: TextBody): boolean {
  for (const p of textBody.paragraphs) {
    for (const r of p.runs) {
      if (r.text != null && r.text.trim().length > 0) return true;
    }
  }
  return false;
}

function svgDashArrayForKind(dashKind: string, strokeWidth: number): string | null {
  const w = Math.max(strokeWidth, 1);
  switch (dashKind) {
    case 'dot':
    case 'sysDot':
      return `${w},${w * 2}`;
    case 'dash':
    case 'sysDash':
      return `${w * 4},${w * 2}`;
    case 'lgDash':
      return `${w * 8},${w * 3}`;
    case 'dashDot':
    case 'sysDashDot':
      return `${w * 4},${w * 2},${w},${w * 2}`;
    case 'lgDashDot':
      return `${w * 8},${w * 3},${w},${w * 3}`;
    case 'lgDashDotDot':
    case 'sysDashDotDot':
      return `${w * 8},${w * 3},${w},${w * 2},${w},${w * 2}`;
    default:
      return null;
  }
}

function parseCssColorToRgba(
  color: string,
): { r: number; g: number; b: number; a?: number } | null {
  if (!color) return null;
  const hex = color.trim();
  if (hex.startsWith('#')) {
    return hexToRgb(hex);
  }
  const m = hex.match(/rgba?\(([^)]+)\)/i);
  if (!m) return null;
  const parts = m[1].split(',').map((s) => Number.parseFloat(s.trim()));
  if (parts.length < 3 || parts.some((v) => Number.isNaN(v))) return null;
  const result: { r: number; g: number; b: number; a?: number } = {
    r: Math.max(0, Math.min(255, parts[0])),
    g: Math.max(0, Math.min(255, parts[1])),
    b: Math.max(0, Math.min(255, parts[2])),
  };
  if (parts.length >= 4 && parts[3] < 1) {
    result.a = parts[3];
  }
  return result;
}

/** Read headEnd/tailEnd from an OOXML a:ln node (e.g. theme line style). */
function getLineEndsFromLn(ln: SafeXmlNode): { headEnd?: LineEndInfo; tailEnd?: LineEndInfo } {
  const out: { headEnd?: LineEndInfo; tailEnd?: LineEndInfo } = {};
  const he = ln.child('headEnd');
  if (he.exists()) {
    const t = he.attr('type');
    if (t && t !== 'none') out.headEnd = { type: t, w: he.attr('w'), len: he.attr('len') };
  }
  const te = ln.child('tailEnd');
  if (te.exists()) {
    const t = te.attr('type');
    if (t && t !== 'none') out.tailEnd = { type: t, w: te.attr('w'), len: te.attr('len') };
  }
  return out;
}

// ---------------------------------------------------------------------------
// Fill → adapter/types.Fill (no fillMapper — aligned with StyleResolver + ShapeRenderer)
// ---------------------------------------------------------------------------

function ensureHex(color: string): string {
  const s = color.trim();
  if (s.startsWith('#')) return s;
  return `#${s}`;
}

/** Convert structured gradient data to pptxtojson `GradientFill.value` (same mapping as former fillMapper). */
function gradientFillDataToValue(data: GradientFillData): GradientFill['value'] {
  const path: GradientFill['value']['path'] =
    data.type === 'linear'
      ? 'line'
      : data.pathType === 'rect'
        ? 'rect'
        : data.pathType === 'circle' || data.pathType === 'shape'
          ? (data.pathType as 'circle' | 'shape')
          : 'circle';
  const rot = data.type === 'linear' ? data.angle : 0;
  const colors = data.stops.map((s) => ({
    pos: `${s.position.toFixed(1)}%`,
    color: cssColorToFillHex(s.color),
  }));
  return { path, rot, colors };
}

function cssColorToFillHex(css: string): string {
  const s = css.trim();
  if (s === 'transparent' || s === 'none') return 'transparent';
  if (s.startsWith('#')) {
    if (s.length === 4) {
      const r = s[1];
      const g = s[2];
      const b = s[3];
      return `#${r}${r}${g}${g}${b}${b}`;
    }
    return s;
  }
  const rgba = parseCssColorToRgba(s);
  if (rgba) {
    const hex = rgbToHex(rgba.r, rgba.g, rgba.b);
    if (rgba.a !== undefined) {
      const alphaHex = Math.round(rgba.a * 255)
        .toString(16)
        .padStart(2, '0');
      return `${hex}${alphaHex}`;
    }
    return hex;
  }
  return '#000000';
}

function patternFillToFill(pattFill: SafeXmlNode, ctx: RenderContext): Fill {
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

/**
 * Build `Fill` after the same fillCss / gradientFillData / line-like rules as ShapeRenderer.
 */
async function fillToJson(
  spPr: SafeXmlNode,
  ctx: RenderContext,
  fillCss: string,
  gradientFillData: GradientFillData | null,
  isLineLike: boolean,
): Promise<Fill> {
  if (isLineLike) {
    return { type: 'color', value: 'transparent' };
  }

  const blipFill = spPr.child('blipFill');
  if (blipFill.exists()) {
    const blipResult = await resolveShapeBlipFill(blipFill, ctx);
    if (blipResult) {
      const imageFill: ImageFill = {
        type: 'image',
        value: { picBase64: blipResult.url, opacity: blipResult.opacity },
      };
      return imageFill;
    }
  }

  if (gradientFillData && gradientFillData.stops.length > 0) {
    return { type: 'gradient', value: gradientFillDataToValue(gradientFillData) };
  }

  if (fillCss && fillCss !== 'transparent' && fillCss !== 'none') {
    if (!fillCss.includes('gradient')) {
      return { type: 'color', value: cssColorToFillHex(fillCss) };
    }
    const again = resolveGradientFill(spPr, ctx);
    if (again && again.stops.length > 0) {
      return { type: 'gradient', value: gradientFillDataToValue(again) };
    }
  }

  const pattFill = spPr.child('pattFill');
  if (pattFill.exists()) {
    return patternFillToFill(pattFill, ctx);
  }

  const solidFill = spPr.child('solidFill');
  if (solidFill.exists()) {
    const { color } = resolveColor(solidFill, ctx);
    return { type: 'color', value: ensureHex(color) };
  }

  const grpFill = spPr.child('grpFill');
  if (grpFill.exists() && ctx.groupFillNode) {
    const g = ctx.groupFillNode;
    const innerCss = resolveFill(g, ctx);
    const innerGrad = resolveGradientFill(g, ctx);
    return fillToJson(g, ctx, innerCss, innerGrad, false);
  }

  const noFill = spPr.child('noFill');
  if (noFill.exists()) {
    return { type: 'color', value: 'transparent' };
  }

  return { type: 'color', value: 'transparent' };
}

// ---------------------------------------------------------------------------
// Shadow + link (mirror ShapeRenderer tail sections)
// ---------------------------------------------------------------------------

function resolveShapeShadow(node: ShapeNodeData, spPr: SafeXmlNode, ctx: RenderContext): Shadow | undefined {
  let effectiveEffectLst = spPr.child('effectLst');
  if (!effectiveEffectLst.exists()) {
    const effectRef = node.source.child('style').child('effectRef');
    const idx = effectRef.numAttr('idx') ?? 0;
    if (idx > 0 && (ctx.theme.effectStyles?.length ?? 0) >= idx) {
      const themeEffect = ctx.theme.effectStyles[idx - 1];
      if (themeEffect.exists()) {
        const lst = themeEffect.child('effectLst');
        if (lst.exists()) effectiveEffectLst = lst;
      }
    }
  }

  if (!effectiveEffectLst.exists()) return undefined;
  const outerShdw = effectiveEffectLst.child('outerShdw');
  if (!outerShdw.exists()) return undefined;

  const dir = outerShdw.numAttr('dir') ?? 0;
  const dist = outerShdw.numAttr('dist') ?? 0;
  const blurRad = outerShdw.numAttr('blurRad') ?? 0;
  const dirDeg = dir / 60000;
  const distPt = emuToPt(dist);
  const blurPt = emuToPt(blurRad);
  const h = distPt * Math.cos((dirDeg * Math.PI) / 180);
  const v = distPt * Math.sin((dirDeg * Math.PI) / 180);

  let color = 'rgba(0,0,0,0.4)';
  const { color: shdColor, alpha: shdAlpha } = resolveColor(outerShdw, ctx);
  if (shdColor) {
    const hex = shdColor.startsWith('#') ? shdColor : `#${shdColor}`;
    const { r: sr, g: sg, b: sb } = hexToRgb(hex);
    color = `rgba(${sr},${sg},${sb},${shdAlpha.toFixed(3)})`;
  }

  return { h, v, blur: blurPt, color };
}

function resolveShapeLink(node: ShapeNodeData, ctx: RenderContext): string | undefined {
  const h = node.hlinkClick;
  if (!h) return undefined;
  const { action, rId } = h;
  if (action === 'ppaction://hlinksldjump' && rId) {
    const rel = ctx.slide.rels.get(rId);
    if (rel) {
      const match = rel.target.match(/slide(\d+)\.xml/i);
      if (match) return `#slide-${match[1]}`;
    }
  } else if (rId) {
    const rel = ctx.slide.rels.get(rId);
    if (rel && rel.targetMode === 'External' && isAllowedExternalUrl(rel.target)) {
      return rel.target;
    }
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Text / bodyPr helpers (align with ShapeRenderer text overlay)
// ---------------------------------------------------------------------------

function getVAlignFromBodyPr(
  bodyPr: SafeXmlNode | undefined,
  fallbackBp: SafeXmlNode | undefined,
): string {
  const anchor = (bodyPr ? bodyPr.attr('anchor') : null) || (fallbackBp ? fallbackBp.attr('anchor') : null);
  if (anchor === 'ctr' || anchor === 'mid' || anchor === 'middle') return 'mid';
  if (anchor === 'b' || anchor === 'bottom') return 'down';
  return 'up';
}

function getIsVertical(bodyPr: SafeXmlNode | undefined, fallbackBp: SafeXmlNode | undefined): boolean {
  const vert = (bodyPr ? bodyPr.attr('vert') : null) || (fallbackBp ? fallbackBp.attr('vert') : null);
  return vert === 'eaVert' || vert === 'vert' || vert === 'wordArtVert' || vert === 'vert270';
}

function computeAutoFit(textBody: TextBody | undefined): AutoFit | undefined {
  if (!textBody?.bodyProperties) return undefined;
  const bp = textBody.bodyProperties;
  if (bp.child('spAutoFit').exists()) {
    return { type: 'shape' };
  }
  const norm = bp.child('normAutofit');
  if (norm.exists()) {
    const fs = norm.numAttr('fontScale');
    if (fs !== undefined) {
      return { type: 'text', fontScale: fs / 1000 };
    }
    return { type: 'text' };
  }
  return undefined;
}

/**
 * Resolve fill from layout or master placeholder when the slide shape has no explicit fill.
 * Checks if the shape's spPr explicitly declares <a:noFill/> — if so, returns null (transparent).
 * Otherwise falls back to layout placeholder spPr, then master placeholder spPr.
 */
function resolveInheritedPlaceholderFill(
  placeholder: PlaceholderInfo,
  slideSpPr: SafeXmlNode,
  ctx: RenderContext,
): { fillCss: string; gradientFillData: GradientFillData | null } | null {
  // If shape explicitly opted out of fill, respect it
  if (slideSpPr.child('noFill').exists()) return null;

  // Try layout placeholder first, then master placeholder
  const sources: SafeXmlNode[] = [];
  const layoutPh = findPlaceholderNode(
    ctx.layout.placeholders.map((e) => e.node),
    placeholder,
  );
  if (layoutPh) sources.push(layoutPh);
  const masterPh = findPlaceholderNode(ctx.master.placeholders, placeholder);
  if (masterPh) sources.push(masterPh);

  for (const phNode of sources) {
    const phSpPr = phNode.child('spPr');
    if (!phSpPr.exists()) continue;

    // Check solidFill
    const solidFill = phSpPr.child('solidFill');
    if (solidFill.exists()) {
      const colorChild = solidFill.child('srgbClr').exists()
        ? solidFill.child('srgbClr')
        : solidFill.child('schemeClr').exists()
          ? solidFill.child('schemeClr')
          : solidFill.child('scrgbClr').exists()
            ? solidFill.child('scrgbClr')
            : solidFill.child('sysClr').exists()
              ? solidFill.child('sysClr')
              : undefined;
      if (colorChild?.exists()) {
        return { fillCss: resolveColorToCss(colorChild, ctx), gradientFillData: null };
      }
    }

    // Check gradFill
    const gradFill = phSpPr.child('gradFill');
    if (gradFill.exists()) {
      const css = resolveFill(phSpPr, ctx);
      const gradData = resolveGradientFill(phSpPr, ctx);
      if (css || gradData) return { fillCss: css, gradientFillData: gradData };
    }

    // Check blipFill (image fill)
    const blipFill = phSpPr.child('blipFill');
    if (blipFill.exists()) {
      const css = resolveFill(phSpPr, ctx);
      if (css) return { fillCss: css, gradientFillData: null };
    }

    // Check pattFill (pattern fill)
    const pattFill = phSpPr.child('pattFill');
    if (pattFill.exists()) {
      const css = resolveFill(phSpPr, ctx);
      if (css) return { fillCss: css, gradientFillData: null };
    }
  }

  return null;
}

// ---------------------------------------------------------------------------
// Shape Rendering → JSON (same structure as ShapeRenderer.renderShape)
// ---------------------------------------------------------------------------

/**
 * Serialize a shape node to pptxtojson `Shape` or `Text`.
 * Control flow and identifiers follow `renderShape` in `ShapeRenderer.ts`; output is JSON, not DOM.
 */
export async function renderShape(node: ShapeNodeData, ctx: RenderContext, _order: number): Promise<Shape | Text> {
  const order = node.xmlOrder;
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);

  const presetKey = node.presetGeometry?.toLowerCase() ?? '';
  const outlineOnlyPresets = new Set([
    'arc',
    'leftbracket',
    'rightbracket',
    'leftbrace',
    'rightbrace',
    'bracketpair',
    'bracepair',
  ]);
  const presetIsLine =
    !!presetKey &&
    (presetKey === 'line' ||
      presetKey === 'lineinv' ||
      presetKey.includes('connector') ||
      outlineOnlyPresets.has(presetKey));
  const isConnectorShape = node.source.localName === 'cxnSp';
  const flatExtent =
    (node.size.w > 0 && node.size.h === 0) || (node.size.w === 0 && node.size.h > 0);
  const isLineLike = presetIsLine || isConnectorShape || flatExtent;
  const minH = isLineLike && node.size.h === 0 ? 1 : node.size.h;
  const minW = isLineLike && node.size.w === 0 ? 1 : node.size.w;
  const width = pxToPt(minW);
  const height = pxToPt(minH);

  // Path coordinates must be in pt to match the JSON width/height (PPTist viewBox).
  const pathW = width;
  const pathH = height;

  const styleNode = node.source.child('style');
  const lnRef = styleNode.exists() ? styleNode.child('lnRef') : undefined;
  const fillRef = styleNode.exists() ? styleNode.child('fillRef') : undefined;

  // ---- Generate SVG path ----
  let pathD = '';
  let multiPaths: PresetSubPath[] | null = null;
  if (node.presetGeometry) {
    let effectivePreset = node.presetGeometry;
    if (isConnectorShape && effectivePreset === 'line') {
      effectivePreset = 'straightConnector1';
    }
    multiPaths = getMultiPathPreset(effectivePreset, pathW, pathH, node.adjustments);
    if (multiPaths) {
      pathD = multiPaths[0]?.d ?? '';
    } else {
      pathD = getPresetShapePath(effectivePreset, pathW, pathH, node.adjustments);
    }
  } else if (node.customGeometry) {
    const extNode = node.source.child('spPr').child('xfrm').child('ext');
    const sourceExtentEmu = {
      w: extNode.numAttr('cx') ?? 0,
      h: extNode.numAttr('cy') ?? 0,
    };
    pathD = renderCustomGeometry(node.customGeometry, pathW, pathH, sourceExtentEmu);
  }
  if (
    !pathD &&
    isLineLike &&
    (node.line?.exists() ||
      (lnRef?.exists() &&
        (lnRef.numAttr('idx') ?? 0) > 0 &&
        (ctx.theme.lineStyles?.length ?? 0) >= (lnRef.numAttr('idx') ?? 0)))
  ) {
    pathD = getPresetShapePath(
      isConnectorShape ? 'straightConnector1' : 'line',
      pathW,
      pathH,
      undefined,
    );
  }

  // ---- Resolve fill and line styles ----
  const spPr = node.source.child('spPr');
  let fillCss = '';
  let gradientFillData = node.fill ? resolveGradientFill(spPr, ctx) : null;
  if (node.fill && node.fill.exists()) {
    if (node.fill.localName === 'solidFill') {
      const colorChild = node.fill.child('srgbClr').exists()
        ? node.fill.child('srgbClr')
        : node.fill.child('schemeClr').exists()
          ? node.fill.child('schemeClr')
          : node.fill.child('scrgbClr').exists()
            ? node.fill.child('scrgbClr')
            : node.fill.child('sysClr').exists()
              ? node.fill.child('sysClr')
              : undefined;
      if (colorChild?.exists()) fillCss = resolveColorToCss(colorChild, ctx);
    }
    if (!fillCss) fillCss = resolveFill(spPr, ctx);
  }
  if (!fillCss) {
    const solidFill = spPr.child('solidFill');
    if (solidFill.exists()) {
      const colorChild = solidFill.child('srgbClr').exists()
        ? solidFill.child('srgbClr')
        : solidFill.child('schemeClr').exists()
          ? solidFill.child('schemeClr')
          : solidFill.child('scrgbClr').exists()
            ? solidFill.child('scrgbClr')
            : solidFill.child('sysClr').exists()
              ? solidFill.child('sysClr')
              : undefined;
      if (colorChild?.exists()) fillCss = resolveColorToCss(colorChild, ctx);
    }
  }
  if (!fillCss && fillRef && fillRef.exists()) {
    const resolvedThemeFill = resolveThemeFillReference(fillRef, ctx);
    fillCss = resolvedThemeFill.fillCss;
    if (!gradientFillData) gradientFillData = resolvedThemeFill.gradientFillData;
  }
  // Placeholder fill inheritance: when the slide shape has no explicit fill,
  // inherit fill from the matching layout placeholder, then master placeholder.
  if (!fillCss && !gradientFillData && node.placeholder) {
    const phFill = resolveInheritedPlaceholderFill(node.placeholder, spPr, ctx);
    if (phFill) {
      fillCss = phFill.fillCss;
      gradientFillData = phFill.gradientFillData;
    }
  }
  if (isLineLike) {
    fillCss = '';
    gradientFillData = null;
  }

  let strokeColor = 'none';
  let strokeWidth = 0;
  let strokeDash = '';
  let strokeDashKind = 'solid';
  let gradientStroke: ReturnType<typeof resolveGradientStroke> = null;

  const lineIsNoFill = node.line && node.line.child('noFill').exists();
  const hasExplicitLine = node.line && !lineIsNoFill;
  const themeLineFromLnRef =
    !hasExplicitLine &&
    !lineIsNoFill &&
    lnRef?.exists() &&
    (lnRef.numAttr('idx') ?? 0) > 0 &&
    (ctx.theme.lineStyles?.length ?? 0) >= (lnRef.numAttr('idx') ?? 0)
      ? ctx.theme.lineStyles![(lnRef.numAttr('idx') ?? 1) - 1]
      : undefined;
  let effectiveLine = hasExplicitLine ? node.line! : themeLineFromLnRef;
  if (lineIsNoFill) effectiveLine = undefined;

  if (effectiveLine?.exists()) {
    gradientStroke = resolveGradientStroke(effectiveLine, ctx);
    if (!gradientStroke) {
      const lineStyle = resolveLineStyle(effectiveLine, ctx, lnRef);
      strokeColor = lineStyle.color;
      strokeWidth = lineStyle.width;
      strokeDash = lineStyle.dash;
      strokeDashKind = lineStyle.dashKind;
    }

    // Line cap / join: ShapeRenderer maps a:ln@cap and a:ln/* to SVG stroke-linecap / stroke-linejoin.
    // pptxtojson border fields omit cap/join (see adapter/types Border).
  }
  if (lineIsNoFill) {
    strokeColor = 'none';
    strokeWidth = 0;
    gradientStroke = null;
  }

  const isCircularArrow = node.presetGeometry?.toLowerCase() === 'circulararrow';
  if (isCircularArrow) {
    strokeColor = 'none';
    strokeWidth = 0;
    gradientStroke = null;
    if (!fillCss) {
      const solid = spPr.child('solidFill');
      if (solid.exists()) {
        const color = solid.child('srgbClr').exists()
          ? solid.child('srgbClr')
          : solid.child('schemeClr').exists()
            ? solid.child('schemeClr')
            : solid.child('scrgbClr').exists()
              ? solid.child('scrgbClr')
              : solid.child('sysClr').exists()
                ? solid.child('sysClr')
                : undefined;
        if (color?.exists()) fillCss = resolveColorToCss(color, ctx);
      }
    }
  }

  let effectiveHeadEnd = node.headEnd;
  let effectiveTailEnd = node.tailEnd;
  if ((!effectiveHeadEnd || !effectiveTailEnd) && effectiveLine?.exists()) {
    const fromLn = getLineEndsFromLn(effectiveLine);
    if (!effectiveHeadEnd && fromLn.headEnd) effectiveHeadEnd = fromLn.headEnd;
    if (!effectiveTailEnd && fromLn.tailEnd) effectiveTailEnd = fromLn.tailEnd;
  }

  let effectiveStrokeWidth = gradientStroke ? gradientStroke.width : strokeWidth;
  if (isLineLike && (effectiveHeadEnd || effectiveTailEnd) && effectiveStrokeWidth <= 0) {
    effectiveStrokeWidth = 1;
  }

  const mainPathStrokeSuppressed = multiPaths && multiPaths[0]?.stroke === false;

  // ---- Border JSON (stroke → pptxtojson border fields) ----
  let borderResult: BorderResult;
  if (isCircularArrow || lineIsNoFill || !effectiveLine?.exists()) {
    borderResult = {
      border: { borderColor: '#000000', borderWidth: 0, borderType: 'solid' },
      borderStrokeDasharray: '0',
    };
  } else if (
    !mainPathStrokeSuppressed &&
    gradientStroke &&
    gradientStroke.stops.length > 0
  ) {
    const c0 = gradientStroke.stops[0]?.color || '#000000';
    borderResult = {
      border: {
        borderColor: cssColorToFillHex(c0),
        borderWidth: pxToPt(Math.max(gradientStroke.width, 1)),
        borderType: 'solid',
      },
      borderStrokeDasharray: '0',
    };
  } else if (!mainPathStrokeSuppressed && effectiveStrokeWidth > 0 && strokeColor !== 'transparent') {
    const lnNode = effectiveLine!;
    const br = lineStyleToBorder(lnNode, ctx, lnRef);
    const widthPx = effectiveStrokeWidth;
    const svgDashArray = svgDashArrayForKind(strokeDashKind, widthPx);
    let dashStr = br.borderStrokeDasharray || '';
    if (svgDashArray) {
      const parts = svgDashArray.split(',').map((x) => pxToPt(Number.parseFloat(x.trim())));
      dashStr = parts.map((x) => x.toFixed(2)).join(',');
    } else if (strokeDash === 'dashed') {
      dashStr = `${pxToPt(widthPx * 4).toFixed(2)},${pxToPt(widthPx * 2).toFixed(2)}`;
    } else if (strokeDash === 'dotted') {
      dashStr = `${pxToPt(widthPx).toFixed(2)},${pxToPt(widthPx * 2).toFixed(2)}`;
    }
    borderResult = {
      border: {
        ...br.border,
        borderWidth: pxToPt(widthPx),
      },
      borderStrokeDasharray: dashStr,
    };
  } else {
    borderResult = {
      border: { borderColor: '#000000', borderWidth: 0, borderType: 'solid' },
      borderStrokeDasharray: '0',
    };
  }

  const fillJson = await fillToJson(spPr, ctx, fillCss, gradientFillData, isLineLike);

  const shadowJson = resolveShapeShadow(node, spPr, ctx);
  const linkStr = resolveShapeLink(node, ctx);

  const placeholder = node.placeholder;
  const content = node.textBody
    ? renderTextBody(node.textBody, placeholder, ctx, textBodyRenderOptions(node, ctx))
    : '';
  const hasContent = node.textBody ? hasVisibleText(node.textBody) : false;

  const bodyPr = node.textBody?.bodyProperties;
  const fallbackBp = node.textBody?.layoutBodyProperties;
  const vAlign = getVAlignFromBodyPr(bodyPr, fallbackBp);
  const isVertical = getIsVertical(bodyPr, fallbackBp);
  const autoFit = computeAutoFit(node.textBody);

  const shapType =
    isConnectorShape && node.presetGeometry === 'line'
      ? 'straightConnector1'
      : node.presetGeometry || (node.customGeometry ? 'custom' : 'rect');

  // Icon overlay in ShapeRenderer is a separate SVG <path>; JSON carries main geometry only (`pathD`).
  const pathOut: string | undefined = pathD || undefined;

  // PPTist expects keypoints normalized by /50000 (OOXML raw value / 50000).
  let keypoints: Record<string, number> | undefined;
  if (node.adjustments.size > 0) {
    keypoints = {};
    for (const [k, v] of node.adjustments) {
      keypoints[k] = v / 50000;
    }
  }

  const baseCommon = {
    left,
    top,
    width,
    height,
    name: node.name || '',
    order,
    borderColor: borderResult.border.borderColor,
    borderWidth: borderResult.border.borderWidth,
    borderType: borderResult.border.borderType,
    borderStrokeDasharray: borderResult.borderStrokeDasharray || '0',
    fill: fillJson,
    isFlipV: node.flipV,
    isFlipH: node.flipH,
    rotate: node.rotation,
    content: content || '',
    ...(shadowJson ? { shadow: shadowJson } : {}),
    ...(linkStr ? { link: linkStr } : {}),
    ...(autoFit ? { autoFit } : {}),
  };

  // --- Shape vs Text type detection ---
  // Mirrors src1/pptxtojson.js logic:
  // 1. cNvSpPr@txBox="1" → text box → output as "text"
  // 2. Placeholder body/title/ctrTitle/subTitle → output as "text"
  // 3. Custom geometry (non-diagram) → shape
  // 4. Preset geometry != 'rect', or type is 'obj'/undefined → shape
  // 5. No visible text but has fill/border → shape
  // 6. Fallthrough → text
  const isTxBox = node.source.child('nvSpPr').child('cNvSpPr').attr('txBox') === '1';
  const isPlaceholderText =
    placeholder?.type === 'body' ||
    placeholder?.type === 'title' ||
    placeholder?.type === 'ctrTitle' ||
    placeholder?.type === 'subTitle';
  const hasCustomGeom = !!node.customGeometry;
  const hasPresetGeom = !!node.presetGeometry;
  const isNonRectPreset = hasPresetGeom && shapType !== 'rect';
  const noPlaceholderType = !placeholder?.type;
  const hasFillOrBorder = !!fillCss || borderResult.border.borderWidth > 0;

  let outputAsText = false;
  if (isTxBox || isPlaceholderText) {
    outputAsText = true;
  } else if (hasCustomGeom) {
    outputAsText = false;
  } else if (isNonRectPreset || noPlaceholderType) {
    if (hasPresetGeom && !hasContent && hasFillOrBorder) {
      outputAsText = false;
    } else if (hasPresetGeom && isNonRectPreset) {
      outputAsText = false;
    } else if (!hasPresetGeom && !hasContent) {
      outputAsText = false;
    } else {
      outputAsText = true;
    }
  } else {
    outputAsText = true;
  }

  if (outputAsText) {
    const textEl: Text = {
      ...baseCommon,
      type: 'text',
      isVertical,
      vAlign,
    };
    return textEl;
  }

  const shapeEl: Shape = {
    ...baseCommon,
    type: 'shape',
    shapType,
    vAlign,
    path: pathOut,
    ...(keypoints ? { keypoints } : {}),
  };
  return shapeEl;
}

function textBodyRenderOptions(
  node: ShapeNodeData,
  ctx: RenderContext,
): RenderTextBodyOptions | undefined {
  const shapeStyle = node.source.child('style');
  if (!shapeStyle.exists()) return undefined;
  const fontRef = shapeStyle.child('fontRef');
  if (fontRef.exists() && fontRef.allChildren().length > 0) {
    return { fontRefColor: resolveColorToCss(fontRef, ctx) };
  }
  return undefined;
}

/** @deprecated Use `renderShape` — same name as `ShapeRenderer` for diff-friendly comparison. */
export async function shapeToElement(node: ShapeNodeData, ctx: RenderContext, order: number): Promise<Shape | Text> {
  return renderShape(node, ctx, order);
}
