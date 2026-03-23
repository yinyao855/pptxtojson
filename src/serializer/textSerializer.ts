/**
 * Text serializer — maps TextBody to HTML for pptxtojson `content` fields.
 * Mirrors pptx-renderer `TextRenderer` (7-level style inheritance); emits HTML string instead of DOM.
 */

import type { RenderContext } from './RenderContext';
import type { TextBody, TextParagraph, TextRun } from '../model/nodes/ShapeNode';
import type { PlaceholderInfo } from '../model/nodes/BaseNode';
import { SafeXmlNode } from '../parser/XmlParser';
import { resolveColor, resolveColorToCss } from './StyleResolver';
import { emuToPx, pctToDecimal, angleToDeg } from '../parser/units';

// ---------------------------------------------------------------------------
// URL safety (pptx-renderer uses utils/urlSafety; main package inlines minimal check)
// ---------------------------------------------------------------------------

function isAllowedExternalUrl(url: string): boolean {
  try {
    const u = new URL(url);
    return u.protocol === 'http:' || u.protocol === 'https:';
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// Style inheritance (aligned with TextRenderer)
// ---------------------------------------------------------------------------

function findStyleAtLevel(styleNode: SafeXmlNode | undefined, level: number): SafeXmlNode {
  if (!styleNode || !styleNode.exists()) {
    return new SafeXmlNode(null);
  }
  const lvlNode = styleNode.child(`lvl${level + 1}pPr`);
  if (lvlNode.exists()) return lvlNode;
  return styleNode.child('defPPr');
}

function getPlaceholderCategory(
  placeholder: PlaceholderInfo | undefined,
): 'title' | 'body' | 'other' {
  if (!placeholder || !placeholder.type) return 'other';
  const t = placeholder.type;
  if (t === 'title' || t === 'ctrTitle') return 'title';
  if (
    t === 'body' ||
    t === 'subTitle' ||
    t === 'obj' ||
    t === 'dt' ||
    t === 'ftr' ||
    t === 'sldNum'
  ) {
    return 'body';
  }
  return 'other';
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

function getPlaceholderLstStyle(phNode: SafeXmlNode): SafeXmlNode | undefined {
  const txBody = phNode.child('txBody');
  if (!txBody.exists()) return undefined;
  const lstStyle = txBody.child('lstStyle');
  return lstStyle.exists() ? lstStyle : undefined;
}

interface MergedParagraphStyle {
  align?: string;
  marginLeft?: number;
  textIndent?: number;
  lineHeight?: string;
  lineHeightAbsolute?: boolean;
  spaceBefore?: number;
  spaceBeforePct?: number;
  spaceAfter?: number;
  spaceAfterPct?: number;
  bulletChar?: string;
  bulletFont?: string;
  bulletAutoNum?: string;
  bulletNone?: boolean;
  bulletColorNode?: SafeXmlNode;
  defRPr?: SafeXmlNode;
}

function mergeParagraphProps(target: MergedParagraphStyle, pPr: SafeXmlNode): void {
  if (!pPr.exists()) return;

  const algn = pPr.attr('algn');
  if (algn) target.align = algn;

  const marL = pPr.numAttr('marL');
  if (marL !== undefined) target.marginLeft = emuToPx(marL);

  const indent = pPr.numAttr('indent');
  if (indent !== undefined) target.textIndent = emuToPx(indent);

  const lnSpc = pPr.child('lnSpc');
  if (lnSpc.exists()) {
    const spcPct = lnSpc.child('spcPct');
    if (spcPct.exists()) {
      const val = spcPct.numAttr('val');
      if (val !== undefined) {
        target.lineHeight = `${(val / 100000).toFixed(3)}`;
      }
    }
    const spcPts = lnSpc.child('spcPts');
    if (spcPts.exists()) {
      const val = spcPts.numAttr('val');
      if (val !== undefined) {
        target.lineHeight = `${val / 100}pt`;
        target.lineHeightAbsolute = true;
      }
    }
  }

  const spcBef = pPr.child('spcBef');
  if (spcBef.exists()) {
    const spcPts = spcBef.child('spcPts');
    if (spcPts.exists()) {
      const val = spcPts.numAttr('val');
      if (val !== undefined) target.spaceBefore = val / 100;
    }
    const spcPct = spcBef.child('spcPct');
    if (spcPct.exists()) {
      const val = spcPct.numAttr('val');
      if (val !== undefined) target.spaceBeforePct = val / 100000;
    }
  }

  const spcAft = pPr.child('spcAft');
  if (spcAft.exists()) {
    const spcPts = spcAft.child('spcPts');
    if (spcPts.exists()) {
      const val = spcPts.numAttr('val');
      if (val !== undefined) target.spaceAfter = val / 100;
    }
    const spcPct = spcAft.child('spcPct');
    if (spcPct.exists()) {
      const val = spcPct.numAttr('val');
      if (val !== undefined) target.spaceAfterPct = val / 100000;
    }
  }

  const buChar = pPr.child('buChar');
  if (buChar.exists()) {
    target.bulletChar = buChar.attr('char') || '';
    target.bulletNone = false;
  }
  const buAutoNum = pPr.child('buAutoNum');
  if (buAutoNum.exists()) {
    target.bulletAutoNum = buAutoNum.attr('type') || 'arabicPeriod';
    target.bulletNone = false;
  }
  const buNone = pPr.child('buNone');
  if (buNone.exists()) {
    target.bulletNone = true;
    target.bulletChar = undefined;
    target.bulletAutoNum = undefined;
  }
  const buFont = pPr.child('buFont');
  if (buFont.exists()) {
    target.bulletFont = buFont.attr('typeface');
  }
  const buClr = pPr.child('buClr');
  if (buClr.exists()) {
    target.bulletColorNode = buClr;
  }

  const defRPr = pPr.child('defRPr');
  if (defRPr.exists()) {
    target.defRPr = defRPr;
  }
}

interface MergedRunStyle {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  fontFamily?: string;
  hlinkClick?: string;
  letterSpacingPt?: number;
  kern?: number;
  cap?: string;
  baseline?: number;
  textGradientCss?: string;
  textNoFill?: boolean;
  textOutlineWidth?: number;
  textOutlineColor?: string;
  textOutlineGradientCss?: string;
}

function mergeRunProps(target: MergedRunStyle, rPr: SafeXmlNode, ctx: RenderContext): void {
  if (!rPr.exists()) return;

  const sz = rPr.numAttr('sz');
  if (sz !== undefined) target.fontSize = sz / 100;

  const b = rPr.attr('b');
  if (b !== undefined) target.bold = b === '1' || b === 'true';

  const i = rPr.attr('i');
  if (i !== undefined) target.italic = i === '1' || i === 'true';

  const u = rPr.attr('u');
  if (u !== undefined && u !== 'none') target.underline = true;
  if (u === 'none') target.underline = false;

  const strike = rPr.attr('strike');
  if (strike !== undefined && strike !== 'noStrike') target.strikethrough = true;
  if (strike === 'noStrike') target.strikethrough = false;

  const solidFill = rPr.child('solidFill');
  if (solidFill.exists()) {
    const { color, alpha } = resolveColor(solidFill, ctx);
    const hex = color.startsWith('#') ? color : `#${color}`;
    if (alpha < 1) {
      const { r, g, b: bl } = hexToRgbInternal(hex);
      target.color = `rgba(${r},${g},${bl},${alpha.toFixed(3)})`;
    } else {
      target.color = hex;
    }
  }
  if (!target.color) {
    const sc = rPr.child('schemeClr');
    if (sc.exists()) {
      target.color = resolveColorToCss(sc, ctx);
    }
  }
  const gradFill = rPr.child('gradFill');
  if (gradFill.exists()) {
    const css = resolveGradientForText(gradFill, ctx);
    if (css) target.textGradientCss = css;
  }

  const latin = rPr.child('latin');
  if (latin.exists()) {
    const typeface = latin.attr('typeface');
    if (typeface) {
      target.fontFamily = resolveThemeFont(typeface, ctx);
    }
  }
  if (!target.fontFamily) {
    const ea = rPr.child('ea');
    if (ea.exists()) {
      const typeface = ea.attr('typeface');
      if (typeface) {
        target.fontFamily = resolveThemeFont(typeface, ctx);
      }
    }
  }
  if (!target.fontFamily) {
    const cs = rPr.child('cs');
    if (cs.exists()) {
      const typeface = cs.attr('typeface');
      if (typeface) {
        target.fontFamily = resolveThemeFont(typeface, ctx);
      }
    }
  }

  const hlinkClick = rPr.child('hlinkClick');
  if (hlinkClick.exists()) {
    const rId = hlinkClick.attr('id') ?? hlinkClick.attr('r:id');
    if (rId) {
      const rel = ctx.slide.rels.get(rId);
      if (rel && rel.targetMode === 'External' && isAllowedExternalUrl(rel.target)) {
        target.hlinkClick = rel.target;
      }
    }
  }

  const spc = rPr.numAttr('spc');
  if (spc !== undefined) target.letterSpacingPt = spc / 100;

  const kern = rPr.numAttr('kern');
  if (kern !== undefined) target.kern = kern / 100;

  const cap = rPr.attr('cap');
  if (cap !== undefined) target.cap = cap;

  const baseline = rPr.numAttr('baseline');
  if (baseline !== undefined) target.baseline = baseline;

  if (rPr.child('noFill').exists()) {
    target.textNoFill = true;
  }

  const ln = rPr.child('ln');
  if (ln.exists() && !ln.child('noFill').exists()) {
    const lnW = ln.numAttr('w');
    target.textOutlineWidth = lnW ? emuToPx(lnW) : 0.75;
    const lnSolid = ln.child('solidFill');
    if (lnSolid.exists()) {
      const { color: c, alpha: a } = resolveColor(lnSolid, ctx);
      target.textOutlineColor = colorToCssLocal(c, a);
    }
    const lnGrad = ln.child('gradFill');
    if (lnGrad.exists()) {
      target.textOutlineGradientCss = resolveGradientForText(lnGrad, ctx);
    }
  }
}

function resolveThemeFont(typeface: string, ctx: RenderContext): string {
  if (typeface === '+mj-lt' || typeface === '+mj-ea' || typeface === '+mj-cs') {
    const key = typeface.slice(3) as 'lt' | 'ea' | 'cs';
    const mapping: Record<string, 'latin' | 'ea' | 'cs'> = { lt: 'latin', ea: 'ea', cs: 'cs' };
    return ctx.theme.majorFont[mapping[key] || 'latin'] || typeface;
  }
  if (typeface === '+mn-lt' || typeface === '+mn-ea' || typeface === '+mn-cs') {
    const key = typeface.slice(3) as 'lt' | 'ea' | 'cs';
    const mapping: Record<string, 'latin' | 'ea' | 'cs'> = { lt: 'latin', ea: 'ea', cs: 'cs' };
    return ctx.theme.minorFont[mapping[key] || 'latin'] || typeface;
  }
  return typeface;
}

function hexToRgbInternal(hex: string): { r: number; g: number; b: number } {
  const cleaned = hex.replace(/^#/, '');
  const num = parseInt(
    cleaned.length === 3
      ? cleaned[0] + cleaned[0] + cleaned[1] + cleaned[1] + cleaned[2] + cleaned[2]
      : cleaned,
    16,
  );
  return { r: (num >> 16) & 0xff, g: (num >> 8) & 0xff, b: num & 0xff };
}

function colorToCssLocal(color: string, alpha: number): string {
  const hex = color.startsWith('#') ? color : `#${color}`;
  if (alpha >= 1) return hex;
  const { r, g, b } = hexToRgbInternal(hex);
  return `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
}

function resolveGradientForText(gradFill: SafeXmlNode, ctx: RenderContext): string {
  const gsLst = gradFill.child('gsLst');
  const stops: { position: number; color: string }[] = [];
  for (const gs of gsLst.children('gs')) {
    const pos = gs.numAttr('pos') ?? 0;
    const posPercent = pctToDecimal(pos) * 100;
    const { color, alpha } = resolveColor(gs, ctx);
    stops.push({ position: posPercent, color: colorToCssLocal(color, alpha) });
  }
  if (stops.length === 0) return '';
  stops.sort((a, b) => a.position - b.position);
  const stopsStr = stops.map((s) => `${s.color} ${s.position.toFixed(1)}%`).join(', ');
  const lin = gradFill.child('lin');
  if (lin.exists()) {
    const angle = angleToDeg(lin.numAttr('ang') ?? 0);
    const cssAngle = (angle + 90) % 360;
    return `linear-gradient(${cssAngle.toFixed(1)}deg, ${stopsStr})`;
  }
  return `linear-gradient(180deg, ${stopsStr})`;
}

function generateAutoNumber(type: string, index: number): string {
  const num = index + 1;
  switch (type) {
    case 'arabicPeriod':
      return `${num}.`;
    case 'arabicParenR':
      return `${num})`;
    case 'arabicParenBoth':
      return `(${num})`;
    case 'arabicPlain':
      return `${num}`;
    case 'romanUcPeriod':
      return `${toRoman(num)}.`;
    case 'romanLcPeriod':
      return `${toRoman(num).toLowerCase()}.`;
    case 'alphaUcPeriod':
      return `${String.fromCharCode(64 + (((num - 1) % 26) + 1))}.`;
    case 'alphaLcPeriod':
      return `${String.fromCharCode(96 + (((num - 1) % 26) + 1))}.`;
    case 'alphaUcParenR':
      return `${String.fromCharCode(64 + (((num - 1) % 26) + 1))})`;
    case 'alphaLcParenR':
      return `${String.fromCharCode(96 + (((num - 1) % 26) + 1))})`;
    default:
      return `${num}.`;
  }
}

function toRoman(num: number): string {
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I'];
  let result = '';
  let remaining = num;
  for (let i = 0; i < vals.length; i++) {
    while (remaining >= vals[i]) {
      result += syms[i];
      remaining -= vals[i];
    }
  }
  return result;
}

// ---------------------------------------------------------------------------
// HTML helpers
// ---------------------------------------------------------------------------

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escapeHtmlAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
}

/** Preserve consecutive spaces (align with TextRenderer). */
function formatRunTextForHtml(raw: string): string {
  if (!raw) return '';
  if (raw.includes('\t')) {
    return escapeHtml(raw);
  }
  let t = escapeHtml(raw);
  if (/ {2}/.test(raw)) {
    t = t.replace(/ {2}/g, ' \u00a0');
  }
  return t;
}

export interface TextToHtmlOptions {
  cellTextColor?: string;
  cellTextBold?: boolean;
  cellTextItalic?: boolean;
  cellTextFontFamily?: string;
  fontRefColor?: string;
}

function paragraphOpenTag(useLineWrappers: boolean, style: string): string {
  const s = style ? ` style="${escapeHtmlAttr(style)}"` : '';
  return useLineWrappers ? `<div${s}>` : `<p${s}>`;
}

function paragraphCloseTag(useLineWrappers: boolean): string {
  return useLineWrappers ? '</div>' : '</p>';
}

function buildMergedParagraphStyle(
  paragraph: TextParagraph,
  textBody: TextBody,
  placeholder: PlaceholderInfo | undefined,
  ctx: RenderContext,
): MergedParagraphStyle {
  const level = paragraph.level;
  const category = getPlaceholderCategory(placeholder);
  const merged: MergedParagraphStyle = {};

  mergeParagraphProps(merged, findStyleAtLevel(ctx.master.defaultTextStyle, level));

  const masterTextStyle =
    category === 'title'
      ? ctx.master.textStyles.titleStyle
      : category === 'body'
        ? ctx.master.textStyles.bodyStyle
        : ctx.master.textStyles.otherStyle;
  mergeParagraphProps(merged, findStyleAtLevel(masterTextStyle, level));

  if (placeholder) {
    const masterPh = findPlaceholderNode(ctx.master.placeholders, placeholder);
    if (masterPh) {
      const lstStyle = getPlaceholderLstStyle(masterPh);
      mergeParagraphProps(merged, findStyleAtLevel(lstStyle, level));
    }
  }

  if (placeholder) {
    const layoutPh = findPlaceholderNode(
      ctx.layout.placeholders.map((e) => e.node),
      placeholder,
    );
    if (layoutPh) {
      const lstStyle = getPlaceholderLstStyle(layoutPh);
      mergeParagraphProps(merged, findStyleAtLevel(lstStyle, level));
    }
  }

  mergeParagraphProps(merged, findStyleAtLevel(textBody.listStyle, level));

  if (paragraph.properties) {
    mergeParagraphProps(merged, paragraph.properties);
  }

  return merged;
}

function mergedParagraphCss(
  merged: MergedParagraphStyle,
  effectiveLineHeight: string | undefined,
  effectiveFontSize: number,
): string {
  const parts: string[] = [];

  if (merged.align) {
    const alignMap: Record<string, string> = {
      l: 'left',
      ctr: 'center',
      r: 'right',
      just: 'justify',
      dist: 'justify',
    };
    parts.push(`text-align: ${alignMap[merged.align] || 'left'}`);
  }
  if (merged.marginLeft !== undefined) {
    parts.push(`margin-left: ${merged.marginLeft}px`);
  }
  if (merged.textIndent !== undefined) {
    parts.push(`text-indent: ${merged.textIndent}px`);
  }
  if (effectiveLineHeight) {
    parts.push(`line-height: ${effectiveLineHeight}`);
  }
  if (merged.spaceBefore !== undefined) {
    parts.push(`margin-top: ${merged.spaceBefore}pt`);
  } else if (merged.spaceBeforePct !== undefined) {
    parts.push(`margin-top: ${merged.spaceBeforePct * effectiveFontSize}pt`);
  }
  if (merged.spaceAfter !== undefined) {
    parts.push(`margin-bottom: ${merged.spaceAfter}pt`);
  } else if (merged.spaceAfterPct !== undefined) {
    parts.push(`margin-bottom: ${merged.spaceAfterPct * effectiveFontSize}pt`);
  }

  return parts.join('; ');
}

function applyLnSpcReduction(
  lineHeight: string | undefined,
  lnSpcReduction: number,
): string | undefined {
  if (!lineHeight || lnSpcReduction <= 0) return lineHeight;
  const parsed = parseFloat(lineHeight);
  if (isNaN(parsed)) return lineHeight;
  if (lineHeight.includes('pt')) {
    return `${(parsed * (1 - lnSpcReduction)).toFixed(2)}pt`;
  }
  return `${(parsed * (1 - lnSpcReduction)).toFixed(3)}`;
}

function buildRunStyleString(
  runStyle: MergedRunStyle,
  run: TextRun,
  fontScale: number,
  options: TextToHtmlOptions | undefined,
  ctx: RenderContext,
): string {
  const runProps = run.properties;
  const fontSize = runStyle.fontSize || 12;
  const effectivePt = fontSize * fontScale;

  const parts: string[] = [];
  parts.push(`font-size: ${effectivePt.toFixed(4)}pt`);

  const hasExplicitRunBold = runProps?.attr('b') !== undefined;
  if (hasExplicitRunBold ? runStyle.bold : (options?.cellTextBold ?? runStyle.bold)) {
    parts.push('font-weight: bold');
  }

  const hasExplicitRunItalic = runProps?.attr('i') !== undefined;
  if (hasExplicitRunItalic ? runStyle.italic : (options?.cellTextItalic ?? runStyle.italic)) {
    parts.push('font-style: italic');
  }

  const decorations: string[] = [];
  if (runStyle.underline) decorations.push('underline');
  if (runStyle.strikethrough) decorations.push('line-through');
  if (decorations.length > 0) {
    parts.push(`text-decoration: ${decorations.join(' ')}`);
  }

  const hasExplicitRunColor =
    runProps?.child('solidFill').exists() || runProps?.child('gradFill').exists();
  let effectiveColor: string | undefined;
  if (options?.fontRefColor) {
    effectiveColor = hasExplicitRunColor ? runStyle.color : options.fontRefColor;
  } else if (options?.cellTextColor && !hasExplicitRunColor) {
    effectiveColor = options.cellTextColor;
  } else {
    effectiveColor = runStyle.color;
  }

  if (runStyle.hlinkClick && !hasExplicitRunColor) {
    const hlinkHex = ctx.theme.colorScheme.get('hlink');
    if (hlinkHex) {
      effectiveColor = hlinkHex.startsWith('#') ? hlinkHex : `#${hlinkHex}`;
    }
  }

  if (runStyle.textGradientCss) {
    parts.push(`background: ${runStyle.textGradientCss}`);
    parts.push('-webkit-background-clip: text');
    parts.push('background-clip: text');
    parts.push('color: transparent');
  } else if (effectiveColor) {
    parts.push(`color: ${effectiveColor}`);
  } else {
    parts.push('color: #000000');
  }

  if (runStyle.textNoFill || runStyle.textOutlineWidth) {
    const strokeW = runStyle.textOutlineWidth ?? 0.75;
    if (runStyle.textNoFill && runStyle.textOutlineGradientCss) {
      parts.push('color: transparent');
      parts.push(`-webkit-text-stroke-width: ${strokeW}px`);
      parts.push('-webkit-text-stroke-color: #ffffff');
      parts.push('paint-order: stroke fill');
      parts.push(`mask-image: ${runStyle.textOutlineGradientCss}`);
      parts.push(`-webkit-mask-image: ${runStyle.textOutlineGradientCss}`);
    } else if (runStyle.textNoFill && runStyle.textOutlineColor) {
      parts.push('color: transparent');
      parts.push(`-webkit-text-stroke-width: ${strokeW}px`);
      parts.push(`-webkit-text-stroke-color: ${runStyle.textOutlineColor}`);
      parts.push('paint-order: stroke fill');
    } else if (runStyle.textNoFill) {
      parts.push('color: transparent');
    } else if (runStyle.textOutlineColor) {
      parts.push(`-webkit-text-stroke-width: ${strokeW}px`);
      parts.push(`-webkit-text-stroke-color: ${runStyle.textOutlineColor}`);
      parts.push('paint-order: stroke fill');
    }
  }

  const hasExplicitRunFont =
    runProps?.child('latin').exists() ||
    runProps?.child('ea').exists() ||
    runProps?.child('cs').exists();
  const effectiveFont = hasExplicitRunFont
    ? runStyle.fontFamily
    : (options?.cellTextFontFamily ?? runStyle.fontFamily);
  if (effectiveFont) {
    parts.push(`font-family: "${effectiveFont.replace(/"/g, '\\"')}"`);
  } else {
    const fallback = ctx.theme.minorFont.latin || ctx.theme.minorFont.ea;
    if (fallback) {
      parts.push(`font-family: "${fallback.replace(/"/g, '\\"')}"`);
    }
  }

  if (runStyle.letterSpacingPt !== undefined) {
    parts.push(`letter-spacing: ${runStyle.letterSpacingPt}pt`);
  }
  if (runStyle.kern !== undefined) {
    const pt = (runStyle.fontSize || 12) * fontScale;
    parts.push(`font-kerning: ${pt >= runStyle.kern ? 'normal' : 'none'}`);
  }
  if (runStyle.cap === 'all') {
    parts.push('text-transform: uppercase');
  } else if (runStyle.cap === 'small') {
    parts.push('font-variant: small-caps');
  }
  if (runStyle.baseline !== undefined && runStyle.baseline !== 0) {
    const shiftPct = runStyle.baseline / 1000;
    parts.push(`vertical-align: ${shiftPct}%`);
    if (Math.abs(shiftPct) >= 20) {
      parts.push(`font-size: ${fontSize * fontScale * 0.65}pt`);
    }
  }

  return parts.join('; ');
}

/**
 * Serialize TextBody to HTML for `Shape.content` / `Text.content`.
 * Implements the same inheritance order as `TextRenderer.renderTextBody`.
 */
export function textToHtml(
  ctx: RenderContext,
  textBody: TextBody | undefined,
  placeholder?: PlaceholderInfo,
  options?: TextToHtmlOptions,
): string {
  if (!textBody?.paragraphs?.length) return '';

  let fontScale = 1;
  let lnSpcReduction = 0;
  if (textBody.bodyProperties) {
    const normAutofit = textBody.bodyProperties.child('normAutofit');
    if (normAutofit.exists()) {
      const fs = normAutofit.numAttr('fontScale');
      if (fs !== undefined) fontScale = fs / 100000;
      const lsr = normAutofit.numAttr('lnSpcReduction');
      if (lsr !== undefined) lnSpcReduction = lsr / 100000;
    }
  }

  let bulletCounter = 0;
  let html = '';

  for (const paragraph of textBody.paragraphs) {
    const merged = buildMergedParagraphStyle(paragraph, textBody, placeholder, ctx);

    let effectiveLineHeight = merged.lineHeight;
    if (merged.lineHeight) {
      effectiveLineHeight = applyLnSpcReduction(merged.lineHeight, lnSpcReduction);
    }

    let effectiveFontSize = 12;
    if (merged.defRPr) {
      const sz = merged.defRPr.numAttr('sz');
      if (sz !== undefined) effectiveFontSize = sz / 100;
    }
    if (paragraph.runs.length > 0 && paragraph.runs[0].properties) {
      const sz = paragraph.runs[0].properties.numAttr('sz');
      if (sz !== undefined) effectiveFontSize = sz / 100;
    }

    const paraCss = mergedParagraphCss(merged, effectiveLineHeight, effectiveFontSize);

    const hasVisibleRuns = paragraph.runs.some((r) => r.text != null && r.text.length > 0);
    const suppressBullet =
      !hasVisibleRuns ||
      placeholder?.type === 'sldNum' ||
      placeholder?.type === 'dt' ||
      placeholder?.type === 'ftr' ||
      placeholder?.type === 'title' ||
      placeholder?.type === 'ctrTitle' ||
      placeholder?.type === 'subTitle';

    let bulletPrefix = '';
    if (!suppressBullet && merged.bulletNone !== true) {
      if (merged.bulletChar) {
        bulletPrefix = merged.bulletChar;
      } else if (merged.bulletAutoNum) {
        bulletPrefix = generateAutoNumber(merged.bulletAutoNum, bulletCounter);
        bulletCounter++;
      }
    }

    const hasLineBreaks = paragraph.runs.some((r) => r.text === '\n');
    const useLineWrappers =
      !!(merged.lineHeightAbsolute && hasLineBreaks && effectiveLineHeight);

    let openExtra = '';
    if (paragraph.runs.some((r) => r.text?.includes('\t'))) {
      const defaultTabPx = 96;
      openExtra = `tab-size: ${defaultTabPx}px`;
    }

    const fullParaStyle = [paraCss, openExtra].filter(Boolean).join('; ');

    html += paragraphOpenTag(useLineWrappers, fullParaStyle);

    if (bulletPrefix) {
      let bulletColor: string | undefined;
      if (merged.bulletColorNode && merged.bulletColorNode.exists()) {
        bulletColor = resolveColorToCss(merged.bulletColorNode, ctx);
      }
      if (bulletColor === undefined && merged.defRPr && merged.defRPr.exists()) {
        const bs: MergedRunStyle = {};
        mergeRunProps(bs, merged.defRPr, ctx);
        bulletColor = bs.color;
      }
      if (bulletColor === undefined && paragraph.runs.length > 0) {
        const rs: MergedRunStyle = {};
        if (merged.defRPr) mergeRunProps(rs, merged.defRPr, ctx);
        if (paragraph.runs[0].properties) mergeRunProps(rs, paragraph.runs[0].properties, ctx);
        bulletColor = rs.color;
      }
      if (bulletColor === undefined && textBody.listStyle) {
        const lstStyleLevel = findStyleAtLevel(textBody.listStyle, paragraph.level);
        if (lstStyleLevel.exists()) {
          const lstDefRPr = lstStyleLevel.child('defRPr');
          if (lstDefRPr.exists()) {
            const fb: MergedRunStyle = {};
            mergeRunProps(fb, lstDefRPr, ctx);
            if (fb.color !== undefined) bulletColor = fb.color;
          }
        }
      }
      const bColor =
        bulletColor ?? options?.fontRefColor ?? options?.cellTextColor ?? '#000000';
      const bFont = merged.bulletFont ? `font-family: "${merged.bulletFont.replace(/"/g, '\\"')}"; ` : '';
      html += `<span style="${bFont}color: ${bColor}">${escapeHtml(bulletPrefix)} </span>`;
    }

    const level = paragraph.level;

    if (paragraph.runs.length === 0) {
      html += '<br/>';
    }

    let currentLineHeightForWrapper = effectiveLineHeight || merged.lineHeight || '';
    let lineWrapperOpen = false;

    const openLineWrapper = () => {
      if (!useLineWrappers) return;
      const h = escapeHtmlAttr(currentLineHeightForWrapper);
      html += `<div style="height: ${h}; overflow: visible">`;
      lineWrapperOpen = true;
    };

    const closeLineWrapper = () => {
      if (lineWrapperOpen) {
        html += '</div>';
        lineWrapperOpen = false;
      }
    };

    if (useLineWrappers) {
      openLineWrapper();
    }

    for (const run of paragraph.runs) {
      if (run.text === '\n') {
        if (useLineWrappers) {
          closeLineWrapper();
          openLineWrapper();
        } else {
          html += '<br/>';
        }
        continue;
      }

      const runStyle: MergedRunStyle = {};
      if (merged.defRPr) {
        mergeRunProps(runStyle, merged.defRPr, ctx);
      }
      if (run.properties) {
        mergeRunProps(runStyle, run.properties, ctx);
      }

      if (runStyle.color === undefined && textBody.listStyle) {
        const lstStyleLevel = findStyleAtLevel(textBody.listStyle, level);
        if (lstStyleLevel.exists()) {
          const lstDefRPr = lstStyleLevel.child('defRPr');
          if (lstDefRPr.exists()) {
            const fallbackStyle: MergedRunStyle = {};
            mergeRunProps(fallbackStyle, lstDefRPr, ctx);
            if (fallbackStyle.color !== undefined) {
              runStyle.color = fallbackStyle.color;
            }
          }
        }
      }

      const styleStr = buildRunStyleString(runStyle, run, fontScale, options, ctx);
      const tabStyle = run.text?.includes('\t') ? `${styleStr}; white-space: pre` : styleStr;

      let inner: string;
      if (run.text && run.text.includes('\t')) {
        inner = formatRunTextForHtml(run.text);
      } else {
        inner = formatRunTextForHtml(run.text ?? '');
      }

      if (runStyle.hlinkClick) {
        const href = escapeHtmlAttr(runStyle.hlinkClick);
        html += `<a href="${href}" target="_blank" rel="noopener noreferrer" style="${escapeHtmlAttr(tabStyle)}">${inner}</a>`;
      } else {
        html += `<span style="${escapeHtmlAttr(tabStyle)}">${inner}</span>`;
      }
    }

    if (useLineWrappers) {
      closeLineWrapper();
    }

    if (paragraph.endParaRPr) {
      const lastRun = paragraph.runs[paragraph.runs.length - 1];
      if (lastRun?.text === '\n') {
        const epSz = paragraph.endParaRPr.numAttr('sz');
        if (epSz !== undefined) {
          const fs = (epSz / 100) * fontScale;
          html += `<span style="font-size: ${fs.toFixed(4)}pt">&#x200B;</span>`;
        }
      }
    }

    html += paragraphCloseTag(useLineWrappers);
  }

  return html;
}
