/**
 * Math serializer — converts MathNodeData into a Math element with LaTeX.
 *
 * Pipeline: OMML XML → omml2mathml (MathML DOM) → mathml-to-latex (LaTeX string)
 */

import type { MathNodeData } from '../model/nodes/MathNode';
import type { RenderContext } from './RenderContext';
import type { Math as MathElement } from '../adapter/types';
import { resolveMediaToUrl } from '../utils/mediaWebConvert';
import { resolveMediaPath } from '../utils/media';
import { DOMParser } from '@xmldom/xmldom';

// @ts-expect-error — omml2mathml has no type declarations
import omml2mathml from 'omml2mathml';
import mathmlToLatex from 'mathml-to-latex';
const MathMLToLaTeX = mathmlToLatex.MathMLToLaTeX ?? mathmlToLatex;

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return Number((px * PX_TO_PT).toFixed(4));
}

// ---------------------------------------------------------------------------
// Unicode Mathematical Alphanumeric Symbols → ASCII / LaTeX normalization
// ---------------------------------------------------------------------------
// PPTX stores math variables using Unicode Mathematical Italic/Bold codepoints
// (e.g. U+1D465 "𝑥" instead of "x"). LaTeX math mode already italicizes
// variables, so we must normalize these back to plain ASCII / LaTeX commands.

const GREEK_LOWER_LATEX: Record<number, string> = {
  0x03B1: '\\alpha', 0x03B2: '\\beta', 0x03B3: '\\gamma', 0x03B4: '\\delta',
  0x03B5: '\\epsilon', 0x03B6: '\\zeta', 0x03B7: '\\eta', 0x03B8: '\\theta',
  0x03B9: '\\iota', 0x03BA: '\\kappa', 0x03BB: '\\lambda', 0x03BC: '\\mu',
  0x03BD: '\\nu', 0x03BE: '\\xi', 0x03C0: '\\pi', 0x03C1: '\\rho',
  0x03C2: '\\varsigma', 0x03C3: '\\sigma', 0x03C4: '\\tau', 0x03C5: '\\upsilon',
  0x03C6: '\\varphi', 0x03C7: '\\chi', 0x03C8: '\\psi', 0x03C9: '\\omega',
};
const GREEK_UPPER_LATEX: Record<number, string> = {
  0x0391: 'A', 0x0392: 'B', 0x0393: '\\Gamma', 0x0394: '\\Delta',
  0x0395: 'E', 0x0396: 'Z', 0x0397: 'H', 0x0398: '\\Theta',
  0x0399: 'I', 0x039A: 'K', 0x039B: '\\Lambda', 0x039C: 'M',
  0x039D: 'N', 0x039E: '\\Xi', 0x039F: 'O', 0x03A0: '\\Pi',
  0x03A1: 'P', 0x03A3: '\\Sigma', 0x03A4: 'T', 0x03A5: '\\Upsilon',
  0x03A6: '\\Phi', 0x03A7: 'X', 0x03A8: '\\Psi', 0x03A9: '\\Omega',
};

function normalizeMathChar(cp: number): string | undefined {
  // Mathematical Bold A-Z / a-z
  if (cp >= 0x1D400 && cp <= 0x1D419) return String.fromCharCode(cp - 0x1D400 + 0x41);
  if (cp >= 0x1D41A && cp <= 0x1D433) return String.fromCharCode(cp - 0x1D41A + 0x61);
  // Mathematical Italic A-Z / a-z (h = U+210E is a gap)
  if (cp >= 0x1D434 && cp <= 0x1D44D) return String.fromCharCode(cp - 0x1D434 + 0x41);
  if (cp === 0x210E) return 'h';
  if (cp >= 0x1D44E && cp <= 0x1D467) return String.fromCharCode(cp - 0x1D44E + 0x61);
  // Mathematical Bold Italic A-Z / a-z
  if (cp >= 0x1D468 && cp <= 0x1D481) return String.fromCharCode(cp - 0x1D468 + 0x41);
  if (cp >= 0x1D482 && cp <= 0x1D49B) return String.fromCharCode(cp - 0x1D482 + 0x61);
  // Mathematical Sans-Serif / Bold Sans-Serif / Monospace (covers 0x1D5A0–0x1D6A3)
  if (cp >= 0x1D5A0 && cp <= 0x1D5B9) return String.fromCharCode(cp - 0x1D5A0 + 0x41);
  if (cp >= 0x1D5BA && cp <= 0x1D5D3) return String.fromCharCode(cp - 0x1D5BA + 0x61);
  if (cp >= 0x1D5D4 && cp <= 0x1D5ED) return String.fromCharCode(cp - 0x1D5D4 + 0x41);
  if (cp >= 0x1D5EE && cp <= 0x1D607) return String.fromCharCode(cp - 0x1D5EE + 0x61);
  if (cp >= 0x1D670 && cp <= 0x1D689) return String.fromCharCode(cp - 0x1D670 + 0x41);
  if (cp >= 0x1D68A && cp <= 0x1D6A3) return String.fromCharCode(cp - 0x1D68A + 0x61);

  // Mathematical Bold / Italic / Bold-Italic Greek → LaTeX commands
  // Bold Greek Capitals (Α-Ω): U+1D6A8–U+1D6C0 → base 0x0391
  if (cp >= 0x1D6A8 && cp <= 0x1D6C0) return GREEK_UPPER_LATEX[cp - 0x1D6A8 + 0x0391];
  // Bold Greek Small (α-ω): U+1D6C2–U+1D6DA → base 0x03B1
  if (cp >= 0x1D6C2 && cp <= 0x1D6DA) return GREEK_LOWER_LATEX[cp - 0x1D6C2 + 0x03B1];
  // Italic Greek Capitals: U+1D6E2–U+1D6FA
  if (cp >= 0x1D6E2 && cp <= 0x1D6FA) return GREEK_UPPER_LATEX[cp - 0x1D6E2 + 0x0391];
  // Italic Greek Small: U+1D6FC–U+1D714
  if (cp >= 0x1D6FC && cp <= 0x1D714) return GREEK_LOWER_LATEX[cp - 0x1D6FC + 0x03B1];
  // Bold Italic Greek Capitals: U+1D71C–U+1D734
  if (cp >= 0x1D71C && cp <= 0x1D734) return GREEK_UPPER_LATEX[cp - 0x1D71C + 0x0391];
  // Bold Italic Greek Small: U+1D736–U+1D74E
  if (cp >= 0x1D736 && cp <= 0x1D74E) return GREEK_LOWER_LATEX[cp - 0x1D736 + 0x03B1];

  // Mathematical Bold Digits 0-9: U+1D7CE–U+1D7D7
  if (cp >= 0x1D7CE && cp <= 0x1D7D7) return String.fromCharCode(cp - 0x1D7CE + 0x30);

  // Basic Greek that mathml-to-latex might pass through as-is
  if (GREEK_LOWER_LATEX[cp]) return GREEK_LOWER_LATEX[cp];
  if (GREEK_UPPER_LATEX[cp]) return GREEK_UPPER_LATEX[cp];

  return undefined;
}

function normalizeMathUnicode(latex: string): string {
  const chars = Array.from(latex);
  const out: string[] = [];
  for (const ch of chars) {
    const cp = ch.codePointAt(0)!;
    const mapped = normalizeMathChar(cp);
    if (mapped !== undefined) {
      // Add space after LaTeX commands that start with \ to prevent merging
      if (mapped.startsWith('\\') && out.length > 0) out.push(' ');
      out.push(mapped);
      if (mapped.startsWith('\\')) out.push(' ');
    } else {
      out.push(ch);
    }
  }
  return out.join('').replace(/ {2,}/g, ' ').trim();
}

// ---------------------------------------------------------------------------
// Core conversion
// ---------------------------------------------------------------------------

/**
 * Pre-process OMML XML: normalize Unicode Mathematical Alphanumeric chars
 * to plain ASCII BEFORE passing to omml2mathml.
 * jsdom (used by omml2mathml) splits surrogate pairs, so we must normalize first.
 */
function normalizeOmmlXml(xml: string): string {
  return Array.from(xml).map(ch => {
    const cp = ch.codePointAt(0)!;
    return normalizeMathChar(cp) ?? ch;
  }).join('');
}

/**
 * Convert OMML XML string to LaTeX via omml2mathml + mathml-to-latex.
 */
function ommlToLatex(ommlXml: string): string {
  try {
    const normalized = normalizeOmmlXml(ommlXml);
    const parser = new DOMParser();
    const doc = parser.parseFromString(normalized, 'application/xml');
    const mathmlNode = omml2mathml(doc);
    if (!mathmlNode) return '';
    const mathmlStr: string = mathmlNode.outerHTML ?? mathmlNode.toString();
    return MathMLToLaTeX.convert(mathmlStr);
  } catch (err) {
    console.warn('[mathSerializer] OMML→LaTeX conversion failed:', err);
    return '';
  }
}

/**
 * Resolve fallback image from the mc:Fallback branch blipFill embed.
 */
async function resolveFallbackImage(
  embed: string | undefined,
  ctx: RenderContext,
): Promise<string> {
  if (!embed) return '';
  const rel = ctx.slide.rels.get(embed);
  if (!rel) return '';
  const mediaPath = resolveMediaPath(rel.target);
  const data = ctx.presentation.media.get(mediaPath);
  if (!data) return '';
  return resolveMediaToUrl(mediaPath, data, ctx.mediaMode, ctx.mediaUrlCache);
}

/**
 * Serialize a math node to a Math element.
 */
export async function mathToElement(
  node: MathNodeData,
  ctx: RenderContext,
  _order: number,
): Promise<MathElement> {
  const order = node.xmlOrder;
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);

  const latex = ommlToLatex(node.ommlXml);
  const picBase64 = await resolveFallbackImage(node.fallbackBlipEmbed, ctx);

  return {
    type: 'math',
    left,
    top,
    width,
    height,
    latex,
    picBase64,
    order,
    text: node.plainText || undefined,
  };
}
