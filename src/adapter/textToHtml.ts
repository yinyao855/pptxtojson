/**
 * Convert TextBody (paragraphs/runs) to HTML for pptxtojson content field.
 * Uses RenderContext for theme/master color resolution.
 */

import type { RenderContext } from '../resolve/RenderContext';
import type { TextBody, TextParagraph, TextRun } from '../model/nodes/ShapeNode';
import { resolveColorToCss } from '../resolve/StyleResolver';
import { hundredthPtToPt } from '../parser/units';
import { SafeXmlNode } from '../parser/XmlParser';

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function runStyle(rPr: SafeXmlNode | undefined, ctx: RenderContext): string {
  if (!rPr?.exists()) return '';
  const parts: string[] = [];
  const sz = rPr.numAttr('sz');
  if (sz !== undefined && sz > 0) {
    const pt = hundredthPtToPt(sz);
    parts.push(`font-size: ${pt.toFixed(1)}pt`);
  }
  const solidFill = rPr.child('solidFill');
  const schemeClr = rPr.child('schemeClr');
  if (solidFill.exists()) {
    const color = resolveColorToCss(solidFill, ctx);
    if (color) parts.push(`color: ${color}`);
  } else if (schemeClr.exists()) {
    const color = resolveColorToCss(schemeClr, ctx);
    if (color) parts.push(`color: ${color}`);
  }
  const b = rPr.attr('b');
  if (b === '1' || b === 'true') parts.push('font-weight: bold');
  const i = rPr.attr('i');
  if (i === '1' || i === 'true') parts.push('font-style: italic');
  const u = rPr.attr('u');
  if (u && u !== 'none') parts.push('text-decoration: underline');
  return parts.join('; ');
}

export function textToHtml(ctx: RenderContext, textBody: TextBody | undefined): string {
  if (!textBody?.paragraphs?.length) return '';
  let html = '';
  for (const para of textBody.paragraphs) {
    const pStyle = paragraphStyle(para);
    html += `<p${pStyle ? ` style="${pStyle}"` : ''}>`;
    for (const run of para.runs) {
      const style = runStyle(run.properties, ctx);
      const text = escapeHtml(run.text);
      if (run.text === '\n') {
        html += '<br/>';
      } else if (style) {
        html += `<span style="${style}">${text}</span>`;
      } else {
        html += text;
      }
    }
    html += '</p>';
  }
  return html;
}

function paragraphStyle(para: TextParagraph): string {
  const parts: string[] = [];
  const pPr = para.properties;
  if (!pPr?.exists()) return '';
  const align = pPr.attr('al') ?? pPr.attr('marL');
  if (align === 'ctr' || align === 'center') parts.push('text-align: center');
  if (align === 'r' || align === 'right') parts.push('text-align: right');
  if (align === 'j' || align === 'just') parts.push('text-align: justify');
  return parts.join('; ');
}
