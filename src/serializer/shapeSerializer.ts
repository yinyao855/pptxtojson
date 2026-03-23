/**
 * Serializes ShapeNodeData to pptxtojson Shape or Text element.
 * Uses same resolution as ShapeRenderer: spPr fill/line, presets/customGeometry path, textSerializer for content.
 */

import type { ShapeNodeData, TextBody } from '../model/nodes/ShapeNode';
import type { RenderContext } from './RenderContext';
import { getPresetShapePath } from '../shapes/presets';
import { renderCustomGeometry } from '../shapes/customGeometry';
import { spPrToFill } from './fillMapper';
import { lineStyleToBorder } from './borderMapper';
import { textToHtml } from './textSerializer';
import type { Shape, Text, Fill } from '../adapter/types';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

function hasVisibleText(textBody: TextBody | undefined): boolean {
  if (!textBody?.paragraphs?.length) return false;
  for (const p of textBody.paragraphs) {
    for (const r of p.runs) {
      if (r.text != null && r.text.trim().length > 0) return true;
    }
  }
  return false;
}

/**
 * Decide vertical alignment from bodyPr (a:bodyPr anchor).
 * PPTist expects: 'top' | 'middle' | 'bottom'.
 */
function getVAlign(bodyPr: { attr: (n: string) => string | undefined } | undefined): string {
  if (!bodyPr) return 'top';
  const anchor = bodyPr.attr('anchor') ?? bodyPr.attr('vert');
  if (anchor === 't' || anchor === 'top') return 'top';
  if (anchor === 'ctr' || anchor === 'mid' || anchor === 'middle') return 'middle';
  if (anchor === 'b' || anchor === 'bottom') return 'bottom';
  return 'top';
}

/**
 * Serialize a shape node to Shape or Text element.
 * When the shape has visible text and is effectively a text box, returns Text; otherwise Shape.
 */
export function shapeToElement(
  node: ShapeNodeData,
  ctx: RenderContext,
  order: number,
): Shape | Text {
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);
  const spPr = node.source.child('spPr');
  const fill: Fill = spPr.exists() ? spPrToFill(spPr, ctx) : { type: 'color', value: '#ffffff' };
  const ln = spPr.exists() ? spPr.child('ln') : node.source.child('__none__');
  const borderResult = ln.exists() ? lineStyleToBorder(ln, ctx) : {
    border: { borderColor: '#000000', borderWidth: 0, borderType: 'solid' as const },
    borderStrokeDasharray: '',
  };
  const content = node.textBody ? textToHtml(ctx, node.textBody) : '';
  const hasContent = hasVisibleText(node.textBody);
  const bodyPr = node.textBody?.bodyProperties;
  const vAlign = bodyPr?.attr ? getVAlign(bodyPr) : 'top';

  const base = {
    left,
    top,
    width,
    height,
    name: node.name || '',
    order,
    borderColor: borderResult.border.borderColor,
    borderWidth: borderResult.border.borderWidth,
    borderType: borderResult.border.borderType,
    borderStrokeDasharray: borderResult.borderStrokeDasharray || '',
    fill,
    isFlipV: node.flipV,
    isFlipH: node.flipH,
    rotate: node.rotation,
    content: content || '',
  };

  const path = getShapePath(node);
  const shapType = node.presetGeometry || 'rect';

  if (hasContent && (node.placeholder?.type === 'body' || node.placeholder?.type === 'title' || node.placeholder?.type === 'ctrTitle')) {
    return {
      ...base,
      type: 'text',
      isVertical: false,
      vAlign,
    } as Text;
  }

  return {
    ...base,
    type: 'shape',
    shapType,
    path: path || undefined,
    vAlign,
  } as Shape;
}

function getShapePath(node: ShapeNodeData): string | undefined {
  const w = node.size.w;
  const h = node.size.h;
  if (node.customGeometry?.exists()) {
    return renderCustomGeometry(node.customGeometry, w, h);
  }
  const preset = node.presetGeometry || 'rect';
  return getPresetShapePath(preset, w, h, node.adjustments);
}
