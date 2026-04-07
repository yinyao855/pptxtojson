/**
 * Serializes GroupNodeData to pptxtojson Group element.
 * Recursively serializes children via nodeToElement; flattens nested groups into elements.
 *
 * Child positions are output relative to the group (consistent with pptxtojson).
 * Applies chOff/chExt coordinate space transformation and scaling.
 */

import type { GroupNodeData } from '../model/nodes/GroupNode';
import type { SlideNode } from '../model/Slide';
import { parseChildNode } from '../model/Slide';
import type { RenderContext } from './RenderContext';
import type { PptxFiles } from '../parser/ZipParser';
import type { Group, Element, BaseElement } from '../adapter/types';
import { SafeXmlNode } from '../parser/XmlParser';

function isGroup(e: Element): e is Group {
  return (e as Group).type === 'group';
}

function isShape(e: Element): e is import('../adapter/types').Shape {
  return (e as import('../adapter/types').Shape).type === 'shape';
}

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return Number((px * PX_TO_PT).toFixed(4));
}

function toFixed(n: number): number {
  return Number(n.toFixed(4));
}

/**
 * Scale all coordinate values in an SVG path string by (sx, sy).
 * Handles M, L, C, Q, S, T, A, H, V, Z and their lowercase (relative) variants.
 */
function scaleSvgPath(d: string, sx: number, sy: number): string {
  if (sx === 1 && sy === 1) return d;
  const tokens = d.match(/[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g);
  if (!tokens) return d;
  const out: string[] = [];
  let cmd = '';
  let argIdx = 0;
  for (const tok of tokens) {
    if (/^[A-Za-z]$/.test(tok)) {
      cmd = tok;
      argIdx = 0;
      out.push(tok);
      continue;
    }
    const v = parseFloat(tok);
    const upper = cmd.toUpperCase();
    let scaled: number;
    if (upper === 'H') {
      scaled = v * sx;
    } else if (upper === 'V') {
      scaled = v * sy;
    } else if (upper === 'A') {
      // A rx ry x-rotation large-arc-flag sweep-flag x y
      const ai = argIdx % 7;
      if (ai === 0) scaled = v * sx;
      else if (ai === 1) scaled = v * sy;
      else if (ai >= 2 && ai <= 4) scaled = v;
      else if (ai === 5) scaled = v * sx;
      else scaled = v * sy;
    } else {
      // M, L, C, Q, S, T: alternating x, y
      scaled = (argIdx % 2 === 0) ? v * sx : v * sy;
    }
    out.push(String(toFixed(scaled)));
    argIdx++;
  }
  // Reconstruct: command letter directly followed by its coordinates
  let result = '';
  for (let i = 0; i < out.length; i++) {
    const t = out[i];
    if (/^[A-Za-z]$/.test(t)) {
      result += t;
    } else {
      if (i > 0 && !/^[A-Za-z]$/.test(out[i - 1])) result += ',';
      result += t;
    }
  }
  return result;
}

const FILL_TAGS = ['solidFill', 'gradFill', 'blipFill', 'pattFill'] as const;

function findGroupFillNode(grpSpPr: SafeXmlNode): SafeXmlNode | undefined {
  for (const tag of FILL_TAGS) {
    const n = grpSpPr.child(tag);
    if (n.exists()) return grpSpPr;
  }
  return undefined;
}

export type NodeToElement = (
  node: SlideNode,
  ctx: RenderContext,
  order: number,
  files?: PptxFiles,
) => Promise<Element>;

export async function groupToElement(
  node: GroupNodeData,
  ctx: RenderContext,
  _order: number,
  files: PptxFiles | undefined,
  nodeToElement: NodeToElement,
): Promise<Group> {
  const order = node.xmlOrder;
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);

  const chOffX = pxToPt(node.childOffset.x);
  const chOffY = pxToPt(node.childOffset.y);
  const chExtW = pxToPt(node.childExtent.w);
  const chExtH = pxToPt(node.childExtent.h);

  const ws = chExtW > 0 ? width / chExtW : 1;
  const hs = chExtH > 0 ? height / chExtH : 1;

  const rels = ctx.slide.rels;
  const slidePath = ctx.slide.slidePath;
  const diagramDrawings = files?.diagramDrawings;
  const elements: BaseElement[] = [];
  let idx = 0;

  const grpSpPr = node.source.child('grpSpPr');
  const groupFillSource = findGroupFillNode(grpSpPr);
  const childCtx: RenderContext = groupFillSource
    ? { ...ctx, groupFillNode: groupFillSource }
    : ctx;

  for (const childXml of node.children) {
    const childNode = parseChildNode(childXml, rels, slidePath, diagramDrawings);
    if (childNode) {
      const el = await nodeToElement(childNode, childCtx, idx, files);
      if (isGroup(el)) {
        const gLeft = toFixed((el.left - chOffX) * ws);
        const gTop = toFixed((el.top - chOffY) * hs);
        for (const child of el.elements) {
          const c = child as BaseElement & { left: number; top: number; width: number; height: number };
          const scaled: any = {
            ...c,
            left: toFixed(gLeft + c.left * ws),
            top: toFixed(gTop + c.top * hs),
            width: toFixed(c.width * ws),
            height: toFixed(c.height * hs),
          };
          if (isShape(c as Element) && (c as any).path) {
            scaled.path = scaleSvgPath((c as any).path, ws, hs);
          }
          elements.push(scaled as BaseElement);
        }
      } else {
        const be = el as BaseElement & { left: number; top: number; width: number; height: number };
        const scaled: any = {
          ...be,
          left: toFixed((be.left - chOffX) * ws),
          top: toFixed((be.top - chOffY) * hs),
          width: toFixed(be.width * ws),
          height: toFixed(be.height * hs),
        };
        if (isShape(el) && (el as any).path) {
          scaled.path = scaleSvgPath((el as any).path, ws, hs);
        }
        elements.push(scaled as BaseElement);
      }
      idx++;
    }
  }

  return {
    type: 'group',
    left,
    top,
    width,
    height,
    rotate: node.rotation,
    elements,
    order,
    isFlipH: node.flipH,
    isFlipV: node.flipV,
  };
}
