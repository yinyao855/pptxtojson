/**
 * Serializes GroupNodeData to pptxtojson Group element.
 * Recursively serializes children via nodeToElement; flattens nested groups into elements.
 */

import type { GroupNodeData } from '../model/nodes/GroupNode';
import type { SlideNode } from '../model/Slide';
import { parseChildNode } from '../model/Slide';
import type { RenderContext } from './RenderContext';
import type { PptxFiles } from '../parser/ZipParser';
import type { Group, Element, BaseElement } from '../adapter/types';

// Type guard for Group (Element = BaseElement | Group)
function isGroup(e: Element): e is Group {
  return (e as Group).type === 'group';
}

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

/**
 * Flatten a group into BaseElement[] with positions in parent (slide) space.
 * Offsets each element by (baseLeft + group.left, baseTop + group.top).
 */
function flattenGroupInto(group: Group, baseLeft: number, baseTop: number, out: BaseElement[]): void {
  const offsetLeft = baseLeft + group.left;
  const offsetTop = baseTop + group.top;
  for (const el of group.elements) {
    const e = el as BaseElement & { left: number; top: number };
    out.push({ ...e, left: offsetLeft + e.left, top: offsetTop + e.top } as BaseElement);
  }
}

export type NodeToElement = (
  node: SlideNode,
  ctx: RenderContext,
  order: number,
  files?: PptxFiles,
) => Element;

/**
 * Serialize group node to Group element.
 * Children are parsed with parseChildNode and serialized with nodeToElement.
 * Nested groups are flattened so Group.elements is BaseElement[].
 */
export function groupToElement(
  node: GroupNodeData,
  ctx: RenderContext,
  order: number,
  files: PptxFiles | undefined,
  nodeToElement: NodeToElement,
): Group {
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);
  const rels = ctx.slide.rels;
  const slidePath = ctx.slide.slidePath;
  const diagramDrawings = files?.diagramDrawings;
  const elements: BaseElement[] = [];
  let idx = 0;
  for (const childXml of node.children) {
    const childNode = parseChildNode(childXml, rels, slidePath, diagramDrawings);
    if (childNode) {
      const el = nodeToElement(childNode, ctx, idx, files);
      if (isGroup(el)) {
        flattenGroupInto(el, left, top, elements);
      } else {
        const be = el as BaseElement & { left: number; top: number };
        elements.push({
          ...be,
          left: left + be.left,
          top: top + be.top,
        } as BaseElement);
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
