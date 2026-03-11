/**
 * Base node types and property parser shared by all slide node kinds.
 */

import { SafeXmlNode } from '../../parser/XmlParser';
import { emuToPx, angleToDeg } from '../../parser/units';

export type NodeType = 'shape' | 'picture' | 'table' | 'group' | 'chart' | 'unknown';

export interface Position {
  x: number;
  y: number;
}

export interface Size {
  w: number;
  h: number;
}

export interface PlaceholderInfo {
  type?: string;
  idx?: number;
}

export interface HlinkAction {
  action?: string;
  rId?: string;
  tooltip?: string;
}

export interface BaseNodeData {
  id: string;
  name: string;
  nodeType: NodeType;
  position: Position;
  size: Size;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  placeholder?: PlaceholderInfo;
  hlinkClick?: HlinkAction;
  source: SafeXmlNode;
}

function findNvProps(node: SafeXmlNode): { cNvPr: SafeXmlNode; nvPr: SafeXmlNode } {
  const wrappers = ['nvSpPr', 'nvPicPr', 'nvGrpSpPr', 'nvGraphicFramePr', 'nvCxnSpPr'];
  for (const name of wrappers) {
    const wrapper = node.child(name);
    if (wrapper.exists()) {
      return {
        cNvPr: wrapper.child('cNvPr'),
        nvPr: wrapper.child('nvPr'),
      };
    }
  }
  return {
    cNvPr: node.child('cNvPr'),
    nvPr: node.child('nvPr'),
  };
}

function findXfrm(node: SafeXmlNode): SafeXmlNode {
  const spPr = node.child('spPr');
  if (spPr.exists()) {
    const xfrm = spPr.child('xfrm');
    if (xfrm.exists()) return xfrm;
  }
  const grpSpPr = node.child('grpSpPr');
  if (grpSpPr.exists()) {
    const xfrm = grpSpPr.child('xfrm');
    if (xfrm.exists()) return xfrm;
  }
  const directXfrm = node.child('xfrm');
  if (directXfrm.exists()) return directXfrm;
  return node.child('__nonexistent__');
}

function parsePlaceholder(nvPr: SafeXmlNode): PlaceholderInfo | undefined {
  const ph = nvPr.child('ph');
  if (!ph.exists()) return undefined;
  const type = ph.attr('type');
  const idx = ph.numAttr('idx');
  return { type, idx };
}

export function parseBaseProps(spNode: SafeXmlNode): Omit<BaseNodeData, 'nodeType'> {
  const { cNvPr, nvPr } = findNvProps(spNode);
  const id = cNvPr.attr('id') ?? '';
  const name = cNvPr.attr('name') ?? '';
  const xfrm = findXfrm(spNode);
  const off = xfrm.child('off');
  const ext = xfrm.child('ext');
  const position: Position = {
    x: emuToPx(off.numAttr('x') ?? 0),
    y: emuToPx(off.numAttr('y') ?? 0),
  };
  const size: Size = {
    w: emuToPx(ext.numAttr('cx') ?? 0),
    h: emuToPx(ext.numAttr('cy') ?? 0),
  };
  const rotation = angleToDeg(xfrm.numAttr('rot') ?? 0);
  const flipH = xfrm.attr('flipH') === '1' || xfrm.attr('flipH') === 'true';
  const flipV = xfrm.attr('flipV') === '1' || xfrm.attr('flipV') === 'true';
  const placeholder = parsePlaceholder(nvPr);
  let hlinkClick: HlinkAction | undefined;
  const hlinkNode = cNvPr.child('hlinkClick');
  if (hlinkNode.exists()) {
    hlinkClick = {
      action: hlinkNode.attr('action') ?? undefined,
      rId: hlinkNode.attr('id') ?? hlinkNode.attr('r:id') ?? undefined,
      tooltip: hlinkNode.attr('tooltip') ?? undefined,
    };
  }
  return {
    id,
    name,
    position,
    size,
    rotation,
    flipH,
    flipV,
    placeholder,
    hlinkClick,
    source: spNode,
  };
}
