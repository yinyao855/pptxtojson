/**
 * Slide parser — converts a slide XML into a structured SlideData with typed nodes.
 */

import { SafeXmlNode, parseXml } from '../parser/XmlParser';
import { RelEntry, resolveRelTarget } from '../parser/RelParser';
import { emuToPx } from '../parser/units';
import { parseBaseProps } from './nodes/BaseNode';
import { ShapeNodeData, parseShapeNode } from './nodes/ShapeNode';
import { PicNodeData, parsePicNode } from './nodes/PicNode';
import { TableNodeData, parseTableNode } from './nodes/TableNode';
import { GroupNodeData, parseGroupNode } from './nodes/GroupNode';
import { ChartNodeData, parseChartNode } from './nodes/ChartNode';

export type SlideNode = ShapeNodeData | PicNodeData | TableNodeData | GroupNodeData | ChartNodeData;

export interface SlideData {
  index: number;
  nodes: SlideNode[];
  background?: SafeXmlNode;
  layoutIndex: string;
  rels: Map<string, RelEntry>;
  slidePath: string;
  showMasterSp: boolean;
}

function isTableFrame(node: SafeXmlNode): boolean {
  const graphic = node.child('graphic');
  const graphicData = graphic.child('graphicData');
  return graphicData.child('tbl').exists();
}

function isChartFrame(node: SafeXmlNode): boolean {
  const graphic = node.child('graphic');
  const graphicData = graphic.child('graphicData');
  return (graphicData.attr('uri') || '').includes('chart');
}

function findOleFallbackPic(graphicFrame: SafeXmlNode): SafeXmlNode | null {
  const graphic = graphicFrame.child('graphic');
  const graphicData = graphic.child('graphicData');
  if (!(graphicData.attr('uri') || '').includes('ole')) return null;
  const altContent = graphicData.child('AlternateContent');
  if (!altContent.exists()) return null;
  for (const branch of ['Fallback', 'Choice'] as const) {
    const oleObj = altContent.child(branch).child('oleObj');
    if (!oleObj.exists()) continue;
    const pic = oleObj.child('pic');
    if (!pic.exists()) continue;
    const blipFill = pic.child('blipFill');
    const blip = blipFill.child('blip');
    const embed = blip.attr('embed') ?? blip.attr('r:embed');
    if (embed) return pic;
  }
  return null;
}

export function parseOleFrameAsPicture(graphicFrame: SafeXmlNode): PicNodeData | undefined {
  const pic = findOleFallbackPic(graphicFrame);
  if (!pic) return undefined;
  const base = parseBaseProps(graphicFrame);
  const blipFill = pic.child('blipFill');
  const blip = blipFill.child('blip');
  const blipEmbed = blip.attr('embed') ?? blip.attr('r:embed');
  const blipLink = blip.attr('link') ?? blip.attr('r:link');
  if (!blipEmbed) return undefined;
  return {
    ...base,
    nodeType: 'picture',
    blipEmbed,
    blipLink,
    source: graphicFrame,
  };
}

function isDiagramFrame(node: SafeXmlNode): boolean {
  return (node.child('graphic').child('graphicData').attr('uri') || '').includes('diagram');
}

function readShapeBounds(node: SafeXmlNode): { x: number; y: number; w: number; h: number } | null {
  const spPr = node.child('spPr');
  if (!spPr.exists()) return null;
  const xfrm = spPr.child('xfrm');
  if (!xfrm.exists()) return null;
  const off = xfrm.child('off');
  const ext = xfrm.child('ext');
  return {
    x: emuToPx(off.numAttr('x') ?? 0),
    y: emuToPx(off.numAttr('y') ?? 0),
    w: emuToPx(ext.numAttr('cx') ?? 0),
    h: emuToPx(ext.numAttr('cy') ?? 0),
  };
}

function buildDiagramGroup(
  base: ReturnType<typeof parseBaseProps>,
  drawingXml: string,
): GroupNodeData {
  const drawingRoot = parseXml(drawingXml);
  const spTree = drawingRoot.child('spTree');
  if (!spTree.exists()) {
    return {
      ...base,
      nodeType: 'group',
      childOffset: { x: 0, y: 0 },
      childExtent: { w: base.size.w, h: base.size.h },
      children: [],
    };
  }
  const CHILD_TAGS = new Set(['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp']);
  const children: SafeXmlNode[] = [];
  let minX = Infinity, minY = Infinity, maxRight = -Infinity, maxBottom = -Infinity;
  for (const child of spTree.allChildren()) {
    if (CHILD_TAGS.has(child.localName)) {
      children.push(child);
      const b = readShapeBounds(child);
      if (b) {
        minX = Math.min(minX, b.x);
        minY = Math.min(minY, b.y);
        maxRight = Math.max(maxRight, b.x + b.w);
        maxBottom = Math.max(maxBottom, b.y + b.h);
      }
    }
  }
  return {
    ...base,
    nodeType: 'group',
    childOffset: { x: 0, y: 0 },
    childExtent: { w: Math.max(1, base.size.w), h: Math.max(1, base.size.h) },
    children,
  };
}

function parseDiagramFrame(
  graphicFrame: SafeXmlNode,
  rels: Map<string, RelEntry>,
  slidePath: string,
  diagramDrawings: Map<string, string>,
): GroupNodeData | undefined {
  const base = parseBaseProps(graphicFrame);
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
  const drawingCandidates = Array.from(rels.values())
    .filter(
      (e) => e.type.includes('diagramDrawing') || e.target.includes('diagrams/drawing'),
    )
    .map((e) => ({
      target: e.target,
      num: (e.target.match(/drawing(\d+)/) || [])[1] ? parseInt(e.target.match(/drawing(\d+)/)![1], 10) : undefined,
    }));
  const graphic = graphicFrame.child('graphic');
  const graphicData = graphic.child('graphicData');
  const relIds = graphicData.child('relIds');
  if (relIds.exists()) {
    const dmRId = relIds.attr('r:dm') ?? relIds.attr('dm');
    if (dmRId) {
      const dmRel = rels.get(dmRId);
      if (dmRel) {
        const numMatch = dmRel.target.match(/data(\d+)/);
        if (numMatch) {
          const drawingNum = parseInt(numMatch[1], 10);
          const ordered = drawingCandidates.slice().sort((a, b) => {
            const da = a.num === undefined ? Infinity : Math.abs(a.num - drawingNum);
            const db = b.num === undefined ? Infinity : Math.abs(b.num - drawingNum);
            return da - db;
          });
          for (const candidate of ordered) {
            const drawingPath = resolveRelTarget(slideDir, candidate.target);
            const drawingXml = diagramDrawings.get(drawingPath);
            if (drawingXml) return buildDiagramGroup(base, drawingXml);
          }
        }
      }
    }
  }
  for (const candidate of drawingCandidates) {
    const drawingPath = resolveRelTarget(slideDir, candidate.target);
    const drawingXml = diagramDrawings.get(drawingPath);
    if (drawingXml) return buildDiagramGroup(base, drawingXml);
  }
  return undefined;
}

export function parseChildNode(
  child: SafeXmlNode,
  rels: Map<string, RelEntry>,
  slidePath: string,
  diagramDrawings?: Map<string, string>,
): SlideNode | undefined {
  const tag = child.localName;
  switch (tag) {
    case 'sp':
    case 'cxnSp':
      return parseShapeNode(child);
    case 'pic':
      return parsePicNode(child);
    case 'grpSp':
      return parseGroupNode(child);
    case 'graphicFrame':
      if (isTableFrame(child)) return parseTableNode(child);
      if (isChartFrame(child)) return parseChartNode(child, rels, slidePath);
      if (isDiagramFrame(child) && diagramDrawings) {
        return parseDiagramFrame(child, rels, slidePath, diagramDrawings);
      }
      const olePic = parseOleFrameAsPicture(child);
      if (olePic) return olePic;
      return undefined;
    default:
      return undefined;
  }
}

function findLayoutRel(rels: Map<string, RelEntry>): string {
  for (const [, entry] of rels) {
    if (entry.type.includes('slideLayout')) return entry.target;
  }
  return '';
}

export function parseSlide(
  root: SafeXmlNode,
  index: number,
  rels: Map<string, RelEntry>,
  slidePath: string = '',
  diagramDrawings?: Map<string, string>,
): SlideData {
  const cSld = root.child('cSld');
  const bg = cSld.child('bg');
  const background = bg.exists() ? bg : undefined;
  const spTree = cSld.child('spTree');
  const nodes: SlideNode[] = [];
  for (const child of spTree.allChildren()) {
    const node = parseChildNode(child, rels, slidePath, diagramDrawings);
    if (node) nodes.push(node);
  }
  const layoutIndex = findLayoutRel(rels);
  const showMasterSpAttr = root.attr('showMasterSp');
  const showMasterSp = showMasterSpAttr !== '0';
  return {
    index,
    nodes,
    background,
    layoutIndex,
    rels,
    slidePath,
    showMasterSp,
  };
}
