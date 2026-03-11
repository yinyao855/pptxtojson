/**
 * Shape node parser — handles auto-shapes, text boxes, and connectors.
 */

import { SafeXmlNode } from '../../parser/XmlParser';
import { emuToPx, angleToDeg } from '../../parser/units';
import { BaseNodeData, parseBaseProps } from './BaseNode';

export interface TextRun {
  text: string;
  properties?: SafeXmlNode;
}

export interface TextParagraph {
  properties?: SafeXmlNode;
  runs: TextRun[];
  level: number;
  endParaRPr?: SafeXmlNode;
}

export interface TextBody {
  bodyProperties?: SafeXmlNode;
  layoutBodyProperties?: SafeXmlNode;
  listStyle?: SafeXmlNode;
  paragraphs: TextParagraph[];
}

export interface LineEndInfo {
  type: string;
  w?: string;
  len?: string;
}

export interface TextBoxBounds {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation?: number;
}

export interface ShapeNodeData extends BaseNodeData {
  nodeType: 'shape';
  presetGeometry?: string;
  adjustments: Map<string, number>;
  customGeometry?: SafeXmlNode;
  fill?: SafeXmlNode;
  line?: SafeXmlNode;
  headEnd?: LineEndInfo;
  tailEnd?: LineEndInfo;
  textBody?: TextBody;
  textBoxBounds?: TextBoxBounds;
}

function parseParagraph(pNode: SafeXmlNode): TextParagraph {
  const pPr = pNode.child('pPr');
  const level = pPr.numAttr('lvl') ?? 0;
  const runs: TextRun[] = [];
  for (const rNode of pNode.children('r')) {
    const rPr = rNode.child('rPr');
    const tNode = rNode.child('t');
    runs.push({ text: tNode.text(), properties: rPr.exists() ? rPr : undefined });
  }
  const orderedRuns: TextRun[] = [];
  for (const child of pNode.allChildren()) {
    const ln = child.localName;
    if (ln === 'r') {
      const rPr = child.child('rPr');
      const tNode = child.child('t');
      orderedRuns.push({ text: tNode.text(), properties: rPr.exists() ? rPr : undefined });
    } else if (ln === 'br') {
      const rPr = child.child('rPr');
      orderedRuns.push({ text: '\n', properties: rPr.exists() ? rPr : undefined });
    } else if (ln === 'fld') {
      const rPr = child.child('rPr');
      const tNode = child.child('t');
      orderedRuns.push({ text: tNode.text(), properties: rPr.exists() ? rPr : undefined });
    }
  }
  const endParaRPrNode = pNode.child('endParaRPr');
  return {
    properties: pPr.exists() ? pPr : undefined,
    runs: orderedRuns.length > 0 ? orderedRuns : runs,
    level,
    endParaRPr: endParaRPrNode.exists() ? endParaRPrNode : undefined,
  };
}

export function parseTextBody(txBody: SafeXmlNode): TextBody | undefined {
  if (!txBody.exists()) return undefined;
  const bodyPr = txBody.child('bodyPr');
  const lstStyle = txBody.child('lstStyle');
  const paragraphs: TextParagraph[] = [];
  for (const pNode of txBody.children('p')) {
    paragraphs.push(parseParagraph(pNode));
  }
  return {
    bodyProperties: bodyPr.exists() ? bodyPr : undefined,
    listStyle: lstStyle.exists() ? lstStyle : undefined,
    paragraphs,
  };
}

const FILL_TYPES = ['solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'noFill'] as const;

function findFill(spPr: SafeXmlNode): SafeXmlNode | undefined {
  for (const fillType of FILL_TYPES) {
    const fill = spPr.child(fillType);
    if (fill.exists()) return fill;
  }
  return undefined;
}

function parseAdjustments(avLst: SafeXmlNode): Map<string, number> {
  const adjustments = new Map<string, number>();
  for (const gd of avLst.children('gd')) {
    const name = gd.attr('name');
    const fmla = gd.attr('fmla') ?? '';
    if (!name) continue;
    const match = fmla.match(/val\s+(-?\d+)/);
    if (match) {
      adjustments.set(name, Number(match[1]));
    } else {
      const num = Number(fmla);
      if (!Number.isNaN(num)) adjustments.set(name, num);
    }
  }
  return adjustments;
}

export function parseShapeNode(spNode: SafeXmlNode): ShapeNodeData {
  const base = parseBaseProps(spNode);
  const spPr = spNode.child('spPr');
  const prstGeom = spPr.child('prstGeom');
  const presetGeometry = prstGeom.attr('prst');
  const avLst = prstGeom.child('avLst');
  const adjustments = parseAdjustments(avLst);
  const custGeom = spPr.child('custGeom');
  const customGeometry = custGeom.exists() ? custGeom : undefined;
  const fill = findFill(spPr);
  const ln = spPr.child('ln');
  const line = ln.exists() ? ln : undefined;
  let headEnd: LineEndInfo | undefined;
  let tailEnd: LineEndInfo | undefined;
  if (ln.exists()) {
    const headEndNode = ln.child('headEnd');
    if (headEndNode.exists()) {
      const t = headEndNode.attr('type');
      if (t && t !== 'none') headEnd = { type: t, w: headEndNode.attr('w'), len: headEndNode.attr('len') };
    }
    const tailEndNode = ln.child('tailEnd');
    if (tailEndNode.exists()) {
      const t = tailEndNode.attr('type');
      if (t && t !== 'none') tailEnd = { type: t, w: tailEndNode.attr('w'), len: tailEndNode.attr('len') };
    }
  }
  const txBody = spNode.child('txBody');
  const textBody = parseTextBody(txBody);
  let textBoxBounds: TextBoxBounds | undefined;
  const txXfrm = spNode.child('txXfrm');
  if (txXfrm.exists()) {
    const txOff = txXfrm.child('off');
    const txExt = txXfrm.child('ext');
    const xfrm = spPr.child('xfrm');
    const off = xfrm.child('off');
    const ext = xfrm.child('ext');
    const shapeX = off.numAttr('x') ?? 0;
    const shapeY = off.numAttr('y') ?? 0;
    const shapeW = ext.numAttr('cx') ?? 0;
    const shapeH = ext.numAttr('cy') ?? 0;
    const txX = txOff.numAttr('x') ?? 0;
    const txY = txOff.numAttr('y') ?? 0;
    const txW = txExt.numAttr('cx') ?? 0;
    const txH = txExt.numAttr('cy') ?? 0;
    if (shapeW > 0 && shapeH > 0) {
      const txRotDeg = angleToDeg(txXfrm.numAttr('rot') ?? 0);
      const localX = txX - shapeX;
      const localY = txY - shapeY;
      const isHalfTurn = Math.abs(Math.round(txRotDeg)) % 360 === 180;
      const boxX = isHalfTurn ? shapeW - (localX + txW) : localX;
      const boxY = isHalfTurn ? shapeH - (localY + txH) : localY;
      textBoxBounds = {
        x: emuToPx(boxX),
        y: emuToPx(boxY),
        w: emuToPx(txW),
        h: emuToPx(txH),
        rotation: txRotDeg,
      };
    }
  }
  return {
    ...base,
    nodeType: 'shape',
    presetGeometry,
    adjustments,
    customGeometry,
    fill,
    line,
    headEnd,
    tailEnd,
    textBody,
    textBoxBounds,
  };
}
