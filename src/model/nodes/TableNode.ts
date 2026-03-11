/**
 * Table node parser — handles graphicFrame elements containing a:tbl.
 */

import { SafeXmlNode } from '../../parser/XmlParser';
import { BaseNodeData, parseBaseProps } from './BaseNode';
import { TextBody, parseTextBody } from './ShapeNode';
import { emuToPx } from '../../parser/units';

export interface TableCell {
  gridSpan: number;
  rowSpan: number;
  hMerge: boolean;
  vMerge: boolean;
  textBody?: TextBody;
  properties?: SafeXmlNode;
}

export interface TableRow {
  height: number;
  cells: TableCell[];
}

export interface TableNodeData extends BaseNodeData {
  nodeType: 'table';
  columns: number[];
  rows: TableRow[];
  properties?: SafeXmlNode;
  tableStyleId?: string;
}

function parseCell(tcNode: SafeXmlNode): TableCell {
  const gridSpan = tcNode.numAttr('gridSpan') ?? 1;
  const rowSpan = tcNode.numAttr('rowSpan') ?? 1;
  const hMerge = tcNode.attr('hMerge') === '1' || tcNode.attr('hMerge') === 'true';
  const vMerge = tcNode.attr('vMerge') === '1' || tcNode.attr('vMerge') === 'true';
  const txBody = tcNode.child('txBody');
  const textBody = parseTextBody(txBody);
  const tcPr = tcNode.child('tcPr');
  return {
    gridSpan,
    rowSpan,
    hMerge,
    vMerge,
    textBody,
    properties: tcPr.exists() ? tcPr : undefined,
  };
}

function parseRow(trNode: SafeXmlNode): TableRow {
  const height = emuToPx(trNode.numAttr('h') ?? 0);
  const cells: TableCell[] = [];
  for (const tcNode of trNode.children('tc')) {
    cells.push(parseCell(tcNode));
  }
  return { height, cells };
}

function findTable(frameNode: SafeXmlNode): SafeXmlNode {
  const graphic = frameNode.child('graphic');
  const graphicData = graphic.child('graphicData');
  return graphicData.child('tbl');
}

function extractTableStyleId(tblPr: SafeXmlNode): string | undefined {
  const tableStyleIdNode = tblPr.child('tableStyleId');
  if (tableStyleIdNode.exists()) {
    return tableStyleIdNode.text() || tableStyleIdNode.attr('val') || undefined;
  }
  const tblStyleNode = tblPr.child('tblStyle');
  if (tblStyleNode.exists()) {
    return tblStyleNode.attr('val') ?? (tblStyleNode.text() || undefined);
  }
  return tblPr.attr('tblStyle') ?? undefined;
}

export function parseTableNode(frameNode: SafeXmlNode): TableNodeData {
  const base = parseBaseProps(frameNode);
  const tbl = findTable(frameNode);
  const tblGrid = tbl.child('tblGrid');
  const columns: number[] = [];
  for (const gridCol of tblGrid.children('gridCol')) {
    columns.push(emuToPx(gridCol.numAttr('w') ?? 0));
  }
  const rows: TableRow[] = [];
  for (const trNode of tbl.children('tr')) {
    rows.push(parseRow(trNode));
  }
  const tblPr = tbl.child('tblPr');
  const tableStyleId = extractTableStyleId(tblPr);
  return {
    ...base,
    nodeType: 'table',
    columns,
    rows,
    properties: tblPr.exists() ? tblPr : undefined,
    tableStyleId,
  };
}
