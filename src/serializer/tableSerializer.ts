/**
 * Serializes TableNodeData to pptxtojson Table element.
 * Extracts cell text, rowHeights, colWidths; optional cell fill/borders from tcPr.
 */

import type { TableNodeData } from '../model/nodes/TableNode';
import type { RenderContext } from './RenderContext';
import { renderTextBody } from './textSerializer';
import type { Table, TableCell as OutCell, Border } from '../adapter/types';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return Number((px * PX_TO_PT).toFixed(4));
}

function defaultCellBorders(): OutCell['borders'] {
  return {};
}

/**
 * Serialize table node to Table element.
 * Cell text from textBody; fillColor/borders from tcPr when available.
 */
export function tableToElement(
  node: TableNodeData,
  ctx: RenderContext,
  _order: number,
): Table {
  const order = node.xmlOrder;
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);
  const data: OutCell[][] = node.rows.map((row) =>
    row.cells.map((cell): OutCell => {
      const text = cell.textBody
        ? renderTextBody(cell.textBody, undefined, ctx)
            .replace(/<[^>]+>/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
        : '';
      return {
        text,
        rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
        colSpan: cell.gridSpan > 1 ? cell.gridSpan : undefined,
        vMerge: cell.vMerge ? 1 : undefined,
        hMerge: cell.hMerge ? 1 : undefined,
        borders: defaultCellBorders(),
      };
    }),
  );
  const rowHeights = node.rows.map((r) => pxToPt(r.height));
  const colWidths = node.columns.map((c) => pxToPt(c));
  return {
    type: 'table',
    left,
    top,
    width,
    height,
    data,
    borders: {},
    order,
    rowHeights,
    colWidths,
  };
}
