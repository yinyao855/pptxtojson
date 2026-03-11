/**
 * Chart node — represents a chart embedded in a graphicFrame element.
 */

import { SafeXmlNode } from '../../parser/XmlParser';
import { RelEntry, resolveRelTarget } from '../../parser/RelParser';
import { BaseNodeData, parseBaseProps } from './BaseNode';

export interface ChartNodeData extends BaseNodeData {
  nodeType: 'chart';
  chartPath: string;
}

export function parseChartNode(
  graphicFrame: SafeXmlNode,
  slideRels: Map<string, RelEntry>,
  slidePath: string,
): ChartNodeData | undefined {
  const base = parseBaseProps(graphicFrame);
  const graphic = graphicFrame.child('graphic');
  const graphicData = graphic.child('graphicData');
  let chartRId: string | undefined;
  for (const child of graphicData.allChildren()) {
    if (child.localName === 'chart') {
      chartRId = child.attr('r:id') || child.attr('id');
      break;
    }
  }
  if (!chartRId) return undefined;
  const rel = slideRels.get(chartRId);
  if (!rel) return undefined;
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
  const chartPath = resolveRelTarget(slideDir, rel.target);
  return {
    ...base,
    nodeType: 'chart',
    chartPath,
  };
}
