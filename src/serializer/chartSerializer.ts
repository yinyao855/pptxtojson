/**
 * Serializes ChartNodeData to pptxtojson CommonChart or ScatterChart.
 * Reads chart XML from presentation.charts; extracts chartType and minimal data.
 */

import type { ChartNodeData } from '../model/nodes/ChartNode';
import type { RenderContext } from '../resolve/RenderContext';
import type { ChartType, CommonChart, ScatterChart, ChartItem, ChartValue } from '../adapter/types';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

const OOXML_CHART_TYPES: string[] = [
  'lineChart', 'line3DChart', 'barChart', 'bar3DChart', 'pieChart', 'pie3DChart',
  'doughnutChart', 'areaChart', 'area3DChart', 'scatterChart', 'bubbleChart',
  'radarChart', 'stockChart', 'surfaceChart', 'surface3DChart',
];

function mapToChartType(ooxmlName: string): ChartType {
  const t = ooxmlName as ChartType;
  if (['scatterChart', 'bubbleChart'].includes(ooxmlName)) return t;
  if (['lineChart', 'line3DChart', 'barChart', 'bar3DChart', 'pieChart', 'pie3DChart',
    'doughnutChart', 'areaChart', 'area3DChart', 'radarChart', 'stockChart', 'surfaceChart', 'surface3DChart'].includes(ooxmlName)) {
    return t;
  }
  return 'barChart';
}

/**
 * Get theme accent colors for chart (hex strings).
 */
function getThemeColors(ctx: RenderContext): string[] {
  const colors: string[] = [];
  for (let i = 1; i <= 6; i++) {
    const hex = ctx.theme.colorScheme.get(`accent${i}`) ?? '000000';
    colors.push(hex.startsWith('#') ? hex : `#${hex}`);
  }
  return colors;
}

/**
 * Serialize chart node to Chart element.
 * Uses chart XML from presentation.charts for chartType; data/colors minimal.
 */
export function chartToElement(
  node: ChartNodeData,
  ctx: RenderContext,
  order: number,
): CommonChart | ScatterChart {
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);
  const chartRoot = ctx.presentation.charts.get(node.chartPath);
  let chartType: ChartType = 'barChart';
  if (chartRoot?.exists()) {
    const chart = chartRoot.child('chart');
    const plotArea = chart.exists() ? chart.child('plotArea') : chartRoot.child('plotArea');
    if (plotArea.exists()) {
      for (const name of OOXML_CHART_TYPES) {
        const el = plotArea.child(name);
        if (el.exists()) {
          chartType = mapToChartType(name);
          break;
        }
      }
    }
  }
  const colors = getThemeColors(ctx);
  if (chartType === 'scatterChart' || chartType === 'bubbleChart') {
    return {
      type: 'chart',
      left,
      top,
      width,
      height,
      data: [[], []],
      colors,
      chartType,
      order,
    } as ScatterChart;
  }
  return {
    type: 'chart',
    left,
    top,
    width,
    height,
    data: [] as ChartItem[],
    colors,
    chartType: chartType as CommonChart['chartType'],
    order,
  };
}
