/**
 * Adapter: PresentationData + PptxFiles → pptxtojson/PPTist output format.
 * All dimensions in output are in pt (px * 0.75).
 */

import type { PresentationData } from '../model/Presentation';
import type { SlideData } from '../model/Slide';
import type { SlideNode } from '../model/Slide';
import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import type { PptxFiles } from '../parser/ZipParser';
import type {
  Output,
  Slide,
  Element,
  Fill,
  Size,
} from './types';
import { createRenderContext } from '../resolve/RenderContext';
import { resolveFill, resolveLineStyle, resolveThemeFillReference } from '../resolve/StyleResolver';
import { getPresetShapePath } from '../shapes/presets';
import { renderCustomGeometry } from '../shapes/customGeometry';
import { textToHtml } from './textToHtml';
import { parseChildNode } from '../model/Slide';
import { parseXml, type SafeXmlNode } from '../parser/XmlParser';
import { resolveRelTarget } from '../parser/RelParser';

const PX_TO_PT = 0.75;

function pxToPt(px: number): number {
  return px * PX_TO_PT;
}

function getThemeColors(presentation: PresentationData): string[] {
  const themeColors: string[] = [];
  const firstTheme = presentation.themes.values().next().value;
  if (!firstTheme) return ['#000000', '#000000', '#000000', '#000000', '#000000', '#000000'];
  for (let i = 1; i <= 6; i++) {
    const hex = firstTheme.colorScheme.get(`accent${i}`) ?? '000000';
    themeColors.push(hex.startsWith('#') ? hex : `#${hex}`);
  }
  return themeColors;
}

function defaultSlideFill(): Fill {
  return { type: 'color', value: '#ffffff' };
}

function resolveSlideFill(
  slide: SlideData,
  ctx: ReturnType<typeof createRenderContext>,
): Fill {
  const candidates: (SafeXmlNode | undefined)[] = [
    slide.background,
    ctx.layout.background,
    ctx.master.background,
  ];
  for (const bg of candidates) {
    if (!bg?.exists()) continue;
    const bgPr = bg.child('bgPr');
    if (bgPr.exists()) {
      const fillCss = resolveFill(bgPr, ctx);
      if (fillCss && fillCss !== 'transparent') {
        const fill = cssToPptxtojsonFill(fillCss);
        if (fill) return fill;
      }
    }
    const bgRef = bg.child('bgRef');
    if (bgRef.exists()) {
      const { fillCss } = resolveThemeFillReference(bgRef, ctx);
      if (fillCss && fillCss !== 'transparent') {
        const fill = cssToPptxtojsonFill(fillCss);
        if (fill) return fill;
      }
    }
  }
  return defaultSlideFill();
}

function getNoteForSlide(slide: SlideData, files: PptxFiles): string | undefined {
  for (const [, entry] of slide.rels) {
    if (!entry.type.includes('notesSlide')) continue;
    const basePath = slide.slidePath.replace(/\/[^/]+$/, '');
    const notesPath = resolveRelTarget(basePath, entry.target);
    const notesXml = files.notesSlides.get(notesPath);
    if (!notesXml) continue;
    const root = parseXml(notesXml);
    const cSld = root.child('cSld');
    if (!cSld.exists()) continue;
    const spTree = cSld.child('spTree');
    const parts: string[] = [];
    for (const sp of spTree.allChildren()) {
      if (sp.localName !== 'sp') continue;
      const nvPr = sp.child('nvSpPr').child('nvPr');
      const ph = nvPr.child('ph');
      if (ph.attr('type') !== 'body') continue;
      const txBody = sp.child('txBody');
      if (!txBody.exists()) continue;
      for (const p of txBody.children('p')) {
        for (const r of p.children('r')) {
          const t = r.child('t');
          parts.push(t.text());
        }
      }
    }
    return parts.length > 0 ? parts.join('').trim() : undefined;
  }
  return undefined;
}

function getTransitionForSlide(slide: SlideData, files: PptxFiles): Slide['transition'] {
  const slideXml = files.slides.get(slide.slidePath);
  if (!slideXml) return undefined;
  const root = parseXml(slideXml);
  const transition = root.child('transition');
  if (!transition.exists()) return undefined;
  let type = 'none';
  let duration = 1000;
  let direction: string | null = null;
  for (const child of transition.allChildren()) {
    if (child.localName && child.localName !== 'p14:transition') {
      type = child.localName.replace(/^[a-z0-9]+:/, '') || child.localName;
      const dur = child.numAttr('dur');
      if (dur !== undefined) duration = dur;
      const dir = child.attr('dir');
      if (dir) direction = dir;
      break;
    }
  }
  const spd = transition.attr('spd');
  if (spd === 'fast') duration = 500;
  else if (spd === 'med') duration = 800;
  else if (spd === 'slow') duration = 1000;
  return { type, duration, direction };
}

function cssToPptxtojsonFill(css: string): Fill | undefined {
  if (!css || css === 'transparent') return undefined;
  if (css.includes('gradient')) {
    return { type: 'gradient', value: css } as Fill;
  }
  return { type: 'color', value: css };
}

function getShapePath(node: ShapeNodeData): string {
  const w = node.size.w;
  const h = node.size.h;
  if (node.customGeometry?.exists()) {
    return renderCustomGeometry(node.customGeometry, w, h);
  }
  const preset = node.presetGeometry || 'rect';
  return getPresetShapePath(preset, w, h, node.adjustments);
}

function nodeToElement(
  node: SlideNode,
  ctx: ReturnType<typeof createRenderContext>,
  order: number,
  files?: PptxFiles,
): Element {
  const left = pxToPt(node.position.x);
  const top = pxToPt(node.position.y);
  const width = pxToPt(node.size.w);
  const height = pxToPt(node.size.h);
  const base = { left, top, width, height, name: node.name || undefined, order };

  switch (node.nodeType) {
    case 'shape': {
      const spPr = node.source.child('spPr');
      const fillCss = spPr.exists() ? resolveFill(spPr, ctx) : '';
      const fill = fillCss ? cssToPptxtojsonFill(fillCss) : undefined;
      let borderColor: string | undefined;
      let borderWidth: number | undefined;
      let borderType: string | undefined;
      if (node.line?.exists()) {
        const line = resolveLineStyle(node.line, ctx);
        borderColor = line.color !== 'transparent' ? line.color : undefined;
        borderWidth = line.width > 0 ? pxToPt(line.width) : undefined;
        borderType = line.dashKind !== 'solid' ? line.dashKind : undefined;
      }
      const path = getShapePath(node);
      const content = node.textBody ? textToHtml(ctx, node.textBody) : undefined;
      const hasContent = content && content.replace(/<[^>]+>/g, '').trim().length > 0;
      if (hasContent && content) {
        return {
          ...base,
          type: 'text',
          content,
          fill,
          borderColor,
          borderWidth,
          borderType,
          rotate: node.rotation || undefined,
          isFlipH: node.flipH || undefined,
          isFlipV: node.flipV || undefined,
        };
      }
      return {
        ...base,
        type: 'shape',
        shapType: node.presetGeometry || 'rect',
        path: path || undefined,
        content,
        fill,
        borderColor,
        borderWidth,
        borderType,
        rotate: node.rotation || undefined,
        isFlipH: node.flipH || undefined,
        isFlipV: node.flipV || undefined,
      };
    }
    case 'picture':
      return {
        ...base,
        type: node.isVideo ? 'video' : node.isAudio ? 'audio' : 'image',
        src: '',
        rotate: node.rotation || undefined,
        isFlipH: node.flipH || undefined,
        isFlipV: node.flipV || undefined,
      };
    case 'table': {
      const data = node.rows.map((row) =>
        row.cells.map((cell) => ({
          text: cell.textBody ? textToHtml(ctx, cell.textBody).replace(/<[^>]+>/g, '').trim() : '',
          fillColor: '',
          borders: {} as Record<string, unknown>,
        })),
      );
      return {
        ...base,
        type: 'table',
        data,
        rowHeights: node.rows.map((r) => pxToPt(r.height)),
        colWidths: node.columns.map((c) => pxToPt(c)),
      };
    }
    case 'group': {
      const childElements: Element[] = [];
      const diagramDrawings = files?.diagramDrawings;
      for (let i = 0; i < node.children.length; i++) {
        const childNode = parseChildNode(
          node.children[i],
          ctx.slide.rels,
          ctx.slide.slidePath,
          diagramDrawings,
        );
        if (childNode) {
          childElements.push(nodeToElement(childNode, ctx, i, files));
        }
      }
      return {
        ...base,
        type: 'group',
        elements: childElements,
        rotate: node.rotation || undefined,
        isFlipH: node.flipH || undefined,
        isFlipV: node.flipV || undefined,
      };
    }
    case 'chart': {
      let chartType: string | undefined;
      let chartData: unknown = null;
      const chartRoot = ctx.presentation.charts.get(node.chartPath);
      if (chartRoot?.exists()) {
        const chartSpace = chartRoot.child('chartSpace');
        const chart = chartSpace.child('chart');
        if (chart.exists()) {
          const plotArea = chart.child('plotArea');
          const firstChart = plotArea.allChildren()[0];
          if (firstChart?.localName) {
            chartType = firstChart.localName.replace('Chart', '').toLowerCase();
          }
        }
      }
      return {
        ...base,
        type: 'chart',
        data: chartData,
        chartType,
      };
    }
    default:
      return {
        ...base,
        type: 'shape',
        shapType: 'rect',
      };
  }
}

function slideToPptxtojsonSlide(
  presentation: PresentationData,
  slide: SlideData,
  files: PptxFiles,
): Slide {
  const ctx = createRenderContext(presentation, slide);
  const elements: Element[] = slide.nodes.map((node, i) =>
    nodeToElement(node, ctx, i, files),
  );
  return {
    fill: resolveSlideFill(slide, ctx),
    elements,
    layoutElements: [],
    note: getNoteForSlide(slide, files),
    transition: getTransitionForSlide(slide, files),
  };
}

export function toPptxtojsonFormat(
  presentation: PresentationData,
  files: PptxFiles,
): Output {
  const size: Size = {
    width: pxToPt(presentation.width),
    height: pxToPt(presentation.height),
  };
  const themeColors = getThemeColors(presentation);
  const slides: Slide[] = presentation.slides.map((slide) =>
    slideToPptxtojsonSlide(presentation, slide, files),
  );
  return {
    slides,
    themeColors,
    size,
  };
}
