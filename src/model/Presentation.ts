/**
 * Top-level presentation builder — assembles parsed components into PresentationData.
 */

import { PptxFiles } from '../parser/ZipParser';
import { parseXml, SafeXmlNode } from '../parser/XmlParser';
import { parseRels, RelEntry, resolveRelTarget } from '../parser/RelParser';
import { emuToPx } from '../parser/units';
import { ThemeData, parseTheme } from './Theme';
import { MasterData, parseMaster } from './Master';
import { LayoutData, parseLayout, PlaceholderEntry } from './Layout';
import { SlideData, SlideNode, parseSlide } from './Slide';
import { Position, Size } from './nodes/BaseNode';

export interface PresentationData {
  width: number;
  height: number;
  slides: SlideData[];
  layouts: Map<string, LayoutData>;
  masters: Map<string, MasterData>;
  themes: Map<string, ThemeData>;
  slideToLayout: Map<number, string>;
  layoutToMaster: Map<string, string>;
  masterToTheme: Map<string, string>;
  media: Map<string, Uint8Array>;
  tableStyles?: SafeXmlNode;
  charts: Map<string, SafeXmlNode>;
  isWps: boolean;
}

function basePath(filePath: string): string {
  const idx = filePath.lastIndexOf('/');
  return idx >= 0 ? filePath.substring(0, idx) : '';
}

function relsPathFor(filePath: string): string {
  const dir = basePath(filePath);
  const fileName = filePath.substring(filePath.lastIndexOf('/') + 1);
  return `${dir}/_rels/${fileName}.rels`;
}

function detectWps(presentationXml: string): boolean {
  return (
    presentationXml.includes('wps') ||
    presentationXml.includes('kso') ||
    presentationXml.includes('Kingsoft') ||
    presentationXml.includes('WPS')
  );
}

function findRelByType(rels: Map<string, RelEntry>, typeSubstring: string): RelEntry | undefined {
  for (const [, entry] of rels) {
    if (entry.type.includes(typeSubstring)) return entry;
  }
  return undefined;
}

function findRelsByType(rels: Map<string, RelEntry>, typeSubstring: string): [string, RelEntry][] {
  const results: [string, RelEntry][] = [];
  for (const [rId, entry] of rels) {
    if (entry.type.includes(typeSubstring)) results.push([rId, entry]);
  }
  return results;
}

function getPhInfo(phNode: SafeXmlNode): { type?: string; idx?: number } {
  for (const wrapper of ['nvSpPr', 'nvPicPr', 'nvGrpSpPr', 'nvGraphicFramePr', 'nvCxnSpPr']) {
    const nvPr = phNode.child(wrapper).child('nvPr');
    const ph = nvPr.child('ph');
    if (ph.exists()) {
      const type = ph.attr('type');
      const idxStr = ph.attr('idx');
      const idx = idxStr !== undefined ? Number(idxStr) : undefined;
      return { type, idx: idx !== undefined && !isNaN(idx) ? idx : undefined };
    }
  }
  return {};
}

function getPhXfrm(phNode: SafeXmlNode): { position: Position; size: Size } | undefined {
  const spPr = phNode.child('spPr');
  if (!spPr.exists()) return undefined;
  const xfrm = spPr.child('xfrm');
  if (!xfrm.exists()) return undefined;
  const off = xfrm.child('off');
  const ext = xfrm.child('ext');
  if (off.numAttr('x') === undefined || ext.numAttr('cx') === undefined) return undefined;
  return {
    position: { x: emuToPx(off.numAttr('x') ?? 0), y: emuToPx(off.numAttr('y') ?? 0) },
    size: { w: emuToPx(ext.numAttr('cx') ?? 0), h: emuToPx(ext.numAttr('cy') ?? 0) },
  };
}

function findMatchingPlaceholder(
  placeholders: SafeXmlNode[],
  type?: string,
  idx?: number,
): SafeXmlNode | undefined {
  let typeMatch: SafeXmlNode | undefined;
  for (const ph of placeholders) {
    const info = getPhInfo(ph);
    if (type !== undefined && info.type === type && idx !== undefined && info.idx === idx) return ph;
    if (type !== undefined && info.type === type && !typeMatch) typeMatch = ph;
    if (idx !== undefined && info.idx === idx && type === undefined && info.type === undefined) return ph;
  }
  if (type === undefined && idx !== undefined) {
    for (const ph of placeholders) {
      if (getPhInfo(ph).idx === idx) return ph;
    }
  }
  return typeMatch;
}

function findMatchingLayoutPlaceholder(
  placeholders: PlaceholderEntry[],
  type?: string,
  idx?: number,
): PlaceholderEntry | undefined {
  let typeMatch: PlaceholderEntry | undefined;
  for (const entry of placeholders) {
    const info = getPhInfo(entry.node);
    if (type !== undefined && info.type === type && idx !== undefined && info.idx === idx) return entry;
    if (type !== undefined && info.type === type && !typeMatch) typeMatch = entry;
    if (idx !== undefined && info.idx === idx && type === undefined && info.type === undefined) return entry;
  }
  if (type === undefined && idx !== undefined) {
    for (const entry of placeholders) {
      if (getPhInfo(entry.node).idx === idx) return entry;
    }
  }
  return typeMatch;
}

function getPhBodyPr(phNode: SafeXmlNode): SafeXmlNode | undefined {
  const txBody = phNode.child('txBody');
  if (!txBody.exists()) return undefined;
  const bodyPr = txBody.child('bodyPr');
  return bodyPr.exists() ? bodyPr : undefined;
}

function resolveNodesPlaceholders(
  nodes: SlideNode[],
  layout: LayoutData | undefined,
  master: MasterData | undefined,
): void {
  for (const node of nodes) {
    if (!node.placeholder) continue;
    const { type, idx } = node.placeholder;
    const sizeIsEmpty = node.size.w === 0 && node.size.h === 0;
    const positionLooksDefault = node.position.y < 5;
    if (layout) {
      const layoutMatch = findMatchingLayoutPlaceholder(layout.placeholders, type, idx);
      if (layoutMatch) {
        const xfrm = layoutMatch.absoluteXfrm ?? getPhXfrm(layoutMatch.node);
        if (xfrm) {
          if (sizeIsEmpty) {
            node.position = xfrm.position;
            node.size = xfrm.size;
          } else if (positionLooksDefault) {
            node.position = xfrm.position;
          }
        }
        if ('textBody' in node && node.textBody) {
          const layoutBodyPr = getPhBodyPr(layoutMatch.node);
          if (layoutBodyPr) node.textBody.layoutBodyProperties = layoutBodyPr;
        }
        if (xfrm) continue;
      }
    }
    if (master) {
      const match = findMatchingPlaceholder(master.placeholders, type, idx);
      if (match) {
        const xfrm = getPhXfrm(match);
        if (xfrm) {
          if (sizeIsEmpty) {
            node.position = xfrm.position;
            node.size = xfrm.size;
          } else if (positionLooksDefault) {
            node.position = xfrm.position;
          }
        }
        if ('textBody' in node && node.textBody && !node.textBody.layoutBodyProperties) {
          const masterBodyPr = getPhBodyPr(match);
          if (masterBodyPr) node.textBody.layoutBodyProperties = masterBodyPr;
        }
      }
    }
  }
}

function resolvePlaceholderInheritance(pres: PresentationData): void {
  for (let i = 0; i < pres.slides.length; i++) {
    const slide = pres.slides[i];
    const layoutPath = pres.slideToLayout.get(i);
    const layout = layoutPath ? pres.layouts.get(layoutPath) : undefined;
    const masterPath = layoutPath ? pres.layoutToMaster.get(layoutPath) : undefined;
    const master = masterPath ? pres.masters.get(masterPath) : undefined;
    resolveNodesPlaceholders(slide.nodes, layout, master);
  }
}

export function buildPresentation(files: PptxFiles): PresentationData {
  const presRoot = parseXml(files.presentation);
  const presRels = parseRels(files.presentationRels);
  const sldSz = presRoot.child('sldSz');
  const width = emuToPx(sldSz.numAttr('cx') ?? 9144000);
  const height = emuToPx(sldSz.numAttr('cy') ?? 6858000);
  const isWps = detectWps(files.presentation);

  const themes = new Map<string, ThemeData>();
  for (const [themePath, themeXml] of files.themes) {
    themes.set(themePath, parseTheme(parseXml(themeXml)));
  }

  const masters = new Map<string, MasterData>();
  const masterToTheme = new Map<string, string>();
  for (const [masterPath, masterXml] of files.slideMasters) {
    const masterData = parseMaster(parseXml(masterXml));
    const masterRelsPath = relsPathFor(masterPath);
    const masterRelsXml = files.slideMasterRels.get(masterRelsPath);
    if (masterRelsXml) {
      masterData.rels = parseRels(masterRelsXml);
      const themeRel = findRelByType(masterData.rels, 'theme');
      if (themeRel) {
        masterToTheme.set(masterPath, resolveRelTarget(basePath(masterPath), themeRel.target));
      }
    }
    masters.set(masterPath, masterData);
  }

  const layouts = new Map<string, LayoutData>();
  const layoutToMaster = new Map<string, string>();
  for (const [layoutPath, layoutXml] of files.slideLayouts) {
    const layoutData = parseLayout(parseXml(layoutXml));
    const layoutRelsPath = relsPathFor(layoutPath);
    const layoutRelsXml = files.slideLayoutRels.get(layoutRelsPath);
    if (layoutRelsXml) {
      layoutData.rels = parseRels(layoutRelsXml);
      const masterRel = findRelByType(layoutData.rels, 'slideMaster');
      if (masterRel) {
        layoutToMaster.set(layoutPath, resolveRelTarget(basePath(layoutPath), masterRel.target));
      }
    }
    layouts.set(layoutPath, layoutData);
  }

  const charts = new Map<string, SafeXmlNode>();
  for (const [chartPath, chartXml] of files.charts) {
    const chartRoot = parseXml(chartXml);
    if (chartRoot.exists()) charts.set(chartPath, chartRoot);
  }

  const sldIdLst = presRoot.child('sldIdLst');
  const orderedSlideTargets: string[] = [];
  for (const sldId of sldIdLst.children('sldId')) {
    const rId = sldId.attr('r:id') ?? sldId.attr('id');
    if (rId) {
      const relEntry = presRels.get(rId);
      if (relEntry) {
        orderedSlideTargets.push(resolveRelTarget('ppt', relEntry.target));
      }
    }
  }
  if (orderedSlideTargets.length === 0) {
    const slideRels = findRelsByType(presRels, 'slide');
    slideRels.sort((a, b) => {
      const numA = parseInt(a[0].replace(/\D/g, ''), 10) || 0;
      const numB = parseInt(b[0].replace(/\D/g, ''), 10) || 0;
      return numA - numB;
    });
    for (const [, entry] of slideRels) {
      if (
        entry.type.includes('/slide') &&
        !entry.type.includes('slideLayout') &&
        !entry.type.includes('slideMaster')
      ) {
        orderedSlideTargets.push(resolveRelTarget('ppt', entry.target));
      }
    }
  }

  const slides: SlideData[] = [];
  const slideToLayout = new Map<number, string>();
  for (let i = 0; i < orderedSlideTargets.length; i++) {
    const slidePath = orderedSlideTargets[i];
    const slideXml = files.slides.get(slidePath);
    if (!slideXml) continue;
    const slideRelsPath = relsPathFor(slidePath);
    const slideRelsXml = files.slideRels.get(slideRelsPath);
    const slideRels = slideRelsXml ? parseRels(slideRelsXml) : new Map<string, RelEntry>();
    const slideData = parseSlide(parseXml(slideXml), i, slideRels, slidePath, files.diagramDrawings);
    if (slideData.layoutIndex) {
      slideData.layoutIndex = resolveRelTarget(basePath(slidePath), slideData.layoutIndex);
      slideToLayout.set(i, slideData.layoutIndex);
    }
    slides.push(slideData);
  }

  let tableStyles: SafeXmlNode | undefined;
  if (files.tableStyles) {
    const tsRoot = parseXml(files.tableStyles);
    if (tsRoot.exists()) tableStyles = tsRoot;
  }

  const result: PresentationData = {
    width,
    height,
    slides,
    layouts,
    masters,
    themes,
    slideToLayout,
    layoutToMaster,
    masterToTheme,
    media: files.media,
    tableStyles,
    charts,
    isWps,
  };
  resolvePlaceholderInheritance(result);
  return result;
}
