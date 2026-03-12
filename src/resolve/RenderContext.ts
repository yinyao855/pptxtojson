/**
 * Render context — provides resolved theme/master/layout chain for a given slide.
 */

import { PresentationData } from '../model/Presentation';
import { SlideData } from '../model/Slide';
import { ThemeData } from '../model/Theme';
import { MasterData } from '../model/Master';
import { LayoutData } from '../model/Layout';
import { SafeXmlNode } from '../parser/XmlParser';

export interface RenderContext {
  presentation: PresentationData;
  slide: SlideData;
  theme: ThemeData;
  master: MasterData;
  layout: LayoutData;
  mediaUrlCache: Map<string, string>;
  colorCache: Map<string, { color: string; alpha: number }>;
  groupFillNode?: SafeXmlNode;
  onNavigate?: (target: { slideIndex?: number; url?: string }) => void;
}

export function createRenderContext(
  presentation: PresentationData,
  slide: SlideData,
  mediaUrlCache?: Map<string, string>,
): RenderContext {
  const layoutPath = presentation.slideToLayout.get(slide.index) || '';
  const masterPath = presentation.layoutToMaster.get(layoutPath) || '';
  const themePath = presentation.masterToTheme.get(masterPath) || '';

  const layout: LayoutData = presentation.layouts.get(layoutPath) || ({
    placeholders: [],
    spTree: new SafeXmlNode(null),
    rels: new Map(),
    showMasterSp: true,
  } as unknown as LayoutData);

  const master: MasterData = presentation.masters.get(masterPath) || ({
    colorMap: new Map(),
    textStyles: {},
    placeholders: [],
    spTree: new SafeXmlNode(null),
    rels: new Map(),
  } as unknown as MasterData);

  const theme: ThemeData = presentation.themes.get(themePath) || {
    colorScheme: new Map(),
    majorFont: { latin: 'Calibri', ea: '', cs: '' },
    minorFont: { latin: 'Calibri', ea: '', cs: '' },
    fillStyles: [],
    lineStyles: [],
    effectStyles: [],
  };

  return {
    presentation,
    slide,
    theme,
    master,
    layout,
    mediaUrlCache: mediaUrlCache ?? new Map(),
    colorCache: new Map(),
  };
}
