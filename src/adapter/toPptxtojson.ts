/**
 * Adapter: PresentationData + PptxFiles → pptxtojson/PPTist output format.
 * All dimensions in output are in pt (px * 0.75).
 * Delegates slide serialization to the serializer layer (slideToSlide).
 */

import type { PresentationData } from '../model/Presentation';
import type { PptxFiles } from '../parser/ZipParser';
import type { Output, Slide, Size } from './types';
import { slideToSlide } from '../serializer/slideSerializer';

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
    slideToSlide(presentation, slide, files),
  );
  return {
    slides,
    themeColors,
    size,
  };
}
