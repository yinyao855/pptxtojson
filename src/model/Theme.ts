/**
 * Theme parser — extracts color scheme and font definitions from a:theme XML.
 */

import { SafeXmlNode } from '../parser/XmlParser';

export interface ThemeData {
  colorScheme: Map<string, string>;
  majorFont: { latin: string; ea: string; cs: string };
  minorFont: { latin: string; ea: string; cs: string };
  fillStyles: SafeXmlNode[];
  lineStyles: SafeXmlNode[];
  effectStyles: SafeXmlNode[];
}

const COLOR_SLOTS = [
  'dk1', 'dk2', 'lt1', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
] as const;

function extractColor(node: SafeXmlNode): string | undefined {
  const srgb = node.child('srgbClr');
  if (srgb.exists()) return srgb.attr('val');
  const sys = node.child('sysClr');
  if (sys.exists()) return sys.attr('lastClr') ?? sys.attr('val');
  return undefined;
}

function parseFontInfo(fontNode: SafeXmlNode): { latin: string; ea: string; cs: string } {
  return {
    latin: fontNode.child('latin').attr('typeface') ?? '',
    ea: fontNode.child('ea').attr('typeface') ?? '',
    cs: fontNode.child('cs').attr('typeface') ?? '',
  };
}

export function parseTheme(root: SafeXmlNode): ThemeData {
  const themeElements = root.child('themeElements');
  const clrScheme = themeElements.child('clrScheme');
  const colorScheme = new Map<string, string>();
  for (const slot of COLOR_SLOTS) {
    const slotNode = clrScheme.child(slot);
    if (slotNode.exists()) {
      const hex = extractColor(slotNode);
      if (hex !== undefined) colorScheme.set(slot, hex);
    }
  }
  const fontScheme = themeElements.child('fontScheme');
  const majorFont = parseFontInfo(fontScheme.child('majorFont'));
  const minorFont = parseFontInfo(fontScheme.child('minorFont'));
  const fmtScheme = themeElements.child('fmtScheme');
  const fillStyleLst = fmtScheme.child('fillStyleLst');
  const fillStyles: SafeXmlNode[] = fillStyleLst.allChildren();
  const lnStyleLst = fmtScheme.child('lnStyleLst');
  const lineStyles: SafeXmlNode[] = lnStyleLst.allChildren();
  const effectStyleLst = fmtScheme.child('effectStyleLst');
  const effectStyles: SafeXmlNode[] = effectStyleLst.allChildren();
  return { colorScheme, majorFont, minorFont, fillStyles, lineStyles, effectStyles };
}
