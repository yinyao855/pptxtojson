/**
 * Slide master parser — extracts color map, background, text styles,
 * and placeholder shapes from a p:sldMaster XML.
 */

import { SafeXmlNode } from '../parser/XmlParser';
import type { RelEntry } from '../parser/RelParser';

export interface MasterData {
  colorMap: Map<string, string>;
  background?: SafeXmlNode;
  textStyles: {
    titleStyle?: SafeXmlNode;
    bodyStyle?: SafeXmlNode;
    otherStyle?: SafeXmlNode;
  };
  defaultTextStyle?: SafeXmlNode;
  placeholders: SafeXmlNode[];
  spTree: SafeXmlNode;
  rels: Map<string, RelEntry>;
}

function isPlaceholder(node: SafeXmlNode): boolean {
  const nvSpPr = node.child('nvSpPr');
  if (nvSpPr.exists()) {
    const nvPr = nvSpPr.child('nvPr');
    if (nvPr.child('ph').exists()) return true;
  }
  const nvPicPr = node.child('nvPicPr');
  if (nvPicPr.exists()) {
    const nvPr = nvPicPr.child('nvPr');
    if (nvPr.child('ph').exists()) return true;
  }
  return false;
}

function extractPlaceholders(spTree: SafeXmlNode): SafeXmlNode[] {
  const placeholders: SafeXmlNode[] = [];
  for (const child of spTree.allChildren()) {
    if (isPlaceholder(child)) placeholders.push(child);
  }
  return placeholders;
}

function parseAllAttributes(node: SafeXmlNode): Map<string, string> {
  const result = new Map<string, string>();
  const el = node.element;
  if (!el) return result;
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    result.set(attr.localName, attr.value);
  }
  return result;
}

export function parseMaster(root: SafeXmlNode): MasterData {
  const cSld = root.child('cSld');
  const bg = cSld.child('bg');
  const background = bg.exists() ? bg : undefined;
  const spTree = cSld.child('spTree');
  const clrMap = root.child('clrMap');
  const colorMap = parseAllAttributes(clrMap);
  const txStyles = root.child('txStyles');
  const titleStyle = txStyles.child('titleStyle');
  const bodyStyle = txStyles.child('bodyStyle');
  const otherStyle = txStyles.child('otherStyle');
  const defaultTextStyle = root.child('defaultTextStyle');
  const placeholders = extractPlaceholders(spTree);
  return {
    colorMap,
    background,
    textStyles: {
      titleStyle: titleStyle.exists() ? titleStyle : undefined,
      bodyStyle: bodyStyle.exists() ? bodyStyle : undefined,
      otherStyle: otherStyle.exists() ? otherStyle : undefined,
    },
    defaultTextStyle: defaultTextStyle.exists() ? defaultTextStyle : undefined,
    placeholders,
    spTree,
    rels: new Map(),
  };
}
