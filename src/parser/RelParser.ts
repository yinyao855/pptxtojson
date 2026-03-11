/**
 * Parser for .rels (Relationship) XML files in OOXML packages.
 * These files map relationship IDs (rId1, rId2, ...) to targets.
 */

import { parseXml } from './XmlParser';

export interface RelEntry {
  type: string;
  target: string;
  targetMode?: string;
}

/**
 * Parse a .rels XML string into a Map of relationship ID -> RelEntry.
 */
export function parseRels(xmlString: string): Map<string, RelEntry> {
  const result = new Map<string, RelEntry>();

  if (!xmlString) return result;

  const root = parseXml(xmlString);
  if (!root.exists()) return result;

  const relationships = root.children('Relationship');
  for (const rel of relationships) {
    const id = rel.attr('Id');
    const type = rel.attr('Type');
    const target = rel.attr('Target');
    const targetMode = rel.attr('TargetMode');

    if (id && type !== undefined && target !== undefined) {
      result.set(id, { type, target, targetMode });
    }
  }

  return result;
}

/**
 * Resolve a relative target path against a base path.
 */
export function resolveRelTarget(basePath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1);
  }

  const baseParts = basePath.replace(/\\/g, '/').split('/').filter(Boolean);
  const targetParts = target.replace(/\\/g, '/').split('/').filter(Boolean);

  const resolved = [...baseParts];
  for (const part of targetParts) {
    if (part === '..') {
      resolved.pop();
    } else if (part !== '.') {
      resolved.push(part);
    }
  }

  return resolved.join('/');
}
