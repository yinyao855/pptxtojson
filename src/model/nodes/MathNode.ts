/**
 * Math node — represents a math formula embedded via mc:AlternateContent.
 *
 * PPTX stores math formulas as:
 *   mc:AlternateContent
 *     mc:Choice (Requires="a14") → p:sp with m:oMathPara in txBody
 *     mc:Fallback → p:sp with blipFill (static image preview)
 */

import { SafeXmlNode } from '../../parser/XmlParser';
import { BaseNodeData, parseBaseProps } from './BaseNode';

export interface MathNodeData extends BaseNodeData {
  nodeType: 'math';
  /** Serialized OMML XML string (m:oMathPara or m:oMath element). */
  ommlXml: string;
  /** r:embed of fallback image from mc:Fallback branch. */
  fallbackBlipEmbed?: string;
  /** Plain text extracted from m:t elements inside the OMML. */
  plainText: string;
}

/**
 * Recursively search for an element with localName 'oMathPara' or 'oMath'.
 */
function findOmmlNode(node: SafeXmlNode): SafeXmlNode | null {
  if (node.localName === 'oMathPara' || node.localName === 'oMath') return node;
  for (const child of node.allChildren()) {
    const found = findOmmlNode(child);
    if (found) return found;
  }
  return null;
}

/**
 * Recursively collect all text from m:t elements.
 */
function collectMathText(node: SafeXmlNode): string {
  if (node.localName === 't') return node.text();
  const parts: string[] = [];
  for (const child of node.allChildren()) {
    parts.push(collectMathText(child));
  }
  return parts.join('');
}

/**
 * Serialize a SafeXmlNode's underlying DOM Element to an XML string.
 */
function serializeElement(node: SafeXmlNode): string {
  const el = node.rawElement;
  if (!el) return '';
  // @xmldom/xmldom Element supports toString() which returns outerHTML-equivalent XML
  return el.toString();
}

/**
 * Detect whether an mc:AlternateContent node contains a math formula.
 * Math formulas have mc:Choice with p:sp > p:txBody containing m:oMathPara/m:oMath.
 */
export function isMathAlternateContent(altContent: SafeXmlNode): boolean {
  const choice = altContent.child('Choice');
  if (!choice.exists()) return false;
  const sp = choice.child('sp');
  if (!sp.exists()) return false;
  const txBody = sp.child('txBody');
  if (!txBody.exists()) return false;
  return findOmmlNode(txBody) !== null;
}

/**
 * Parse an mc:AlternateContent node containing a math formula into MathNodeData.
 */
export function parseMathNode(altContent: SafeXmlNode): MathNodeData | undefined {
  const choice = altContent.child('Choice');
  if (!choice.exists()) return undefined;

  const sp = choice.child('sp');
  if (!sp.exists()) return undefined;

  const txBody = sp.child('txBody');
  if (!txBody.exists()) return undefined;

  const ommlNode = findOmmlNode(txBody);
  if (!ommlNode) return undefined;

  // Use the sp from Choice for position/size (it has the xfrm)
  const base = parseBaseProps(sp);
  const ommlXml = serializeElement(ommlNode);
  const plainText = collectMathText(ommlNode);

  // Extract fallback image embed from mc:Fallback > p:sp > p:spPr > a:blipFill > a:blip
  let fallbackBlipEmbed: string | undefined;
  const fallback = altContent.child('Fallback');
  if (fallback.exists()) {
    const fbSp = fallback.child('sp');
    if (fbSp.exists()) {
      const blip = fbSp.child('spPr').child('blipFill').child('blip');
      if (blip.exists()) {
        fallbackBlipEmbed = blip.attr('embed') ?? blip.attr('r:embed');
      }
    }
  }

  return {
    ...base,
    nodeType: 'math' as const,
    ommlXml,
    fallbackBlipEmbed,
    plainText,
  };
}
