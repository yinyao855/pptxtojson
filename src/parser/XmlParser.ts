/**
 * Safe XML parser — uses browser DOMParser when available, falls back to
 * @xmldom/xmldom in Node.js environments.
 * All operations are null-safe — accessing missing elements never crashes.
 */

import { DOMParser as XmlDOMParser } from '@xmldom/xmldom';

/** Iterate only Element children (nodeType === 1) from childNodes.
 *  Works in both browser DOM and @xmldom/xmldom (which may lack `.children`). */
function elementChildren(el: Element): Element[] {
  const out: Element[] = [];
  const nodes = el.childNodes;
  for (let i = 0; i < nodes.length; i++) {
    if (nodes[i].nodeType === 1) out.push(nodes[i] as Element);
  }
  return out;
}

export class SafeXmlNode {
  private readonly el: Element | null;

  constructor(el: Element | null) {
    this.el = el;
  }

  /** Expose raw DOM Element (needed for document-order computation). */
  get rawElement(): Element | null {
    return this.el;
  }

  /** Get a string attribute value, or undefined if missing. */
  attr(name: string): string | undefined {
    if (!this.el) return undefined;
    return this.el.hasAttribute(name) ? this.el.getAttribute(name)! : undefined;
  }

  /** Get a numeric attribute value, or undefined if missing or not a number. */
  numAttr(name: string): number | undefined {
    const raw = this.attr(name);
    if (raw === undefined) return undefined;
    const n = Number(raw);
    return Number.isNaN(n) ? undefined : n;
  }

  /**
   * Find the first child element matching the given localName (namespace-agnostic).
   * Returns an empty SafeXmlNode if not found, so chaining never crashes.
   */
  child(localName: string): SafeXmlNode {
    if (!this.el) return new SafeXmlNode(null);
    const children = elementChildren(this.el);
    for (let i = 0; i < children.length; i++) {
      if (children[i].localName === localName) {
        return new SafeXmlNode(children[i]);
      }
    }
    return new SafeXmlNode(null);
  }

  /**
   * Get child elements, optionally filtered by localName (namespace-agnostic).
   * If no localName is given, returns all direct child elements.
   */
  children(localName?: string): SafeXmlNode[] {
    if (!this.el) return [];
    const result: SafeXmlNode[] = [];
    const children = elementChildren(this.el);
    for (let i = 0; i < children.length; i++) {
      if (localName === undefined || children[i].localName === localName) {
        result.push(new SafeXmlNode(children[i]));
      }
    }
    return result;
  }

  /** Get the text content, or empty string if the element is missing. */
  text(): string {
    if (!this.el) return '';
    return this.el.textContent ?? '';
  }

  /** Whether the underlying element actually exists. */
  exists(): boolean {
    return this.el !== null;
  }

  /** All direct child elements as SafeXmlNode[]. */
  allChildren(): SafeXmlNode[] {
    return this.children();
  }

  /** The localName of the underlying element, or empty string. */
  get localName(): string {
    return this.el?.localName ?? '';
  }

  /** Raw access to the underlying Element (may be null). */
  get element(): Element | null {
    return this.el;
  }
}

// ---------------------------------------------------------------------------
// DOMParser resolution — browser built-in or @xmldom/xmldom fallback
// ---------------------------------------------------------------------------

type DOMParserLike = { parseFromString(s: string, mime: string): Document };
let _cachedParser: DOMParserLike | null = null;
let _resolvePromise: Promise<void> | null = null;

/**
 * Eagerly initialise the DOMParser polyfill (Node.js only).
 * Call once at startup; subsequent `parseXml` calls are synchronous.
 */
export async function initDOMParser(): Promise<void> {
  if (_cachedParser) return;
  if (typeof DOMParser !== 'undefined') {
    _cachedParser = new DOMParser();
    return;
  }
  _cachedParser = new XmlDOMParser() as unknown as DOMParserLike;
}

function getDOMParser(): DOMParserLike {
  if (_cachedParser) return _cachedParser;
  if (typeof DOMParser !== 'undefined') {
    _cachedParser = new DOMParser();
    return _cachedParser;
  }
  throw new Error(
    'DOMParser is not available. Call `await initDOMParser()` before parsing, or install @xmldom/xmldom for Node.js.',
  );
}

/**
 * Ensure the DOMParser is ready (lazy init). Called internally by `parseXml`.
 */
function ensureDOMParser(): DOMParserLike {
  if (_cachedParser) return _cachedParser;
  if (typeof DOMParser !== 'undefined') {
    _cachedParser = new DOMParser();
    return _cachedParser;
  }
  _cachedParser = new XmlDOMParser() as unknown as DOMParserLike;
  return _cachedParser;
}

/**
 * Parse an XML string into a SafeXmlNode wrapping the document element.
 * Uses browser DOMParser when available; falls back to @xmldom/xmldom in Node.js.
 */
export function parseXml(xmlString: string): SafeXmlNode {
  const parser = ensureDOMParser();
  const doc = parser.parseFromString(xmlString, 'application/xml');

  // Check for parser errors — browser DOMParser returns a parsererror document on failure
  const errorNode = doc.getElementsByTagName('parsererror');
  if (errorNode.length > 0) {
    console.warn('XML parse error:', errorNode[0].textContent);
    return new SafeXmlNode(null);
  }

  return new SafeXmlNode(doc.documentElement);
}
