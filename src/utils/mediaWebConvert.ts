/**
 * Convert legacy / non-web-safe image bytes (TIFF, EMF bitmap, etc.) to PNG data URLs
 * so JSON output works in browsers (PPTist). Mirrors pptx-renderer ImageRenderer strategy.
 */

import UTIF from 'utif';
import { parseEmfContent } from './emfParser';
import { rgbaToPngDataUrl } from './rgbaToPng';
import { getMimeType, toDataUrl } from './media';

type UtifPage = {
  width: number;
  height: number;
  data?: Uint8Array;
  [key: string]: unknown;
};

function arrayBufferToBase64(data: Uint8Array): string {
  let binary = '';
  const len = data.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(data[i]);
  }
  if (typeof btoa !== 'undefined') return btoa(binary);
  const NodeBuffer = (typeof globalThis !== 'undefined' &&
    (globalThis as unknown as { Buffer?: { from(a: Uint8Array): { toString(e: string): string } } }).Buffer);
  if (NodeBuffer) return NodeBuffer.from(data).toString('base64');
  return btoa(binary);
}

function extOf(path: string): string {
  return path.split('.').pop()?.toLowerCase() || '';
}

/**
 * Decode TIFF/TIF bytes to RGBA using UTIF.
 */
function tiffToRgba(data: Uint8Array): { width: number; height: number; data: Uint8ClampedArray } | null {
  try {
    const ifds = UTIF.decode(data) as UtifPage[];
    if (!ifds.length) return null;
    UTIF.decodeImage(data, ifds[0], ifds);
    const page = ifds[0];
    const w = page.width;
    const h = page.height;
    if (!w || !h) return null;
    const rgba = UTIF.toRGBA8(page);
    return { width: w, height: h, data: new Uint8ClampedArray(rgba) };
  } catch {
    return null;
  }
}

/** 1×1 transparent PNG — fallback when conversion is not possible */
const TRANSPARENT_PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/**
 * If the media type is not reliably displayable in browsers, convert to PNG data URL.
 * Otherwise return `data:<mime>;base64,<raw>` for the original bytes.
 *
 * - TIFF/TIF: decode → PNG
 * - EMF: embedded DIB → PNG; embedded PDF → transparent placeholder (full PDF render needs pdfjs)
 * - WMF: not supported → transparent placeholder
 * - PNG/JPEG/GIF/WebP/BMP/SVG: pass through with correct MIME
 */
export function encodeMediaForWebDisplay(mediaPath: string, data: Uint8Array): string {
  const ext = extOf(mediaPath);

  if (ext === 'tif' || ext === 'tiff') {
    const rgba = tiffToRgba(data);
    if (rgba) {
      return rgbaToPngDataUrl(rgba.data, rgba.width, rgba.height);
    }
    return TRANSPARENT_PNG_DATA_URL;
  }

  if (ext === 'emf') {
    const content = parseEmfContent(data);
    if (content.type === 'bitmap' && content.bitmap) {
      const { width, height, data: rgba } = content.bitmap;
      return rgbaToPngDataUrl(rgba, width, height);
    }
    if (content.type === 'pdf') {
      // Optional: integrate pdfjs-dist later (see pptx-renderer pdfRenderer.ts)
      return TRANSPARENT_PNG_DATA_URL;
    }
    if (content.type === 'empty') {
      return TRANSPARENT_PNG_DATA_URL;
    }
    return TRANSPARENT_PNG_DATA_URL;
  }

  if (ext === 'wmf') {
    return TRANSPARENT_PNG_DATA_URL;
  }

  const mime = getMimeType(mediaPath);
  return toDataUrl(arrayBufferToBase64(data), mime);
}
