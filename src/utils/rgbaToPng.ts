/**
 * Encode RGBA pixel buffer to PNG data URL (browser + Node).
 */

import { PNG } from 'pngjs';

function uint8ToBase64(u8: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < u8.length; i++) {
    binary += String.fromCharCode(u8[i]);
  }
  if (typeof btoa !== 'undefined') return btoa(binary);
  const NodeBuffer = (typeof globalThis !== 'undefined' &&
    (globalThis as unknown as { Buffer?: { from(a: Uint8Array): { toString(e: string): string } } }).Buffer);
  if (NodeBuffer) return NodeBuffer.from(u8).toString('base64');
  return btoa(binary);
}

/**
 * Encode RGBA (length = width * height * 4) to PNG data URL.
 */
export function rgbaToPngDataUrl(
  rgba: Uint8Array | Uint8ClampedArray,
  width: number,
  height: number,
): string {
  const png = new PNG({ width, height });
  png.data.set(rgba);
  const buf = PNG.sync.write(png) as unknown as Uint8Array;
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf as ArrayLike<number>);
  const b64 = uint8ToBase64(u8);
  return `data:image/png;base64,${b64}`;
}
