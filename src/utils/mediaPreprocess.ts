/**
 * Async media pre-processing — converts non-web-safe image formats
 * (EMF-PDF, WDP/JPEG XR) into PNG/JPEG before the synchronous
 * serialization pass.
 *
 * Call `preprocessMedia(presentation)` after `buildPresentation()`.
 * It mutates `presentation.media` in-place, replacing binary data
 * for formats that require heavy/async decoders.
 *
 * Lightweight conversions (TIFF via UTIF, EMF-bitmap) are already
 * handled synchronously in `encodeMediaForWebDisplay`; this module
 * covers cases that need `jpegxr` or `pdfjs-dist`.
 */

import type { PresentationData } from '../model/Presentation';

function extOf(path: string): string {
  return path.split('.').pop()?.toLowerCase() || '';
}

async function convertWdpToPng(data: Uint8Array): Promise<Uint8Array | null> {
  try {
    const jpegxrModule: any = await import('jpegxr');
    const JpegXR = jpegxrModule.default || jpegxrModule;
    const { rgbaToPngDataUrl } = await import('./rgbaToPng');

    const mod = await new JpegXR();
    const result = mod.decode(data);
    const { width, height, bytes, pixelInfo } = result;
    if (!width || !height || !bytes) return null;

    const channels: number = pixelInfo?.channels ?? 3;
    const isBgr: boolean = pixelInfo?.bgr ?? false;

    const rgba = new Uint8ClampedArray(width * height * 4);
    for (let i = 0, j = 0; i < width * height; i++, j += channels) {
      const dst = i * 4;
      rgba[dst + 0] = isBgr ? bytes[j + 2] : bytes[j + 0];
      rgba[dst + 1] = bytes[j + 1];
      rgba[dst + 2] = isBgr ? bytes[j + 0] : bytes[j + 2];
      rgba[dst + 3] = channels === 4 ? bytes[j + 3] : 255;
    }

    const dataUrl = rgbaToPngDataUrl(rgba, width, height);
    const base64 = dataUrl.split(',')[1];
    const binaryStr = atob(base64);
    const pngBytes = new Uint8Array(binaryStr.length);
    for (let i = 0; i < binaryStr.length; i++) pngBytes[i] = binaryStr.charCodeAt(i);
    return pngBytes;
  } catch {
    return null;
  }
}

async function renderPdfToPng(pdfData: Uint8Array, targetWidth = 1024): Promise<Uint8Array | null> {
  try {
    const pdfjsLib: any = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const canvasModule: any = await import('canvas');

    const doc = await pdfjsLib.getDocument({ data: pdfData, verbosity: 0 }).promise;
    const page = await doc.getPage(1);
    const baseViewport = page.getViewport({ scale: 1 });
    const scale = Math.max(1, targetWidth / baseViewport.width);
    const viewport = page.getViewport({ scale });
    const w = Math.round(viewport.width);
    const h = Math.round(viewport.height);

    const canvas = canvasModule.createCanvas(w, h);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);

    await page.render({ canvasContext: ctx, viewport }).promise;

    const pngBuf: Uint8Array = canvas.toBuffer('image/png');
    await doc.destroy();
    return new Uint8Array(pngBuf.buffer, (pngBuf as any).byteOffset ?? 0, pngBuf.byteLength);
  } catch {
    return null;
  }
}

async function convertEmfPdfToPng(emfData: Uint8Array): Promise<Uint8Array | null> {
  const { parseEmfContent } = await import('./emfParser');
  const content = parseEmfContent(emfData);
  if (content.type !== 'pdf') return null;
  return renderPdfToPng(content.data);
}

/**
 * Pre-process non-web-safe media in-place.
 *
 * - WDP / JXR / HDP → PNG via `jpegxr` (WASM-based JPEG XR decoder)
 * - EMF containing PDF → PNG via `pdfjs-dist` + `canvas`
 *
 * Returns the set of media paths that were successfully converted
 * (callers can use this for logging / debugging).
 */
export async function preprocessMedia(presentation: PresentationData): Promise<Set<string>> {
  const converted = new Set<string>();
  const tasks: Array<{ path: string; promise: Promise<Uint8Array | null> }> = [];

  for (const [mediaPath, data] of presentation.media) {
    const ext = extOf(mediaPath);
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);

    if (ext === 'wdp' || ext === 'jxr' || ext === 'hdp') {
      tasks.push({ path: mediaPath, promise: convertWdpToPng(bytes) });
    } else if (ext === 'emf') {
      tasks.push({ path: mediaPath, promise: convertEmfPdfToPng(bytes) });
    }
  }

  const results = await Promise.allSettled(tasks.map((t) => t.promise));

  for (let i = 0; i < tasks.length; i++) {
    const result = results[i];
    if (result.status === 'fulfilled' && result.value) {
      const oldPath = tasks[i].path;
      const newPath = oldPath.replace(/\.[^.]+$/, '.png');
      presentation.media.delete(oldPath);
      presentation.media.set(newPath, result.value);
      converted.add(oldPath);

      // Patch relationship targets so serializers find the new path
      for (const slide of presentation.slides) {
        patchRels(slide.rels, oldPath, newPath);
      }
      for (const layout of presentation.layouts.values()) {
        patchRels(layout.rels, oldPath, newPath);
      }
      for (const master of presentation.masters.values()) {
        patchRels(master.rels, oldPath, newPath);
      }
    }
  }

  return converted;
}

function patchRels(
  rels: Map<string, { target: string; type: string; targetMode?: string }>,
  oldPath: string,
  newPath: string,
): void {
  const oldFileName = oldPath.split('/').pop() || '';
  const newFileName = newPath.split('/').pop() || '';
  if (!oldFileName || !newFileName) return;

  for (const [, rel] of rels) {
    if (rel.target.endsWith(oldFileName)) {
      rel.target = rel.target.replace(oldFileName, newFileName);
    }
  }
}
