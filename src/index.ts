/**
 * pptxtojson — Parse .pptx to JSON for PPTist
 * New TypeScript implementation; src1 is reference for data format.
 */

import { parseZip } from './parser/ZipParser'
import { buildPresentation } from './model/Presentation'
import { toPptxtojsonFormat } from './adapter/toPptxtojson'
import { preprocessMedia } from './utils/mediaPreprocess'
import type { Output } from './adapter/types'
import type { MediaMode } from './serializer/RenderContext'

export interface ParseOptions {
  /**
   * 'base64' — embed media as data:… URLs (default, portable, large JSON).
   * 'blob'   — use blob: URLs (compact JSON, browser-only, good for development).
   */
  mediaMode?: MediaMode
  /**
   * Run async media pre-processing (EMF-PDF → PNG via pdfjs, WDP → PNG via sharp).
   * Requires `sharp` and/or `pdfjs-dist` + `canvas` to be installed.
   * Default: false (only synchronous conversions like TIFF and EMF-bitmap are done).
   */
  preprocess?: boolean
}

/**
 * Parse a .pptx file (ArrayBuffer) and return pptxtojson/PPTist format.
 * All dimensions in output are in pt.
 */
export async function parse(buffer: ArrayBuffer, options?: ParseOptions): Promise<Output> {
  const files = await parseZip(buffer)
  const presentation = buildPresentation(files)
  if (options?.preprocess !== false) {
    await preprocessMedia(presentation)
  }
  return toPptxtojsonFormat(presentation, files, options?.mediaMode ?? 'base64')
}

export { parseZip, buildPresentation, toPptxtojsonFormat, preprocessMedia }
export type { Output, Slide, Element } from './adapter/types'
export type { PptxFiles } from './parser/ZipParser'
export type { PresentationData } from './model/Presentation'
export type { MediaMode } from './serializer/RenderContext'
