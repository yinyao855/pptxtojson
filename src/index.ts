/**
 * pptxtojson — Parse .pptx to JSON for PPTist
 * New TypeScript implementation; src1 is reference for data format.
 */

import { parseZip } from './parser/ZipParser'
import { buildPresentation } from './model/Presentation'
import { toPptxtojsonFormat } from './adapter/toPptxtojson'
import type { PptxtojsonOutput } from './adapter/types'

/**
 * Parse a .pptx file (ArrayBuffer) and return pptxtojson/PPTist format.
 * All dimensions in output are in pt.
 */
export async function parse(buffer: ArrayBuffer): Promise<PptxtojsonOutput> {
  const files = await parseZip(buffer)
  const presentation = buildPresentation(files)
  return toPptxtojsonFormat(presentation, files)
}

export { parseZip, buildPresentation, toPptxtojsonFormat }
export type { PptxtojsonOutput, PptxtojsonSlide, PptxtojsonElement } from './adapter/types'
export type { PptxFiles } from './parser/ZipParser'
export type { PresentationData } from './model/Presentation'
