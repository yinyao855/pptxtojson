/**
 * pptxtojson — Parse .pptx to JSON for PPTist
 * New TypeScript implementation; src1 is reference for data format.
 */

import { parseZip } from './parser/ZipParser'
import { buildPresentation } from './model/Presentation'
import { toPptxtojsonFormat } from './adapter/toPptxtojson'
import type { Output } from './adapter/types'
import { serializePresentation } from './export/serializePresentation'

/**
 * Parse a .pptx file (ArrayBuffer) and return pptxtojson/PPTist format.
 * All dimensions in output are in pt.
 */
export async function parse(buffer: ArrayBuffer): Promise<Output> {
  const files = await parseZip(buffer)
  const presentation = buildPresentation(files)
  const serializedPresentation = serializePresentation(presentation)
  console.log(serializedPresentation)
  return toPptxtojsonFormat(presentation, files)
}

export { parseZip, buildPresentation, toPptxtojsonFormat }
export type { Output, Slide, Element } from './adapter/types'
export type { PptxFiles } from './parser/ZipParser'
export type { PresentationData } from './model/Presentation'
