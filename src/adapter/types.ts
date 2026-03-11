/**
 * pptxtojson / PPTist 输出格式类型定义
 * 参考 README 与 src1 的字段结构；长度与坐标单位均为 pt。
 */

export interface PptxtojsonSize {
  width: number
  height: number
}

export interface PptxtojsonFillColor {
  type: 'color'
  value: string
}

export interface PptxtojsonFillGradient {
  type: 'gradient'
  [key: string]: unknown
}

export interface PptxtojsonFillImage {
  type: 'image'
  [key: string]: unknown
}

export interface PptxtojsonFillPattern {
  type: 'pattern'
  [key: string]: unknown
}

export type PptxtojsonFill =
  | PptxtojsonFillColor
  | PptxtojsonFillGradient
  | PptxtojsonFillImage
  | PptxtojsonFillPattern

export interface PptxtojsonTransition {
  type: string
  duration: number
  direction: string | null
}

export interface PptxtojsonElementBase {
  left: number
  top: number
  width: number
  height: number
  name?: string
  order?: number
}

export interface PptxtojsonTextElement extends PptxtojsonElementBase {
  type: 'text'
  content: string
  fill?: PptxtojsonFill
  borderColor?: string
  borderWidth?: number
  borderType?: string
  borderStrokeDasharray?: string | number
  shadow?: unknown
  isFlipV?: boolean
  isFlipH?: boolean
  rotate?: number
  vAlign?: string
  isVertical?: boolean
  autoFit?: unknown
  link?: string
}

export interface PptxtojsonShapeElement extends PptxtojsonElementBase {
  type: 'shape'
  shapType: string
  path?: string
  keypoints?: Record<string, number>
  content?: string
  fill?: PptxtojsonFill
  borderColor?: string
  borderWidth?: number
  borderType?: string
  borderStrokeDasharray?: string | number
  shadow?: unknown
  isFlipV?: boolean
  isFlipH?: boolean
  rotate?: number
  vAlign?: string
  autoFit?: unknown
  link?: string
}

export interface PptxtojsonImageElement extends PptxtojsonElementBase {
  type: 'image'
  src: string
  rect?: { t?: number; b?: number; l?: number; r?: number }
  geom?: string
  borderColor?: string
  borderWidth?: number
  borderType?: string
  borderStrokeDasharray?: string | number
  rotate?: number
  isFlipV?: boolean
  isFlipH?: boolean
  filters?: unknown
  link?: string
}

export interface PptxtojsonVideoElement extends PptxtojsonElementBase {
  type: 'video'
  src?: string
  blob?: string
  rotate?: number
}

export interface PptxtojsonAudioElement extends PptxtojsonElementBase {
  type: 'audio'
  blob?: string
  rotate?: number
}

export interface PptxtojsonTableElement extends PptxtojsonElementBase {
  type: 'table'
  data: Array<Array<Record<string, unknown>>>
  rowHeights: number[]
  colWidths: number[]
  borders?: unknown
}

export interface PptxtojsonChartElement extends PptxtojsonElementBase {
  type: 'chart'
  data: unknown
  colors?: string[]
  chartType?: string
  barDir?: string
  marker?: boolean
  holeSize?: number
  grouping?: string
  style?: unknown
}

export interface PptxtojsonDiagramElement extends PptxtojsonElementBase {
  type: 'diagram'
  elements?: PptxtojsonElement[]
  textList?: string[]
}

export interface PptxtojsonGroupElement extends PptxtojsonElementBase {
  type: 'group'
  elements: PptxtojsonElement[]
  rotate?: number
  isFlipV?: boolean
  isFlipH?: boolean
}

export interface PptxtojsonMathElement extends PptxtojsonElementBase {
  type: 'math'
  latex?: string
  picBase64?: string
  text?: string
  order?: number
}

export type PptxtojsonElement =
  | PptxtojsonTextElement
  | PptxtojsonShapeElement
  | PptxtojsonImageElement
  | PptxtojsonVideoElement
  | PptxtojsonAudioElement
  | PptxtojsonTableElement
  | PptxtojsonChartElement
  | PptxtojsonDiagramElement
  | PptxtojsonGroupElement
  | PptxtojsonMathElement

export interface PptxtojsonSlide {
  fill: PptxtojsonFill
  elements: PptxtojsonElement[]
  layoutElements: PptxtojsonElement[]
  note?: string
  transition?: PptxtojsonTransition | null
}

export interface PptxtojsonOutput {
  slides: PptxtojsonSlide[]
  themeColors: string[]
  size: PptxtojsonSize
}
