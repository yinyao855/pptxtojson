#!/usr/bin/env node
/**
 * 命令行脚本：接受一个 .pptx 文件路径，解析并输出 JSON。
 * 用法: node scripts/transvert.js <path-to.pptx>
 * 或:   pnpm run transvert <path-to.pptx>
 */

import { parse } from '../dist/index.js'
import fs from 'fs'
import path from 'path'

const pptxPath = process.argv[2]
if (!pptxPath) {
  console.error('用法: node scripts/transvert.js <path-to.pptx>')
  process.exit(1)
}

const resolved = path.resolve(process.cwd(), pptxPath)
if (!fs.existsSync(resolved)) {
  console.error('文件不存在:', resolved)
  process.exit(1)
}

const buf = fs.readFileSync(resolved)
const arrayBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)

parse(arrayBuffer)
  .then((json) => {
    console.log(JSON.stringify(json, null, 2))
  })
  .catch((err) => {
    console.error('解析失败:', err.message)
    process.exit(1)
  })
