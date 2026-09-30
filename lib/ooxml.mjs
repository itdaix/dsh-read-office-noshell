/**
 * OOXML 只读解析：xlsx / docx / pptx 的文本与内嵌图片。
 *
 * 三种格式都是 zip 包，所以统一走 `zip-read`：取 XML 抽文本，取 media 目录抽图。
 * 图片不解析内容 —— 交给多模态模型看图。
 *
 * @module lib/ooxml
 */

import { readZipDirectory, readZipEntry, readZipText } from './zip-read.mjs'

const IMAGE_EXTENSIONS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tif', 'tiff', 'emf', 'wmf'])

/** 图片扩展名 → media type。 */
function mediaTypeOf(name) {
  const ext = name.split('.').pop()?.toLowerCase() ?? ''
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg'
  if (ext === 'png') return 'image/png'
  if (ext === 'gif') return 'image/gif'
  if (ext === 'webp') return 'image/webp'
  if (ext === 'bmp') return 'image/bmp'
  if (ext === 'tif' || ext === 'tiff') return 'image/tiff'
  return 'application/octet-stream'
}

/** 解 XML 实体。 */
function decodeXml(text) {
  return text
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** 取 zip 里 media 前缀下的所有图片。 */
function collectMedia(buffer, entries, prefix) {
  const images = []
  for (const name of entries.keys()) {
    if (!name.startsWith(prefix) || name.endsWith('/')) continue
    const ext = name.split('.').pop()?.toLowerCase() ?? ''
    if (!IMAGE_EXTENSIONS.has(ext)) continue
    const bytes = readZipEntry(buffer, entries, name)
    if (!bytes?.length) continue
    images.push({ name: name.slice(prefix.length), mediaType: mediaTypeOf(name), bytes })
  }
  return images
}

/* ─────────────────────────── xlsx ─────────────────────────── */

/**
 * 读工作簿。
 * @returns `{ sheets: Array<{ name, text, rows, truncated }>, images }`
 */
export function readXlsx(buffer, { maxRows = 300, maxCols = 40 } = {}) {
  const entries = readZipDirectory(buffer)
  const workbook = readZipText(buffer, entries, 'xl/workbook.xml')
  if (!workbook) throw new Error('这不是一个可读的 xlsx（缺 xl/workbook.xml）。')

  const rels = readZipText(buffer, entries, 'xl/_rels/workbook.xml.rels') ?? ''
  const relTargets = new Map()
  for (const one of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relTargets.set(one[1], one[2].replace(/^\/?xl\//, ''))
  }

  const shared = readSharedStrings(buffer, entries)
  const sheets = []
  for (const sheet of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /name="([^"]*)"/.exec(sheet[0])?.[1] ?? 'Sheet'
    const relId = /r:id="([^"]+)"/.exec(sheet[0])?.[1]
    const target = relId ? relTargets.get(relId) : null
    if (!target) { sheets.push({ name: decodeXml(name), text: '(找不到这个工作表的数据部件)', rows: 0, truncated: false }); continue }
    const xml = readZipText(buffer, entries, `xl/${target}`)
    if (!xml) { sheets.push({ name: decodeXml(name), text: '(工作表数据为空)', rows: 0, truncated: false }); continue }
    sheets.push({ name: decodeXml(name), ...renderSheet(xml, shared, maxRows, maxCols) })
  }

  return { sheets, images: collectMedia(buffer, entries, 'xl/media/') }
}

/** 共享字符串表。 */
function readSharedStrings(buffer, entries) {
  const xml = readZipText(buffer, entries, 'xl/sharedStrings.xml')
  if (!xml) return []
  const out = []
  for (const si of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g)) {
    if (!si[1]) { out.push(''); continue }
    const texts = [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((one) => decodeXml(one[1]))
    out.push(texts.join(''))
  }
  return out
}

/** 把一张工作表渲染成制表符对齐的文本。 */
function renderSheet(xml, shared, maxRows, maxCols) {
  const lines = []
  let rows = 0
  let truncated = false

  for (const row of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
    if (rows >= maxRows) { truncated = true; break }
    const cells = []
    let overflow = false
    for (const cell of row[1].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = cell[1] ?? ''
      const body = cell[2] ?? ''
      const ref = /r="([A-Z]+)\d+"/.exec(attrs)?.[1]
      const column = ref ? columnIndex(ref) : cells.length
      if (column >= maxCols) { overflow = true; continue }
      cells[column] = cellValue(attrs, body, shared)
    }
    lines.push(cells.map((v) => v ?? '').join('\t').replace(/\t+$/, ''))
    rows += 1
    if (overflow) truncated = true
  }

  return { text: lines.join('\n'), rows, truncated }
}

function cellValue(attrs, body, shared) {
  const type = /t="([^"]+)"/.exec(attrs)?.[1]
  if (type === 'inlineStr') {
    return [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((one) => decodeXml(one[1])).join('')
  }
  const value = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1]
  if (value === undefined) return ''
  if (type === 's') return shared[Number(value)] ?? ''
  if (type === 'b') return value === '1' ? 'TRUE' : 'FALSE'
  return decodeXml(value)
}

function columnIndex(letters) {
  let index = 0
  for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64)
  return index - 1
}

/* ─────────────────────────── docx ─────────────────────────── */

/**
 * 读 Word 文档。
 * @returns `{ text, images, paragraphs }`
 */
export function readDocx(buffer) {
  const entries = readZipDirectory(buffer)
  const xml = readZipText(buffer, entries, 'word/document.xml')
  if (!xml) throw new Error('这不是一个可读的 docx（缺 word/document.xml）。')

  const lines = []
  let paragraphs = 0
  const body = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(xml)?.[1] ?? xml

  for (const block of body.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>/g)) {
    const chunk = block[0]
    if (chunk.startsWith('<w:tbl')) {
      for (const row of chunk.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) {
        const cells = [...row[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cell) => paragraphText(cell[0]).replace(/\n/g, ' ').trim())
        lines.push(cells.join(' | '))
        paragraphs += 1
      }
      continue
    }
    const text = paragraphText(chunk)
    if (text.trim()) { lines.push(text.trim()); paragraphs += 1 }
  }

  return { text: lines.join('\n'), paragraphs, images: collectMedia(buffer, entries, 'word/media/') }
}

/** 一个段落 XML → 文本（含制表符、换行、图片占位）。 */
function paragraphText(chunk) {
  return decodeXml(
    chunk
      .replace(/<w:tab\b[^>]*\/>/g, '\t')
      .replace(/<w:br\b[^>]*\/>/g, '\n')
      .replace(/<w:drawing\b[\s\S]*?<\/w:drawing>/g, '[图片]')
      .replace(/<w:pict\b[\s\S]*?<\/w:pict>/g, '[图片]')
      .replace(/<[^>]+>/g, ''),
  )
}

/* ─────────────────────────── pptx ─────────────────────────── */

/**
 * 读 PowerPoint。
 * @returns `{ text, slides, images }`
 */
export function readPptx(buffer) {
  const entries = readZipDirectory(buffer)
  const slideNames = [...entries.keys()]
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b))
  if (slideNames.length === 0) throw new Error('这不是一个可读的 pptx（找不到 ppt/slides）。')

  const out = []
  for (const name of slideNames) {
    const xml = readZipText(buffer, entries, name)
    if (!xml) continue
    const lines = []
    for (const para of xml.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)) {
      const runs = [...para[0].matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)].map((one) => decodeXml(one[1]))
      const text = runs.join('').trim()
      if (text) lines.push(text)
    }
    out.push({ index: slideNumber(name), text: lines.join('\n') })
  }

  return {
    slides: out,
    text: out.map((slide) => `【第 ${slide.index} 页】\n${slide.text}`).join('\n\n'),
    images: collectMedia(buffer, entries, 'ppt/media/'),
  }
}

function slideNumber(name) {
  return Number(/slide(\d+)\.xml$/.exec(name)?.[1] ?? 0)
}
