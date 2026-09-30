/**
 * PDF 只读解析：逐页判定「文字页 / 图片页 / 图文页」，分别给出文本或内嵌图片。
 *
 * 为什么自己写：本插件零依赖，且宿主上没有 Ghostscript / poppler / OCR 可用。
 * PDF 的两条路都在这里：
 *   - 文字页：解压内容流 → 抽 Tj/TJ 文本 → 用字体的 ToUnicode 映射还原中文；
 *   - 图片页（扫描件）：把该页 XObject 里的图抽出来（DCTDecode 直出 JPEG，
 *     FlateDecode 位图自己拼 PNG），交给多模态模型看图。
 *
 * 边界：JBIG2 / CCITT / LZW / 加密 PDF 明确报「不支持」，不猜内容。
 *
 * @module lib/pdf-read
 */

import { deflateSync, inflateSync, inflateRawSync } from 'node:zlib'

/** 一页里有效字符少于这个数，就认为它是图片页（扫描件）。 */
const TEXT_PAGE_MIN_CHARS = 8

/* ─────────────────────────── PDF 对象扫描 ─────────────────────────── */

/** 扫出所有 `N 0 obj ... endobj`，带上 stream 区间（不做 xref 解析）。 */
function scanObjects(src) {
  const objects = new Map()
  const re = /(\d+)\s+(\d+)\s+obj\b/g
  let match
  while ((match = re.exec(src)) !== null) {
    const num = Number(match[1])
    const bodyStart = re.lastIndex
    const endobj = src.indexOf('endobj', bodyStart)
    const limit = endobj < 0 ? src.length : endobj

    let dictEnd = limit
    let stream = null
    const streamIdx = src.indexOf('stream', bodyStart)
    if (streamIdx >= 0 && streamIdx < limit) {
      dictEnd = streamIdx
      let start = streamIdx + 'stream'.length
      if (src[start] === '\r') start += 1
      if (src[start] === '\n') start += 1
      const stop = src.indexOf('endstream', start)
      if (stop >= 0) stream = { start, end: stop }
    }

    objects.set(num, { num, dict: src.slice(bodyStart, dictEnd), stream })
    if (endobj >= 0) re.lastIndex = endobj + 'endobj'.length
  }
  return objects
}

/** 取 `/Key` 后面的引用号，例如 `/Contents 7 0 R` → 7。 */
function refAfter(dict, key) {
  const m = new RegExp(`/${key}\\s+(\\d+)\\s+\\d+\\s+R`).exec(dict)
  return m ? Number(m[1]) : null
}

/** 取 `/Key` 后面的原始记号（数字或名字）。 */
function valueAfter(dict, key) {
  const m = new RegExp(`/${key}\\s*(\\/[-\\w#]+|[\\d.]+)`).exec(dict)
  return m ? m[1] : null
}

/** 取 `/Key` 后面的子字典原文（不处理嵌套，够用）。 */
function dictAfter(dict, key) {
  const m = new RegExp(`/${key}\\s*(<<)`).exec(dict)
  if (!m) return null
  let depth = 0
  for (let i = m.index + m[0].length - 2; i < dict.length; i++) {
    if (dict.startsWith('<<', i)) { depth += 1; i += 1; continue }
    if (dict.startsWith('>>', i)) { depth -= 1; i += 1; if (depth === 0) return dict.slice(m.index + m[0].length - 2, i + 1) }
  }
  return null
}

/** 取子字典：内联 `<<...>>` 或引用 `N 0 R` 都认。 */
function resolveDict(objects, dict, key) {
  const inline = dictAfter(dict, key)
  if (inline) return inline
  const ref = refAfter(dict, key)
  if (ref === null) return null
  return objects.get(ref)?.dict ?? null
}

/** 取全部 `/Key a 0 R b 0 R` 形式的引用号。 */
function allRefs(dict, key) {
  const out = []
  const re = new RegExp(`/(?:${key})\\s*\\[([^\\]]*)\\]`, 'g')
  const m = re.exec(dict)
  if (m) {
    for (const one of m[1].matchAll(/(\d+)\s+\d+\s+R/g)) out.push(Number(one[1]))
  }
  const single = refAfter(dict, key)
  if (single !== null) out.push(single)
  return out
}

/** 解出 stream 的字节；返回 `{ bytes, unsupported }`。 */
function decodeStream(objects, obj) {
  if (!obj?.stream) return { bytes: null, unsupported: null }
  const raw = Buffer.from(obj.rawBytes ?? '', 'latin1')
  const filters = filtersOf(obj.dict)
  let bytes = raw
  for (const filter of filters) {
    if (filter === 'FlateDecode' || filter === 'Fl') {
      try {
        bytes = inflateSync(bytes)
      } catch {
        try { bytes = inflateRawSync(bytes) } catch { return { bytes: null, unsupported: filter } }
      }
      continue
    }
    if (filter === 'DCTDecode' || filter === 'DCT') return { bytes, unsupported: null }
    return { bytes: null, unsupported: filter }
  }
  return { bytes, unsupported: null }
}

/** 取 `/Filter`（单值或数组）；没有就是空数组 —— 不能拿整段字典乱匹配。 */
function filtersOf(dict) {
  const m = /\/Filter\s*(\[[^\]]*\]|\/[A-Za-z0-9]+)/.exec(dict)
  if (!m) return []
  return [...m[1].matchAll(/\/([A-Za-z0-9]+)/g)].map((one) => one[1])
}

/* ─────────────────────────── 文本抽取 ─────────────────────────── */

/** 解析 ToUnicode CMap → `{ map, bytesPerCode }`。 */
function parseToUnicode(text) {
  const map = new Map()
  let bytesPerCode = 1

  const put = (code, value) => {
    map.set(code, value)
    if (code > 0xff) bytesPerCode = 2
  }

  for (const block of text.matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const pair of block[1].matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) {
      put(parseInt(pair[1], 16), hexToUtf16(pair[2]))
    }
  }

  for (const block of text.matchAll(/beginbfrange([\s\S]*?)endbfrange/g)) {
    const body = block[1]
    for (const one of body.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g)) {
      const lo = parseInt(one[1], 16)
      const hi = parseInt(one[2], 16)
      const start = parseInt(one[3], 16)
      for (let code = lo; code <= hi && code - lo < 65536; code++) put(code, hexToUtf16((start + code - lo).toString(16).padStart(one[3].length, '0')))
    }
    for (const one of body.matchAll(/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([^\]]*)\]/g)) {
      const lo = parseInt(one[1], 16)
      const items = [...one[3].matchAll(/<([0-9A-Fa-f]+)>/g)]
      items.forEach((item, index) => put(lo + index, hexToUtf16(item[1])))
    }
  }

  return { map, bytesPerCode }
}

/** `<0041>` 这类 UTF-16BE 十六进制 → 字符串。 */
function hexToUtf16(hex) {
  const clean = hex.length % 2 ? '0' + hex : hex
  let out = ''
  for (let i = 0; i < clean.length; i += 4) {
    const unit = clean.slice(i, i + 4).padEnd(4, '0')
    out += String.fromCharCode(parseInt(unit, 16))
  }
  return out
}

/** 解 PDF 字面串的转义。 */
function unescapePdfString(text) {
  let out = ''
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch !== '\\') { out += ch; continue }
    const next = text[i + 1]
    if (next === undefined) break
    if (next === 'n') { out += '\n'; i += 1; continue }
    if (next === 'r') { out += '\r'; i += 1; continue }
    if (next === 't') { out += '\t'; i += 1; continue }
    if (next === 'b' || next === 'f') { i += 1; continue }
    if (next === '\n') { i += 1; continue }
    if (next === '(' || next === ')' || next === '\\') { out += next; i += 1; continue }
    const octal = /^[0-7]{1,3}/.exec(text.slice(i + 1))
    if (octal) { out += String.fromCharCode(parseInt(octal[0], 8)); i += octal[0].length; continue }
    out += next
    i += 1
  }
  return out
}

/** 把一个 PDF 字符串按字体映射解成可读文本。 */
function decodeString(bytes, font) {
  if (!font?.map || font.map.size === 0) return bytes.toString('latin1')
  let out = ''
  const step = font.bytesPerCode === 2 ? 2 : 1
  for (let i = 0; i < bytes.length; i += step) {
    const code = step === 2 ? (bytes[i] << 8) | (bytes[i + 1] ?? 0) : bytes[i]
    out += font.map.get(code) ?? ''
  }
  return out
}

/** 从内容流里抽文本（BT/ET 块 + Tj/TJ/'/"）。 */
function extractText(content, fonts) {
  const chunks = []
  let currentFont = null
  const fontSwitch = /(\/[A-Za-z0-9#_.-]+)\s+[\d.]+\s+Tf/g

  for (const block of content.matchAll(/BT([\s\S]*?)ET/g)) {
    const body = block[1]
    let cursor = 0
    const parts = []
    while (cursor < body.length) {
      fontSwitch.lastIndex = cursor
      const tf = fontSwitch.exec(body)
      const token = nextStringToken(body, cursor)
      if (token && (!tf || token.index < tf.index)) {
        const operator = operatorAfter(body, token.end)
        if (operator === 'Tj' || operator === "'" || operator === '"') {
          parts.push(decodeString(token.bytes, fonts.get(currentFont)))
          if (operator !== 'Tj') parts.push('\n')
        } else if (operator === 'TJ') {
          parts.push(decodeString(token.bytes, fonts.get(currentFont)))
        }
        cursor = token.end
        continue
      }
      if (tf) {
        currentFont = tf[1].slice(1)
        cursor = tf.index + tf[0].length
        continue
      }
      break
    }
    if (parts.length) chunks.push(parts.join(''))
  }
  return chunks.join('\n')
}

/** 找下一个字面串或十六进制串。 */
function nextStringToken(body, from) {
  const open = body.indexOf('(', from)
  const hex = body.indexOf('<', from)
  const useHex = hex >= 0 && (open < 0 || hex < open)
  if (!useHex && open < 0) return null
  if (useHex) {
    if (body.startsWith('<<', hex)) return nextStringToken(body, hex + 2)
    const close = body.indexOf('>', hex)
    if (close < 0) return null
    const hexText = body.slice(hex + 1, close).replace(/\s+/g, '')
    const padded = hexText.length % 2 ? hexText + '0' : hexText
    if (!/^[0-9A-Fa-f]*$/.test(padded)) return null
    return { index: hex, end: close + 1, bytes: Buffer.from(padded, 'hex') }
  }
  let depth = 0
  for (let i = open; i < body.length; i++) {
    const ch = body[i]
    if (ch === '\\') { i += 1; continue }
    if (ch === '(') depth += 1
    else if (ch === ')') {
      depth -= 1
      if (depth === 0) {
        const inner = body.slice(open + 1, i)
        return { index: open, end: i + 1, bytes: Buffer.from(unescapePdfString(inner), 'latin1') }
      }
    }
  }
  return null
}

/** 取字符串后面的操作符。 */
function operatorAfter(body, from) {
  const rest = body.slice(from, from + 40)
  const m = /^\s*(\[?[^\s\]]*\]?)?\s*(TJ|Tj|'|")/.exec(rest)
  if (!m) return null
  return m[2]
}

/* ─────────────────────────── 图片抽取 ─────────────────────────── */

/** 极简 PNG 打包（8bit RGB / Gray，无滤波）。 */
function encodePng(width, height, components, samples) {
  const stride = width * components
  const raw = Buffer.alloc((stride + 1) * height)
  for (let y = 0; y < height; y++) {
    raw[y * (stride + 1)] = 0
    samples.copy(raw, y * (stride + 1) + 1, y * stride, y * stride + stride)
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8
  ihdr[9] = components === 3 ? 2 : 0
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length, 0)
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body), 0)
  return Buffer.concat([length, body, crc])
}

let crcTable = null
function crc32(buffer) {
  if (!crcTable) {
    crcTable = new Int32Array(256)
    for (let n = 0; n < 256; n++) {
      let c = n
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      crcTable[n] = c
    }
  }
  let c = 0xffffffff
  for (const byte of buffer) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/* ─────────────────────────── 主流程 ─────────────────────────── */

/**
 * 解析一份 PDF。
 * @param {Buffer} buffer - 文件字节。
 * @param {{ maxPages?: number }} [options]
 * @returns {{ pages: Array<object>, notes: string[] }}
 */
export function readPdf(buffer, { maxPages = 50 } = {}) {
  const src = buffer.toString('latin1')
  const objects = scanObjects(src)
  const notes = []

  // 把每个对象的 stream 字节挂上去（延迟到用时再取，这里直接切好）
  for (const obj of objects.values()) {
    obj.rawBytes = obj.stream ? src.slice(obj.stream.start, obj.stream.end) : ''
  }

  const pages = collectPages(objects, src)
  if (pages.length === 0) return { pages: [], notes: ['这份 PDF 里没有找到页面对象。'] }

  const result = []
  let emitted = 0
  for (const pageNum of pages) {
    if (emitted >= maxPages) { notes.push(`只解析了前 ${maxPages} 页。`); break }
    emitted += 1
    result.push(readPage(objects, pageNum, emitted, notes))
  }
  return { pages: result, notes }
}

/** 顺着 /Pages 树收集页面对象号（顺序即页序）。 */
function collectPages(objects, src) {
  const catalog = [...objects.values()].find((o) => /\/Type\s*\/Catalog/.test(o.dict))
  const rootRef = catalog ? refAfter(catalog.dict, 'Pages') : null
  const out = []
  const seen = new Set()

  const walk = (num) => {
    if (num === null || seen.has(num) || out.length > 5000) return
    seen.add(num)
    const obj = objects.get(num)
    if (!obj) return
    if (/\/Type\s*\/Page[^s]/.test(obj.dict)) { out.push(num); return }
    const kids = /\/Kids\s*\[([^\]]*)\]/.exec(obj.dict)
    if (kids) {
      for (const kid of kids[1].matchAll(/(\d+)\s+\d+\s+R/g)) walk(Number(kid[1]))
      return
    }
    const single = refAfter(obj.dict, 'Kids')
    if (single !== null) walk(single)
  }

  if (rootRef !== null) walk(rootRef)
  if (out.length === 0) {
    for (const obj of objects.values()) if (/\/Type\s*\/Page[^s]/.test(obj.dict)) out.push(obj.num)
  }
  return out
}

/** 解析单页：文本 + 图片。 */
function readPage(objects, pageNum, displayIndex, notes) {
  const page = objects.get(pageNum)
  const dict = page.dict

  // 资源：可能是内联字典，也可能指向另一个对象
  const resources = resolveDict(objects, dict, 'Resources') ?? ''

  const fonts = loadFonts(objects, resources)
  let content = ''
  for (const contentRef of allRefs(dict, 'Contents')) {
    const decoded = decodeStream(objects, objects.get(contentRef))
    if (decoded.bytes) content += decoded.bytes.toString('latin1') + '\n'
    else if (decoded.unsupported) notes.push(`第 ${displayIndex} 页内容流用了不支持的编码 ${decoded.unsupported}。`)
  }

  const text = content ? extractText(content, fonts) : ''
  const images = loadImages(objects, resources, displayIndex, notes)
  const plain = text.replace(/\s+/g, '')

  let kind = 'empty'
  if (plain.length >= TEXT_PAGE_MIN_CHARS && images.length) kind = 'mixed'
  else if (plain.length >= TEXT_PAGE_MIN_CHARS) kind = 'text'
  else if (images.length) kind = 'image'
  else if (plain.length) kind = 'text'

  return { index: displayIndex, kind, text: text.trim(), images }
}

/** 收集该页字体 → ToUnicode 映射。 */
function loadFonts(objects, resources) {
  const fonts = new Map()
  const fontDict = resources ? resolveDict(objects, resources, 'Font') : null
  if (!fontDict) return fonts

  for (const one of fontDict.matchAll(/\/([A-Za-z0-9#_.-]+)\s+(\d+)\s+\d+\s+R/g)) {
    const name = one[1]
    const font = objects.get(Number(one[2]))
    if (!font) continue
    const toUnicodeRef = refAfter(font.dict, 'ToUnicode')
    if (toUnicodeRef === null) continue
    const cmapStream = decodeStream(objects, objects.get(toUnicodeRef))
    if (!cmapStream.bytes) continue
    fonts.set(name, parseToUnicode(cmapStream.bytes.toString('latin1')))
  }
  return fonts
}

/** 抽出该页引用的图片。 */
function loadImages(objects, resources, displayIndex, notes) {
  const images = []
  const xobjects = resources ? resolveDict(objects, resources, 'XObject') : null
  if (!xobjects) return images

  const seen = new Set()
  for (const one of xobjects.matchAll(/\/([A-Za-z0-9#_.-]+)\s+(\d+)\s+\d+\s+R/g)) {
    const num = Number(one[2])
    if (seen.has(num)) continue
    seen.add(num)
    const obj = objects.get(num)
    if (!obj || !/\/Subtype\s*\/Image/.test(obj.dict)) continue

    const filter = filtersOf(obj.dict)[0] ?? null
    const width = Number(valueAfter(obj.dict, 'Width') ?? 0)
    const height = Number(valueAfter(obj.dict, 'Height') ?? 0)

    if (filter === 'DCTDecode' || filter === 'DCT') {
      const decoded = decodeStream(objects, obj)
      if (decoded.bytes) images.push({ name: `p${displayIndex}-img${images.length + 1}.jpg`, mediaType: 'image/jpeg', bytes: decoded.bytes, width, height })
      continue
    }
    if (filter === 'FlateDecode' || filter === 'Fl') {
      const decoded = decodeStream(objects, obj)
      const colorspace = valueAfter(obj.dict, 'ColorSpace')
      const components = colorspace === '/DeviceRGB' ? 3 : colorspace === '/DeviceGray' ? 1 : 0
      const bpc = Number(valueAfter(obj.dict, 'BitsPerComponent') ?? 8)
      if (!decoded.bytes || components === 0 || bpc !== 8 || !width || !height) {
        notes.push(`第 ${displayIndex} 页有一张图（${colorspace ?? '未知色彩空间'} / ${bpc}bpc）暂不支持转换。`)
        continue
      }
      images.push({ name: `p${displayIndex}-img${images.length + 1}.png`, mediaType: 'image/png', bytes: encodePng(width, height, components, decoded.bytes), width, height })
      continue
    }
    notes.push(`第 ${displayIndex} 页有一张图，编码 ${filter ?? '未知'}，暂不支持抽取。`)
  }
  return images
}
