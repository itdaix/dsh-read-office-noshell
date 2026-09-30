/**
 * 只读 ZIP 解析：从 zip 容器里按名字取文件。
 *
 * 为什么自己写：xlsx / docx / pptx 都是 zip 包，而本插件保持零依赖 ——
 * Node 自带 `zlib.inflateRawSync`，缺的只是 central directory 的解析。
 *
 * @module lib/zip-read
 */

import { inflateRawSync } from 'node:zlib'

const EOCD_SIGNATURE = 0x06054b50
const CENTRAL_SIGNATURE = 0x02014b50

/**
 * 解析整个 zip 的目录表。
 * @param {Buffer} buffer - 完整 zip 字节。
 * @returns {Map<string, { method: number, offset: number, compressedSize: number, size: number }>}
 */
export function readZipDirectory(buffer) {
  const eocd = findEocd(buffer)
  if (eocd < 0) throw new Error('不是合法的 zip（找不到目录结束标记）。')

  const count = buffer.readUInt16LE(eocd + 10)
  let cursor = buffer.readUInt32LE(eocd + 16)
  const entries = new Map()

  for (let i = 0; i < count; i++) {
    if (buffer.readUInt32LE(cursor) !== CENTRAL_SIGNATURE) break
    const method = buffer.readUInt16LE(cursor + 10)
    const compressedSize = buffer.readUInt32LE(cursor + 20)
    const size = buffer.readUInt32LE(cursor + 24)
    const nameLength = buffer.readUInt16LE(cursor + 28)
    const extraLength = buffer.readUInt16LE(cursor + 30)
    const commentLength = buffer.readUInt16LE(cursor + 32)
    const offset = buffer.readUInt32LE(cursor + 42)
    const name = buffer.toString('utf8', cursor + 46, cursor + 46 + nameLength)
    entries.set(name, { method, offset, compressedSize, size })
    cursor += 46 + nameLength + extraLength + commentLength
  }

  return entries
}

/** 从后往前找 EOCD，允许尾部有注释。 */
function findEocd(buffer) {
  const start = Math.max(0, buffer.length - 65557)
  for (let i = buffer.length - 22; i >= start; i--) {
    if (buffer.readUInt32LE(i) === EOCD_SIGNATURE) return i
  }
  return -1
}

/**
 * 读出某个成员的原始字节。
 * @param {Buffer} buffer - 完整 zip 字节。
 * @param {Map<string, object>} entries - `readZipDirectory` 的结果。
 * @param {string} name - 成员名（zip 内路径，正斜杠）。
 * @returns {Buffer|null}
 */
export function readZipEntry(buffer, entries, name) {
  const entry = entries.get(name)
  if (!entry) return null

  const nameLength = buffer.readUInt16LE(entry.offset + 26)
  const extraLength = buffer.readUInt16LE(entry.offset + 28)
  const start = entry.offset + 30 + nameLength + extraLength
  // 压缩大小在 central directory 里是权威值；本地头可能在流式写入时为 0。
  const end = entry.compressedSize > 0 ? start + entry.compressedSize : buffer.length
  const raw = buffer.subarray(start, Math.min(end, buffer.length))

  if (entry.method === 0) return Buffer.from(raw)
  if (entry.method === 8) return inflateRawSync(raw)
  throw new Error(`不支持的压缩方式 ${entry.method}（成员 ${name}）。`)
}

/** 按名字取文本成员，顺带清掉 BOM。 */
export function readZipText(buffer, entries, name) {
  const bytes = readZipEntry(buffer, entries, name)
  if (!bytes) return null
  return bytes.toString('utf8').replace(/^\uFEFF/, '')
}

/** 列出成员名（调试用）。 */
export function zipNames(entries) {
  return [...entries.keys()]
}
