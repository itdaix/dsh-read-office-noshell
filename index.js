/**
 * dsh-read-office-noshell —— 解析**用户上传的附件**：出文本，出图片。
 *
 * 存在的唯一理由：上传的 xlsx / docx / pptx / pdf 用普通 `read` 读出来是乱码，
 * 而无 shell 的预设又跑不了 Python 解析。所以这里只补这一件事，**不做通用读文件**。
 *
 * 定位：**零第三方依赖的纯 JS 实现** —— 解析全在宿主 Node 进程内做，不需要 Python、
 * 也不要求模型能起命令，所以无 shell 的预设照样能用；有 shell 的环境同样能用，还更快。
 * （老格式 doc/xls/ppt 由插件自己 spawnSync 拉本机 LibreOffice，跟模型有没有 shell 无关。）
 *
 * 一个工具：`read_document_noshell`，只认两类路径 ——
 *   - IM 上传件：`<会话工作区>/.dsh-im/inbound/**`
 *   - GUI 上传件：`<DSH_HOME>/attachments/v1/files/**`（消息里给的只读副本路径）
 *
 * 支持：xlsx / docx / pptx（解包抽文本 + 抽内嵌图）、pdf（逐页判定文字页/扫描页）、
 * csv/txt/md/json/xml/html（直读）、png/jpg 等图片（只返回路径给 `read_image`）、
 * doc/xls/ppt（老二进制，交本机 LibreOffice 转）。
 *
 * 归属：只往宿主 `tools` 注册表注册工具，不发布服务，因此不需要 isolate realm。
 * 边界：只读上传件；抽出的图片落在「调用方会话工作区」的 `.read-office/` 下，别的什么都不写。
 *
 * @module dsh-read-office-noshell
 */

import { existsSync, mkdirSync, readFileSync, readdirSync, realpathSync, statSync, writeFileSync } from 'node:fs'
import { AsyncLocalStorage } from 'node:async_hooks'
import { homedir } from 'node:os'
import { basename, extname, isAbsolute, join, relative, resolve } from 'node:path'

import { readDocx, readPptx, readXlsx } from './lib/ooxml.mjs'
import { readPdf } from './lib/pdf-read.mjs'
import { cleanScratch, convertToText, findSoffice, makeScratch } from './lib/soffice.mjs'

/* ────────────────────────────── 配置 ────────────────────────────── */

const DEFAULTS = {
  /** 兜底工作区根目录：仅在调用方会话拿不到 cwd 时使用；能拿到就以会话工作区为准。 */
  root: '',
  /** 单次返回的文本上限（字符）。 */
  maxChars: 200000,
  /** 表格最多渲染多少行 / 多少列。 */
  maxRows: 300,
  maxCols: 40,
  /** PDF 最多解析多少页。 */
  maxPages: 50,
  /** 单次最多抽出多少张图片。 */
  maxImages: 20,
  /** 抽出图片的落盘目录（相对 root）。 */
  imageDir: '.read-office',
  /** LibreOffice 路径（留空自动探测），只给老格式用。 */
  soffice: '',
  convertTimeoutMs: 120000,
  /** 是否放行 DSH 附件库里的上传件只读副本（GUI / IM 上传都会落到那里）。 */
  allowAttachmentStore: true,
  /** 附件库根目录；留空则按 DSH_HOME（默认 ~/.dsh）推导。 */
  attachmentRoot: '',
  /** 额外放行的目录（绝对路径）。 */
  extraRoots: [],
  /** 单次读取的文件大小上限（字节），防止一次读爆内存。 */
  maxFileBytes: 50 * 1024 * 1024,
}

/** Cordis 只认 Standard Schema。 */
function normalizeConfig(raw) {
  const input = raw !== null && typeof raw === 'object' && !Array.isArray(raw) ? raw : {}
  const issues = []
  const value = { ...DEFAULTS }

  // root 降级为兜底项：读取与落图都优先用调用方会话的工作区，所以不再强制配置。
  const root = typeof input.root === 'string' ? input.root.trim() : ''
  if (root !== '') value.root = resolve(root)

  for (const key of ['maxChars', 'maxRows', 'maxCols', 'maxPages', 'maxImages', 'convertTimeoutMs', 'maxFileBytes']) {
    if (input[key] === undefined || input[key] === null) continue
    const number = Number(input[key])
    if (!Number.isFinite(number) || number <= 0) issues.push({ message: `${key} 必须是正数`, path: [key] })
    else value[key] = Math.floor(number)
  }

  if (typeof input.imageDir === 'string' && input.imageDir.trim() !== '') {
    const cleaned = input.imageDir.trim().replace(/\\/g, '/').replace(/^\.\//, '').replace(/^\/+/, '').replace(/\/+$/, '')
    if (cleaned === '' || cleaned.includes('..')) issues.push({ message: 'imageDir 必须是工作区内的相对目录', path: ['imageDir'] })
    else value.imageDir = cleaned
  }
  if (typeof input.soffice === 'string' && input.soffice.trim() !== '') value.soffice = input.soffice.trim()

  if (input.allowAttachmentStore === false) value.allowAttachmentStore = false
  if (typeof input.attachmentRoot === 'string' && input.attachmentRoot.trim() !== '') value.attachmentRoot = resolve(input.attachmentRoot.trim())
  if (Array.isArray(input.extraRoots)) {
    value.extraRoots = input.extraRoots.filter((one) => typeof one === 'string' && one.trim() !== '').map((one) => resolve(one.trim()))
  }

  return issues.length > 0 ? { issues } : { value }
}

export const Config = {
  '~standard': { version: 1, vendor: 'dsh-read-office-noshell', validate: normalizeConfig },
}

export const inject = ['tools']

let activeConfig = { ...DEFAULTS }

/**
 * 本次工具调用的工作区根目录。
 *
 * 用 AsyncLocalStorage 而不是直接改模块级 activeConfig：并发调用（同一批多个工具、
 * 或多会话同时读附件）时后者会互相覆盖 root，把文件读写到别人的工作区去。
 */
const callRoot = new AsyncLocalStorage()

/** 当前调用应使用的工作区根目录；没有调用上下文时退回配置的兜底 root。 */
function currentRoot() {
  const store = callRoot.getStore()
  return typeof store === 'string' && store !== '' ? store : activeConfig.root
}

/* ─────────────────────────── 取路径与落图 ─────────────────────────── */

/** DSH 附件库根目录：GUI / IM 上传的文件副本都在这里（只读硬链接）。 */
function defaultAttachmentRoot() {
  const home = (process.env.DSH_HOME ?? '').trim() || join(homedir(), '.dsh')
  return join(home, 'attachments', 'v1')
}

/**
 * 允许读取的位置 —— **只有用户上传的附件**，没有别的。
 *
 *   - IM 上传：`<root>/.dsh-im/inbound/**`（IM 插件把文件暂存在会话工作区里）
 *   - GUI 上传：`<DSH_HOME>/attachments/v1/files/**`（DSH 存下的只读副本）
 *
 * 这个插件存在的唯一理由就是「上传的 xlsx/docx/pptx/pdf 用普通 read 读出来是乱码」，
 * 所以它**不做通用读文件**：工作区里的普通文件请用 `read`，图片用 `read_image`。
 */
function allowedRoots() {
  const roots = [
    { path: join(currentRoot(), '.dsh-im'), label: 'IM 上传件' },
    { path: activeConfig.attachmentRoot || defaultAttachmentRoot(), label: 'GUI 上传件' },
  ]
  if (activeConfig.allowAttachmentStore === false) roots.splice(1, 1)
  for (const extra of activeConfig.extraRoots ?? []) roots.push({ path: extra, label: '额外目录' })
  return roots.filter((one) => one.path)
}

/** 取真实路径；不存在或不可解析时返回 null。 */
function realPathOf(path) {
  try {
    return realpathSync(path)
  } catch {
    return null
  }
}

/** 解析并校验路径：只认「用户上传的文件」，其余一律拒绝。 */
function resolveInput(input, uploads) {
  const root = currentRoot()
  if (root === '') return { ok: false, error: '插件未配置 root（工作区根目录）。' }

  const raw = String(input ?? '').trim()
  if (raw === '') return { ok: false, error: 'path 不能为空。' }

  const target = isAbsolute(raw) ? resolve(raw) : resolve(root, raw)
  let inside = null
  let label = null
  let matchedRoot = null
  for (const one of allowedRoots()) {
    const rest = relative(one.path, target)
    if (rest !== '' && !rest.startsWith('..') && !isAbsolute(rest)) { inside = rest; label = one.label; matchedRoot = one.path; break }
  }
  if (inside === null) {
    return {
      ok: false,
      error: '只读用户上传的附件：IM 上传的在 .dsh-im/inbound/ 下，GUI 上传的用消息里给的「只读副本」路径。工作区里的普通文件请用 read 工具，图片用 read_image。',
    }
  }

  // 判据只有一句：文件出现在**用户消息**里 → 读；没出现 → 不读。
  const key = pathKey(target)
  if (uploads === null || !uploads.allowed.has(key)) {
    return {
      ok: false,
      error: '这个文件没有出现在本次会话的用户消息里（用户没上传过它）。本工具只读用户上传过的附件；工作区的文本文件请用 read。',
    }
  }
  const matchedBy = uploads.lastTurn.has(key) ? 'turn' : 'earlier'

  // 链接绕道检查：字面路径在范围内，不代表真实文件也在 —— 上传目录里可能有指向外面的
  // 符号链接 / junction（Windows 上很常见）。这里用 realpath 再判一次。
  const realTarget = realPathOf(target)
  const realRoot = realPathOf(matchedRoot)
  if (realTarget === null) return { ok: false, error: `文件不存在：${raw}` }
  if (realRoot !== null) {
    const realRest = relative(realRoot, realTarget)
    if (realRest === '' || realRest.startsWith('..') || isAbsolute(realRest)) {
      return { ok: false, error: '这个路径经由链接指到了允许范围之外，已拒绝。' }
    }
  }

  if (!existsSync(target)) return { ok: false, error: `文件不存在：${raw}` }
  if (!statSync(target).isFile()) return { ok: false, error: `这不是一个文件：${raw}` }

  // 限制：单次读取的文件大小
  const size = statSync(target).size
  if (size > activeConfig.maxFileBytes) {
    return { ok: false, error: `文件 ${(size / 1048576).toFixed(1)} MB，超过单次读取上限 ${Math.round(activeConfig.maxFileBytes / 1048576)} MB。` }
  }

  return { ok: true, target, relativePath: inside.split('\\').join('/'), source: label, matchedBy }
}

/**
 * 把抽出的图片写到 `<root>/<imageDir>/<文件>-<本地时间戳>/`。
 * 目录**延迟到真有图片时**才建 —— 否则每读一次文本文件都会留下一个空目录。
 */
function imageWriter(sourceName) {
  const now = new Date()
  const pad = (value) => String(value).padStart(2, '0')
  const stamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`
  const stem = basename(sourceName, extname(sourceName)).replace(/[\\/:*?"<>|\s]/g, '_').slice(0, 40)
  let dir = null

  return (image) => {
    if (dir === null) {
      dir = join(currentRoot(), activeConfig.imageDir, `${stem}-${stamp}`)
      mkdirSync(dir, { recursive: true })
    }
    const path = join(dir, image.name)
    writeFileSync(path, image.bytes)
    return {
      name: image.name,
      mediaType: image.mediaType,
      bytes: image.bytes.length,
      ...(image.width ? { width: image.width } : {}),
      ...(image.height ? { height: image.height } : {}),
      path,
      relativePath: relative(currentRoot(), path).split('\\').join('/'),
    }
  }
}

/** 文本截断。 */
function clip(text, limit = activeConfig.maxChars) {
  if (typeof text !== 'string') return { text: '', chars: 0, truncated: false }
  if (text.length <= limit) return { text, chars: text.length, truncated: false }
  return { text: text.slice(0, limit), chars: limit, truncated: true }
}

/* ──────────────────── 「这一轮用户传了什么」 ──────────────────── */

/** 这个绝对路径是否落在「上传区」里（目录级白名单，仅作兜底）。 */
function isUploadPath(target) {
  for (const one of allowedRoots()) {
    const rest = relative(one.path, target)
    if (rest !== '' && !rest.startsWith('..') && !isAbsolute(rest)) return true
  }
  return false
}

/** 统一成可比较的键：优先真实路径（能对上短名、大小写、链接）。 */
function pathKey(target) {
  return (realPathOf(target) ?? target).toLowerCase()
}

/**
 * 从**本次会话里用户发过的所有消息**中提取「用户上传过的文件」。
 *
 * 判据只有一句：**文件出现在用户消息里 → 能读；没出现 → 不读。**
 *   - IM 上传：消息里带 `<dsh_im_files>` 清单（路径相对工作区）
 *   - GUI 上传：消息里是结构化文件块（`attachmentId` = sha256 摘要 + 文件名）
 *
 * 为什么扫全会话而不只看本轮：真实用法是「第一轮传文件、第二轮说『把上个文件解析我看一下』」——
 * 只看本轮会把这种正常请求全部拒掉。仍然只认用户自己上传过的文件，不是通用读文件。
 *
 * 拿不到会话上下文时返回 `null`（此时才退回「最近上传」窗口）。
 * @returns {{ allowed: Set<string>, lastTurn: Set<string> }|null}
 */
function sessionUploads(exec) {
  const session = exec?.agent?.session
  const events = typeof session?.snapshotEvents === 'function' ? session.snapshotEvents() : null
  if (!Array.isArray(events)) return null

  const userMessages = events.filter((event) => event?.type === 'user/message')
  if (userMessages.length === 0) return null

  const allowed = new Set()
  const lastTurn = new Set()
  const recent = userMessages.slice(-100)
  let sawManifest = false

  recent.forEach((message, index) => {
    const text = textOfUserMessage(message)
    const isLast = index === recent.length - 1
    if (text) {
      if (/<dsh_im_files>/.test(text) || /read-only copy saved at\s*"/.test(text)) sawManifest = true
      for (const key of uploadKeysInText(text)) {
        allowed.add(key)
        if (isLast) lastTurn.add(key)
      }
    }
    const blockKeys = uploadKeysInBlocks(message)
    if (blockKeys.size > 0) sawManifest = true
    for (const key of blockKeys) {
      allowed.add(key)
      if (isLast) lastTurn.add(key)
    }
  })

  return { allowed, lastTurn, sawManifest }
}

/**
 * GUI 上传在会话里是**结构化文件块**（不是文本）：
 *   `{"type":"file","attachment":{"attachmentId":"sha256:<digest>","name":"abc.xlsx"}}`
 * 里面只有摘要和文件名，没有路径 —— 按附件库布局反推：
 *   `<store>/files/<digest 前 2 位>/<digest>/<name>`
 */
function uploadKeysInBlocks(message) {
  const keys = new Set()
  const blocks = message?.data?.content ?? message?.data?.message?.content
  if (!Array.isArray(blocks)) return keys

  const store = activeConfig.attachmentRoot || defaultAttachmentRoot()
  for (const block of blocks) {
    const attachment = block?.attachment
    const id = typeof attachment?.attachmentId === 'string' ? attachment.attachmentId : null
    if (!id) continue
    const digest = id.startsWith('sha256:') ? id.slice('sha256:'.length) : id
    if (!/^[0-9a-f]{8,}$/i.test(digest)) continue
    const name = typeof attachment?.name === 'string' ? attachment.name : null

    const candidates = []
    if (name) candidates.push(join(store, 'files', digest.slice(0, 2), digest, name))
    candidates.push(join(store, 'objects', digest.slice(0, 2), digest))
    for (const candidate of candidates) {
      const absolute = resolve(candidate)
      if (existsSync(absolute) && isUploadPath(absolute)) keys.add(pathKey(absolute))
    }
  }
  return keys
}

/** 取一条用户消息里的纯文本。 */
function textOfUserMessage(message) {
  const data = message?.data ?? {}
  const blocks = data.content ?? data.message?.content ?? []
  if (typeof blocks === 'string') return blocks
  if (!Array.isArray(blocks)) return null
  const texts = blocks.filter((block) => typeof block?.text === 'string').map((block) => block.text)
  return texts.length ? texts.join('\n') : null
}

/** 从一段消息文本里解析出「上传件路径键」。 */
function uploadKeysInText(text) {
  const keys = new Set()
  const add = (candidate) => {
    const raw = String(candidate).trim()
    // 消息文本里可能是单反斜杠，也可能被转义成双反斜杠；两种都试
    for (const variant of [raw, raw.replace(/\\\\/g, '\\')]) {
      const absolute = resolve(currentRoot(), variant)
      if (isUploadPath(absolute)) keys.add(pathKey(absolute))
    }
  }

  // IM 清单：优先按 JSON 解析；解析不了就用正则兜底（清单里的 path 是 Windows 反斜杠，
  // 万一转义被处理坏了，JSON.parse 会整个失败，不能因此丢掉「上传了哪些文件」）
  for (const manifest of text.matchAll(/<dsh_im_files>\s*([\s\S]*?)\s*<\/dsh_im_files>/g)) {
    const body = manifest[1]
    let parsed = false
    try {
      const document = JSON.parse(body)
      for (const file of document?.files ?? []) {
        const path = typeof file === 'string' ? file : (file?.path ?? file?.relativePath)
        if (typeof path === 'string' && path.trim()) add(path)
      }
      parsed = true
    } catch {
      /* 落到下面的正则兜底 */
    }
    if (!parsed) {
      for (const match of body.matchAll(/"path"\s*:\s*"((?:[^"\\]|\\.)*)"/g)) {
        add(match[1].replace(/\\\\/g, '\\').replace(/\\"/g, '"'))
      }
    }
  }

  // GUI 只读副本
  for (const match of text.matchAll(/read-only copy saved at\s*"([^"]+)"/g)) add(match[1])

  return keys
}

/* ─────────────────────────── 各格式处理 ─────────────────────────── */

const TEXT_EXTENSIONS = new Set(['.csv', '.tsv', '.txt', '.md', '.json', '.xml', '.html', '.htm', '.log', '.yml', '.yaml', '.ini'])
const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp'])

/** 统一的执行入口：按扩展名分派。 */
function readDocument(args, uploads) {
  const resolved = resolveInput(args.path, uploads)
  if (!resolved.ok) return resolved

  const ext = extname(resolved.target).toLowerCase()
  const notes = []
  const writeImage = imageWriter(resolved.target)
  let images = []
  let text = ''
  let kind = 'text'
  let pages = null

  try {
    if (IMAGE_EXTENSIONS.has(ext)) {
      return {
        ok: true,
        kind: 'image',
        source: resolved.source,
        matchedBy: resolved.matchedBy,
        path: resolved.target,
        relativePath: resolved.relativePath,
        mediaType: `image/${ext === '.jpg' ? 'jpeg' : ext.slice(1)}`,
        note: '这是图片文件：用 read_image 查看内容。',
      }
    }

    if (ext === '.xlsx' || ext === '.xlsm') {
      kind = 'sheet'
      const book = readXlsx(readFileSync(resolved.target), { maxRows: activeConfig.maxRows, maxCols: activeConfig.maxCols })
      text = book.sheets
        .map((sheet) => `【工作表：${sheet.name}】${sheet.rows} 行${sheet.truncated ? '（已截断）' : ''}\n${sheet.text}`)
        .join('\n\n')
      images = book.images
    } else if (ext === '.docx' || ext === '.docm') {
      kind = 'document'
      const doc = readDocx(readFileSync(resolved.target))
      text = doc.text
      images = doc.images
    } else if (ext === '.pptx' || ext === '.pptm') {
      kind = 'slides'
      const deck = readPptx(readFileSync(resolved.target))
      text = deck.text
      images = deck.images
    } else if (ext === '.pdf') {
      kind = 'pdf'
      const pdf = readPdf(readFileSync(resolved.target), { maxPages: activeConfig.maxPages })
      notes.push(...pdf.notes)
      text = pdf.pages
        .map((page) => {
          const head = `【第 ${page.index} 页 · ${labelOfKind(page.kind)}】`
          if (page.text) return `${head}\n${page.text}`
          return page.images.length ? `${head}（无文本层，已抽出 ${page.images.length} 张图片）` : `${head}（这页没有可读内容）`
        })
        .join('\n\n')
      pages = pdf.pages.map((page) => ({ index: page.index, kind: page.kind, chars: page.text.replace(/\s/g, '').length, images: page.images.map((one) => one.name) }))
      images = pdf.pages.flatMap((page) => page.images)
    } else if (TEXT_EXTENSIONS.has(ext)) {
      text = readFileSync(resolved.target, 'utf8')
    } else if (ext === '.doc' || ext === '.xls' || ext === '.ppt') {
      kind = 'legacy'
      const found = findSoffice(activeConfig.soffice)
      if (found.error) return { ok: false, error: found.error }
      const scratch = makeScratch()
      try {
        const filter = ext === '.xls' ? 'csv' : 'txt'
        const converted = convertToText({
          soffice: found.path,
          source: resolved.target,
          outDir: scratch,
          timeoutMs: activeConfig.convertTimeoutMs,
          filter,
        })
        if (!converted.ok) return { ok: false, error: converted.error }
        text = readFileSync(converted.text, 'utf8')
        notes.push('老格式（.doc/.xls/.ppt）由 LibreOffice 转换，版式可能有差异。')
      } finally {
        cleanScratch(scratch)
      }
    } else {
      return { ok: false, error: `不支持的格式 ${ext || '(无扩展名)'}。支持：xlsx/docx/pptx/pdf/csv/txt/md/json/xml/html 与常见图片。` }
    }
  } catch (error) {
    return { ok: false, error: `读取失败：${error instanceof Error ? error.message : String(error)}` }
  }

  const clipped = clip(text)
  if (clipped.truncated) notes.push(`文本超过 ${activeConfig.maxChars} 字，已截断。`)
  if (resolved.matchedBy === 'earlier') notes.push('这个文件是本次会话较早前用户上传的（「上个文件」那类用法），按用户消息记录放行。')

  const kept = images.slice(0, activeConfig.maxImages)
  if (images.length > kept.length) notes.push(`图片共 ${images.length} 张，只抽出前 ${kept.length} 张。`)

  return {
    ok: true,
    kind,
    source: resolved.source,
    matchedBy: resolved.matchedBy,
    path: resolved.target,
    relativePath: resolved.relativePath,
    chars: clipped.chars,
    truncated: clipped.truncated,
    text: clipped.text,
    images: kept.map(writeImage),
    ...(pages ? { pages } : {}),
    ...(notes.length ? { notes } : {}),
  }
}

function labelOfKind(kind) {
  if (kind === 'text') return '文字页'
  if (kind === 'image') return '图片页（扫描件）'
  if (kind === 'mixed') return '图文页'
  return '无内容'
}

/* ─────────────────────────── 工具定义 ─────────────────────────── */

const OUTPUT = {
  schema: { type: 'object', additionalProperties: true },
  render: (_args, value) => {
    const body = JSON.stringify(value, null, 2)
    const count = Array.isArray(value?.images) ? value.images.length : 0
    if (!value?.ok || count === 0) return [{ type: 'text', text: body }]
    return [{ type: 'text', text: `${body}\n\n图片提醒：还有 ${count} 张图片需要用 read_image 逐张查看后再回答。` }]
  },
}

const readTool = {
  name: 'read_document_noshell',
  description: [
    '解析用户上传的附件（二进制 xlsx/docx/pptx/pdf 等），返回文本并抽出内嵌图片 —— 普通 read 工具读这些只会得到乱码。',
    '本工具的实现：解析全程在宿主 Node 进程内用纯 JS 完成，零第三方依赖、不需要 Python，也不要求模型能起命令；只有老格式 .doc/.xls/.ppt 由插件自己拉起本机 LibreOffice。',
    '什么时候调：只有本轮用户消息里出现 <dsh_im_files> 清单（IM 上传）或「只读副本」路径（GUI 上传）时才调，路径照抄清单里那一个；没有这两种信号就别调。',
    '限制：只放行本轮清单里的文件；工作区里的普通文本文件请用 read，图片用 read_image。清单外的路径（含同目录的历史上传件）会被拒绝。',
    '支持 .xlsx/.xlsm（按工作表输出文本表格）、.docx、.pptx、.pdf（逐页判定：文字页给文本、扫描页抽图）、.csv/.txt/.md/.json/.xml/.html，以及 .png/.jpg 等图片（只返回路径）。',
    '返回里若带图片，请用 read_image 逐张查看后再回答；抽出的图片落在 .read-office/ 目录。',
  ].join(' '),
  parameters: {
    type: 'object',
    properties: {
      path: { type: 'string', description: '用户这一轮上传的附件路径：IM 上传的给 .dsh-im/inbound/ 下的相对路径，GUI 上传的给消息里那个「只读副本」绝对路径' },
      maxChars: { type: 'number', description: '本次返回的文本上限（字符），默认 200000' },
      maxRows: { type: 'number', description: '表格最多渲染多少行，默认 300' },
      maxPages: { type: 'number', description: 'PDF 最多解析多少页，默认 50' },
    },
    required: ['path'],
  },
  output: OUTPUT,
  async execute(args, exec) {
    const previous = { ...activeConfig }
    const cwd = exec?.agent?.session?.header?.cwd
    try {
      if (Number.isFinite(Number(args.maxChars)) && Number(args.maxChars) > 0) activeConfig.maxChars = Math.floor(Number(args.maxChars))
      if (Number.isFinite(Number(args.maxRows)) && Number(args.maxRows) > 0) activeConfig.maxRows = Math.floor(Number(args.maxRows))
      if (Number.isFinite(Number(args.maxPages)) && Number(args.maxPages) > 0) activeConfig.maxPages = Math.floor(Number(args.maxPages))
      const call = () => readDocument(args, sessionUploads(exec))
      return typeof cwd === 'string' && cwd.trim() !== '' ? callRoot.run(resolve(cwd), call) : call()
    } finally {
      activeConfig = previous
    }
  },
}

/* ────────────────────────────── 装载 ────────────────────────────── */

/** 注册工具。 */
export function apply(ctx, config) {
  activeConfig = config ?? { ...DEFAULTS }
  ctx.tools.register(readTool)
}
