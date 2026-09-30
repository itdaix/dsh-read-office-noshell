/**
 * LibreOffice 无头转换的调度：只给老格式（.doc / .xls / .ppt）兜底。
 *
 * 新格式（xlsx / docx / pptx / pdf）我们自己解析，不依赖 LibreOffice；
 * 这里只处理那三种老二进制格式 —— 自己写解析器不划算。
 *
 * @module lib/soffice
 */

import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

const CANDIDATES = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/usr/bin/soffice',
  '/usr/local/bin/soffice',
]

/** 找 LibreOffice；找不到就返回 `{ error }`。 */
export function findSoffice(configured) {
  if (configured) {
    return existsSync(configured) ? { path: configured } : { error: `配置的 LibreOffice 不存在：${configured}` }
  }
  const found = CANDIDATES.find((candidate) => existsSync(candidate))
  return found ? { path: found } : { error: `没找到 LibreOffice（找过：${CANDIDATES.join(' / ')}）。` }
}

/** 临时目录，用完即删。 */
export function makeScratch() {
  const dir = join(tmpdir(), `dsh-read-office-noshell-${process.pid}-${Date.now().toString(36)}`)
  mkdirSync(dir, { recursive: true })
  return dir
}

export function cleanScratch(dir) {
  try {
    rmSync(dir, { recursive: true, force: true })
  } catch {
    /* 清不掉不影响结果 */
  }
}

/**
 * 把一个文件转成文本类产物。
 * @returns `{ ok: true, text }` 或 `{ ok: false, error }`
 */
export function convertToText({ soffice, source, outDir, timeoutMs = 120000, filter, capture = true }) {
  const profile = join(tmpdir(), 'dsh-read-office-noshell-loprofile')
  const args = [
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    '--headless',
    '--norestore',
    '--convert-to',
    filter,
    '--outdir',
    outDir,
    source,
  ]

  const result = capture
    ? spawnSync(soffice, args, { timeout: timeoutMs, encoding: 'utf8' })
    : spawnSync(soffice, args, { timeout: timeoutMs, stdio: 'ignore' })

  // 受限沙箱里开不了管道时退一次不捕获输出。
  if (result.error?.code === 'EPERM' && capture) {
    return convertToText({ soffice, source, outDir, timeoutMs, filter, capture: false })
  }
  if (result.error) return { ok: false, error: `调用 LibreOffice 失败：${result.error.message}` }

  const wanted = filter.split(':')[0].toLowerCase()
  const produced = readdirSync(outDir).find((name) => name.toLowerCase().endsWith(`.${wanted}`))
  if (!produced) {
    const detail = `${result.stdout ?? ''}${result.stderr ?? ''}`.trim().split('\n').slice(-2).join(' ')
    return { ok: false, error: `LibreOffice 没有产出 .${wanted}（退出码 ${result.status}）。${detail}` }
  }
  return { ok: true, text: join(outDir, produced) }
}
