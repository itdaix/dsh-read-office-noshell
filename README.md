# dsh-read-office-noshell

> **零第三方依赖的纯 JS 实现**：解析全在宿主 Node 进程内完成 —— 无 shell 的预设能用，有 shell 的环境同样能用（比现写脚本更快）。

解析**用户上传的附件**（Word / Excel / PowerPoint / PDF / 图片）：出文本、出内嵌图。图片不读内容，抽出后交 `read_image` 让多模态模型看图。零第三方依赖，宿主 Node 进程内运行，只往 `tools` 注册表注册一个工具 `read_document_noshell`。

存在的唯一理由：上传的 xlsx / docx / pptx / pdf 用普通 `read` 读出来是乱码，而无 shell 的预设又跑不了 Python 解析。所以它**不做通用读文件**——工作区里的普通文本请用 `read`，图片用 `read_image`。

## 技术选型

| 维度 | 选择 | 说明 |
|---|---|---|
| 语言 | 纯 JS（ESM） | 无编译，宿主进程内运行 |
| 依赖 | 零第三方依赖 | 只用 Node 内置 `fs` / `zlib` / `child_process` / `async_hooks` / `os` / `path` / `url` |
| OOXML（xlsx/docx/pptx） | 自写 zip central-directory 解析 + XML 正则 | 取 XML 抽文本、取 `media/` 目录抽图 |
| PDF | 自写对象扫描 + 内容流 `Tj`/`TJ` 抽文本 | 字体 `ToUnicode` 映射还原中文；扫描页抽图交多模态 |
| 老格式（doc/xls/ppt） | 本机 LibreOffice 无头转换 | `spawnSync --convert-to txt/csv` |
| 工具注册 | Cordis 插件，注入 `tools` 注册表 | 只注册工具，不发布服务，无需 isolate realm |
| 并发隔离 | `AsyncLocalStorage` | 传调用方工作区，防 root 互相覆盖 |

## 核心流程（动作清单）

一次 `read_document_noshell` 调用 = 接管参数 → 认路径、验权限 → 按格式解析 → 裁剪、落图、交回。

### 阶段一：接管参数、隔离工作区

1. 拿到工具入参 `path`（必填）、`maxChars`、`maxRows`、`maxPages`，以及宿主执行上下文 `exec`
2. 取出调用方会话的工作区目录 `cwd`（来自 `exec.agent.session.header.cwd`）
3. 把当前配置快照存一份（`finally` 里回滚用）
4. 把 `maxChars` / `maxRows` / `maxPages` 里合法且为正数的值，临时覆盖进当前配置
5. 把会话工作区塞进 `AsyncLocalStorage` 上下文（隔离并发调用的 root）
6. 同步执行 `readDocument`，把「本次会话上传清单」`sessionUploads(exec)` 一并传进去

### 阶段二：认路径、验权限（`resolveInput`）

7. 取出当前工作区根 `root`（有会话上下文用 `cwd`，没有则退回配置的 `root` 兜底）
8. 把传入 `path` 拼成绝对路径（相对路径基于 `root` 解析）
9. 拼出白名单目录三处：IM 上传区 `root/.dsh-im`、GUI 附件库 `DSH_HOME/attachments/v1`、额外目录 `extraRoots`
10. 逐个比对：目标路径是否落在某个白名单内（`relative` 结果不含 `..` 且非绝对）
11. 不在白名单 → 拒绝「只读用户上传的附件」
12. 扫本次会话最近 100 条用户消息，提取用户上传过的文件键（IM 的 `<dsh_im_files>` 清单 / GUI 的结构化文件块）
13. 比对目标文件键是否在上传集合里；不在 → 拒绝「用户没上传过它」（拿不到会话上下文也直接拒绝）
14. 用 `realpath` 取真实路径再判一次，防符号链接 / junction 把真实文件指到白名单外
15. 校验文件存在、是普通文件（非目录）、大小 ≤ `maxFileBytes`（默认 50 MB）
16. 通过后交回给「按扩展名分派」

### 阶段三：按格式解析出文本 + 内嵌图

17. 取扩展名（小写）
18. 图片扩展名（png/jpg/webp/gif/bmp）→ 直接返回路径 + `mediaType`，不读内容，交 `read_image` 看图
19. xlsx/xlsm → 读 zip 目录表 → 取 `xl/workbook.xml` 工作表清单 → 取共享字符串表 → 逐表渲染成制表符文本（截断 `maxRows=300` 行 / `maxCols=40` 列）→ 抽 `xl/media/` 下的图
20. docx/docm → 取 `word/document.xml` → 按段落 + 表格顺序抽文本（表格单元格用 `|` 连接、内嵌图占位 `[图片]`）→ 抽 `word/media/` 下的图
21. pptx/pptm → 按 `slideN.xml` 序号升序 → 逐页取 `a:p` 段落文本 → 抽 `ppt/media/` 下的图
22. pdf → 扫 PDF 对象表 → 沿 `/Pages` 树收集页号 → 逐页（上限 `maxPages=50`）：解压内容流用 `ToUnicode` 映射还原文本；`XObject` 里抽图（DCTDecode 直出 jpg、FlateDecode 位图自拼 png）→ 有效字符 ≥ 8 判文字页 / 有图判图片页 / 都有判图文页
23. csv/txt/md/json/xml/html 等 → 直接按 utf8 读全文
24. doc/xls/ppt 老格式 → 找 LibreOffice（配置值或固定候选路径）→ 建临时目录 → 无头转 txt/csv（超时 `convertTimeoutMs=120s`）→ 读转换结果 → 删临时目录

### 阶段四：裁剪、落图、交回

25. 把文本按 `maxChars`（默认 200000 字符）截断，超了记「已截断」注记
26. 图片只留前 `maxImages`（默认 20）张，超了记注记
27. 逐张把图片字节写到 `<root>/<imageDir>/<源文件名>-<本地时间戳>/` 目录（目录延迟到真有图才建，避免纯文本读取留空目录）
28. 拼出返回对象：`kind` / `source` / `path` / `chars` / `text` / `images` / `pages` / `notes`
29. 输出渲染时若 `images > 0`，追加提醒「还有 N 张图需用 read_image 逐张查看」
30. `finally` 恢复配置到调用前快照，防参数污染下一次调用

## 配置项

| 键 | 默认 | 说明 |
|---|---|---|
| `root` | 空 | **兜底**工作区根目录（优先用会话 cwd） |
| `maxChars` | 200000 | 单次返回文本上限（字符） |
| `maxRows` | 300 | xlsx 表格最多渲染行数 |
| `maxCols` | 40 | xlsx 表格最多渲染列数 |
| `maxPages` | 50 | PDF 最多解析页数 |
| `maxImages` | 20 | 单次最多抽出图片张数 |
| `imageDir` | `.read-office` | 抽图落盘目录（相对 root） |
| `soffice` | 空 | LibreOffice 路径（留空自动探测） |
| `convertTimeoutMs` | 120000 | 老格式转换超时（毫秒） |
| `allowAttachmentStore` | true | 是否放行 GUI 附件库只读副本 |
| `attachmentRoot` | 空 | 附件库根目录，留空按 `DSH_HOME`（默认 `~/.dsh`）推导 |
| `extraRoots` | `[]` | 额外放行的绝对路径 |
| `maxFileBytes` | 52428800 | 单文件大小上限（字节，即 50 MB） |

## 工具入参（`read_document_noshell`）

| 参数 | 类型 | 必填 | 默认 | 说明 |
|---|---|---|---|---|
| `path` | string | 是 | — | 上传件路径：IM 给 `.dsh-im/inbound/` 下相对路径，GUI 给「只读副本」绝对路径 |
| `maxChars` | number | 否 | 200000 | 本次返回文本上限 |
| `maxRows` | number | 否 | 300 | 表格最多渲染行数 |
| `maxPages` | number | 否 | 50 | PDF 最多解析页数 |

## 关键设计点

- **只读上传件、不做通用读文件**：双保险——目录白名单 + 用户消息上传清单，再加 `realpath` 判一次，防链接绕道指到白名单外。
- **工作区归属与并发隔离**：root 优先取会话 `cwd`，用 `AsyncLocalStorage` 传递；改模块级配置会让并发调用 root 互相覆盖、写错工作区。
- **PDF 零依赖自解析**：文字页 / 图片页按有效字符数（阈值 8）判定；文字用字体 `ToUnicode` 还原中文；扫描页抽图交多模态看，不猜 OCR。JBIG2 / CCITT / LZW / 加密 PDF 明确报「不支持」。
- **图片不读内容**：抽出来只落盘 + 返回路径 / 尺寸，交给 `read_image` 多模态看图。
- **支持「上个文件」用法**：扫全会话最近 100 条用户消息，第一轮传文件、第二轮说「解析上个文件」也能放行。
- **参数临时覆盖 + finally 回滚**：本次调用临时改 `maxChars` 等，结束后恢复，不污染后续调用。
