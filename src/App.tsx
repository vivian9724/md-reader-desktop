import { isValidElement, type CSSProperties, type ReactNode, useEffect, useMemo, useRef, useState } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import { asBlob } from 'html-docx-js-typescript'
import { saveAs } from 'file-saver'
import { marked } from 'marked'
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Clock3,
  Download,
  FileText,
  FolderOpen,
  GripVertical,
  LayoutGrid,
  ListTree,
  Minus,
  Moon,
  PencilLine,
  Pin,
  Plus,
  Search,
  Settings2,
  Sidebar,
  TableProperties,
  X,
} from 'lucide-react'
import './App.css'

const recentStorageKey = 'md-reader-recent-files'

const sampleMarkdown = `# 第三章 结果一 写作框架（主线版）

> **核心问题：** Fur 如何影响 2P24 的生防能力？
>
> **主线逻辑：** Fur 缺失 -> 铁稳态紊乱 + pHlA 去抑制 -> 铁载体暴增 + 2,4-DAPG 变化 -> 抗菌活性 + 根际定殖变化 -> 生防能力净效应
>
> **标注：** [已有直接写] [待补] 需实验 [预测] 基于文献推测，补数据后修正

## 3.1 fur 基因的鉴定、突变体与互补菌株的构建

目的：证明你用的遗传工具是正确的。本节要快，不展开。

### 3.1.1 2P24 中 fur 同源基因的生物信息学分析

| 菌株 | 培养基 | 处理 | 说明 |
| :-- | :-- | :-- | :-- |
| WT(p970Km-phlA) | KB | 标准 Miller 法 | 野生型，对照组 |
| Δfur(p970Km-phlA) | KB | 标准 Miller 法 | fur 缺失突变体 |
| Δfur(pBBRKm-fur)(p970Km-phlA) | KB + Km | 标准 Miller 法 | 互补菌株（+ Km） |

## 3.2 fur 缺失对 2P24 铁稳态的影响及铁载体介导的生防功能

### 3.2.1 fur 缺失对 2P24 生长的影响

### 3.2.2 fur 缺失导致 pyoverdine 产量大幅增加

### 3.2.3 fur 缺失对胞内铁含量的影响

### 3.2.4 本节小结

## 3.3 fur 缺失对 2,4-DAPG 合成的影响

### 3.3.1 fur 缺失导致 phlA 启动子活性上调

### 3.3.2 2,4-dapg 的定量分析

### 3.3.3 本节小结

## 3.1.2 30 天实验安排（按天、含穿插）

| 天数 | 阶段 | 任务 | 关键产出 | 备注 |
| :-- | :-- | :-- | :-- | :-- |
| 第1天（周一） | 准备工作 | 菌株活化（第1-3天，必做） | 单克隆菌落 | - |
| 第2天（周二） | 准备工作 | 药剂准备清单（Day 1-2 同步准备） | 试剂/配方就绪 | 通用试剂到位 |
| 第3天（周三） | 实验执行 | 30天实验安排（按天、含穿插） | 完整计划表 | 打印放实验台 |
| 第4天（周四） | 实验执行 | H2O2 敏感性 + 运动性 | 表型数据 | 注意对照组 |
| 第5天（周五） | 结果记录 | 记录结果 + 接菌 | 原始记录表 | 及时备份 |
| 第6-7天（周末） | 休息 | 休息 / 数据初步整理 | 初步观察汇总 | 适当调整计划 |

\`\`\`text
第1周：
| 天数 | 一 | 二 | 三 | 四 | 五 | 六 | 日 |
| 配菌线 | 配药 | 配药 | 接菌 | H2O2 | 运动性 | 休息 | 休息 |
| 铁代谢线 | 待机 | 待机 | 接菌 | 铁测 D1 | 铁测 D2 | 休息 | 休息 |
| 记录 | 建表 | 复核 | 贴签 | 录入 | 复盘 | - | - |
\`\`\`
`

type Section = {
  id: string
  depth: number
  text: string
}

type TableSummary = {
  id: string
  title: string
  headers: string[]
  rowCount: number
}

type ScheduleMatrix = {
  title: string | null
  rows: string[][]
}

type OutlineGroup = {
  parent: Section
  children: Section[]
}

type DocTab = {
  id: string
  fileName: string
  filePath: string | null
  markdown: string
  isDirty: boolean
  isSample?: boolean
}

type RecentFile = {
  fileName: string
  filePath: string
  openedAt: number
}

type ViewMode = 'preview' | 'split' | 'edit'
type DragKind = 'left' | 'right' | null

type ReadingFontPreset = {
  id: string
  label: string
  bodyFont: string
  bodySize: string
  lineHeight: string
  h1Size: string
  h2Size: string
  h3Size: string
}

const readingFontPresets: ReadingFontPreset[] = [
  {
    id: 'serif-default',
    label: '衬线 默认',
    bodyFont: '"Georgia", "Times New Roman", "STSong", serif',
    bodySize: '0.92rem',
    lineHeight: '1.62',
    h1Size: '2.15rem',
    h2Size: '1.45rem',
    h3Size: '1rem',
  },
  {
    id: 'song-compact',
    label: '宋体 紧凑',
    bodyFont: '"STSong", "Songti SC", "SimSun", serif',
    bodySize: '0.88rem',
    lineHeight: '1.56',
    h1Size: '2.02rem',
    h2Size: '1.36rem',
    h3Size: '0.96rem',
  },
  {
    id: 'sans-readable',
    label: '黑体 清爽',
    bodyFont: '"PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif',
    bodySize: '0.9rem',
    lineHeight: '1.58',
    h1Size: '2.08rem',
    h2Size: '1.4rem',
    h3Size: '0.98rem',
  },
]

function slugify(value: string) {
  return value
    .toLowerCase()
    .trim()
    .replace(/[^\w\u4e00-\u9fa5\s-]/g, '')
    .replace(/\s+/g, '-')
}

function extractText(node: unknown): string {
  if (typeof node === 'string') return node
  if (Array.isArray(node)) return node.map(extractText).join('')
  if (node && typeof node === 'object' && 'props' in node) {
    const props = (node as { props?: { children?: unknown } }).props
    return extractText(props?.children)
  }
  return ''
}

function inspectMarkdown(markdown: string) {
  const tokens = marked.lexer(markdown, { gfm: true }) as Array<Record<string, unknown>>
  const sections: Section[] = []
  const tables: TableSummary[] = []
  let tableIndex = 0

  for (const token of tokens) {
    if (token.type === 'heading') {
      const text = String(token.text ?? '')
      sections.push({ id: slugify(text), depth: Number(token.depth ?? 1), text })
    }

    if (token.type === 'table') {
      tableIndex += 1
      const headers = Array.isArray(token.header)
        ? token.header.map((cell) => String((cell as { text?: string }).text ?? ''))
        : []
      const rows = Array.isArray(token.rows) ? token.rows : []
      tables.push({
        id: `table-${tableIndex}`,
        title: headers.join(' / ') || `表 ${tableIndex}`,
        headers,
        rowCount: rows.length,
      })
    }
  }

  return { sections, tables }
}

function buildOutlineGroups(sections: Section[]) {
  const root = sections.find((section) => section.depth === 1) ?? null
  const groups: OutlineGroup[] = []
  let currentGroup: OutlineGroup | null = null

  for (const section of sections) {
    if (section.depth === 2) {
      currentGroup = { parent: section, children: [] }
      groups.push(currentGroup)
      continue
    }

    if (section.depth === 3 && currentGroup) {
      currentGroup.children.push(section)
    }
  }

  return { root, groups }
}

function normalizeGridRow(row: string[], width: number) {
  return [...row, ...Array(Math.max(0, width - row.length)).fill('')].slice(0, width)
}

function parseScheduleMatrix(value: string): ScheduleMatrix | null {
  const lines = value
    .replace(/\r/g, '')
    .split('\n')
    .map((line) => line.trimEnd())
    .filter(Boolean)

  if (lines.length < 3) return null

  let title: string | null = null
  let bodyLines = lines
  const firstLine = lines[0]?.trim()

  if (
    firstLine &&
    !firstLine.includes('|') &&
    /[：:]$/.test(firstLine) &&
    (/[周天日一二三四五六]/.test(firstLine) || /day|week/i.test(firstLine))
  ) {
    title = firstLine.replace(/[：:]$/, '')
    bodyLines = lines.slice(1)
  }

  const rawRows = bodyLines
    .map((line) =>
      line
        .split('|')
        .map((cell) => cell.trim())
        .filter((cell, index, cells) => !((index === 0 || index === cells.length - 1) && cell === '')),
    )
    .filter((row) => row.length >= 3)

  if (rawRows.length < 2) return null

  const maxWidth = Math.max(...rawRows.map((row) => row.length))
  const rows = rawRows.map((row) => normalizeGridRow(row, maxWidth))
  const looksLikeSchedule = rows.some((row) =>
    row.some((cell) => /(?:周|day|第\d+天|[一二三四五六日天])/.test(cell)),
  )

  return looksLikeSchedule ? { title, rows } : null
}

function dedupeRecentFiles(items: RecentFile[]) {
  const seen = new Set<string>()
  return items.filter((item) => {
    if (seen.has(item.filePath)) return false
    seen.add(item.filePath)
    return true
  })
}

function loadRecentFiles() {
  const raw = localStorage.getItem(recentStorageKey)
  if (!raw) return [] as RecentFile[]
  try {
    const parsed = JSON.parse(raw) as RecentFile[]
    return Array.isArray(parsed) ? parsed : []
  } catch {
    return []
  }
}

function persistRecentFiles(items: RecentFile[]) {
  localStorage.setItem(recentStorageKey, JSON.stringify(items.slice(0, 10)))
}

function formatRecentTime(openedAt: number) {
  const diff = Date.now() - openedAt
  const minutes = Math.max(1, Math.floor(diff / 60000))
  if (minutes < 3) return '刚刚'
  if (minutes < 60) return `${minutes} 分钟前`
  const hours = Math.floor(minutes / 60)
  if (hours < 24) return `${hours} 小时前`
  const days = Math.floor(hours / 24)
  if (days < 7) return `${days} 天前`
  return '上周'
}

function createSampleTab(): DocTab {
  return {
    id: 'sample-home',
    fileName: '第三章写作框架.md',
    filePath: null,
    markdown: sampleMarkdown,
    isDirty: false,
    isSample: true,
  }
}

function fileToTab(file: OpenedMarkdownFile): DocTab {
  return {
    id: file.filePath || `${file.fileName}-${Date.now()}`,
    fileName: file.fileName,
    filePath: file.filePath,
    markdown: file.content,
    isDirty: false,
  }
}

async function exportDocx(fileName: string, html: string) {
  const fullHtml = `<!DOCTYPE html>
  <html lang="zh-CN">
    <head>
      <meta charset="UTF-8" />
      <style>
        body { font-family: "Georgia", "Times New Roman", serif; color: #1f2937; line-height: 1.68; font-size: 12pt; margin: 0; }
        h1, h2, h3, h4, h5, h6 { font-family: "Segoe UI", "PingFang SC", sans-serif; color: #111827; margin: 18pt 0 8pt; }
        p, ul, ol, blockquote, pre, table { margin: 0 0 12pt; }
        blockquote { border-left: 3px solid #8fb6bc; padding-left: 12pt; color: #475569; }
        pre { background: #f3f4f6; padding: 12pt; border-radius: 8px; white-space: pre-wrap; }
        code { font-family: "Consolas", monospace; }
        table { border-collapse: collapse; width: 100%; font-size: 10.5pt; }
        th, td { border: 1px solid #d6dbe4; padding: 6pt 8pt; vertical-align: top; }
      </style>
    </head>
    <body>${html}</body>
  </html>`

  const output = await asBlob(fullHtml, {
    margins: { top: 720, right: 720, bottom: 720, left: 720 },
  })

  const blob =
    output instanceof Blob
      ? output
      : new Blob([new Uint8Array(output as ArrayLike<number>)], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })

  saveAs(blob, fileName.replace(/\.(md|markdown|txt)$/i, '') + '.docx')
}

function ScheduleMatrixBlock({ block }: { block: ScheduleMatrix }) {
  return (
    <section className="schedule-card">
      {block.title ? <div className="schedule-title">{block.title}</div> : null}
      <div className="schedule-rows">
        {block.rows.map((row, rowIndex) => (
          <div
            key={`row-${rowIndex}`}
            className={`schedule-row ${rowIndex === 0 ? 'schedule-row-header' : ''}`}
            style={{ gridTemplateColumns: `repeat(${row.length}, minmax(88px, 1fr))` }}
          >
            {row.map((cell, cellIndex) => (
              <div
                key={`${rowIndex}-${cellIndex}`}
                className={`schedule-cell ${rowIndex === 0 ? 'schedule-cell-header' : ''}`}
              >
                {cell || ' '}
              </div>
            ))}
          </div>
        ))}
      </div>
    </section>
  )
}

function ActionButton({
  icon,
  title,
  description,
  onClick,
}: {
  icon: ReactNode
  title: string
  description: string
  onClick: () => void
}) {
  return (
    <button className="action-card" type="button" onClick={onClick}>
      <span className="action-icon">{icon}</span>
      <span className="action-copy">
        <strong>{title}</strong>
        <span>{description}</span>
      </span>
    </button>
  )
}

function App() {
  const [tabs, setTabs] = useState<DocTab[]>([createSampleTab()])
  const [activeTabId, setActiveTabId] = useState('sample-home')
  const [viewMode, setViewMode] = useState<ViewMode>('preview')
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>([])
  const [isExporting, setIsExporting] = useState(false)
  const [leftSidebarWidth, setLeftSidebarWidth] = useState(286)
  const [rightSidebarWidth, setRightSidebarWidth] = useState(284)
  const [dragging, setDragging] = useState<DragKind>(null)
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({})
  const [fontMenuOpen, setFontMenuOpen] = useState(false)
  const [fontPresetId, setFontPresetId] = useState(readingFontPresets[0].id)
  const [tableFitEnabled, setTableFitEnabled] = useState(false)
  const previewRef = useRef<HTMLElement | null>(null)
  const isDesktopMode = Boolean(window.desktopBridge)

  const activeTab = useMemo(() => tabs.find((tab) => tab.id === activeTabId) ?? tabs[0], [activeTabId, tabs])
  const stats = useMemo(() => inspectMarkdown(activeTab?.markdown ?? ''), [activeTab?.markdown])
  const outline = useMemo(() => buildOutlineGroups(stats.sections), [stats.sections])
  const activeFontPreset = useMemo(
    () => readingFontPresets.find((preset) => preset.id === fontPresetId) ?? readingFontPresets[0],
    [fontPresetId],
  )
  const paperStyle = useMemo(
    () =>
      ({
        '--reader-font': activeFontPreset.bodyFont,
        '--reader-font-size': activeFontPreset.bodySize,
        '--reader-line-height': activeFontPreset.lineHeight,
        '--reader-h1-size': activeFontPreset.h1Size,
        '--reader-h2-size': activeFontPreset.h2Size,
        '--reader-h3-size': activeFontPreset.h3Size,
      }) as CSSProperties,
    [activeFontPreset],
  )

  useEffect(() => {
    document.title = 'MD Reader'
    setRecentFiles(loadRecentFiles())
  }, [])

  useEffect(() => {
    function handleMouseMove(event: MouseEvent) {
      if (!dragging) return
      if (dragging === 'left') {
        setLeftSidebarWidth(Math.max(250, Math.min(320, event.clientX)))
      }
      if (dragging === 'right') {
        setRightSidebarWidth(Math.max(250, Math.min(320, window.innerWidth - event.clientX)))
      }
    }

    function handleMouseUp() {
      setDragging(null)
    }

    window.addEventListener('mousemove', handleMouseMove)
    window.addEventListener('mouseup', handleMouseUp)
    return () => {
      window.removeEventListener('mousemove', handleMouseMove)
      window.removeEventListener('mouseup', handleMouseUp)
    }
  }, [dragging])

  useEffect(() => {
    if (!window.desktopBridge) return
    return window.desktopBridge.onMarkdownFileOpened((file) => {
      openFileInTab(file)
    })
  }, [recentFiles])

  function updateRecentFiles(file: OpenedMarkdownFile) {
    if (!file.filePath) return
    const nextRecent = dedupeRecentFiles([
      { fileName: file.fileName, filePath: file.filePath, openedAt: Date.now() },
      ...recentFiles,
    ]).slice(0, 10)
    setRecentFiles(nextRecent)
    persistRecentFiles(nextRecent)
  }

  function openFileInTab(file: OpenedMarkdownFile) {
    const nextTab = fileToTab(file)
    setTabs((currentTabs) => {
      const existing = currentTabs.find((tab) => tab.filePath && tab.filePath === file.filePath)
      if (existing) {
        return currentTabs.map((tab) =>
          tab.id === existing.id
            ? { ...tab, fileName: file.fileName, filePath: file.filePath, markdown: file.content, isDirty: false }
            : tab,
        )
      }
      const withoutSample = currentTabs.length === 1 && currentTabs[0]?.isSample ? [] : currentTabs
      return [...withoutSample, nextTab]
    })
    setActiveTabId(nextTab.id)
    updateRecentFiles(file)
  }

  function updateActiveTab(updater: (tab: DocTab) => DocTab) {
    if (!activeTab) return
    setTabs((currentTabs) => currentTabs.map((tab) => (tab.id === activeTab.id ? updater(tab) : tab)))
  }

  async function handleOpenClick() {
    if (!window.desktopBridge) return
    const file = await window.desktopBridge.openMarkdownFile()
    if (file) openFileInTab(file)
  }

  async function handleRecentOpen(filePath: string) {
    if (!window.desktopBridge) return
    const file = await window.desktopBridge.readMarkdownFile(filePath)
    if (file) openFileInTab(file)
  }

  async function handleSave(saveAsNew = false) {
    if (!activeTab || !window.desktopBridge) return
    const saved = await window.desktopBridge.saveMarkdownFile({
      filePath: saveAsNew ? null : activeTab.filePath,
      defaultName: activeTab.fileName,
      content: activeTab.markdown,
    })
    if (!saved) return

    setTabs((currentTabs) =>
      currentTabs.map((tab) =>
        tab.id === activeTab.id
          ? {
              ...tab,
              id: saved.filePath || tab.id,
              fileName: saved.fileName,
              filePath: saved.filePath,
              markdown: saved.content,
              isDirty: false,
              isSample: false,
            }
          : tab,
      ),
    )

    setActiveTabId(saved.filePath || activeTab.id)
    updateRecentFiles(saved)
  }

  async function handleExport() {
    if (!previewRef.current || !activeTab) return
    setIsExporting(true)
    try {
      await exportDocx(activeTab.fileName, previewRef.current.innerHTML)
    } finally {
      setIsExporting(false)
    }
  }

  function handleCloseTab(tabId: string) {
    setTabs((currentTabs) => {
      const nextTabs = currentTabs.filter((tab) => tab.id !== tabId)
      if (nextTabs.length === 0) {
        const sample = createSampleTab()
        setActiveTabId(sample.id)
        return [sample]
      }
      if (activeTabId === tabId) {
        setActiveTabId(nextTabs[Math.max(0, nextTabs.length - 1)].id)
      }
      return nextTabs
    })
  }

  function handleNewDraft() {
    const draftId = `draft-${Date.now()}`
    const draftTab: DocTab = {
      id: draftId,
      fileName: '未命名.md',
      filePath: null,
      markdown: '# 新文档\n\n在这里开始写 Markdown。',
      isDirty: true,
    }
    setTabs((currentTabs) => {
      const withoutSample = currentTabs.length === 1 && currentTabs[0]?.isSample ? [] : currentTabs
      return [...withoutSample, draftTab]
    })
    setActiveTabId(draftId)
    setViewMode('split')
  }

  function handleLoadSample() {
    const sample = createSampleTab()
    setTabs((currentTabs) => {
      const existing = currentTabs.find((tab) => tab.isSample)
      if (existing) return currentTabs.map((tab) => (tab.id === existing.id ? sample : tab))
      return [sample, ...currentTabs]
    })
    setActiveTabId(sample.id)
    setViewMode('preview')
  }

  function toggleOutlineGroup(groupId: string) {
    setCollapsedGroups((current) => ({
      ...current,
      [groupId]: !current[groupId],
    }))
  }

  const layoutStyle = {
    gridTemplateColumns: `${leftSidebarWidth}px 12px minmax(0, 1fr) 12px ${rightSidebarWidth}px`,
  }

  return (
    <div className="app-shell" style={layoutStyle}>
      <header className="topbar app-topbar">
        <div className="topbar-brand">
          <div className="window-dots" aria-hidden="true">
            <span />
            <span />
            <span />
          </div>
          <div className="brand-mark">
            <div className="brand-logo">MD</div>
            <span>MD Reader</span>
          </div>
        </div>

        <div className="topbar-center">
          <div className="tabs-strip">
            {tabs.map((tab) => (
              <button
                key={tab.id}
                className={`doc-tab ${tab.id === activeTabId ? 'active' : ''}`}
                type="button"
                onClick={() => setActiveTabId(tab.id)}
              >
                <FileText size={13} />
                <span>{tab.fileName}</span>
                {tab.isDirty ? <em className="dirty-dot" /> : null}
                {tabs.length > 1 ? (
                  <span
                    className="tab-close"
                    onClick={(event) => {
                      event.stopPropagation()
                      handleCloseTab(tab.id)
                    }}
                  >
                    <X size={11} />
                  </span>
                ) : null}
              </button>
            ))}
          </div>
          <button className="toolbar-icon-button plus-button" type="button" onClick={handleNewDraft} title="新建草稿">
            <Plus size={18} />
          </button>
        </div>

        <div className="view-toggle">
          <button className={viewMode === 'preview' ? 'selected' : ''} type="button" onClick={() => setViewMode('preview')}>
            阅读
          </button>
          <button className={viewMode === 'split' ? 'selected' : ''} type="button" onClick={() => setViewMode('split')}>
            双栏
          </button>
          <button className={viewMode === 'edit' ? 'selected' : ''} type="button" onClick={() => setViewMode('edit')}>
            编辑
          </button>
        </div>

        <div className="toolbar-icons">
          <button className="toolbar-icon-button" type="button" onClick={() => void handleOpenClick()} title="打开 Markdown">
            <Search size={16} />
          </button>
          <button className="toolbar-icon-button" type="button" onClick={handleNewDraft} title="新建草稿">
            <LayoutGrid size={16} />
          </button>
          <button className="toolbar-icon-button" type="button" onClick={() => void handleSave(false)} title="保存当前文档">
            <Settings2 size={16} />
          </button>
          <button className="toolbar-icon-button" type="button" onClick={() => setViewMode((current) => (current === 'split' ? 'preview' : 'split'))} title="切换双栏">
            <Sidebar size={16} />
          </button>
        </div>
      </header>

      <aside className="sidebar left-sidebar">
        <section className="panel-card current-doc-card">
          <div className="card-title">
            <FileText size={16} />
            <h2>当前文档</h2>
          </div>

          <div className="document-summary">
            <div className="document-chip">
              <FileText size={17} />
            </div>
            <div className="document-copy">
              <strong>{activeTab?.fileName ?? '未选择文档'}</strong>
              <span>{activeTab?.filePath ?? '默认主页示例'}</span>
            </div>
          </div>

          <div className="stats-grid">
            <div>
              <span>章节</span>
              <strong>{stats.sections.length}</strong>
            </div>
            <div>
              <span>表格</span>
              <strong>{stats.tables.length}</strong>
            </div>
            <div>
              <span>字数</span>
              <strong>{(activeTab?.markdown ?? '').replace(/\s+/g, '').length}</strong>
            </div>
          </div>
        </section>

        <section className="panel-card recent-card">
          <div className="card-title">
            <Clock3 size={16} />
            <h2>最近打开</h2>
          </div>

          <div className="recent-list">
            {recentFiles.length === 0 ? <span className="empty-text">还没有历史文件记录。</span> : null}
            {recentFiles.map((item, index) => {
              const isCurrent = item.filePath === activeTab?.filePath
              return (
                <button
                  key={item.filePath}
                  className={`recent-item ${isCurrent ? 'active' : ''}`}
                  type="button"
                  onClick={() => void handleRecentOpen(item.filePath)}
                >
                  <div className="recent-item-icon">
                    <FileText size={14} />
                  </div>
                  <div className="recent-item-main">
                    <strong>{item.fileName}</strong>
                    {index === 0 || isCurrent ? <span>{item.filePath}</span> : null}
                  </div>
                  <div className="recent-item-meta">
                    {isCurrent ? <Pin size={13} /> : null}
                    <em>{index === 0 ? '刚刚' : formatRecentTime(item.openedAt)}</em>
                  </div>
                </button>
              )
            })}
          </div>
        </section>

        <section className="actions-panel">
          <ActionButton
            icon={<FolderOpen size={17} />}
            title="打开 Markdown"
            description="调用系统文件选择器读取本地文件"
            onClick={() => void handleOpenClick()}
          />
          <ActionButton
            icon={<Download size={17} />}
            title={isExporting ? '正在导出 Word...' : '导出为 Word'}
            description="按当前阅读视图生成 .docx"
            onClick={() => void handleExport()}
          />
        </section>

        <div className="sidebar-footer">
          <button className="footer-button" type="button">
            <Settings2 size={16} />
            <span>设置</span>
          </button>
          <button className="footer-icon-button" type="button" onClick={handleLoadSample} title="返回主页示例">
            <PencilLine size={16} />
          </button>
        </div>
      </aside>

      <div className={`resize-rail ${dragging === 'left' ? 'active' : ''}`} onMouseDown={() => setDragging('left')}>
        <GripVertical size={14} />
      </div>

      <main className="workspace-panel">
        <section className={`workspace-body mode-${viewMode}`}>
          {viewMode !== 'preview' ? (
            <section className="editor-pane surface-card">
              <div className="pane-header">
                <div className="pane-title">
                  <PencilLine size={15} />
                  <span>Markdown 编辑</span>
                </div>
                <span className="pane-hint">
                  {isDesktopMode ? '支持本地保存和另存为' : '浏览器模式仅编辑当前会话内容'}
                </span>
              </div>
              <textarea
                className="editor-textarea"
                value={activeTab?.markdown ?? ''}
                onChange={(event) => {
                  updateActiveTab((tab) => ({ ...tab, markdown: event.target.value, isDirty: true, isSample: false }))
                }}
                spellCheck={false}
              />
            </section>
          ) : null}

          {viewMode !== 'edit' ? (
            <section className="preview-pane surface-card">
              <article className="paper" ref={previewRef} style={paperStyle}>
                <ReactMarkdown
                  remarkPlugins={[remarkGfm]}
                  components={{
                    h1: ({ children }) => {
                      const text = extractText(children)
                      return <h1 id={slugify(text)}>{children}</h1>
                    },
                    h2: ({ children }) => {
                      const text = extractText(children)
                      return <h2 id={slugify(text)}>{children}</h2>
                    },
                    h3: ({ children }) => {
                      const text = extractText(children)
                      return <h3 id={slugify(text)}>{children}</h3>
                    },
                    a: ({ href, children }) => (
                      <a href={href} target="_blank" rel="noreferrer">
                        {children}
                      </a>
                    ),
                    code: ({ children, ...props }) => {
                      const isInline = 'inline' in props ? Boolean(props.inline) : false
                      return isInline ? <code>{children}</code> : <code>{children}</code>
                    },
                    pre: ({ children }) => {
                      const rawValue = extractText(children).replace(/\n$/, '')
                      const scheduleMatrix = parseScheduleMatrix(rawValue)
                      if (scheduleMatrix) return <ScheduleMatrixBlock block={scheduleMatrix} />

                      if (isValidElement(children)) {
                        const element = children as { props?: { className?: string } }
                        const codeClassName = element.props?.className ? String(element.props.className) : ''
                        return (
                          <pre className="code-block">
                            <code className={codeClassName}>{rawValue}</code>
                          </pre>
                        )
                      }

                      return (
                        <pre className="code-block">
                          <code>{rawValue}</code>
                        </pre>
                      )
                    },
                    table: ({ children }) => (
                      <section className="table-card">
                        <div className="table-scroll">
                          <table className={`reading-table ${tableFitEnabled ? 'table-fit' : ''}`}>{children}</table>
                        </div>
                      </section>
                    ),
                    blockquote: ({ children }) => <blockquote className="quote-block">{children}</blockquote>,
                  }}
                >
                  {activeTab?.markdown ?? ''}
                </ReactMarkdown>
              </article>
            </section>
          ) : null}
        </section>
      </main>

      <div className={`resize-rail ${dragging === 'right' ? 'active' : ''}`} onMouseDown={() => setDragging('right')}>
        <GripVertical size={14} />
      </div>

      <aside className="sidebar right-sidebar">
        <section className="panel-card outline-card">
          <div className="card-title">
            <ListTree size={16} />
            <h2>文档结构</h2>
          </div>

          <div className="outline-list">
            {outline.root ? (
              <a className="outline-root" href={`#${outline.root.id}`}>
                {outline.root.text}
              </a>
            ) : null}

            {outline.groups.map((group) => (
              <div className="outline-group" key={group.parent.id}>
                <button
                  className="outline-group-title"
                  type="button"
                  onClick={() => toggleOutlineGroup(group.parent.id)}
                >
                  <ChevronDown size={14} className={collapsedGroups[group.parent.id] ? 'collapsed' : ''} />
                  <span>{group.parent.text}</span>
                </button>
                {!collapsedGroups[group.parent.id] ? (
                  <div className="outline-children">
                    {group.children.map((child) => (
                      <a className="outline-child" href={`#${child.id}`} key={child.id}>
                        <span>{child.text}</span>
                        <em className="outline-dot" />
                      </a>
                    ))}
                  </div>
                ) : null}
              </div>
            ))}
          </div>
        </section>

        <section className="panel-card table-nav-card">
          <div className="card-title table-nav-heading">
            <div className="card-title">
              <TableProperties size={16} />
              <h2>表格导航</h2>
            </div>
            <div className="mini-nav-buttons">
              <button className="mini-nav-button" type="button">
                <ChevronLeft size={14} />
              </button>
              <button className="mini-nav-button" type="button">
                <ChevronRight size={14} />
              </button>
            </div>
          </div>

          <div className="table-list">
            {stats.tables.length === 0 ? <span className="empty-text">当前文档没有表格。</span> : null}
            {stats.tables.map((table, index) => (
              <div className="table-list-item" key={table.id}>
                <strong>表 {index + 1} · {table.title}</strong>
                <span>第 {table.rowCount} 行 / 共 {table.headers.length} 列</span>
              </div>
            ))}
          </div>
        </section>
      </aside>

      <footer className="statusbar">
        <div className="statusbar-left">
          <span className="status-indicator">{activeTab?.isDirty ? '未保存' : '已保存'}</span>
        </div>
        <div className="statusbar-center">
          <span>{(activeTab?.markdown ?? '').replace(/\s+/g, '').length} 字</span>
          <span>约 {Math.max(1, Math.round((activeTab?.markdown ?? '').replace(/\s+/g, '').length / 450))} 分钟阅读</span>
          <span>Markdown</span>
        </div>
        <div className="statusbar-right">
          <div className="status-font-picker">
            <button className="status-button" type="button" onClick={() => setFontMenuOpen((current) => !current)}>
              <span>Aa</span>
            </button>
            {fontMenuOpen ? (
              <div className="font-menu">
                {readingFontPresets.map((preset) => (
                  <button
                    key={preset.id}
                    className={`font-menu-item ${preset.id === activeFontPreset.id ? 'active' : ''}`}
                    type="button"
                    onClick={() => {
                      setFontPresetId(preset.id)
                      setFontMenuOpen(false)
                    }}
                  >
                    {preset.label}
                  </button>
                ))}
              </div>
            ) : null}
          </div>
          <button className="status-button" type="button">
            <Moon size={15} />
          </button>
          <button
            className={`status-button ${tableFitEnabled ? 'active' : ''}`}
            type="button"
            onClick={() => setTableFitEnabled((current) => !current)}
            title="表格自适应"
          >
            <Minus size={15} />
          </button>
          <span className="zoom-label">100%</span>
          <button className="status-button" type="button">
            <Plus size={15} />
          </button>
        </div>
      </footer>
    </div>
  )
}

export default App
