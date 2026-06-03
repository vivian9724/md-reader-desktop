import { useEffect, useMemo, useRef, useState, type ChangeEvent } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import { asBlob } from 'html-docx-js-typescript'
import { saveAs } from 'file-saver'
import { marked } from 'marked'
import {
  Download,
  FileText,
  FolderOpen,
  LayoutTemplate,
  ListTree,
  MonitorSmartphone,
  TableProperties,
} from 'lucide-react'
import './App.css'

const sampleMarkdown = `# 产品周报

这是一个更适合阅读的 Markdown 视图示例。目标不是改写内容，而是把原始 Markdown 用更接近 Word 文档的可视方式展示出来。

## 本周摘要

- 阅读区尽量像纸面文档
- 表格做成更强的视觉卡片
- 可以直接导出为 \`.docx\`
- 安装后支持双击 \`.md\` 文件打开

## 重点数据

| 指标 | 本周 | 上周 | 变化 |
| :-- | --: | --: | :-- |
| 活跃用户 | 18,240 | 16,980 | +7.4% |
| 付费转化率 | 4.8% | 4.4% | +0.4pp |
| 工单完成率 | 92% | 88% | +4pp |

## 备注

> 这个版本更偏“阅读器”而不是“编辑器”。

\`\`\`ts
export function keepContentStable(markdown: string) {
  return markdown
}
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

function slugify(value: string) {
  return value
    .toLowerCase()
    .trim()
    .replace(/[^\w\u4e00-\u9fa5\s-]/g, '')
    .replace(/\s+/g, '-')
}

function extractText(node: unknown): string {
  if (typeof node === 'string') {
    return node
  }

  if (Array.isArray(node)) {
    return node.map(extractText).join('')
  }

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
      sections.push({
        id: slugify(text),
        depth: Number(token.depth ?? 1),
        text,
      })
    }

    if (token.type === 'table') {
      tableIndex += 1
      const header = Array.isArray(token.header)
        ? token.header.map((cell) => String((cell as { text?: string }).text ?? ''))
        : []
      const rows = Array.isArray(token.rows) ? token.rows : []
      const title = header[0] ? `${header[0]} 表` : `表格 ${tableIndex}`

      tables.push({
        id: `table-${tableIndex}`,
        title,
        headers: header,
        rowCount: rows.length,
      })
    }
  }

  return { sections, tables }
}

async function exportDocx(fileName: string, html: string) {
  const fullHtml = `<!DOCTYPE html>
  <html lang="zh-CN">
    <head>
      <meta charset="UTF-8" />
      <style>
        body {
          font-family: "Georgia", "Times New Roman", serif;
          color: #1f2937;
          line-height: 1.65;
          font-size: 12pt;
          margin: 0;
        }
        h1, h2, h3, h4, h5, h6 {
          font-family: "Segoe UI", "PingFang SC", sans-serif;
          color: #111827;
          margin: 18pt 0 8pt;
        }
        p, ul, ol, blockquote, pre, table {
          margin: 0 0 12pt;
        }
        blockquote {
          border-left: 3px solid #94a3b8;
          padding-left: 12pt;
          color: #475569;
        }
        pre {
          background: #f3f4f6;
          padding: 12pt;
          border-radius: 8px;
          white-space: pre-wrap;
        }
        code {
          font-family: "Consolas", monospace;
        }
        table {
          border-collapse: collapse;
          width: 100%;
          font-size: 10.5pt;
        }
        th, td {
          border: 1px solid #d1d5db;
          padding: 6pt 8pt;
          vertical-align: top;
        }
        th {
          background: #eff6ff;
          font-family: "Segoe UI", sans-serif;
        }
        hr {
          border: none;
          border-top: 1px solid #d1d5db;
          margin: 16pt 0;
        }
      </style>
    </head>
    <body>${html}</body>
  </html>`

  const output = await asBlob(fullHtml, {
    margins: {
      top: 720,
      right: 720,
      bottom: 720,
      left: 720,
    },
  })

  const blob =
    output instanceof Blob
      ? output
      : new Blob([new Uint8Array(output as ArrayLike<number>)], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })

  saveAs(blob, fileName.replace(/\.(md|markdown|txt)$/i, '') + '.docx')
}

function applyOpenedFile(file: OpenedMarkdownFile, setMarkdown: (value: string) => void, setFileName: (value: string) => void, setFilePath: (value: string) => void) {
  setMarkdown(file.content)
  setFileName(file.fileName)
  setFilePath(file.filePath)
}

function App() {
  const [markdown, setMarkdown] = useState(sampleMarkdown)
  const [fileName, setFileName] = useState('demo.md')
  const [filePath, setFilePath] = useState('未绑定到本地文件')
  const [isExporting, setIsExporting] = useState(false)
  const [isDesktopMode] = useState(Boolean(window.desktopBridge))
  const previewRef = useRef<HTMLElement | null>(null)

  const stats = useMemo(() => inspectMarkdown(markdown), [markdown])

  useEffect(() => {
    if (!window.desktopBridge) {
      return
    }

    return window.desktopBridge.onMarkdownFileOpened((file) => {
      applyOpenedFile(file, setMarkdown, setFileName, setFilePath)
    })
  }, [])

  async function handleBrowserFileOpen(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    const nextMarkdown = await file.text()
    setMarkdown(nextMarkdown)
    setFileName(file.name)
    setFilePath(file.name)
  }

  async function handleOpenClick() {
    if (!window.desktopBridge) {
      return
    }

    const file = await window.desktopBridge.openMarkdownFile()
    if (!file) {
      return
    }

    applyOpenedFile(file, setMarkdown, setFileName, setFilePath)
  }

  async function handleExport() {
    if (!previewRef.current) {
      return
    }

    setIsExporting(true)
    try {
      await exportDocx(fileName, previewRef.current.innerHTML)
    } finally {
      setIsExporting(false)
    }
  }

  return (
    <div className="shell">
      <aside className="panel panel-left">
        <div className="panel-header">
          <span className="eyebrow">Markdown Reader</span>
          <h1>读 Markdown，像读 Word 一样顺。</h1>
          <p>保留原文内容，用更稳定的文档预览和更强的表格阅读方式来展示。</p>
        </div>

        <div className="actions">
          {isDesktopMode ? (
            <button className="action-card" type="button" onClick={handleOpenClick}>
              <FolderOpen size={18} />
              <div>
                <strong>打开 Markdown</strong>
                <span>调用系统文件选择器读取本地文件</span>
              </div>
            </button>
          ) : (
            <label className="action-card upload-card">
              <input type="file" accept=".md,.markdown,.txt" onChange={handleBrowserFileOpen} />
              <FolderOpen size={18} />
              <div>
                <strong>打开 Markdown</strong>
                <span>读取本地 .md 文件</span>
              </div>
            </label>
          )}

          <button className="action-card" type="button" onClick={handleExport} disabled={isExporting}>
            <Download size={18} />
            <div>
              <strong>{isExporting ? '正在导出 DOCX...' : '导出为 Word'}</strong>
              <span>按当前阅读视图生成 .docx</span>
            </div>
          </button>

          <button
            className="action-card"
            type="button"
            onClick={() => {
              setMarkdown(sampleMarkdown)
              setFileName('demo.md')
              setFilePath('内置示例文档')
            }}
          >
            <LayoutTemplate size={18} />
            <div>
              <strong>载入示例</strong>
              <span>查看默认排版与表格样式</span>
            </div>
          </button>
        </div>

        <div className="meta-card">
          <div className="meta-row">
            <span>当前文件</span>
            <strong>{fileName}</strong>
          </div>
          <div className="meta-row meta-row-stack">
            <span>来源路径</span>
            <strong className="path-text">{filePath}</strong>
          </div>
          <div className="meta-grid">
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
              <strong>{markdown.replace(/\s+/g, '').length}</strong>
            </div>
          </div>
        </div>
      </aside>

      <main className="center-stage">
        <div className="stage-topbar">
          <div className="window-dots" aria-hidden="true">
            <span />
            <span />
            <span />
          </div>
          <div className="document-name">
            <FileText size={16} />
            <span>{fileName}</span>
          </div>
        </div>

        <article className="paper" ref={previewRef}>
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
              code: ({ className, children }) => {
                const isBlock = Boolean(className)
                if (!isBlock) {
                  return <code>{children}</code>
                }

                return (
                  <pre className="code-block">
                    <code className={className}>{children}</code>
                  </pre>
                )
              },
              table: ({ children }) => (
                <section className="table-card">
                  <header className="table-card-header">
                    <div>
                      <strong>表格阅读视图</strong>
                      <span>横向信息更清楚，导出时仍保留为表格。</span>
                    </div>
                  </header>
                  <div className="table-scroll">
                    <table>{children}</table>
                  </div>
                </section>
              ),
              blockquote: ({ children }) => <blockquote className="quote-block">{children}</blockquote>,
            }}
          >
            {markdown}
          </ReactMarkdown>
        </article>
      </main>

      <aside className="panel panel-right">
        <section className="outline-card">
          <div className="section-title">
            <ListTree size={16} />
            <h2>文档结构</h2>
          </div>
          <div className="outline-list">
            {stats.sections.length === 0 ? <span className="empty">没有检测到标题</span> : null}
            {stats.sections.map((section) => (
              <a
                key={section.id}
                className={`outline-link depth-${section.depth}`}
                href={`#${section.id}`}
              >
                {section.text}
              </a>
            ))}
          </div>
        </section>

        <section className="outline-card">
          <div className="section-title">
            <TableProperties size={16} />
            <h2>表格导航</h2>
          </div>
          <div className="table-list">
            {stats.tables.length === 0 ? <span className="empty">当前文档没有表格</span> : null}
            {stats.tables.map((table, index) => (
              <div className="table-list-item" key={table.id}>
                <strong>
                  {index + 1}. {table.title}
                </strong>
                <span>
                  {table.headers.length} 列 / {table.rowCount} 行
                </span>
                <span className="muted">{table.headers.slice(0, 3).join(' / ') || '无表头'}</span>
              </div>
            ))}
          </div>
        </section>

        <section className="outline-card">
          <div className="section-title">
            <MonitorSmartphone size={16} />
            <h2>桌面能力</h2>
          </div>
          <div className="table-list">
            <div className="table-list-item">
              <strong>{isDesktopMode ? '当前是桌面模式' : '当前是浏览器模式'}</strong>
              <span>安装版可直接接收系统传入的 .md 文件路径。</span>
            </div>
            <div className="table-list-item">
              <strong>文件关联</strong>
              <span>安装后可在“打开方式”里选择 MD Reader 作为默认 Markdown 阅读器。</span>
            </div>
          </div>
        </section>
      </aside>
    </div>
  )
}

export default App
