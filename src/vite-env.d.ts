/// <reference types="vite/client" />

type OpenedMarkdownFile = {
  fileName: string
  filePath: string | null
  content: string
}

type SaveMarkdownPayload = {
  filePath: string | null
  defaultName: string
  content: string
}

interface Window {
  desktopBridge?: {
    openMarkdownFile: () => Promise<OpenedMarkdownFile | null>
    readMarkdownFile: (filePath: string) => Promise<OpenedMarkdownFile | null>
    saveMarkdownFile: (payload: SaveMarkdownPayload) => Promise<OpenedMarkdownFile | null>
    onMarkdownFileOpened: (callback: (payload: OpenedMarkdownFile) => void) => () => void
  }
}
