/// <reference types="vite/client" />

type OpenedMarkdownFile = {
  fileName: string
  filePath: string
  content: string
}

interface Window {
  desktopBridge?: {
    openMarkdownFile: () => Promise<OpenedMarkdownFile | null>
    onMarkdownFileOpened: (callback: (payload: OpenedMarkdownFile) => void) => () => void
  }
}
