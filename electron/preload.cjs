const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('desktopBridge', {
  openMarkdownFile: () => ipcRenderer.invoke('dialog:open-markdown-file'),
  readMarkdownFile: (filePath) => ipcRenderer.invoke('file:read-markdown', filePath),
  saveMarkdownFile: (payload) => ipcRenderer.invoke('file:save-markdown', payload),
  onMarkdownFileOpened: (callback) => {
    const handler = (_event, payload) => callback(payload)
    ipcRenderer.on('markdown-file-opened', handler)
    return () => {
      ipcRenderer.removeListener('markdown-file-opened', handler)
    }
  },
})
