const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const path = require('node:path')
const fs = require('node:fs/promises')

let mainWindow = null
let pendingFilePath = extractMarkdownPath(process.argv)

function extractMarkdownPath(argv) {
  const candidates = argv.slice(1)
  return (
    candidates.find((value) => typeof value === 'string' && /\.(md|markdown|txt)$/i.test(value)) ?? null
  )
}

async function readMarkdownFile(filePath) {
  const content = await fs.readFile(filePath, 'utf8')
  return {
    fileName: path.basename(filePath),
    filePath,
    content,
  }
}

async function deliverFileToRenderer(filePath) {
  if (!mainWindow || !filePath) {
    return
  }

  const payload = await readMarkdownFile(filePath)
  pendingFilePath = null
  mainWindow.webContents.send('markdown-file-opened', payload)
}

function showOpenError(error) {
  const message = error instanceof Error ? error.message : 'Unknown error'
  dialog.showErrorBox('无法打开 Markdown 文件', message)
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 960,
    minWidth: 1080,
    minHeight: 720,
    title: 'MD Reader',
    backgroundColor: '#f4efe6',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  const devServerUrl = process.env.VITE_DEV_SERVER_URL
  if (devServerUrl) {
    mainWindow.loadURL(devServerUrl)
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'))
  }

  mainWindow.webContents.on('did-finish-load', () => {
    if (pendingFilePath) {
      deliverFileToRenderer(pendingFilePath).catch(showOpenError)
    }
  })

  mainWindow.on('closed', () => {
    mainWindow = null
  })
}

app.whenReady().then(() => {
  const gotLock = app.requestSingleInstanceLock()
  if (!gotLock) {
    app.quit()
    return
  }

  app.on('second-instance', (_event, argv) => {
    const filePath = extractMarkdownPath(argv)
    if (mainWindow) {
      if (mainWindow.isMinimized()) {
        mainWindow.restore()
      }
      mainWindow.focus()
    }
    if (filePath) {
      pendingFilePath = filePath
      deliverFileToRenderer(filePath).catch(showOpenError)
    }
  })

  ipcMain.handle('dialog:open-markdown-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: '打开 Markdown 文件',
      properties: ['openFile'],
      filters: [
        { name: 'Markdown Files', extensions: ['md', 'markdown', 'txt'] },
        { name: 'All Files', extensions: ['*'] },
      ],
    })

    if (result.canceled || result.filePaths.length === 0) {
      return null
    }

    return readMarkdownFile(result.filePaths[0])
  })

  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
