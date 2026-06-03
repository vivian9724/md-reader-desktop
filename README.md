# MD Reader Desktop

A desktop Markdown reader built with React, Vite, and Electron. It focuses on a Word-like reading experience, clearer table rendering, and one-click export to `.docx`.

## Features

- Word-style reading layout for Markdown documents
- Enhanced table cards with sticky headers and better scanability
- Desktop app packaging for Windows with Electron
- Open local `.md`, `.markdown`, and `.txt` files
- Export the current reading view to `.docx`
- Windows file association metadata for Markdown files

## Tech Stack

- React 19
- TypeScript
- Vite
- Electron
- electron-builder

## Development

```bash
npm install
npm run dev:electron
```

This starts the Vite dev server and the Electron desktop shell together.

## Production Build

```bash
npm run build:desktop
```

The Windows installer output is written to `release/`.

## Project Structure

- `src/`: React UI
- `electron/`: Electron main/preload processes
- `public/`: static assets

## Notes

- `dist/`, `release/`, and `node_modules/` are ignored in Git.
- The packaged app is intended for local Markdown reading rather than full Markdown editing.
