# MD Reader Desktop

A desktop Markdown reader focused on comfortable reading instead of heavy editing.

It turns plain `.md` files into a more document-like experience with a Word-style reading layout, stronger table presentation, and direct `.docx` export.

## What It Does

- Open local `.md`, `.markdown`, and `.txt` files
- Read Markdown in a cleaner paper-style layout
- Render tables as stronger visual cards for easier scanning
- Export the current reading view to `.docx`
- Package as a Windows desktop app with Electron
- Register Markdown file associations for Windows installer builds

## Why This Exists

Most Markdown tools are either:

- editor-first, where the source view dominates the reading experience
- export-first, where reading tables and long-form content still feels rough

MD Reader Desktop is aimed at the middle ground:

- keep the original Markdown content intact
- make the reading experience feel closer to Word
- make tables easier to read
- still let you export to a `.docx` file when needed

## Tech Stack

- React 19
- TypeScript
- Vite
- Electron
- electron-builder

## Local Development

```bash
npm install
npm run dev:electron
```

This starts the Vite dev server and the Electron shell together.

## Build Installer

```bash
npm run build:desktop
```

Installer output is written to `release/`.

## Project Structure

- `src/` - React UI and Markdown reader experience
- `electron/` - Electron main process and preload bridge
- `public/` - static assets

## Git Notes

Generated build outputs are ignored:

- `node_modules/`
- `dist/`
- `release/`

## Status

The app currently focuses on:

- reading Markdown comfortably
- improving table visibility
- exporting preview content to Word

It is intentionally not a full Markdown editor.
