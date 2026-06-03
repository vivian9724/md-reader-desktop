import { mkdir, writeFile } from 'node:fs/promises'
import path from 'node:path'
import sharp from 'sharp'
import pngToIco from 'png-to-ico'

const root = process.cwd()
const buildDir = path.join(root, 'build')
const publicDir = path.join(root, 'public')
const svgPath = path.join(buildDir, 'icon.svg')

await mkdir(buildDir, { recursive: true })

const iconSizes = [16, 32, 48, 64, 128, 256]
const pngBuffers = []

for (const size of iconSizes) {
  const outputPath = path.join(buildDir, `icon-${size}.png`)
  const buffer = await sharp(svgPath).resize(size, size).png().toBuffer()
  pngBuffers.push(buffer)
  await writeFile(outputPath, buffer)
}

const icoBuffer = await pngToIco(pngBuffers)
await writeFile(path.join(buildDir, 'icon.ico'), icoBuffer)

const appIconBuffer = await sharp(svgPath).resize(512, 512).png().toBuffer()
await writeFile(path.join(publicDir, 'app-icon.png'), appIconBuffer)
