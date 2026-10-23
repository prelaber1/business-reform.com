import ExcelJS from 'exceljs'
import { type ReportChart } from './types.mts'

export function addWrappedParagraphs(sheet: ExcelJS.Worksheet, text: string, startRow: number) {
  let row = startRow
  for (const para of text.split('\n').map((p) => p.trim()).filter(Boolean)) {
    const cell = sheet.getCell(`A${row}`)
    cell.value = para
    cell.alignment = { wrapText: true, vertical: 'top' }
    sheet.getRow(row).height = Math.max(20, Math.ceil(para.length / 105) * 15)
    row += 1
  }
  return row
}

export function sanitizeSheetName(name: string, used: Set<string>) {
  let base = name.replace(/[\\/*?:[\]]/g, ' ').trim().slice(0, 31) || 'Sheet'
  let candidate = base
  let suffix = 1
  while (used.has(candidate)) {
    suffix += 1
    candidate = `${base.slice(0, 28)} ${suffix}`
  }
  used.add(candidate)
  return candidate
}

// Renders executive summary / sections / next steps / charts into an existing
// worksheet starting at `startRow`. Shared by both the legacy fully-AI-driven
// workbook and the new deterministic workbook's Summary sheet.
// Draws each chart's title + cached PNG image into `sheet`, stacked
// vertically starting at `startRow`. Shared by the narrative Summary sheet
// (contextual, interleaved with prose) and the dedicated Charts sheet
// (all charts collected together).
export function addChartImagesToSheet(
  workbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  charts: ReportChart[],
  chartPngCache: Map<ReportChart, Buffer>,
  startRow: number,
) {
  let row = startRow
  for (const chart of charts) {
    const pngBuffer = chartPngCache.get(chart)
    if (!pngBuffer) continue
    const imageId = workbook.addImage({ buffer: pngBuffer, extension: 'png' })
    sheet.addImage(imageId, { tl: { col: 0, row }, ext: { width: 640, height: 299 } })
    row += 17 // reserve rows so later content doesn't overlap the floating image
  }
  return row
}

// @netlify/blobs' set() only accepts string | ArrayBuffer | Blob — a Node
// Buffer is a Uint8Array view, not itself an ArrayBuffer, so slice it out.
export function bufferToArrayBuffer(buffer: Buffer): ArrayBuffer {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer
}

export function workbookToArrayBuffer(buffer: ExcelJS.Buffer): ArrayBuffer {
  return bufferToArrayBuffer(buffer as unknown as Buffer)
}
