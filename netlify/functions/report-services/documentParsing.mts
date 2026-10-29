import * as XLSX from 'xlsx'
import mammoth from 'mammoth'
import { MAX_TEXT_CHARS } from './constants.mts'

export function getMimeType(filename: string, fallback?: string) {
  if (fallback) return fallback
  const ext = (filename.split('.').pop() || '').toLowerCase()
  const map: Record<string, string> = {
    pdf: 'application/pdf',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
    csv: 'text/plain',
    txt: 'text/plain',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }
  return map[ext] || 'application/octet-stream'
}

export function isNativeAnthropicType(mimeType: string) {
  return ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp'].includes(mimeType)
}

export function isSpreadsheetType(mimeType: string) {
  return [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ].includes(mimeType)
}

export function isWordType(mimeType: string) {
  return [
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ].includes(mimeType)
}

export function parseSpreadsheetToText(fileBuffer: Buffer) {
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' })
  const sheets = workbook.SheetNames.map((sheetName) => {
    const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName])
    return `--- Sheet: ${sheetName} ---\n${csv}`
  })
  return sheets.join('\n\n').slice(0, MAX_TEXT_CHARS)
}

export async function parseDocxToText(fileBuffer: Buffer) {
  const result = await mammoth.extractRawText({ buffer: fileBuffer })
  return result.value.slice(0, MAX_TEXT_CHARS)
}
