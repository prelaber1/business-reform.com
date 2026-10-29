import ExcelJS from 'exceljs'
import { type ReportChart, type LineItem } from './types.mts'

export const RESULT_FILL = 'FFC6E0B4'
export const CAVEAT_FILL = 'FFFCE4D6'
export const SUBTOTAL_FILL = 'FFDDEBF7'
export const BORDER_COLOR = 'FFB7B7B7'
export const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: BORDER_COLOR } },
  left: { style: 'thin', color: { argb: BORDER_COLOR } },
  bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
  right: { style: 'thin', color: { argb: BORDER_COLOR } },
}

export type Formula = { formula: string; result: number }
// 'ratio' = a plain multiple (e.g. current ratio 1.85), shown as "1.85x";
// 'number' = a plain count such as months of runway.
export type MetricFormat = 'currency' | 'percent' | 'ratio' | 'number'
export function fx(formula: string, result: number): Formula {
  return { formula, result }
}
// A metric cell is a live formula, a plain editable numeric input (e.g. a
// target ratio or growth-rate assumption the user may want to tweak), or a
// plain-text placeholder (e.g. "N/A (no revenue data)") when the underlying
// figure genuinely can't be computed — never a broken formula / #DIV0! left
// for the user to puzzle over.
export type MetricValue = Formula | string | number

export function sumGroup(items: LineItem[]): number {
  return items.reduce((total, item) => total + (Number.isFinite(item.amount) ? item.amount : 0), 0)
}

export function groupItems(lineItems: LineItem[], tag: string): LineItem[] {
  return lineItems.filter((item) => item.group === tag && Number.isFinite(item.amount))
}

export function refCell(sheetName: string, row: number, col: string): string {
  return `'${sheetName}'!${col}${row}`
}

// ---- Raw-data sheet model: account-level line items grouped into contiguous
// blocks (one per classification tag), each followed by a SUM() subtotal row.
// Fixed column layout: A=Account, B=Description, C=Group, D=Amount.

export type RawDataItemRow = { row: number; kind: 'item'; account?: string; description: string; groupLabel: string; amount: number }
export type RawDataTotalRow = { row: number; kind: 'subtotal' | 'total'; label: string; value: MetricValue }
export type RawDataRow = RawDataItemRow | RawDataTotalRow
export type RawDataBlock = { tag: string; label: string; subtotalRow: number; subtotalValue: number }
export type RawDataSheetModel = { sheetName: string; title: string; rows: RawDataRow[]; blocks: RawDataBlock[]; lastRow: number }

export const RAW_DATA_FIRST_DATA_ROW = 4 // row 1 = title, row 2 = blank, row 3 = column headers

export function layoutRawDataSheet(
  sheetName: string,
  title: string,
  lineItems: LineItem[],
  groupOrder: { tag: string; label: string }[],
): RawDataSheetModel {
  const rows: RawDataRow[] = []
  const blocks: RawDataBlock[] = []
  let row = RAW_DATA_FIRST_DATA_ROW

  for (const { tag, label } of groupOrder) {
    const items = groupItems(lineItems, tag)
    if (items.length === 0) continue
    const blockStart = row
    for (const item of items) {
      rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: label, amount: item.amount })
      row += 1
    }
    const blockEnd = row - 1
    const subtotalValue = sumGroup(items)
    rows.push({ row, kind: 'subtotal', label: `Total ${label}`, value: fx(`=SUM(D${blockStart}:D${blockEnd})`, subtotalValue) })
    blocks.push({ tag, label, subtotalRow: row, subtotalValue })
    row += 1
  }

  return { sheetName, title, rows, blocks, lastRow: row - 1 }
}

export function findBlock(model: RawDataSheetModel, tag: string) {
  return model.blocks.find((b) => b.tag === tag)
}

// Appends a derived row (e.g. "Total Operating Expenses", "Net Result") below
// the last block/row on a raw-data sheet, referencing other rows on that same
// sheet. Returns the row number so calculators can reference it in turn.
export function addTotalRow(model: RawDataSheetModel, label: string, value: MetricValue): number {
  const row = model.lastRow + 1
  model.rows.push({ row, kind: 'total', label, value })
  model.lastRow = row
  return row
}

// ---- Calculations sheet model: label/value metric rows grouped into
// optionally-headed sections (e.g. "Scenario A", "Scenario B"). Fixed column
// layout: A=spacer, B=label, C=value.

export type CalcRow =
  | { row: number; kind: 'sectionHeading'; label: string }
  | { row: number; kind: 'columnHeader' }
  | { row: number; kind: 'metric'; label: string; value: MetricValue; bold?: boolean; highlight?: boolean; format?: MetricFormat }

export type CalcSheetModel = { sheetName: string; title: string; rows: CalcRow[]; lastRow: number }

export function newCalcSheet(sheetName: string, title: string): CalcSheetModel {
  return { sheetName, title, rows: [], lastRow: 2 }
}

export function addSection(model: CalcSheetModel, heading: string) {
  model.lastRow += 1
  model.rows.push({ row: model.lastRow, kind: 'sectionHeading', label: heading })
  model.lastRow += 1
  model.rows.push({ row: model.lastRow, kind: 'columnHeader' })
}

export function addMetric(
  model: CalcSheetModel,
  label: string,
  value: MetricValue,
  opts: { bold?: boolean; highlight?: boolean; format?: MetricFormat } = {},
): number {
  model.lastRow += 1
  model.rows.push({ row: model.lastRow, kind: 'metric', label, value, format: 'currency', ...opts })
  return model.lastRow
}

export type KeyMetric = { label: string; cellRef: string; value: number; format: MetricFormat }
export type ReportModelBase = {
  rawSheets: RawDataSheetModel[]
  calcSheets: CalcSheetModel[]
  keyMetrics: KeyMetric[]
  charts: ReportChart[]
  extraCaveats: string[]
  // Escape hatch for calculators whose output doesn't fit the generic
  // account-list (RawDataSheetModel) / label-value (CalcSheetModel) shapes —
  // e.g. a multi-year, multi-column projection grid. Each builder is invoked
  // once the real workbook/usedSheetNames exist, after the raw/calc sheets.
  customSheetBuilders?: ((workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) => void)[]
}

// ── Platform-controlled financial assumptions ───────────────────────────────
// These are NOT user-facing and NOT AI-supplied — they're fixed inputs the
// analysis engine applies deterministically to every report of the relevant
// type, regardless of what's in the uploaded file. Update the values here
// (and redeploy) whenever the macro outlook or benchmark data changes; the
// frontend, the AI prompt, and the client experience are completely
// unaffected — the only visible difference is the numbers in the report.
export const PLATFORM_ASSUMPTIONS = {
  // Applied in the Financial Projections calculator.
  annualInflationRate: 0.025, // 2.5%/year — escalates operating expenses across all scenarios
  baseRevenueCagr: 0.08, // 8%/year — base-case revenue growth
  optimisticCagrOffset: 0.05, // optimistic scenario = base + 5 percentage points
  pessimisticCagrOffset: 0.05, // pessimistic scenario = base - 5 percentage points
  // Applied in the Cash Flow Forecast calculator to convert a run-rate from
  // one granularity to the other when the source data only supports one.
  weeksPerMonth: 4.33, // 52 weeks / 12 months
  // Applied in the FP&A Digest calculator: a metric whose current period moved
  // this much (as a fraction) against trailing average or last year is flagged.
  digestFlagThreshold: 0.1,
}
