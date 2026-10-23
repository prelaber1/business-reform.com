import ExcelJS from 'exceljs'
import { LOGO_PNG_BASE64, LOGO_WIDTH, LOGO_HEIGHT } from '../logo-data.mts'
import { aiMarkLine, applyAiMarkToWorkbook } from '../ai-disclosure.mts'
import { NAVY, LIGHT_BAND } from './constants.mts'
import { type ReportTable, type ReportChart, type ReportAssumptions, type ReportData } from './types.mts'
import { RESULT_FILL, CAVEAT_FILL, SUBTOTAL_FILL, THIN_BORDER, type MetricFormat, fx, type MetricValue, type RawDataSheetModel, type CalcSheetModel, type KeyMetric, type ReportModelBase } from './calculationSupport.mts'
import { sanitizeSheetName, addChartImagesToSheet, workbookToArrayBuffer } from './excelSheetHelpers.mts'
import { renderNarrativeSummary } from './narrativeRendering.mts'

// A dedicated tab collecting every chart from the report (deterministic
// model charts + any supplementary AI-authored charts) in one place, so the
// final-outcome visuals can be reviewed or presented without hunting through
// the narrative Summary sheet.
export function buildChartsSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, charts: ReportChart[], chartPngCache: Map<ReportChart, Buffer>) {
  if (!charts.length) return
  const sheet = workbook.addWorksheet(sanitizeSheetName('Charts', usedSheetNames))
  sheet.getColumn(1).width = 14

  sheet.getCell('A1').value = 'Charts'
  sheet.getCell('A1').font = { size: 16, bold: true, color: { argb: NAVY } }

  let row = 3
  for (const chart of charts) {
    if (!chartPngCache.has(chart)) continue
    sheet.getCell(`A${row}`).value = chart.title
    sheet.getCell(`A${row}`).font = { size: 12, bold: true, color: { argb: NAVY } }
    row += 1
    row = addChartImagesToSheet(workbook, sheet, [chart], chartPngCache, row)
    row += 2
  }
}

// One sheet per supplementary AI-authored table — the escape hatch for
// qualitative/non-deterministic content that doesn't fit the line-item model.
export function renderTableSheets(workbook: ExcelJS.Workbook, tables: ReportTable[], usedSheetNames: Set<string>) {
  for (const table of tables) {
    if (!table.columns?.length) continue
    if (!Array.isArray(table.rows)) table.rows = []
    const sheet = workbook.addWorksheet(sanitizeSheetName(table.title || 'Data', usedSheetNames))

    sheet.addRow(table.columns)
    sheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    })
    sheet.views = [{ state: 'frozen', ySplit: 1 }]
    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: table.columns.length } }

    for (const dataRow of table.rows) sheet.addRow(dataRow)

    table.columns.forEach((col, ci) => {
      const headerLower = col.toLowerCase()
      const maxContentLength = Math.max(col.length, ...table.rows.map((r) => String(r[ci] ?? '').length), 1)
      const column = sheet.getColumn(ci + 1)
      column.width = Math.min(40, Math.max(12, maxContentLength + 3))

      const isNumericColumn = table.rows.length > 0 && table.rows.every((r) => typeof r[ci] === 'number')
      if (isNumericColumn) {
        if (/%|percent|margin|rate|ratio/.test(headerLower)) {
          column.numFmt = '0.0%'
        } else if (/€|\$|revenue|cost|profit|cash|amount|ebitda|price|budget|expense|income|funding|burn/.test(headerLower)) {
          column.numFmt = '#,##0.00'
        } else {
          column.numFmt = '#,##0'
        }

        if (table.rows.length > 1) {
          const colLetter = column.letter
          sheet.addConditionalFormatting({
            ref: `${colLetter}2:${colLetter}${table.rows.length + 1}`,
            rules: [
              {
                type: 'dataBar',
                priority: 1,
                gradient: true,
                cfvo: [{ type: 'min' }, { type: 'max' }],
              },
            ],
          })
        }
      }
    })

    for (let r = 2; r <= table.rows.length + 1; r += 2) {
      sheet.getRow(r).eachCell({ includeEmpty: true }, (cell) => {
        if (!cell.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BAND } }
      })
    }
  }
}

// ── Branding: cover page + a discrete logo mark on every other sheet ───────

export const LOGO_ASPECT = LOGO_HEIGHT / LOGO_WIDTH

// A dedicated title/cover sheet, always added first. Carries the logo at a
// modest, tasteful size — the "prominent" cover treatment is still restrained
// (this platform's reports are meant to look like they belong on a CFO's
// desk, not a pitch deck), consistent with the small corner mark used on
// every other sheet.
export function buildCoverSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, logoImageId: number, reportLabel: string, sourceFileName: string) {
  const sheet = workbook.addWorksheet(sanitizeSheetName('Cover', usedSheetNames))
  sheet.getColumn(1).width = 100

  const logoWidth = 200
  const logoHeight = Math.round(logoWidth * LOGO_ASPECT)
  sheet.addImage(logoImageId, { tl: { col: 0.15, row: 1.3 }, ext: { width: logoWidth, height: logoHeight } })

  const titleRow = 9
  sheet.getCell(`A${titleRow}`).value = reportLabel
  sheet.getCell(`A${titleRow}`).font = { size: 22, bold: true, color: { argb: NAVY } }

  sheet.getCell(`A${titleRow + 2}`).value = `Source file: ${sourceFileName}`
  sheet.getCell(`A${titleRow + 2}`).font = { size: 12, color: { argb: 'FF555555' } }

  sheet.getCell(`A${titleRow + 3}`).value = `Generated: ${new Date().toLocaleString()}`
  sheet.getCell(`A${titleRow + 3}`).font = { size: 12, color: { argb: 'FF555555' } }

  sheet.getCell(`A${titleRow + 5}`).value = 'Prepared by the Business Reform AI CFO platform.'
  sheet.getCell(`A${titleRow + 5}`).font = { size: 11, italic: true, color: { argb: 'FF8A98A8' } }

  sheet.getCell(`A${titleRow + 7}`).value = aiMarkLine('en')
  sheet.getCell(`A${titleRow + 7}`).font = { size: 10, bold: true, color: { argb: NAVY } }
}

// Stamps a small logo mark in the top-right corner of every sheet in the
// workbook except the Cover sheet (which already carries the logo at a more
// prominent size) — run once, after all sheets exist, so it doesn't need to
// be threaded through every individual sheet-builder function.
export function addDiscreteLogoToAllSheets(workbook: ExcelJS.Workbook, logoImageId: number) {
  const logoWidth = 60
  const logoHeight = Math.round(logoWidth * LOGO_ASPECT)
  for (const sheet of workbook.worksheets) {
    if (sheet.name === 'Cover') continue
    const anchorCol = Math.max(sheet.columnCount || 1, 1)
    sheet.addImage(logoImageId, { tl: { col: anchorCol, row: 0.1 }, ext: { width: logoWidth, height: logoHeight } })
  }
}

export async function buildExcelWorkbook(reportData: ReportData, reportLabel: string, sourceFileName: string, chartPngCache: Map<ReportChart, Buffer>) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Business Reform AI CFO'
  workbook.created = new Date()

  const usedSheetNames = new Set<string>()
  const logoImageId = workbook.addImage({ base64: LOGO_PNG_BASE64, extension: 'png' })
  buildCoverSheet(workbook, usedSheetNames, logoImageId, reportLabel, sourceFileName)

  // --- Summary sheet: narrative + embedded chart images ---
  const summary = workbook.addWorksheet(sanitizeSheetName('Summary', usedSheetNames))
  summary.getColumn(1).width = 100

  summary.getCell('A1').value = reportLabel
  summary.getCell('A1').font = { size: 18, bold: true, color: { argb: NAVY } }

  summary.getCell('A2').value = `Source file: ${sourceFileName}  |  Generated: ${new Date().toLocaleString()}`
  summary.getCell('A2').font = { size: 10, italic: true, color: { argb: 'FF555555' } }

  await renderNarrativeSummary(workbook, summary, reportData, 4, chartPngCache)
  buildChartsSheet(workbook, usedSheetNames, reportData.charts, chartPngCache)

  // --- One sheet per data table, properly formatted ---
  renderTableSheets(workbook, reportData.tables, usedSheetNames)

  addDiscreteLogoToAllSheets(workbook, logoImageId)
  applyAiMarkToWorkbook(workbook, 'en')

  const buffer = await workbook.xlsx.writeBuffer()
  return workbookToArrayBuffer(buffer)
}

// ── Deterministic Excel workbook rendering ──────────────────────────────────
//
// Renders a ReportModelBase (raw-data sheets + calculations sheets, all
// formula-driven) plus the narrative/assumptions content, into a workbook
// whose sheet order is always: Summary → raw-data sheet(s) → calculations
// sheet(s) → any supplementary AI tables/charts (the escape hatch) →
// Assumptions & Notes last.

export function numFmtFor(format?: MetricFormat) {
  if (format === 'percent') return '0.00%'
  if (format === 'ratio') return '0.00"x"'
  if (format === 'number') return '#,##0.0'
  return '#,##0.00" €"'
}

export function setMetricCell(cell: ExcelJS.Cell, value: MetricValue, format?: MetricFormat) {
  if (typeof value === 'string') {
    cell.value = value
    cell.font = { italic: true, color: { argb: 'FF884400' } }
    return
  }
  if (typeof value === 'number') {
    cell.value = value
    cell.numFmt = numFmtFor(format)
    return
  }
  cell.value = { formula: value.formula.replace(/^=/, ''), result: value.result }
  cell.numFmt = numFmtFor(format)
}

export function buildKeyMetricsTable(summary: ExcelJS.Worksheet, startRow: number, keyMetrics: KeyMetric[]) {
  let row = startRow
  summary.getCell(`A${row}`).value = 'Key Metrics'
  summary.getCell(`A${row}`).font = { size: 13, bold: true, color: { argb: NAVY } }
  row += 1

  const headerRow = row
  summary.getCell(`A${headerRow}`).value = 'Metric'
  summary.getCell(`B${headerRow}`).value = 'Value'
  ;['A', 'B'].forEach((col) => {
    const cell = summary.getCell(`${col}${headerRow}`)
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.border = THIN_BORDER
    cell.alignment = { horizontal: 'center' }
  })
  row += 1

  for (const metric of keyMetrics) {
    const labelCell = summary.getCell(`A${row}`)
    labelCell.value = metric.label
    labelCell.border = THIN_BORDER
    const valueCell = summary.getCell(`B${row}`)
    setMetricCell(valueCell, fx(metric.cellRef, metric.value), metric.format)
    valueCell.border = THIN_BORDER
    row += 1
  }
  return row + 1
}

export function buildCaveatCallout(summary: ExcelJS.Worksheet, startRow: number, assumptions: ReportAssumptions) {
  const caveat = assumptions.dataCaveats[0] || assumptions.limitations[0]
  if (!caveat) return startRow
  const row = startRow
  summary.mergeCells(`A${row}:B${row + 2}`)
  const cell = summary.getCell(`A${row}`)
  cell.value = `Caveat: ${caveat} See the Assumptions & Notes sheet for full detail.`
  cell.font = { italic: true, size: 10, color: { argb: 'FF663C00' } }
  cell.alignment = { wrapText: true, vertical: 'top' }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CAVEAT_FILL } }
  return row + 4
}

export async function buildDeterministicSummarySheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  reportData: ReportData,
  model: ReportModelBase,
  reportLabel: string,
  sourceFileName: string,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  const summary = workbook.addWorksheet(sanitizeSheetName('Summary', usedSheetNames))
  summary.getColumn(1).width = 55
  summary.getColumn(2).width = 45

  summary.getCell('A1').value = reportLabel
  summary.getCell('A1').font = { size: 18, bold: true, color: { argb: NAVY } }
  summary.mergeCells('A1:B1')

  summary.getCell('A2').value = `Source file: ${sourceFileName}  |  Generated: ${new Date().toLocaleString()}`
  summary.getCell('A2').font = { size: 10, italic: true, color: { argb: 'FF555555' } }
  summary.mergeCells('A2:B2')

  let row = 4
  if (model.keyMetrics.length) row = buildKeyMetricsTable(summary, row, model.keyMetrics)
  row = buildCaveatCallout(summary, row, reportData.assumptions)
  await renderNarrativeSummary(workbook, summary, { ...reportData, charts: [...model.charts, ...reportData.charts] }, row, chartPngCache)
}

export function buildRawDataSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, model: RawDataSheetModel) {
  const sheet = workbook.addWorksheet(sanitizeSheetName(model.sheetName, usedSheetNames))
  sheet.getColumn(1).width = 12
  sheet.getColumn(2).width = 42
  sheet.getColumn(3).width = 18
  sheet.getColumn(4).width = 18

  sheet.getCell('A1').value = model.title
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:D1')

  const headerRow = 3
  ;['Account', 'Description', 'Group', 'Amount'].forEach((label, i) => {
    const cell = sheet.getCell(headerRow, i + 1)
    cell.value = label
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.alignment = { horizontal: 'center' }
    cell.border = THIN_BORDER
  })
  sheet.views = [{ state: 'frozen', ySplit: headerRow }]

  for (const dataRow of model.rows) {
    if (dataRow.kind === 'item') {
      sheet.getCell(`A${dataRow.row}`).value = dataRow.account || ''
      sheet.getCell(`B${dataRow.row}`).value = dataRow.description
      sheet.getCell(`C${dataRow.row}`).value = dataRow.groupLabel
      const amountCell = sheet.getCell(`D${dataRow.row}`)
      amountCell.value = dataRow.amount
      amountCell.numFmt = '#,##0.00" €"'
      ;['A', 'B', 'C', 'D'].forEach((col) => (sheet.getCell(`${col}${dataRow.row}`).border = THIN_BORDER))
    } else {
      sheet.mergeCells(`A${dataRow.row}:C${dataRow.row}`)
      const labelCell = sheet.getCell(`A${dataRow.row}`)
      labelCell.value = dataRow.label
      labelCell.font = { bold: true }
      const valueCell = sheet.getCell(`D${dataRow.row}`)
      setMetricCell(valueCell, dataRow.value)
      const fillColor = dataRow.kind === 'total' ? RESULT_FILL : SUBTOTAL_FILL
      ;['A', 'B', 'C', 'D'].forEach((col) => {
        const cell = sheet.getCell(`${col}${dataRow.row}`)
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } }
        cell.border = THIN_BORDER
      })
    }
  }
}

export function buildCalculationsSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, model: CalcSheetModel) {
  const sheet = workbook.addWorksheet(sanitizeSheetName(model.sheetName, usedSheetNames))
  sheet.getColumn(1).width = 4
  sheet.getColumn(2).width = 48
  sheet.getColumn(3).width = 22

  sheet.getCell('A1').value = model.title
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:C1')

  for (const dataRow of model.rows) {
    if (dataRow.kind === 'sectionHeading') {
      sheet.mergeCells(`B${dataRow.row}:C${dataRow.row}`)
      const cell = sheet.getCell(`B${dataRow.row}`)
      cell.value = dataRow.label
      cell.font = { size: 12, bold: true, color: { argb: NAVY } }
    } else if (dataRow.kind === 'columnHeader') {
      ;['B', 'C'].forEach((col, i) => {
        const cell = sheet.getCell(`${col}${dataRow.row}`)
        cell.value = i === 0 ? 'Metric' : 'Value'
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
        cell.alignment = { horizontal: 'center' }
        cell.border = THIN_BORDER
      })
    } else {
      const labelCell = sheet.getCell(`B${dataRow.row}`)
      labelCell.value = dataRow.label
      labelCell.font = { bold: !!dataRow.bold }
      labelCell.border = THIN_BORDER
      const valueCell = sheet.getCell(`C${dataRow.row}`)
      setMetricCell(valueCell, dataRow.value, dataRow.format)
      if (dataRow.bold) valueCell.font = { ...valueCell.font, bold: true }
      valueCell.border = THIN_BORDER
      if (dataRow.highlight) {
        ;['B', 'C'].forEach((col) => {
          sheet.getCell(`${col}${dataRow.row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RESULT_FILL } }
        })
      }
    }
  }
}

export function buildAssumptionsSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, assumptions: ReportAssumptions, extraCaveats: string[]) {
  const sheet = workbook.addWorksheet(sanitizeSheetName('Assumptions & Notes', usedSheetNames))
  sheet.getColumn(1).width = 100

  sheet.getCell('A1').value = 'Assumptions, Methodology & Limitations'
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }

  let row = 3
  const sections: { heading: string; items: string[] }[] = [
    { heading: 'Data Caveats', items: assumptions.dataCaveats || [] },
    { heading: 'Excluded Items', items: assumptions.excludedItems || [] },
    { heading: 'Classification Rationale', items: assumptions.classificationRationale || [] },
    { heading: 'Limitations', items: [...(assumptions.limitations || []), ...extraCaveats] },
  ]

  for (const section of sections) {
    if (!section.items.length) continue
    const headingCell = sheet.getCell(`A${row}`)
    headingCell.value = section.heading
    headingCell.font = { size: 12, bold: true, color: { argb: NAVY } }
    row += 1
    for (const item of section.items) {
      const cell = sheet.getCell(`A${row}`)
      cell.value = `•  ${item}`
      cell.alignment = { wrapText: true, vertical: 'top' }
      sheet.getRow(row).height = Math.max(20, Math.ceil(item.length / 105) * 15)
      row += 1
    }
    row += 1
  }
}

export async function buildDeterministicExcelWorkbook(
  reportData: ReportData,
  model: ReportModelBase,
  reportLabel: string,
  sourceFileName: string,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Business Reform AI CFO'
  workbook.created = new Date()
  workbook.calcProperties.fullCalcOnLoad = true

  const usedSheetNames = new Set<string>()
  const allCharts = [...model.charts, ...reportData.charts]
  const logoImageId = workbook.addImage({ base64: LOGO_PNG_BASE64, extension: 'png' })
  buildCoverSheet(workbook, usedSheetNames, logoImageId, reportLabel, sourceFileName)

  await buildDeterministicSummarySheet(workbook, usedSheetNames, reportData, model, reportLabel, sourceFileName, chartPngCache)
  for (const rawSheet of model.rawSheets) buildRawDataSheet(workbook, usedSheetNames, rawSheet)
  for (const calcSheet of model.calcSheets) buildCalculationsSheet(workbook, usedSheetNames, calcSheet)
  for (const buildCustomSheet of model.customSheetBuilders || []) buildCustomSheet(workbook, usedSheetNames)
  buildChartsSheet(workbook, usedSheetNames, allCharts, chartPngCache)
  renderTableSheets(workbook, reportData.tables, usedSheetNames)
  buildAssumptionsSheet(workbook, usedSheetNames, reportData.assumptions, model.extraCaveats)

  addDiscreteLogoToAllSheets(workbook, logoImageId)
  applyAiMarkToWorkbook(workbook, 'en')

  const buffer = await workbook.xlsx.writeBuffer()
  return workbookToArrayBuffer(buffer)
}

// ── PowerPoint deck rendering ────────────────────────────────────────────────
//
// A slide-form companion to the Excel workbook: title, executive summary, key
// metrics table, one slide per chart (reusing the same cached PNGs as the
// Excel Charts sheet), and next steps — generated from the same ReportData /
// ReportModelBase the Excel workbook is built from, so the two artifacts
// never disagree.
