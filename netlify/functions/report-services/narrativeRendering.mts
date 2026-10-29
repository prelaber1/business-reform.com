import ExcelJS from 'exceljs'
import { NAVY, LIGHT_BAND } from './constants.mts'
import { type ReportChart, type ReportData } from './types.mts'
import { addWrappedParagraphs, addChartImagesToSheet } from './excelSheetHelpers.mts'

export async function renderNarrativeSummary(
  workbook: ExcelJS.Workbook,
  summary: ExcelJS.Worksheet,
  reportData: ReportData,
  startRow: number,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  let row = startRow
  summary.getCell(`A${row}`).value = 'Executive Summary'
  summary.getCell(`A${row}`).font = { size: 14, bold: true, color: { argb: NAVY } }
  row += 1
  row = addWrappedParagraphs(summary, reportData.executiveSummary, row)
  row += 1

  for (const section of reportData.sections) {
    const headingCell = summary.getCell(`A${row}`)
    headingCell.value = section.heading
    headingCell.font = { size: 13, bold: true, color: { argb: NAVY } }
    headingCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BAND } }
    row += 1

    if (section.body) row = addWrappedParagraphs(summary, section.body, row)

    if (section.bullets?.length) {
      for (const bullet of section.bullets) {
        const cell = summary.getCell(`A${row}`)
        cell.value = `•  ${bullet}`
        cell.alignment = { wrapText: true, vertical: 'top' }
        row += 1
      }
    }
    row += 1
  }

  if (reportData.nextSteps?.length) {
    summary.getCell(`A${row}`).value = 'Next Steps'
    summary.getCell(`A${row}`).font = { size: 13, bold: true, color: { argb: NAVY } }
    row += 1
    reportData.nextSteps.forEach((step, i) => {
      const cell = summary.getCell(`A${row}`)
      cell.value = `${i + 1}. ${step}`
      cell.alignment = { wrapText: true, vertical: 'top' }
      row += 1
    })
    row += 1
  }

  row = addChartImagesToSheet(workbook, summary, reportData.charts, chartPngCache, row)

  return row
}

// A dedicated tab collecting every chart from the report (deterministic
// model charts + any supplementary AI-authored charts) in one place, so the
// final-outcome visuals can be reviewed or presented without hunting through
