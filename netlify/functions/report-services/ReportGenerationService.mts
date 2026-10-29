import { renderChartPngCache } from './chartRendering.mts'
import { buildExcelWorkbook, buildDeterministicExcelWorkbook } from './ExcelReportBuilder.mts'
import { buildPowerPointPresentation } from './PowerPointReportBuilder.mts'
import type { ReportData } from './types.mts'
import type { ReportModelBase } from './calculationSupport.mts'

export type GeneratedReportFiles = {
  xlsx: ArrayBuffer
  // The PowerPoint deck is a companion artifact: if it fails to generate, the
  // Excel report (the primary deliverable) still completes, just without it.
  pptx?: ArrayBuffer
}

// Assembles the final deliverables from verified data: renders every chart
// once (shared by the workbook and the deck), then builds the Excel workbook
// with its tables, charts, assumptions and disclaimers, and the PowerPoint deck.
export async function generateReportFiles(
  reportData: ReportData,
  model: ReportModelBase | null,
  reportLabel: string,
  sourceFileName: string,
): Promise<GeneratedReportFiles> {
  const allCharts = [...(model?.charts ?? []), ...reportData.charts]
  const chartPngCache = await renderChartPngCache(allCharts)
  const xlsx = model
    ? await buildDeterministicExcelWorkbook(reportData, model, reportLabel, sourceFileName, chartPngCache)
    : await buildExcelWorkbook(reportData, reportLabel, sourceFileName, chartPngCache)

  let pptx: ArrayBuffer | undefined
  try {
    pptx = await buildPowerPointPresentation(reportData, model, reportLabel, sourceFileName, chartPngCache)
  } catch (error) {
    console.error('PowerPoint generation failed, continuing with Excel-only report', error)
  }

  return { xlsx, pptx }
}
