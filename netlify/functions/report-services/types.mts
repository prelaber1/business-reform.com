export type JobStatus = 'processing' | 'complete' | 'failed'

export type ReportJob = {
  id: string
  userId: string
  fileKey: string
  fileName: string
  reportType: string
  reportLabel: string
  status: JobStatus
  createdAt: string
  updatedAt: string
  workerToken: string
  downloadKey?: string
  downloadKeyPptx?: string
  error?: string
}

export type ReportSection = { heading: string; body?: string; bullets?: string[] }
export type ReportTable = { title: string; columns: string[]; rows: (string | number)[][] }
export type ReportChartSeries = { name: string; values: number[] }
export type ReportChart = { type: 'line' | 'bar'; title: string; categories: string[]; series: ReportChartSeries[] }

// A single atomic figure extracted from the source document, classified with a
// semantic `group` tag (vocabulary varies per report type — see REPORT_DEFINITIONS).
// The AI supplies these raw and unaggregated; all arithmetic on top of them is
// performed deterministically in code (see "Deterministic calculation engine" below)
// and rendered as real cross-referencing Excel formulas, never as AI-computed numbers.
export type LineItem = { label: string; account?: string; group: string; amount: number; period?: string }
export type ReportAssumptions = {
  dataCaveats: string[]
  excludedItems: string[]
  classificationRationale: string[]
  limitations: string[]
}
export type ReportData = {
  executiveSummary: string
  sections: ReportSection[]
  lineItems: LineItem[]
  assumptions: ReportAssumptions
  tables: ReportTable[]
  charts: ReportChart[]
  nextSteps: string[]
}
