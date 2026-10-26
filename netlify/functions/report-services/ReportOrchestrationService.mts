import { getStore } from '@netlify/blobs'
import { REPORTS_STORE, UPLOADS_STORE } from './constants.mts'
import { REPORT_DEFINITIONS } from './reportDefinitions.mts'
import { getMimeType } from './documentParsing.mts'
import { extractDocument } from './DocumentExtractionService.mts'
import { assertNoBlockingIssues, validateFinancialInputs } from './FinancialValidationService.mts'
import { researchGrowthAssumptions } from './GrowthAssumptionService.mts'
import { calculateReportModel } from './ReportCalculationRegistry.mts'
import { buildReportNarrative } from './ReportNarrativeService.mts'
import { generateReportFiles } from './ReportGenerationService.mts'
import type { ReportJob } from './types.mts'

export async function saveJob(job: ReportJob) {
  const store = getStore({ name: REPORTS_STORE, consistency: 'strong' })
  await store.setJSON(`${job.userId}/${job.id}.json`, job)
}

export async function getJob(userId: string, jobId: string) {
  const store = getStore({ name: REPORTS_STORE, consistency: 'strong' })
  return (await store.get(`${userId}/${jobId}.json`, { type: 'json' })) as ReportJob | null
}

// Runs one report job end to end. Each stage is its own service; this function
// only sequences them and persists the outcome:
//   extract (Claude) -> validate -> calculate (per report type) -> narrate -> generate files
export async function processReport(job: ReportJob) {
  const reportsStore = getStore({ name: REPORTS_STORE, consistency: 'strong' })
  const uploadsStore = getStore(UPLOADS_STORE)

  try {
    const upload = await uploadsStore.getWithMetadata(job.fileKey, { type: 'arrayBuffer' })
    if (!upload) throw new Error('The uploaded file could not be found.')

    const meta = upload.metadata as Record<string, string>
    const filename = meta.originalName || job.fileName
    const mimeType = getMimeType(filename, meta.contentType)
    const fileBuffer = Buffer.from(upload.data as ArrayBuffer)
    const reportDef = REPORT_DEFINITIONS[job.reportType]

    const extracted = await extractDocument({ fileBuffer, filename, mimeType, reportType: job.reportType })

    const validation = validateFinancialInputs(job.reportType, extracted)
    for (const issue of validation.issues) {
      if (issue.severity === 'warning') console.warn(`Report ${job.id} validation: ${issue.message}`)
    }
    assertNoBlockingIssues(validation)

    const growth = job.reportType === 'financial-projections' ? await researchGrowthAssumptions(extracted, filename) : undefined
    const model = calculateReportModel(job.reportType, validation.lineItems, growth)
    const reportData = await buildReportNarrative(extracted, validation.lineItems, model, { reportLabel: reportDef.label, issues: validation.issues, growth })
    const files = await generateReportFiles(reportData, model, reportDef.label, filename)

    const downloadKey = `${job.userId}/${job.id}.xlsx`
    await reportsStore.set(downloadKey, files.xlsx)

    let downloadKeyPptx: string | undefined
    if (files.pptx) {
      downloadKeyPptx = `${job.userId}/${job.id}.pptx`
      await reportsStore.set(downloadKeyPptx, files.pptx)
    }

    await saveJob({
      ...job,
      status: 'complete',
      fileName: filename,
      updatedAt: new Date().toISOString(),
      downloadKey,
      downloadKeyPptx,
    })
  } catch (error) {
    console.error('Report worker failed', error)
    await saveJob({
      ...job,
      status: 'failed',
      updatedAt: new Date().toISOString(),
      error: 'Report generation failed. Please try again.',
    })
  }
}
