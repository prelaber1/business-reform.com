import { getStore } from '@netlify/blobs'
import { getUser } from '@netlify/identity'
import type { Config, Context } from '@netlify/functions'
import { REPORTS_STORE, UPLOADS_STORE } from './report-services/constants.mts'
import { REPORT_DEFINITIONS } from './report-services/reportDefinitions.mts'
import { getJob, saveJob } from './report-services/ReportOrchestrationService.mts'
import type { ReportJob } from './report-services/types.mts'

// HTTP entry point. The report pipeline itself lives in ./report-services; the
// background worker (report-worker-background.mts) runs it via processReport.
export { getJob, saveJob, processReport } from './report-services/ReportOrchestrationService.mts'
export type { LineItem, ReportAssumptions, ReportData } from './report-services/types.mts'
export { calculateReportModel } from './report-services/ReportCalculationRegistry.mts'

function json(data: unknown, status = 200) {
  return Response.json(data, { status })
}

function getReportDownloadFilename(reportLabel: string, extension: 'xlsx' | 'pptx' = 'xlsx') {
  const safeName = reportLabel.replace(/[^\w.\- ]+/g, '_').trim() || 'Generated Report'
  return `${safeName}.${extension}`
}

async function startBackgroundReportWorker(req: Request, job: ReportJob) {
  // report-worker-background.mts declares a custom config.path, which in
  // Netlify Functions v2 replaces its default /.netlify/functions/<name>
  // route rather than adding to it — so it must be invoked at that path.
  const workerUrl = new URL('/api/portal/report-worker-background', req.url)
  const response = await fetch(workerUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      jobId: job.id,
      userId: job.userId,
      workerToken: job.workerToken,
    }),
  })

  if (!response.ok) {
    const bodyText = await response.text().catch(() => '')
    console.error('report-worker-background invocation failed', {
      url: workerUrl.toString(),
      status: response.status,
      statusText: response.statusText,
      body: bodyText,
    })
    await saveJob({
      ...job,
      status: 'failed',
      updatedAt: new Date().toISOString(),
      error: 'Report generation could not start. Please try again.',
    })
    throw new Error('Report generation could not start. Please try again.')
  }
}

const ALLOWED_REPORT_EMAILS = new Set([
  'a.romosios@gmail.com',
  'a_romosios@yahoo.gr',
'anna.eftychidou@gmail.com',
  'themis.eftychidoy@gmail.com',
'vasilis.tsolis@cognitivplus.com',
'gianniszetas@gmail.com',
'angelos.p.eftychidis@gmail.com',
])

async function handlePost(req: Request, _context: Context, userId: string, userEmail: string | undefined) {
  if (!userEmail || !ALLOWED_REPORT_EMAILS.has(userEmail.toLowerCase())) {
    return json({ error: 'Not authorized user request. Please contact site administrators' }, 403)
  }

  let body: { fileKey?: string; reportType?: string }
  try {
    body = await req.json()
  } catch {
    return json({ error: 'Invalid JSON body.' }, 400)
  }

  const fileKey = body.fileKey
  const reportType = body.reportType
  if (!fileKey || !fileKey.startsWith(`${userId}/`)) {
    return json({ error: 'Choose one of your uploaded files first.' }, 400)
  }
  if (!reportType || !REPORT_DEFINITIONS[reportType]) {
    return json({ error: 'Choose a valid report type.' }, 400)
  }

  const uploadsStore = getStore(UPLOADS_STORE)
  const uploadMeta = await uploadsStore.getMetadata(fileKey)
  if (!uploadMeta) return json({ error: 'The selected file was not found.' }, 404)

  const meta = uploadMeta.metadata as Record<string, string>
  const now = new Date().toISOString()
  const job: ReportJob = {
    id: crypto.randomUUID(),
    userId,
    fileKey,
    fileName: meta.originalName || fileKey.slice(`${userId}/`.length),
    reportType,
    reportLabel: REPORT_DEFINITIONS[reportType].label,
    status: 'processing',
    createdAt: now,
    updatedAt: now,
    workerToken: crypto.randomUUID(),
  }

  await saveJob(job)
  try {
    await startBackgroundReportWorker(req, job)
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : 'Report generation could not start. Please try again.' }, 502)
  }

  return json({ jobId: job.id, status: job.status, reportLabel: job.reportLabel }, 202)
}

async function handleGet(req: Request, userId: string) {
  const url = new URL(req.url)
  const jobId = url.searchParams.get('jobId')
  if (!jobId) return json({ error: 'Missing jobId.' }, 400)

  const job = await getJob(userId, jobId)
  if (!job) return json({ error: 'Report job not found.' }, 404)

  if (url.searchParams.get('download') === '1') {
    const wantsPptx = url.searchParams.get('format') === 'pptx'
    const downloadKey = wantsPptx ? job.downloadKeyPptx : job.downloadKey
    if (job.status !== 'complete' || !downloadKey) {
      return json({ error: 'Report is not ready yet.' }, 409)
    }
    const reportsStore = getStore({ name: REPORTS_STORE, consistency: 'strong' })
    const report = await reportsStore.get(downloadKey, { type: 'arrayBuffer' })
    if (!report) return json({ error: 'Report file not found.' }, 404)

    const filename = getReportDownloadFilename(job.reportLabel, wantsPptx ? 'pptx' : 'xlsx')
    return new Response(report, {
      headers: {
        'Content-Type': wantsPptx
          ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    })
  }

  return json({
    jobId: job.id,
    status: job.status,
    reportLabel: job.reportLabel,
    fileName: job.fileName,
    downloadFileName: getReportDownloadFilename(job.reportLabel),
    downloadFileNamePptx: job.downloadKeyPptx ? getReportDownloadFilename(job.reportLabel, 'pptx') : null,
    error: job.error,
    downloadUrl: job.status === 'complete' && job.downloadKey ? `/api/portal/generate-report?jobId=${encodeURIComponent(job.id)}&download=1` : null,
    downloadUrlPptx:
      job.status === 'complete' && job.downloadKeyPptx
        ? `/api/portal/generate-report?jobId=${encodeURIComponent(job.id)}&download=1&format=pptx`
        : null,
  })
}

export default async (req: Request, context: Context) => {
  const user = await getUser()
  if (!user) return json({ error: 'You must be logged in.' }, 401)

  if (req.method === 'POST') return handlePost(req, context, user.id, user.email)
  if (req.method === 'GET') return handleGet(req, user.id)
  return json({ error: 'Method not allowed.' }, 405)
}

export const config: Config = {
  path: '/api/portal/generate-report',
  method: ['POST', 'GET'],
}