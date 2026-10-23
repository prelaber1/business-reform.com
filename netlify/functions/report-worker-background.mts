import type { Config } from '@netlify/functions'
import { getJob, processReport } from './generate-report.mts'

function json(data: unknown, status = 200) {
  return Response.json(data, { status })
}

export default async (req: Request) => {
  if (req.method !== 'POST') return json({ error: 'Method not allowed.' }, 405)

  let body: { jobId?: string; userId?: string; workerToken?: string }
  try {
    body = await req.json()
  } catch {
    return json({ error: 'Invalid JSON body.' }, 400)
  }

  const { jobId, userId, workerToken } = body
  if (!jobId || !userId || !workerToken) {
    return json({ error: 'Missing worker payload.' }, 400)
  }

  const job = await getJob(userId, jobId)
  if (!job || job.workerToken !== workerToken) {
    return json({ error: 'Report job not found.' }, 404)
  }

  if (job.status !== 'processing') {
    return json({ ok: true, status: job.status })
  }

  await processReport(job)
  return json({ ok: true })
}

export const config: Config = {
  path: '/api/portal/report-worker-background',
  method: 'POST',
}
