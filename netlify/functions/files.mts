import { getStore } from '@netlify/blobs'
import { getUser } from '@netlify/identity'
import type { Config, Context } from '@netlify/functions'

const STORE_NAME = 'client-uploads'

export default async (req: Request, _context: Context) => {
  const user = await getUser()
  if (!user) {
    return Response.json({ error: 'You must be logged in.' }, { status: 401 })
  }

  const store = getStore(STORE_NAME)
  const prefix = `${user.id}/`
  const key = new URL(req.url).searchParams.get('key')

  // Guard: a user may only ever touch keys under their own prefix.
  if (key && !key.startsWith(prefix)) {
    return Response.json({ error: 'Not found.' }, { status: 404 })
  }

  // Delete a single file.
  if (req.method === 'DELETE') {
    if (!key) return Response.json({ error: 'Missing key.' }, { status: 400 })
    await store.delete(key)
    return Response.json({ ok: true })
  }

  // Download a single file.
  if (key) {
    const result = await store.getWithMetadata(key, { type: 'blob' })
    if (!result) return Response.json({ error: 'Not found.' }, { status: 404 })
    const meta = result.metadata as Record<string, string>
    const name = meta.originalName || key.slice(prefix.length)
    return new Response(result.data, {
      headers: {
        'Content-Type': meta.contentType || 'application/octet-stream',
        'Content-Disposition': `attachment; filename="${name.replace(/"/g, '')}"`,
      },
    })
  }

  // List all of the current user's files.
  const { blobs } = await store.list({ prefix })
  const files = await Promise.all(
    blobs.map(async (blob) => {
      const meta = (await store.getMetadata(blob.key))?.metadata as Record<string, string> | undefined
      return {
        key: blob.key,
        name: meta?.originalName || blob.key.slice(prefix.length),
        size: Number(meta?.size || 0),
        contentType: meta?.contentType || 'application/octet-stream',
        uploadedAt: meta?.uploadedAt || null,
      }
    }),
  )
  files.sort((a, b) => (b.uploadedAt || '').localeCompare(a.uploadedAt || ''))

  return Response.json({ files })
}

export const config: Config = {
  path: '/api/portal/files',
  method: ['GET', 'DELETE'],
}
