import { getStore } from '@netlify/blobs'
import { getUser } from '@netlify/identity'
import type { Config, Context } from '@netlify/functions'

const STORE_NAME = 'client-uploads'
const MAX_BYTES = 25 * 1024 * 1024 // 25 MB per file

export default async (req: Request, _context: Context) => {
  // Only authenticated portal users may upload.
  const user = await getUser()
  if (!user) {
    return Response.json({ error: 'You must be logged in to upload files.' }, { status: 401 })
  }

  let form: FormData
  try {
    form = await req.formData()
  } catch {
    return Response.json({ error: 'Expected a multipart form upload.' }, { status: 400 })
  }

  const file = form.get('file')
  if (!(file instanceof File) || file.size === 0) {
    return Response.json({ error: 'No file was provided.' }, { status: 400 })
  }
  if (file.size > MAX_BYTES) {
    return Response.json({ error: 'File is too large. The limit is 25 MB.' }, { status: 413 })
  }

  // Scope every blob to the owner so users only ever see their own files.
  const safeName = (file.name || 'upload').replace(/[^\w.\- ]+/g, '_').slice(0, 200)
  const key = `${user.id}/${Date.now()}-${safeName}`

  const store = getStore(STORE_NAME)
  const uploadedAt = new Date().toISOString()
  const contentType = file.type || 'application/octet-stream'
  const originalName = file.name || safeName

  await store.set(key, await file.arrayBuffer(), {
    metadata: {
      contentType,
      originalName,
      size: String(file.size),
      uploadedAt,
    },
  })

  return Response.json({
    ok: true,
    file: {
      key,
      name: originalName,
      size: file.size,
      contentType,
      uploadedAt,
    },
    key,
    name: originalName,
    size: file.size,
  })
}

export const config: Config = {
  path: '/api/portal/upload',
  method: 'POST',
}
