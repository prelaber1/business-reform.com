import Anthropic from '@anthropic-ai/sdk'
import { getStore } from '@netlify/blobs'
import { getUser } from '@netlify/identity'
import type { Config, Context } from '@netlify/functions'

const CHAT_STORE = 'client-chat-history'
const MAX_STORED_MESSAGES = 40
const MAX_CONTEXT_MESSAGES = 20
const MAX_MESSAGE_CHARS = 4000

type ChatMessage = { role: 'user' | 'assistant'; content: string; timestamp: string }

const SYSTEM_PROMPT = `You are the AI CFO helper for Business Reform, a platform that gives businesses AI-powered financial intelligence (financial projections, break-even analysis, burn rate, capital optimization, use of funds reports). Answer the user's questions about their business finances, the platform's features, or general financial/business guidance clearly and concisely, in a professional but approachable tone. You are not a substitute for personalized licensed financial, legal, or tax advice — say so if asked something that clearly requires it. Keep answers focused; expand only if the user asks for more depth.`

function json(data: unknown, status = 200) {
  return Response.json(data, { status })
}

async function loadHistory(userId: string): Promise<ChatMessage[]> {
  const store = getStore({ name: CHAT_STORE, consistency: 'strong' })
  const history = await store.get(`${userId}.json`, { type: 'json' })
  return Array.isArray(history) ? (history as ChatMessage[]) : []
}

async function saveHistory(userId: string, history: ChatMessage[]) {
  const store = getStore({ name: CHAT_STORE, consistency: 'strong' })
  await store.setJSON(`${userId}.json`, history.slice(-MAX_STORED_MESSAGES))
}

async function handleGet(userId: string) {
  const history = await loadHistory(userId)
  return json({ messages: history })
}

async function handlePost(req: Request, userId: string) {
  let body: { message?: string }
  try {
    body = await req.json()
  } catch {
    return json({ error: 'Invalid JSON body.' }, 400)
  }

  const message = (body.message || '').trim()
  if (!message) return json({ error: 'Message cannot be empty.' }, 400)
  if (message.length > MAX_MESSAGE_CHARS) return json({ error: 'Message is too long.' }, 400)

  const history = await loadHistory(userId)
  history.push({ role: 'user', content: message, timestamp: new Date().toISOString() })

  try {
    const anthropic = new Anthropic()
    const contextMessages = history.slice(-MAX_CONTEXT_MESSAGES).map((m) => ({ role: m.role, content: m.content }))
    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 1024,
      system: SYSTEM_PROMPT,
      messages: contextMessages,
    })

    const replyText =
      response.content
        .map((part) => (part.type === 'text' ? part.text : ''))
        .join('\n')
        .trim() || "Sorry, I couldn't generate a response. Please try again."

    const assistantMessage: ChatMessage = { role: 'assistant', content: replyText, timestamp: new Date().toISOString() }
    history.push(assistantMessage)
    await saveHistory(userId, history)

    return json({ reply: assistantMessage, messages: history.slice(-MAX_STORED_MESSAGES) })
  } catch (error) {
    console.error('Chat AI call failed', error)
    // Keep the user's message even if the AI call failed, so it isn't lost.
    await saveHistory(userId, history)
    return json({ error: 'Could not get a response right now. Please try again.' }, 502)
  }
}

export default async (req: Request, _context: Context) => {
  const user = await getUser()
  if (!user) return json({ error: 'You must be logged in.' }, 401)

  if (req.method === 'GET') return handleGet(user.id)
  if (req.method === 'POST') return handlePost(req, user.id)
  return json({ error: 'Method not allowed.' }, 405)
}

export const config: Config = {
  path: '/api/portal/chat',
  method: ['GET', 'POST'],
}
