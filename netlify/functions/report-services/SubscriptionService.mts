// Server-side check that a user has an active Stripe subscription.
// Uses the Stripe REST API directly (no SDK) with a restricted key stored in
// the BR_STRIPE_SECRET_KEY Netlify environment variable (read-only: Customers, Subscriptions).

const STRIPE_API = 'https://api.stripe.com/v1'
const CACHE_TTL_MS = 60_000
const ACTIVE_STATUSES = new Set(['active', 'trialing'])

const cache = new Map<string, { active: boolean; expires: number }>()

export type KeyInfo = {
  defined: boolean
  length: number
  prefix: string
  hadWhitespaceOrQuotes: boolean
  // Names only, never values: env vars visible to this function that look relevant.
  visibleEnvNames?: string[]
  totalEnvVars?: number
  netlifyEnvAvailable?: boolean
  runtime?: string
}

export class SubscriptionCheckError extends Error {
  constructor(message: string, public reason = 'unknown', public keyInfo?: KeyInfo, public stripeCode?: string) {
    super(message)
  }
}

// Reads the key from the environment, tolerating pasted whitespace, newlines,
// surrounding quotes or a leading "Bearer ". Returns only non-secret metadata
// alongside it (never the key itself).
function readKey(): { key: string; info: KeyInfo } {
  const nodeEnv = (globalThis as any).process?.env?.BR_STRIPE_SECRET_KEY as string | undefined
  const netlifyEnv = (globalThis as any).Netlify?.env?.get?.('BR_STRIPE_SECRET_KEY') as string | undefined
  const raw = nodeEnv ?? netlifyEnv ?? ''
  const key = raw.trim().replace(/^bearer\s+/i, '').replace(/^["']+|["']+$/g, '').trim()
  const m = key.match(/^(rk_live|rk_test|sk_live|sk_test|pk_live|pk_test)_/)
  const g = globalThis as any
  const names = new Set<string>([...Object.keys(g.process?.env ?? {})])
  let netlifyNames: string[] = []
  try { netlifyNames = Object.keys(g.Netlify?.env?.toObject?.() ?? {}) } catch { /* ignore */ }
  netlifyNames.forEach((n) => names.add(n))
  const relevant = [...names].filter((n) => /stripe/i.test(n) || n.startsWith('BR_')).sort()
  return {
    key,
    info: {
      visibleEnvNames: relevant,
      totalEnvVars: names.size,
      netlifyEnvAvailable: Boolean(g.Netlify?.env),
      runtime: g.Netlify ? 'netlify-functions-v2' : 'other',
      defined: raw.length > 0,
      length: key.length,
      prefix: m ? m[1] : key ? 'unexpected' : 'none',
      hadWhitespaceOrQuotes: key !== raw,
    },
  }
}

async function stripeGet(path: string, params: Record<string, string>, key: string, info?: KeyInfo) {
  const res = await fetch(`${STRIPE_API}${path}?${new URLSearchParams(params)}`, {
    headers: { Authorization: `Bearer ${key}` },
  })
  if (!res.ok) {
    const text = await res.text().catch(() => '')
    let stripeCode: string | undefined
    try { stripeCode = JSON.parse(text)?.error?.code || JSON.parse(text)?.error?.type } catch { /* not JSON */ }
    console.error('Stripe API error', { path, status: res.status, stripeCode, keyInfo: info })
    throw new SubscriptionCheckError('Subscription check failed.', `stripe_${res.status}`, info, stripeCode)
  }
  return (await res.json()) as { data: Array<Record<string, any>> }
}

export type SubscriptionDebug = {
  mode: 'live' | 'test' | 'unknown'
  keyInfo: KeyInfo
  customers: number
  subscriptionStatuses: string[]
  checkoutSessions: number
  searchedEmail: string
  fallback: { subscriptionsScanned: number; sessionsScanned: number; matched: number } | null
  lookupErrors: string[]
}

export async function hasActiveSubscription(email: string): Promise<{ active: boolean; debug: SubscriptionDebug }> {
  const { key, info } = readKey()
  if (!key) {
    console.error('BR_STRIPE_SECRET_KEY is not configured', { keyInfo: info })
    throw new SubscriptionCheckError('Subscription check is not configured.', 'missing_key', info)
  }

  const debug: SubscriptionDebug = {
    mode: info.prefix.endsWith('_live') ? 'live' : info.prefix.endsWith('_test') ? 'test' : 'unknown',
    keyInfo: info,
    customers: 0,
    subscriptionStatuses: [],
    checkoutSessions: 0,
    searchedEmail: email.trim(),
    fallback: null,
    lookupErrors: [],
  }

  const trimmed = email.trim()
  const normalized = trimmed.toLowerCase()
  const cached = cache.get(normalized)
  if (cached && cached.expires > Date.now()) return { active: cached.active, debug }

  const note = (label: string, r: PromiseRejectedResult) => {
    debug.lookupErrors.push(`${label}:${r.reason instanceof SubscriptionCheckError ? r.reason.reason : 'error'}`)
  }

  // Stripe email filters are case-sensitive: try lowercase and as-typed, plus
  // a search, and checkout sessions (covers customers whose email differs in case
  // or checkouts where the customer record has no email).
  const emails = [...new Set([normalized, trimmed])]
  const q = normalized.replace(/["\\]/g, '')
  const results = await Promise.allSettled([
    ...emails.map((e) => stripeGet('/customers', { email: e, limit: '10' }, key, info)),
    stripeGet('/customers/search', { query: `email:"${q}"`, limit: '10' }, key, info),
    ...emails.map((e) => stripeGet('/checkout/sessions', { 'customer_details[email]': e, limit: '20' }, key, info)),
  ])
  const labels = [...emails.map(() => 'customers'), 'customer_search', ...emails.map(() => 'checkout_sessions')]
  if (results.every((r) => r.status === 'rejected')) throw (results[0] as PromiseRejectedResult).reason

  const customerIds = new Set<string>()
  const subscriptionIds = new Set<string>()
  results.forEach((r, i) => {
    if (r.status === 'rejected') return note(labels[i], r)
    for (const item of r.value.data) {
      if (labels[i] === 'checkout_sessions') {
        if (item.mode !== 'subscription') continue
        debug.checkoutSessions++
        if (typeof item.customer === 'string') customerIds.add(item.customer)
        if (typeof item.subscription === 'string') subscriptionIds.add(item.subscription)
      } else if (typeof item.email === 'string' && item.email.toLowerCase() === normalized) {
        customerIds.add(item.id)
      }
    }
  })
  debug.customers = customerIds.size

  const statuses = new Set<string>()
  for (const customerId of customerIds) {
    try {
      const subs = await stripeGet('/subscriptions', { customer: customerId, status: 'all', limit: '20' }, key, info)
      for (const sub of subs.data) statuses.add(sub.status)
    } catch (error) {
      debug.lookupErrors.push(`subscriptions:${error instanceof SubscriptionCheckError ? error.reason : 'error'}`)
    }
  }
  for (const subId of subscriptionIds) {
    try {
      const res = await fetch(`${STRIPE_API}/subscriptions/${encodeURIComponent(subId)}`, { headers: { Authorization: `Bearer ${key}` } })
      if (res.ok) statuses.add(((await res.json()) as { status: string }).status)
      else debug.lookupErrors.push(`subscription:stripe_${res.status}`)
    } catch {
      debug.lookupErrors.push('subscription:error')
    }
  }
  // Last-resort fallback (bounded to the 100 most recent): match the email
  // case-insensitively against recent subscriptions' customers and recent
  // checkout sessions, in case Stripe's own email filters missed it.
  if (statuses.size === 0) {
    const fb = { subscriptionsScanned: 0, sessionsScanned: 0, matched: 0 }
    debug.fallback = fb
    const [subsR, sessR] = await Promise.allSettled([
      stripeGet('/subscriptions', { status: 'all', limit: '100', 'expand[]': 'data.customer' }, key, info),
      stripeGet('/checkout/sessions', { limit: '100' }, key, info),
    ])
    if (subsR.status === 'fulfilled') {
      fb.subscriptionsScanned = subsR.value.data.length
      for (const sub of subsR.value.data) {
        const e = typeof sub.customer === 'object' && sub.customer ? sub.customer.email : undefined
        if (typeof e === 'string' && e.trim().toLowerCase() === normalized) { statuses.add(sub.status); fb.matched++ }
      }
    } else note('fallback_subscriptions', subsR)
    if (sessR.status === 'fulfilled') {
      fb.sessionsScanned = sessR.value.data.length
      for (const sess of sessR.value.data) {
        const e = sess.customer_details?.email ?? sess.customer_email
        if (sess.mode !== 'subscription' || typeof e !== 'string' || e.trim().toLowerCase() !== normalized) continue
        fb.matched++
        if (typeof sess.subscription === 'string') {
          try {
            const res = await fetch(`${STRIPE_API}/subscriptions/${encodeURIComponent(sess.subscription)}`, { headers: { Authorization: `Bearer ${key}` } })
            if (res.ok) statuses.add(((await res.json()) as { status: string }).status)
          } catch { /* ignore */ }
        }
      }
    } else note('fallback_sessions', sessR)
  }
  debug.subscriptionStatuses = [...statuses]

  const active = debug.subscriptionStatuses.some((st) => ACTIVE_STATUSES.has(st))
  // Only cache positive results so a new purchase unlocks immediately.
  if (active) cache.set(normalized, { active, expires: Date.now() + CACHE_TTL_MS })
  return { active, debug }
}
