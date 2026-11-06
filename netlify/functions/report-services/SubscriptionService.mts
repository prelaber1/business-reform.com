// Server-side check that a user has an active Stripe subscription.
// Uses the Stripe REST API directly (no SDK) with a restricted key stored in
// the BR_STRIPE_SECRET_KEY Netlify environment variable (read-only: Customers, Subscriptions).

const STRIPE_API = 'https://api.stripe.com/v1'
const CACHE_TTL_MS = 60_000
const ACTIVE_STATUSES = new Set(['active', 'trialing'])

const cache = new Map<string, { active: boolean; expires: number }>()

export class SubscriptionCheckError extends Error {
  constructor(message: string, public reason = 'unknown') {
    super(message)
  }
}

async function stripeGet(path: string, params: Record<string, string>, key: string) {
  const res = await fetch(`${STRIPE_API}${path}?${new URLSearchParams(params)}`, {
    headers: { Authorization: `Bearer ${key}` },
  })
  if (!res.ok) {
    const text = await res.text().catch(() => '')
    console.error('Stripe API error', { path, status: res.status, body: text.slice(0, 300) })
    throw new SubscriptionCheckError('Subscription check failed.', `stripe_${res.status}`)
  }
  return (await res.json()) as { data: Array<Record<string, any>> }
}

export type SubscriptionDebug = {
  mode: 'live' | 'test' | 'unknown'
  customers: number
  subscriptionStatuses: string[]
  checkoutSessions: number
  lookupErrors: string[]
}

export async function hasActiveSubscription(email: string): Promise<{ active: boolean; debug: SubscriptionDebug }> {
  const key = process.env.BR_STRIPE_SECRET_KEY
  if (!key) {
    console.error('BR_STRIPE_SECRET_KEY is not configured')
    throw new SubscriptionCheckError('Subscription check is not configured.', 'missing_key')
  }

  const debug: SubscriptionDebug = {
    mode: key.startsWith('rk_live_') || key.startsWith('sk_live_') ? 'live' : key.startsWith('rk_test_') || key.startsWith('sk_test_') ? 'test' : 'unknown',
    customers: 0,
    subscriptionStatuses: [],
    checkoutSessions: 0,
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
    ...emails.map((e) => stripeGet('/customers', { email: e, limit: '10' }, key)),
    stripeGet('/customers/search', { query: `email:"${q}"`, limit: '10' }, key),
    ...emails.map((e) => stripeGet('/checkout/sessions', { 'customer_details[email]': e, limit: '20' }, key)),
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
      const subs = await stripeGet('/subscriptions', { customer: customerId, status: 'all', limit: '20' }, key)
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
  debug.subscriptionStatuses = [...statuses]

  const active = debug.subscriptionStatuses.some((st) => ACTIVE_STATUSES.has(st))
  // Only cache positive results so a new purchase unlocks immediately.
  if (active) cache.set(normalized, { active, expires: Date.now() + CACHE_TTL_MS })
  return { active, debug }
}
