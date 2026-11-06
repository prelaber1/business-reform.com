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

export async function hasActiveSubscription(email: string): Promise<boolean> {
  const key = process.env.BR_STRIPE_SECRET_KEY
  if (!key) {
    console.error('BR_STRIPE_SECRET_KEY is not configured')
    throw new SubscriptionCheckError('Subscription check is not configured.', 'missing_key')
  }

  const normalized = email.trim().toLowerCase()
  const cached = cache.get(normalized)
  if (cached && cached.expires > Date.now()) return cached.active

  // Stripe email filters are case-sensitive, so try the address as typed in
  // Stripe-friendly forms: lowercase, plus a case-insensitive search.
  const customerIds = new Set<string>()
  // The exact-email list is required; the case-insensitive search is best
  // effort (a failure there must not block a customer found by the list).
  const [listed, searched] = await Promise.allSettled([
    stripeGet('/customers', { email: normalized, limit: '10' }, key),
    stripeGet('/customers/search', { query: `email:"${normalized.replace(/["\\]/g, '')}"`, limit: '10' }, key),
  ])
  if (listed.status === 'rejected' && searched.status === 'rejected') throw listed.reason
  for (const result of [listed, searched]) {
    if (result.status !== 'fulfilled') continue
    for (const c of result.value.data) {
      if (typeof c.email === 'string' && c.email.toLowerCase() === normalized) customerIds.add(c.id)
    }
  }

  let active = false
  for (const customerId of customerIds) {
    const subs = await stripeGet('/subscriptions', { customer: customerId, status: 'all', limit: '20' }, key)
    if (subs.data.some((s) => ACTIVE_STATUSES.has(s.status))) {
      active = true
      break
    }
  }

  // Only cache positive results so a new purchase unlocks immediately.
  if (active) cache.set(normalized, { active, expires: Date.now() + CACHE_TTL_MS })
  return active
}
