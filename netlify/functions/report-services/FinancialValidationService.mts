import type { LineItem, ReportData } from './types.mts'
import type { ExtractedDocument } from './DocumentExtractionService.mts'

export type ValidationIssue = { severity: 'error' | 'warning'; message: string }

export type ValidationResult = {
  lineItems: LineItem[]
  issues: ValidationIssue[]
}

// Group tags each report's calculation service cannot do without. Mirrors the
// vocabulary in REPORT_DEFINITIONS; a missing group is reported as a warning
// (the calculators already degrade with their own on-sheet caveats).
const REQUIRED_GROUPS: Record<string, string[]> = {
  'break-even': ['revenue'],
  'use-of-funds': ['allocation_category'],
  'burn-rate': ['starting_cash_balance'],
  'cash-flow-forecast': ['starting_cash_balance'],
  'capital-optimization': ['equity'],
  'financial-projections': ['revenue'],
  'budget-vs-actual': ['actual_revenue', 'actual_cost', 'budget_revenue', 'budget_cost'],
  'fpa-digest': ['current_period', 'trailing_period', 'prior_year_period'],
}

const isValidLineItem = (li: unknown): li is LineItem => {
  const item = li as Partial<LineItem> | null
  return !!item && typeof item.label === 'string' && typeof item.group === 'string' && Number.isFinite(item.amount)
}

// Checks the extracted figures before any calculation runs: drops items that
// are missing required fields or have non-numeric amounts, rejects an
// extraction with nothing usable, and flags business-rule problems. Only the
// first two affect the report; warnings are surfaced to logs so they can be
// reviewed without changing report output.
export function validateFinancialInputs(reportType: string, extracted: ExtractedDocument): ValidationResult {
  const issues: ValidationIssue[] = []
  const lineItems = extracted.lineItems.filter(isValidLineItem)

  const dropped = extracted.lineItems.length - lineItems.length
  if (dropped > 0) {
    issues.push({ severity: 'warning', message: `${dropped} extracted line item(s) were dropped for a missing label/group or non-numeric amount.` })
  }

  const { executiveSummary, sections, tables } = extracted.narrative
  if (!executiveSummary && !sections.length && !tables.length && !lineItems.length) {
    issues.push({ severity: 'error', message: 'The AI worker returned an empty or incomplete report.' })
    return { lineItems, issues }
  }

  const groups = new Set(lineItems.map((li) => li.group))
  for (const required of REQUIRED_GROUPS[reportType] ?? []) {
    if (!groups.has(required)) issues.push({ severity: 'warning', message: `No line items were classified as "${required}".` })
  }

  if (groups.has('starting_cash_balance') && lineItems.filter((li) => li.group === 'starting_cash_balance').length > 1) {
    issues.push({ severity: 'warning', message: 'More than one starting_cash_balance item was extracted; only one is expected.' })
  }

  const periods = lineItems.map((li) => li.period).filter((p): p is string => !!p)
  if (periods.some((p) => /week/i.test(p)) && periods.some((p) => /month/i.test(p))) {
    issues.push({ severity: 'warning', message: 'Line items mix weekly and monthly period labels.' })
  }

  issues.push(...runModelSanityChecks(reportType, lineItems))

  const revenueNegative = lineItems.some((li) => li.group === 'revenue' && li.amount < 0)
  if (revenueNegative) issues.push({ severity: 'warning', message: 'A revenue line item has a negative amount.' })

  return { lineItems, issues }
}

export function assertNoBlockingIssues(result: ValidationResult) {
  const blocking = result.issues.find((issue) => issue.severity === 'error')
  if (blocking) throw new Error(blocking.message)
}

const sumOf = (items: LineItem[], group: string) =>
  items.filter((li) => li.group === group).reduce((total, li) => total + li.amount, 0)

// Balance sheet group tags per report type: [asset groups, liability groups, equity groups].
const BALANCE_SHEET_GROUPS: Record<string, [string[], string[], string[]]> = {
  'financial-projections': [['current_asset', 'fixed_asset'], ['current_liability', 'long_term_debt'], ['equity']],
  'capital-optimization': [
    ['current_asset', 'fixed_asset'],
    ['short_term_debt', 'long_term_debt', 'accounts_payable', 'current_liability', 'other_liability'],
    ['equity'],
  ],
}

const TIE_TOLERANCE = 0.01 // 1% of total assets
const MAGNITUDE_RATIO = 1e4 // same-group items this far apart usually mean a units mix-up (units vs thousands)

// Mechanical checks adapted from the model-sanity-checker rubric (ties &
// totals, cross-item consistency, hygiene), applied to the extracted line
// items. They flag, never block: the extraction is the model being reviewed.
export function runModelSanityChecks(reportType: string, lineItems: LineItem[]): ValidationIssue[] {
  const issues: ValidationIssue[] = []
  const warn = (message: string) => issues.push({ severity: 'warning', message })

  // Ties: assets = liabilities + equity.
  const bs = BALANCE_SHEET_GROUPS[reportType]
  if (bs) {
    const [assetGroups, liabilityGroups, equityGroups] = bs
    const present = [...assetGroups, ...liabilityGroups, ...equityGroups].some((g) => lineItems.some((li) => li.group === g))
    if (present) {
      const assets = assetGroups.reduce((t, g) => t + sumOf(lineItems, g), 0)
      const liabilitiesAndEquity = [...liabilityGroups, ...equityGroups].reduce((t, g) => t + sumOf(lineItems, g), 0)
      const gap = assets - liabilitiesAndEquity
      if (assets > 0 && Math.abs(gap) > assets * TIE_TOLERANCE) {
        warn(`Balance sheet does not tie: assets ${assets.toFixed(0)} vs liabilities + equity ${liabilitiesAndEquity.toFixed(0)} (gap ${gap.toFixed(0)}). Items may be missing or misclassified.`)
      }
    }
  }

  // Ties: cash flow sections should sum to a plausible net change; flag a lone section.
  if (reportType === 'financial-projections') {
    const cf = ['operating_cash_flow_item', 'investing_cash_flow_item', 'financing_cash_flow_item']
    const present = cf.filter((g) => lineItems.some((li) => li.group === g))
    if (present.length > 0 && present.length < 3) warn(`Cash flow statement is incomplete: only ${present.join(', ')} extracted.`)

    const revenue = sumOf(lineItems, 'revenue')
    const cogs = Math.abs(sumOf(lineItems, 'cogs'))
    if (revenue > 0 && cogs > revenue) warn('COGS exceeds revenue (negative gross margin); check the sign or classification.')
    if (revenue > 0 && lineItems.some((li) => li.group === 'tax') && sumOf(lineItems, 'revenue') - cogs - Math.abs(sumOf(lineItems, 'opex')) < 0 && sumOf(lineItems, 'tax') > 0) {
      warn('A tax charge is recorded although operating profit appears negative; check for a tax credit or a sign error.')
    }
  }

  if (reportType === 'capital-optimization') {
    const past = sumOf(lineItems, 'past_due_payable')
    const payables = sumOf(lineItems, 'accounts_payable')
    if (past > payables && payables > 0) warn('Past-due payables exceed total accounts payable; past_due_payable is a memo item within accounts_payable.')
  }

  if (reportType === 'burn-rate' || reportType === 'cash-flow-forecast') {
    const flows = lineItems.filter((li) => li.group === 'cash_inflow' || li.group === 'cash_outflow')
    const dated = flows.filter((li) => li.period)
    if (flows.length > 0 && dated.length === 0) warn('No cash flow item has a period; run-rate figures will be spread evenly, which hides timing.')
    if (dated.length > 0 && dated.length < flows.length) warn(`${flows.length - dated.length} of ${flows.length} cash flow items have no period and may be dropped from the schedule.`)
    for (const li of dated) {
      const match = /(week|month)\s*(\d+)/i.exec(li.period ?? '')
      if (!match) warn(`Unrecognised period label "${li.period}" on "${li.label}".`)
      else if (Number(match[2]) > (/week/i.test(match[1]) ? 13 : 12) || Number(match[2]) < 1) warn(`Period "${li.period}" on "${li.label}" is outside the forecast horizon.`)
    }
    if (flows.length > 0 && !flows.some((li) => li.group === 'cash_outflow')) warn('No cash outflows were extracted; burn and runway cannot be meaningful.')
  }

  if (reportType === 'fpa-digest') {
    const current = new Set(lineItems.filter((li) => li.group === 'current_period').map((li) => li.label.trim().toLowerCase()))
    for (const group of ['trailing_period', 'prior_year_period']) {
      const orphans = [...new Set(lineItems.filter((li) => li.group === group).map((li) => li.label.trim().toLowerCase()))].filter((l) => !current.has(l))
      if (orphans.length) warn(`${group} has metrics with no matching current_period metric (labels must match exactly): ${orphans.join(', ')}.`)
    }
    const trailing = lineItems.filter((li) => li.group === 'trailing_period')
    if (trailing.length && trailing.some((li) => !li.period)) warn('Some trailing_period items have no period label; the number of periods averaged cannot be verified.')
    const priorYear = lineItems.filter((li) => li.group === 'prior_year_period')
    if (new Set(priorYear.map((li) => li.label.trim().toLowerCase())).size < priorYear.length) warn('prior_year_period has more than one figure for the same metric; they will be summed.')
  }

  // Cross-item consistency: sign conventions and units.
  for (const li of lineItems) {
    if ((li.group === 'cash_outflow' || li.group === 'cogs' || li.group === 'opex') && li.amount < 0) {
      warn(`"${li.label}" (${li.group}) is negative; costs should be entered as positive amounts.`)
    }
  }
  const byGroup = new Map<string, number[]>()
  for (const li of lineItems) {
    const magnitude = Math.abs(li.amount)
    if (magnitude > 0) byGroup.set(li.group, [...(byGroup.get(li.group) ?? []), magnitude])
  }
  for (const [group, values] of byGroup) {
    if (values.length > 1 && Math.max(...values) / Math.min(...values) > MAGNITUDE_RATIO) {
      warn(`Items in "${group}" differ by more than ${MAGNITUDE_RATIO.toExponential(0)}x; check for a units mix-up (units vs thousands).`)
    }
  }

  // Hygiene: the same label twice in a group is usually a double count.
  const seen = new Set<string>()
  for (const li of lineItems) {
    const key = `${li.group}|${li.period ?? ''}|${li.label.trim().toLowerCase()}|${li.amount}`
    if (seen.has(key)) warn(`"${li.label}" appears more than once in "${li.group}" with the same amount; possible double count.`)
    seen.add(key)
  }

  return issues
}
