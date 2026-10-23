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

  const revenueNegative = lineItems.some((li) => li.group === 'revenue' && li.amount < 0)
  if (revenueNegative) issues.push({ severity: 'warning', message: 'A revenue line item has a negative amount.' })

  return { lineItems, issues }
}

export function assertNoBlockingIssues(result: ValidationResult) {
  const blocking = result.issues.find((issue) => issue.severity === 'error')
  if (blocking) throw new Error(blocking.message)
}
