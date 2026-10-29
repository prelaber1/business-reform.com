// ---- FP&A Digest calculator ----
//
// Compares each metric's current-period figure against its trailing history
// and against the same period a year ago (the weekly-fpa-digest rubric). The
// AI only classifies atomic figures into three groups; every average,
// variance and flag below is computed here and rendered as live formulas.
import { type ReportChart, type LineItem } from './types.mts'
import {
  fx, groupItems, refCell, layoutRawDataSheet, newCalcSheet, addSection, addMetric,
  PLATFORM_ASSUMPTIONS, type RawDataItemRow, type KeyMetric, type ReportModelBase, type MetricValue,
} from './calculationSupport.mts'

export type FpaDigestModel = ReportModelBase & { kind: 'fpa-digest' }

const CURRENT_SHEET = 'Current Period'
const TRAILING_SHEET = 'Trailing Periods'
const PRIOR_YEAR_SHEET = 'Prior Year Period'
const MAX_DIGEST_METRICS = 12 // keeps the digest short and the narrative prompt small

const metricKey = (label: string) => label.trim().toLowerCase()

// Trailing items are listed with their period in the description so the raw
// sheet shows which periods feed each average.
const withPeriodInLabel = (items: LineItem[]): LineItem[] =>
  items.map((i) => ({ ...i, label: i.period ? `${i.label} (${i.period})` : i.label }))

type SheetRef = { refs: string[]; values: number[] }
const emptyRef = (): SheetRef => ({ refs: [], values: [] })

function collectRefs(rows: RawDataItemRow[], sheetName: string, items: LineItem[]): Map<string, SheetRef> {
  const out = new Map<string, SheetRef>()
  const shown = withPeriodInLabel(items)
  const used = new Set<number>()
  items.forEach((item, idx) => {
    const match = rows.find((r) => !used.has(r.row) && r.description === shown[idx].label && r.amount === item.amount)
    if (!match) return
    used.add(match.row)
    const entry = out.get(metricKey(item.label)) ?? emptyRef()
    entry.refs.push(refCell(sheetName, match.row, 'D'))
    entry.values.push(item.amount)
    out.set(metricKey(item.label), entry)
  })
  return out
}

const mean = (values: number[]) => values.reduce((a, b) => a + b, 0) / values.length

export function calculateFpaDigest(lineItems: LineItem[]): FpaDigestModel {
  const extraCaveats: string[] = []
  const threshold = PLATFORM_ASSUMPTIONS.digestFlagThreshold

  const current = groupItems(lineItems, 'current_period')
  const trailing = groupItems(lineItems, 'trailing_period')
  const priorYear = groupItems(lineItems, 'prior_year_period')

  // Raw sheets keep one block per source group; trailing rows carry their period.
  const currentSheet = layoutRawDataSheet(CURRENT_SHEET, 'Current Period Figures', current, [{ tag: 'current_period', label: 'Current Period' }])
  const trailingSheet = layoutRawDataSheet(TRAILING_SHEET, 'Trailing Period Figures', withPeriodInLabel(trailing).map((i) => ({ ...i, group: 'trailing_period' })), [
    { tag: 'trailing_period', label: 'Trailing Period' },
  ])
  const priorYearSheet = layoutRawDataSheet(PRIOR_YEAR_SHEET, 'Same Period Last Year Figures', priorYear, [{ tag: 'prior_year_period', label: 'Prior Year Period' }])
  const rawSheets = [currentSheet, trailingSheet, priorYearSheet].filter((s) => s.blocks.length > 0)

  const itemRows = (sheet: typeof currentSheet) => sheet.rows.filter((r): r is RawDataItemRow => r.kind === 'item')
  const currentRefs = collectRefs(itemRows(currentSheet), CURRENT_SHEET, current)
  const trailingRefs = collectRefs(itemRows(trailingSheet), TRAILING_SHEET, trailing)
  const priorRefs = collectRefs(itemRows(priorYearSheet), PRIOR_YEAR_SHEET, priorYear)

  if (!current.length) extraCaveats.push('No current-period figures were classified; the digest could not be computed.')
  if (current.length && !trailing.length) extraCaveats.push('No trailing-period history was found; variances against recent trend could not be computed.')
  if (current.length && !priorYear.length) extraCaveats.push('No same-period-last-year figures were found; year-over-year variances could not be computed.')

  // Metric names come from the current period, in source order.
  const metrics: { key: string; label: string }[] = []
  for (const item of current) {
    const key = metricKey(item.label)
    if (!metrics.some((m) => m.key === key)) metrics.push({ key, label: item.label.trim() })
  }
  if (metrics.length > MAX_DIGEST_METRICS) {
    extraCaveats.push(`${metrics.length} metrics were found; the digest covers the first ${MAX_DIGEST_METRICS} to stay concise.`)
    metrics.length = MAX_DIGEST_METRICS
  }
  for (const key of new Set([...trailingRefs.keys(), ...priorRefs.keys()])) {
    if (!metrics.some((m) => m.key === key)) extraCaveats.push(`"${key}" has history but no current-period figure and was left out of the digest.`)
  }

  const calcSheet = newCalcSheet('Digest Analysis', 'FP&A Digest: Variance Analysis')
  const keyMetrics: KeyMetric[] = []
  const flagged: string[] = []
  const chartRows: { label: string; vsTrailing: number | null; vsPriorYear: number | null }[] = []

  const pct = (change: number, base: number) => change / Math.abs(base)
  const na = (why: string): MetricValue => `N/A (${why})`

  for (const { key, label } of metrics) {
    const cur = currentRefs.get(key)
    const trail = trailingRefs.get(key)
    const prior = priorRefs.get(key)
    if (!cur) continue
    if (cur.refs.length > 1) extraCaveats.push(`"${label}" has ${cur.refs.length} current-period figures; they are summed.`)

    addSection(calcSheet, label)
    const curValue = cur.values.reduce((a, b) => a + b, 0)
    const curRow = addMetric(calcSheet, 'Current Period', fx(`=${cur.refs.join('+')}`, curValue), { bold: true })
    keyMetrics.push({ label: `${label}: current period`, cellRef: refCell('Digest Analysis', curRow, 'C'), value: curValue, format: 'currency' })

    let vsTrailing: number | null = null
    if (trail) {
      const avg = mean(trail.values)
      const avgRow = addMetric(calcSheet, `Trailing Average (${trail.values.length} periods)`, fx(`=AVERAGE(${trail.refs.join(',')})`, avg))
      const chgRow = addMetric(calcSheet, 'Change vs Trailing Average', fx(`=C${curRow}-C${avgRow}`, curValue - avg))
      if (avg !== 0) {
        vsTrailing = pct(curValue - avg, avg)
        const pctRow = addMetric(calcSheet, '% Change vs Trailing Average', fx(`=C${chgRow}/ABS(C${avgRow})`, vsTrailing), { format: 'percent', bold: true })
        keyMetrics.push({ label: `${label}: % change vs trailing average`, cellRef: refCell('Digest Analysis', pctRow, 'C'), value: vsTrailing, format: 'percent' })
      } else {
        addMetric(calcSheet, '% Change vs Trailing Average', na('zero trailing average'))
      }
      if (trail.values.length < 3) extraCaveats.push(`"${label}" has only ${trail.values.length} trailing period(s); the trend baseline is thin.`)
    } else {
      addMetric(calcSheet, 'Trailing Average', na('no trailing history'))
    }

    let vsPrior: number | null = null
    if (prior) {
      const priorValue = prior.values.reduce((a, b) => a + b, 0)
      const priorRow = addMetric(calcSheet, 'Same Period Last Year', fx(`=${prior.refs.join('+')}`, priorValue))
      const chgRow = addMetric(calcSheet, 'Change vs Last Year', fx(`=C${curRow}-C${priorRow}`, curValue - priorValue))
      if (priorValue !== 0) {
        vsPrior = pct(curValue - priorValue, priorValue)
        const pctRow = addMetric(calcSheet, '% Change vs Last Year', fx(`=C${chgRow}/ABS(C${priorRow})`, vsPrior), { format: 'percent', bold: true })
        keyMetrics.push({ label: `${label}: % change vs same period last year`, cellRef: refCell('Digest Analysis', pctRow, 'C'), value: vsPrior, format: 'percent' })
      } else {
        addMetric(calcSheet, '% Change vs Last Year', na('zero prior-year figure'))
      }
    } else {
      addMetric(calcSheet, 'Same Period Last Year', na('no prior-year figure'))
    }

    const moves = [vsTrailing, vsPrior].filter((v): v is number => v !== null)
    const isFlag = moves.some((v) => Math.abs(v) >= threshold)
    if (isFlag) flagged.push(label)
    addMetric(
      calcSheet,
      `Flag (move of ${Math.round(threshold * 100)}% or more)`,
      isFlag ? 'REVIEW' : moves.length ? 'In range' : 'N/A (no comparison)',
      { highlight: isFlag },
    )
    chartRows.push({ label, vsTrailing, vsPriorYear: vsPrior })
  }

  // The count comes from the flags above, so it is a plain number rather than a formula.
  addSection(calcSheet, 'Summary')
  const flagRow = addMetric(calcSheet, 'Metrics flagged for review', flagged.length, { format: 'number', bold: true, highlight: true })
  keyMetrics.unshift({
    label: `Metrics flagged for review (moved ${Math.round(threshold * 100)}% or more)`,
    cellRef: refCell('Digest Analysis', flagRow, 'C'),
    value: flagged.length,
    format: 'number',
  })

  const charts: ReportChart[] = []
  const comparable = chartRows.filter((r) => r.vsTrailing !== null || r.vsPriorYear !== null)
  if (comparable.length) {
    charts.push({
      type: 'bar',
      title: 'Change in Each Metric (% vs Trailing Average and vs Last Year)',
      categories: comparable.map((r) => r.label),
      series: [
        { name: 'vs Trailing Average (%)', values: comparable.map((r) => Math.round((r.vsTrailing ?? 0) * 1000) / 10) },
        { name: 'vs Last Year (%)', values: comparable.map((r) => Math.round((r.vsPriorYear ?? 0) * 1000) / 10) },
      ],
    })
  }

  return { kind: 'fpa-digest', rawSheets, calcSheets: [calcSheet], keyMetrics, charts, extraCaveats }
}
