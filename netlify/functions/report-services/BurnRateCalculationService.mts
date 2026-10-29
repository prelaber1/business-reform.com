import { type ReportChart, type LineItem } from './types.mts'
import { fx, sumGroup, groupItems, refCell, type RawDataRow, type RawDataBlock, type RawDataSheetModel, RAW_DATA_FIRST_DATA_ROW, newCalcSheet, addSection, addMetric, type KeyMetric, type ReportModelBase } from './calculationSupport.mts'

export type BurnRateModel = ReportModelBase & { kind: 'burn-rate' }

export function periodSortKey(period: string | undefined): number {
  const match = period?.match(/(\d+)/)
  return match ? parseInt(match[1], 10) : 0
}

export function layoutBurnRateSheet(lineItems: LineItem[]) {
  const inflowItems = groupItems(lineItems, 'cash_inflow')
  const outflowItems = groupItems(lineItems, 'cash_outflow')
  const startingItems = groupItems(lineItems, 'starting_cash_balance')

  const periodsSet = new Set<string>()
  for (const item of [...inflowItems, ...outflowItems]) periodsSet.add(item.period || 'All Periods')
  const periods = Array.from(periodsSet).sort((a, b) => periodSortKey(a) - periodSortKey(b))

  const rows: RawDataRow[] = []
  const blocks: RawDataBlock[] = []
  let row = RAW_DATA_FIRST_DATA_ROW
  let startingBalanceRow: number | null = null
  const startingBalanceValue = sumGroup(startingItems)

  if (startingItems.length) {
    const item = startingItems[0]
    rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: 'Starting Cash Balance', amount: item.amount })
    startingBalanceRow = row
    row += 1
  }

  const inflowRowByPeriod: Record<string, number> = {}
  const outflowRowByPeriod: Record<string, number> = {}
  const inflowValueByPeriod: Record<string, number> = {}
  const outflowValueByPeriod: Record<string, number> = {}

  for (const period of periods) {
    const periodInflow = inflowItems.filter((i) => (i.period || 'All Periods') === period)
    const periodOutflow = outflowItems.filter((i) => (i.period || 'All Periods') === period)

    if (periodInflow.length) {
      const start = row
      for (const item of periodInflow) {
        rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: `Inflow — ${period}`, amount: item.amount })
        row += 1
      }
      const subtotalValue = sumGroup(periodInflow)
      rows.push({ row, kind: 'subtotal', label: `Total Inflow — ${period}`, value: fx(`=SUM(D${start}:D${row - 1})`, subtotalValue) })
      inflowRowByPeriod[period] = row
      inflowValueByPeriod[period] = subtotalValue
      blocks.push({ tag: `cash_inflow:${period}`, label: `Inflow — ${period}`, subtotalRow: row, subtotalValue })
      row += 1
    }
    if (periodOutflow.length) {
      const start = row
      for (const item of periodOutflow) {
        rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: `Outflow — ${period}`, amount: item.amount })
        row += 1
      }
      const subtotalValue = sumGroup(periodOutflow)
      rows.push({ row, kind: 'subtotal', label: `Total Outflow — ${period}`, value: fx(`=SUM(D${start}:D${row - 1})`, subtotalValue) })
      outflowRowByPeriod[period] = row
      outflowValueByPeriod[period] = subtotalValue
      blocks.push({ tag: `cash_outflow:${period}`, label: `Outflow — ${period}`, subtotalRow: row, subtotalValue })
      row += 1
    }
  }

  const model: RawDataSheetModel = { sheetName: 'Cash Flow Detail', title: 'Cash Flow Detail', rows, blocks, lastRow: row - 1 }
  return { model, periods, inflowRowByPeriod, outflowRowByPeriod, inflowValueByPeriod, outflowValueByPeriod, startingBalanceRow, startingBalanceValue, outflowItems }
}

export function calculateBurnRate(lineItems: LineItem[]): BurnRateModel {
  const extraCaveats: string[] = []
  const {
    model: rawSheet,
    periods,
    inflowRowByPeriod,
    outflowRowByPeriod,
    inflowValueByPeriod,
    outflowValueByPeriod,
    startingBalanceRow,
    startingBalanceValue,
    outflowItems,
  } = layoutBurnRateSheet(lineItems)

  if (!periods.length) extraCaveats.push('No cash inflow/outflow line items were classified; burn rate metrics below could not be computed.')
  if (periods.length === 1 && periods[0] === 'All Periods') {
    extraCaveats.push(
      'Source data did not distinguish individual months — the figures below treat the classified data as a single representative period rather than a true 12-month average.',
    )
  }
  if (startingBalanceRow === null) extraCaveats.push('No starting cash balance was identified; runway could not be computed.')

  const calcSheet = newCalcSheet('Burn Analysis', 'Burn Rate & Runway Analysis')
  const keyMetrics: KeyMetric[] = []

  addSection(calcSheet, 'Burn Rate & Runway')
  const outflowRefs = periods.filter((p) => outflowRowByPeriod[p]).map((p) => refCell('Cash Flow Detail', outflowRowByPeriod[p], 'D'))
  const outflowValues = periods.filter((p) => outflowRowByPeriod[p]).map((p) => outflowValueByPeriod[p])
  const inflowRefs = periods.filter((p) => inflowRowByPeriod[p]).map((p) => refCell('Cash Flow Detail', inflowRowByPeriod[p], 'D'))
  const inflowValues = periods.filter((p) => inflowRowByPeriod[p]).map((p) => inflowValueByPeriod[p])

  const grossBurnValue = outflowValues.length ? outflowValues.reduce((a, b) => a + b, 0) / outflowValues.length : 0
  const grossBurnRow = addMetric(
    calcSheet,
    'Gross Burn Rate (monthly operating expenses, average per period)',
    outflowRefs.length ? fx(`=AVERAGE(${outflowRefs.join(',')})`, grossBurnValue) : 'N/A (no outflow data)',
  )
  const grossInflowValue = inflowValues.length ? inflowValues.reduce((a, b) => a + b, 0) / inflowValues.length : 0
  const grossInflowRow = addMetric(
    calcSheet,
    'Gross Inflow (average per period)',
    inflowRefs.length ? fx(`=AVERAGE(${inflowRefs.join(',')})`, grossInflowValue) : 'N/A (no inflow data)',
  )
  const netBurnValue = grossBurnValue - grossInflowValue
  const netBurnRow = addMetric(calcSheet, 'Net Burn (average per period)', fx(`=C${grossBurnRow}-C${grossInflowRow}`, netBurnValue), {
    bold: true,
  })

  let currentBalanceRow: number | null = null
  if (startingBalanceRow !== null) {
    currentBalanceRow = addMetric(
      calcSheet,
      'Current Cash Balance',
      fx(refCell('Cash Flow Detail', startingBalanceRow, 'D'), startingBalanceValue),
    )
    keyMetrics.push({ label: 'Current Cash Balance', cellRef: refCell('Burn Analysis', currentBalanceRow, 'C'), value: startingBalanceValue, format: 'currency' })
  }

  if (currentBalanceRow !== null && netBurnValue > 0) {
    const runway = startingBalanceValue / netBurnValue
    const runwayRow = addMetric(calcSheet, 'Runway (periods)', fx(`=C${currentBalanceRow}/C${netBurnRow}`, runway), {
      bold: true,
      highlight: true,
      format: 'number',
    })
    keyMetrics.push({ label: 'Runway (periods)', cellRef: refCell('Burn Analysis', runwayRow, 'C'), value: runway, format: 'number' })
  } else if (currentBalanceRow !== null) {
    addMetric(calcSheet, 'Runway (periods)', 'N/A — cash flow break-even or positive', { bold: true, highlight: true })
  } else {
    addMetric(calcSheet, 'Runway (periods)', 'N/A (no starting cash balance)', { bold: true })
  }

  // Gross runway = company cash / monthly operating expenses: how long cash
  // lasts if no cash comes in at all.
  if (currentBalanceRow !== null && grossBurnValue > 0) {
    const grossRunway = startingBalanceValue / grossBurnValue
    const grossRunwayRow = addMetric(calcSheet, 'Gross Runway (cash ÷ monthly operating expenses, periods)', fx(`=C${currentBalanceRow}/C${grossBurnRow}`, grossRunway), {
      format: 'number',
    })
    keyMetrics.push({ label: 'Gross Runway (periods, no inflows)', cellRef: refCell('Burn Analysis', grossRunwayRow, 'C'), value: grossRunway, format: 'number' })
  } else {
    addMetric(calcSheet, 'Gross Runway (cash ÷ monthly operating expenses, periods)', currentBalanceRow === null ? 'N/A (no starting cash balance)' : 'N/A (no outflow data)')
  }

  keyMetrics.push({ label: 'Gross Burn Rate (average per period)', cellRef: refCell('Burn Analysis', grossBurnRow, 'C'), value: grossBurnValue, format: 'currency' })
  keyMetrics.push({ label: 'Net Burn (average per period)', cellRef: refCell('Burn Analysis', netBurnRow, 'C'), value: netBurnValue, format: 'currency' })

  addSection(calcSheet, 'Period-by-Period Cash Position')
  let prevRef: string | null = startingBalanceRow !== null ? refCell('Cash Flow Detail', startingBalanceRow, 'D') : null
  let runningBalance = startingBalanceValue
  const balanceSeries: number[] = []
  for (const period of periods) {
    const inflowVal = inflowValueByPeriod[period] ?? 0
    const outflowVal = outflowValueByPeriod[period] ?? 0
    runningBalance = runningBalance + inflowVal - outflowVal
    const inflowRef = inflowRowByPeriod[period] ? refCell('Cash Flow Detail', inflowRowByPeriod[period], 'D') : null
    const outflowRef = outflowRowByPeriod[period] ? refCell('Cash Flow Detail', outflowRowByPeriod[period], 'D') : null
    let formulaStr = prevRef ? `=${prevRef}` : '=0'
    if (inflowRef) formulaStr += `+${inflowRef}`
    if (outflowRef) formulaStr += `-${outflowRef}`
    const balRow = addMetric(calcSheet, `Cash Balance — ${period}`, fx(formulaStr, runningBalance))
    prevRef = `C${balRow}`
    balanceSeries.push(runningBalance)
  }

  const categoryTotals = new Map<string, number>()
  for (const item of outflowItems) categoryTotals.set(item.label, (categoryTotals.get(item.label) ?? 0) + item.amount)
  const topCategories = [...categoryTotals.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8)

  const charts: ReportChart[] = []
  if (periods.length) {
    charts.push({
      type: 'line',
      title: 'Period-by-Period Cash Balance',
      categories: periods,
      series: [{ name: 'Cash Balance', values: balanceSeries }],
    })
  }
  if (topCategories.length) {
    charts.push({
      type: 'bar',
      title: 'Largest Cash Outflow Categories',
      categories: topCategories.map(([label]) => label),
      series: [{ name: 'Total Outflow', values: topCategories.map(([, value]) => value) }],
    })
  }

  return { kind: 'burn-rate', rawSheets: [rawSheet], calcSheets: [calcSheet], keyMetrics, charts, extraCaveats }
}

// ---- Cash Flow Forecast calculator (13-week / 12-month rolling) ----
//
// Reuses the burn-rate raw sheet layout (same "cash_inflow"/"cash_outflow"/
// "starting_cash_balance" group vocabulary, period-dated), then rolls the
// balance forward across two fixed horizons. Whichever granularity the
// source data actually supports (weekly or monthly periods) is used as the
// baseline run-rate; the other horizon is derived from it via the
// platform-controlled weeks-per-month conversion factor, exactly like the
// Financial Projections calculator derives Years 2-5 from Year-1 data.
