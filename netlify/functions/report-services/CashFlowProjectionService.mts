import ExcelJS from 'exceljs'
import { NAVY } from './constants.mts'
import { type ReportChart, type LineItem } from './types.mts'
import { THIN_BORDER, fx, type MetricValue, refCell, newCalcSheet, addSection, addMetric, type KeyMetric, type ReportModelBase, PLATFORM_ASSUMPTIONS } from './calculationSupport.mts'
import { periodSortKey, layoutBurnRateSheet } from './BurnRateCalculationService.mts'
import { sanitizeSheetName } from './excelSheetHelpers.mts'
import { setMetricCell } from './ExcelReportBuilder.mts'

export type CashFlowForecastModel = ReportModelBase & { kind: 'cash-flow-forecast' }

export const CF_WEEK_HORIZON = 13
export const CF_MONTH_HORIZON = 12
export const CF_GRID_ROWS = { header: 4, beginning: 5, inflow: 6, outflow: 7, net: 8, ending: 9 }

export function cfColLetter(indexFromB: number): string {
  // indexFromB is 0-based; 0 -> B, 1 -> C, ... (stays within single letters for our horizons)
  return String.fromCharCode('B'.charCodeAt(0) + indexFromB)
}

export function averageOf(values: number[]): number {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0
}

export function classifyPeriodGranularity(period: string): 'week' | 'month' | 'unlabeled' {
  if (/week/i.test(period)) return 'week'
  if (/month/i.test(period)) return 'month'
  return 'unlabeled'
}

export function buildCashFlowGridSheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  title: string,
  periodLabelPrefix: string,
  horizon: number,
  startingBalance: MetricValue,
  startingBalanceValue: number,
  periodInflow: (index: number) => MetricValue,
  periodInflowValue: (index: number) => number,
  periodOutflow: (index: number) => MetricValue,
  periodOutflowValue: (index: number) => number,
) {
  const R = CF_GRID_ROWS
  const sheet = workbook.addWorksheet(sanitizeSheetName(title, usedSheetNames))
  sheet.getColumn(1).width = 22
  for (let i = 0; i < horizon; i++) sheet.getColumn(2 + i).width = 15

  sheet.getCell('A1').value = title
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells(`A1:${cfColLetter(horizon - 1)}1`)

  sheet.getCell(`A${R.header}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
  sheet.getCell(`A${R.header}`).border = THIN_BORDER
  for (let i = 0; i < horizon; i++) {
    const col = cfColLetter(i)
    const cell = sheet.getCell(`${col}${R.header}`)
    cell.value = `${periodLabelPrefix} ${i + 1}`
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.alignment = { horizontal: 'center' }
    cell.border = THIN_BORDER
  }
  sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: R.header }]

  const labelSpecs: [number, string, boolean?][] = [
    [R.beginning, 'Beginning Balance'],
    [R.inflow, 'Cash Inflows'],
    [R.outflow, 'Cash Outflows'],
    [R.net, 'Net Cash Flow', true],
    [R.ending, 'Ending Balance', true],
  ]
  for (const [row, label, bold] of labelSpecs) {
    const cell = sheet.getCell(`A${row}`)
    cell.value = label
    cell.font = { bold: !!bold }
  }

  function setCell(row: number, col: string, value: MetricValue) {
    const cell = sheet.getCell(`${col}${row}`)
    setMetricCell(cell, value, 'currency')
    cell.border = THIN_BORDER
  }

  let runningBalance = startingBalanceValue
  const endingSeries: number[] = []
  for (let i = 0; i < horizon; i++) {
    const col = cfColLetter(i)
    const prevCol = i > 0 ? cfColLetter(i - 1) : null

    setCell(R.beginning, col, i === 0 ? startingBalance : fx(`=${prevCol}${R.ending}`, runningBalance))

    const inflowVal = periodInflowValue(i)
    setCell(R.inflow, col, periodInflow(i))
    const outflowVal = periodOutflowValue(i)
    setCell(R.outflow, col, periodOutflow(i))

    const netVal = inflowVal - outflowVal
    setCell(R.net, col, fx(`=${col}${R.inflow}-${col}${R.outflow}`, netVal))

    const beginningVal = runningBalance
    runningBalance = beginningVal + netVal
    setCell(R.ending, col, fx(`=${col}${R.beginning}+${col}${R.net}`, runningBalance))

    endingSeries.push(runningBalance)
  }

  return { sheetName: sheet.name, endingSeries }
}

export function calculateCashFlowForecast(lineItems: LineItem[]): CashFlowForecastModel {
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
  } = layoutBurnRateSheet(lineItems)

  if (startingBalanceRow === null) extraCaveats.push('No starting cash balance was identified; the forecast below assumes a starting balance of zero.')

  const weekPeriods = periods.filter((p) => classifyPeriodGranularity(p) === 'week').sort((a, b) => periodSortKey(a) - periodSortKey(b))
  const monthPeriods = periods.filter((p) => classifyPeriodGranularity(p) === 'month').sort((a, b) => periodSortKey(a) - periodSortKey(b))
  const unlabeledPeriods = periods.filter((p) => classifyPeriodGranularity(p) === 'unlabeled')

  const { weeksPerMonth } = PLATFORM_ASSUMPTIONS
  const rawSheetName = rawSheet.sheetName

  const assumptionsSheet = newCalcSheet('Run-Rate Assumptions', 'Baseline Run-Rate Assumptions')
  addSection(assumptionsSheet, 'Recurring Cash Flow Run-Rate')
  const weeksPerMonthRow = addMetric(assumptionsSheet, 'Weeks per Month (conversion factor, editable)', weeksPerMonth)

  let weeklyInflowRow: number, weeklyOutflowRow: number, monthlyInflowRow: number, monthlyOutflowRow: number
  let weeklyInflowValue: number, weeklyOutflowValue: number, monthlyInflowValue: number, monthlyOutflowValue: number

  if (monthPeriods.length) {
    const inflowRefs = monthPeriods.filter((p) => inflowRowByPeriod[p]).map((p) => refCell(rawSheetName, inflowRowByPeriod[p], 'D'))
    monthlyInflowValue = averageOf(monthPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    monthlyInflowRow = addMetric(assumptionsSheet, 'Monthly Inflow Run-Rate', inflowRefs.length ? fx(`=AVERAGE(${inflowRefs.join(',')})`, monthlyInflowValue) : 'N/A (no inflow data)')
    const outflowRefs = monthPeriods.filter((p) => outflowRowByPeriod[p]).map((p) => refCell(rawSheetName, outflowRowByPeriod[p], 'D'))
    monthlyOutflowValue = averageOf(monthPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    monthlyOutflowRow = addMetric(assumptionsSheet, 'Monthly Outflow Run-Rate', outflowRefs.length ? fx(`=AVERAGE(${outflowRefs.join(',')})`, monthlyOutflowValue) : 'N/A (no outflow data)')
  } else if (weekPeriods.length) {
    const weeklyInflowForConversion = averageOf(weekPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    const weeklyOutflowForConversion = averageOf(weekPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    monthlyInflowValue = weeklyInflowForConversion * weeksPerMonth
    monthlyInflowRow = addMetric(assumptionsSheet, 'Monthly Inflow Run-Rate (derived from weekly data)', monthlyInflowValue)
    monthlyOutflowValue = weeklyOutflowForConversion * weeksPerMonth
    monthlyOutflowRow = addMetric(assumptionsSheet, 'Monthly Outflow Run-Rate (derived from weekly data)', monthlyOutflowValue)
    extraCaveats.push('Source data was only weekly-dated — the 12-month forecast derives its monthly run-rate from the weekly figures using a fixed 4.33 weeks/month conversion factor.')
  } else {
    monthlyInflowValue = averageOf(unlabeledPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    monthlyInflowRow = addMetric(assumptionsSheet, 'Monthly Inflow Run-Rate (assumed from undated total)', monthlyInflowValue)
    monthlyOutflowValue = averageOf(unlabeledPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    monthlyOutflowRow = addMetric(assumptionsSheet, 'Monthly Outflow Run-Rate (assumed from undated total)', monthlyOutflowValue)
    if (unlabeledPeriods.length) extraCaveats.push('Source data did not distinguish individual weeks or months — the undated total is treated as a single representative month, and both horizons below are extrapolated flat from it.')
  }

  if (weekPeriods.length) {
    const inflowRefs = weekPeriods.filter((p) => inflowRowByPeriod[p]).map((p) => refCell(rawSheetName, inflowRowByPeriod[p], 'D'))
    weeklyInflowValue = averageOf(weekPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    weeklyInflowRow = addMetric(assumptionsSheet, 'Weekly Inflow Run-Rate', inflowRefs.length ? fx(`=AVERAGE(${inflowRefs.join(',')})`, weeklyInflowValue) : 'N/A (no inflow data)')
    const outflowRefs = weekPeriods.filter((p) => outflowRowByPeriod[p]).map((p) => refCell(rawSheetName, outflowRowByPeriod[p], 'D'))
    weeklyOutflowValue = averageOf(weekPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    weeklyOutflowRow = addMetric(assumptionsSheet, 'Weekly Outflow Run-Rate', outflowRefs.length ? fx(`=AVERAGE(${outflowRefs.join(',')})`, weeklyOutflowValue) : 'N/A (no outflow data)')
  } else {
    weeklyInflowValue = monthlyInflowValue / weeksPerMonth
    weeklyInflowRow = addMetric(assumptionsSheet, 'Weekly Inflow Run-Rate (derived from monthly data)', fx(`=C${monthlyInflowRow}/C${weeksPerMonthRow}`, weeklyInflowValue))
    weeklyOutflowValue = monthlyOutflowValue / weeksPerMonth
    weeklyOutflowRow = addMetric(assumptionsSheet, 'Weekly Outflow Run-Rate (derived from monthly data)', fx(`=C${monthlyOutflowRow}/C${weeksPerMonthRow}`, weeklyOutflowValue))
    if (monthPeriods.length) extraCaveats.push('Source data was only monthly-dated — the 13-week forecast derives its weekly run-rate from the monthly figures using a fixed 4.33 weeks/month conversion factor.')
  }

  const assumptionsSheetName = assumptionsSheet.sheetName

  const startingBalanceRef: MetricValue =
    startingBalanceRow !== null ? fx(refCell(rawSheetName, startingBalanceRow, 'D'), startingBalanceValue) : fx('=0', 0)

  function makeAccessors(
    periodsForHorizon: string[],
    rowByPeriod: Record<string, number>,
    valueByPeriod: Record<string, number>,
    flatRunRateRow: number,
    flatRunRateValue: number,
    labelPrefix: string,
  ) {
    return {
      value: (index: number) => valueByPeriod[`${labelPrefix} ${index + 1}`] ?? flatRunRateValue,
      cell: (index: number): MetricValue => {
        const label = `${labelPrefix} ${index + 1}`
        if (rowByPeriod[label]) return fx(refCell(rawSheetName, rowByPeriod[label], 'D'), valueByPeriod[label])
        return fx(refCell(assumptionsSheetName, flatRunRateRow, 'C'), flatRunRateValue)
      },
    }
  }

  const weekInflowAccessor = makeAccessors(weekPeriods, inflowRowByPeriod, inflowValueByPeriod, weeklyInflowRow, weeklyInflowValue, 'Week')
  const weekOutflowAccessor = makeAccessors(weekPeriods, outflowRowByPeriod, outflowValueByPeriod, weeklyOutflowRow, weeklyOutflowValue, 'Week')
  const monthInflowAccessor = makeAccessors(monthPeriods, inflowRowByPeriod, inflowValueByPeriod, monthlyInflowRow, monthlyInflowValue, 'Month')
  const monthOutflowAccessor = makeAccessors(monthPeriods, outflowRowByPeriod, outflowValueByPeriod, monthlyOutflowRow, monthlyOutflowValue, 'Month')

  const customSheetBuilders = [
    (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildCashFlowGridSheet(
        workbook,
        usedSheetNames,
        '13-Week Cash Flow Forecast',
        'Week',
        CF_WEEK_HORIZON,
        startingBalanceRef,
        startingBalanceValue,
        weekInflowAccessor.cell,
        weekInflowAccessor.value,
        weekOutflowAccessor.cell,
        weekOutflowAccessor.value,
      ),
    (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildCashFlowGridSheet(
        workbook,
        usedSheetNames,
        '12-Month Cash Flow Forecast',
        'Month',
        CF_MONTH_HORIZON,
        startingBalanceRef,
        startingBalanceValue,
        monthInflowAccessor.cell,
        monthInflowAccessor.value,
        monthOutflowAccessor.cell,
        monthOutflowAccessor.value,
      ),
  ]

  // Precompute the same ending-balance series numerically (independent of the
  // deferred sheet builders above) so key metrics and charts are available
  // immediately, mirroring the Financial Projections calculator's approach.
  function projectEndingSeries(horizon: number, inflowValue: (i: number) => number, outflowValue: (i: number) => number) {
    let balance = startingBalanceValue
    const series: number[] = []
    for (let i = 0; i < horizon; i++) {
      balance = balance + inflowValue(i) - outflowValue(i)
      series.push(balance)
    }
    return series
  }

  const weekSeries = projectEndingSeries(CF_WEEK_HORIZON, weekInflowAccessor.value, weekOutflowAccessor.value)
  const monthSeries = projectEndingSeries(CF_MONTH_HORIZON, monthInflowAccessor.value, monthOutflowAccessor.value)

  const minWeekBalance = Math.min(...weekSeries)
  const minWeekIndex = weekSeries.indexOf(minWeekBalance)
  const minMonthBalance = Math.min(...monthSeries)
  const minMonthIndex = monthSeries.indexOf(minMonthBalance)

  const keyMetrics: KeyMetric[] = [
    {
      label: 'Starting Cash Balance',
      cellRef: refCell('13-Week Cash Flow Forecast', CF_GRID_ROWS.beginning, cfColLetter(0)),
      value: startingBalanceValue,
      format: 'currency',
    },
    {
      label: `Lowest Projected Balance — 13-Week (Week ${minWeekIndex + 1})`,
      cellRef: refCell('13-Week Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(minWeekIndex)),
      value: minWeekBalance,
      format: 'currency',
    },
    {
      label: 'Ending Balance — Week 13',
      cellRef: refCell('13-Week Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(CF_WEEK_HORIZON - 1)),
      value: weekSeries[CF_WEEK_HORIZON - 1],
      format: 'currency',
    },
    {
      label: `Lowest Projected Balance — 12-Month (Month ${minMonthIndex + 1})`,
      cellRef: refCell('12-Month Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(minMonthIndex)),
      value: minMonthBalance,
      format: 'currency',
    },
    {
      label: 'Ending Balance — Month 12',
      cellRef: refCell('12-Month Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(CF_MONTH_HORIZON - 1)),
      value: monthSeries[CF_MONTH_HORIZON - 1],
      format: 'currency',
    },
  ]

  if (minWeekBalance < 0 || minMonthBalance < 0) {
    extraCaveats.push('The projected cash balance goes negative within the forecast horizon under the current run-rate assumptions — treat this as a liquidity risk flag, not a certainty, since it is driven by a flat extrapolated run-rate wherever period-specific data was not available.')
  }

  const charts: ReportChart[] = [
    {
      type: 'line',
      title: '13-Week Rolling Cash Balance',
      categories: Array.from({ length: CF_WEEK_HORIZON }, (_, i) => `Week ${i + 1}`),
      series: [{ name: 'Ending Balance', values: weekSeries.map((v) => Math.round(v)) }],
    },
    {
      type: 'line',
      title: '12-Month Rolling Cash Balance',
      categories: Array.from({ length: CF_MONTH_HORIZON }, (_, i) => `Month ${i + 1}`),
      series: [{ name: 'Ending Balance', values: monthSeries.map((v) => Math.round(v)) }],
    },
  ]

  return {
    kind: 'cash-flow-forecast',
    rawSheets: [rawSheet],
    calcSheets: [assumptionsSheet],
    keyMetrics,
    charts,
    extraCaveats,
    customSheetBuilders,
  }
}

// ---- Capital Optimization calculator ----
