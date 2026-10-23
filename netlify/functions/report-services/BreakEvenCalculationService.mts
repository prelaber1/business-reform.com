import { type ReportChart, type LineItem } from './types.mts'
import { fx, type MetricValue, refCell, layoutRawDataSheet, findBlock, addTotalRow, newCalcSheet, addSection, addMetric, type KeyMetric, type ReportModelBase } from './calculationSupport.mts'

export type BreakEvenModel = ReportModelBase & { kind: 'break-even' }

export function buildBreakEvenChart(revenue: number, fixed: number, variable: number): ReportChart {
  const base = revenue > 0 ? revenue : fixed + variable || 1
  const points = [0, 0.5, 1, 1.5, 2]
  const categories = points.map((p) => `${Math.round(p * 100)}%`)
  return {
    type: 'line',
    title: 'Revenue vs. Total Cost at Increasing Volume',
    categories,
    series: [
      { name: 'Revenue', values: points.map((p) => Math.round(base * p)) },
      { name: 'Total Cost', values: points.map((p) => Math.round(fixed + variable * p)) },
    ],
  }
}

export function calculateBreakEven(lineItems: LineItem[]): BreakEvenModel {
  const extraCaveats: string[] = []
  const rawSheet = layoutRawDataSheet('P&L', 'Reconstructed Income Statement', lineItems, [
    { tag: 'revenue', label: 'Revenue' },
    { tag: 'fixed_cost', label: 'Fixed Costs' },
    { tag: 'variable_cost', label: 'Variable Costs' },
  ])

  const revenueBlock = findBlock(rawSheet, 'revenue')
  const fixedBlock = findBlock(rawSheet, 'fixed_cost')
  const variableBlock = findBlock(rawSheet, 'variable_cost')
  if (!revenueBlock) extraCaveats.push('No line items were classified as revenue; break-even metrics below could not be fully computed.')
  if (!fixedBlock && !variableBlock) extraCaveats.push('No cost line items were classified as fixed or variable; break-even metrics below could not be fully computed.')

  const revenue = revenueBlock?.subtotalValue ?? 0
  const fixed = fixedBlock?.subtotalValue ?? 0
  const variable = variableBlock?.subtotalValue ?? 0

  const opexValue: MetricValue =
    fixedBlock || variableBlock
      ? fx(
          `=${fixedBlock ? `D${fixedBlock.subtotalRow}` : '0'}+${variableBlock ? `D${variableBlock.subtotalRow}` : '0'}`,
          fixed + variable,
        )
      : 'N/A (no cost data)'
  const opexRow = addTotalRow(rawSheet, 'Total Operating Expenses', opexValue)
  const netResultValue: MetricValue = revenueBlock
    ? fx(`=D${revenueBlock.subtotalRow}-D${opexRow}`, revenue - (fixed + variable))
    : 'N/A (no revenue data)'
  addTotalRow(rawSheet, 'Net Result', netResultValue)

  const calcSheet = newCalcSheet('Break-Even', 'Break-Even Calculation')
  const keyMetrics: KeyMetric[] = []

  addSection(calcSheet, 'Scenario A — Realistic (fixed vs. variable as classified)')
  const revRowA = addMetric(
    calcSheet,
    'Revenue',
    revenueBlock ? fx(refCell('P&L', revenueBlock.subtotalRow, 'D'), revenue) : 'N/A (no revenue data)',
  )
  const fixedRowA = addMetric(calcSheet, 'Fixed Costs', fixedBlock ? fx(refCell('P&L', fixedBlock.subtotalRow, 'D'), fixed) : fx('=0', 0))
  const varRowA = addMetric(
    calcSheet,
    'Variable Costs',
    variableBlock ? fx(refCell('P&L', variableBlock.subtotalRow, 'D'), variable) : fx('=0', 0),
  )

  if (revenueBlock && revenue !== 0) {
    const cm = revenue - variable
    const cmRow = addMetric(calcSheet, 'Contribution Margin (Revenue − Variable Costs)', fx(`=C${revRowA}-C${varRowA}`, cm))
    const cmRatio = cm / revenue
    const cmRatioRow = addMetric(calcSheet, 'Contribution Margin Ratio', fx(`=C${cmRow}/C${revRowA}`, cmRatio), { format: 'percent' })
    if (cmRatio !== 0) {
      const be = fixed / cmRatio
      const beRow = addMetric(calcSheet, 'Break-Even Revenue', fx(`=C${fixedRowA}/C${cmRatioRow}`, be), { bold: true, highlight: true })
      const mos = revenue - be
      const mosRow = addMetric(calcSheet, 'Margin of Safety (€)', fx(`=C${revRowA}-C${beRow}`, mos))
      addMetric(calcSheet, 'Margin of Safety (%)', fx(`=C${mosRow}/C${revRowA}`, mos / revenue), { format: 'percent' })
      keyMetrics.push({ label: 'Break-Even Revenue (Scenario A)', cellRef: refCell('Break-Even', beRow, 'C'), value: be, format: 'currency' })
      keyMetrics.push({
        label: 'Margin of Safety % (Scenario A)',
        cellRef: refCell('Break-Even', mosRow + 1, 'C'),
        value: mos / revenue,
        format: 'percent',
      })
    } else {
      addMetric(calcSheet, 'Break-Even Revenue', 'N/A (zero contribution margin)', { bold: true })
      extraCaveats.push('Contribution margin ratio is zero — break-even revenue could not be computed for Scenario A.')
    }
  } else {
    addMetric(calcSheet, 'Break-Even Revenue', 'N/A (no revenue data)', { bold: true })
  }

  addSection(calcSheet, 'Scenario B — Conservative (all operating costs treated as fixed)')
  const fixedRowB = addMetric(
    calcSheet,
    'Fixed Costs (= Total Operating Expenses)',
    fx(refCell('P&L', opexRow, 'D'), fixed + variable),
  )
  const beRowB = addMetric(calcSheet, 'Break-Even Revenue', fx(`=C${fixedRowB}`, fixed + variable), { bold: true, highlight: true })
  if (revenueBlock && revenue !== 0) {
    const mosBRow = addMetric(calcSheet, 'Margin of Safety (€)', fx(`=C${revRowA}-C${beRowB}`, revenue - (fixed + variable)))
    addMetric(calcSheet, 'Margin of Safety (%)', fx(`=C${mosBRow}/C${revRowA}`, (revenue - (fixed + variable)) / revenue), {
      format: 'percent',
    })
  } else {
    addMetric(calcSheet, 'Margin of Safety (€)', 'N/A (no revenue data)')
  }

  const headlineMetrics: KeyMetric[] = [
    { label: 'Total Operating Expenses', cellRef: refCell('P&L', opexRow, 'D'), value: fixed + variable, format: 'currency' },
  ]
  if (revenueBlock) {
    headlineMetrics.unshift({
      label: 'Total Revenue',
      cellRef: refCell('P&L', revenueBlock.subtotalRow, 'D'),
      value: revenue,
      format: 'currency',
    })
  }

  return {
    kind: 'break-even',
    rawSheets: [rawSheet],
    calcSheets: [calcSheet],
    keyMetrics: [...headlineMetrics, ...keyMetrics],
    charts: [buildBreakEvenChart(revenue, fixed, variable)],
    extraCaveats,
  }
}

// ---- Use of Funds calculator ----
