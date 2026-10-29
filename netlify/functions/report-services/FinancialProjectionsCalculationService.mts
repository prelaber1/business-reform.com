import ExcelJS from 'exceljs'
import { NAVY } from './constants.mts'
import { type ReportChart, type LineItem } from './types.mts'
import type { GrowthAssumptions } from './GrowthAssumptionService.mts'
import { THIN_BORDER, type MetricFormat, fx, type MetricValue, refCell, layoutRawDataSheet, findBlock, type KeyMetric, type ReportModelBase, PLATFORM_ASSUMPTIONS } from './calculationSupport.mts'
import { sanitizeSheetName } from './excelSheetHelpers.mts'
import { setMetricCell } from './ExcelReportBuilder.mts'

export type FinancialProjectionsModel = ReportModelBase & { kind: 'financial-projections' }

export const PROJECTION_YEARS = 5
export const PROJECTION_YEAR_COLS = ['B', 'C', 'D', 'E', 'F']
export const PROJECTION_ROWS = {
  header: 4,
  revenue: 5,
  cogs: 6,
  grossProfit: 7,
  opex: 8,
  ebitda: 9,
  depreciation: 10,
  interest: 11,
  pretax: 12,
  tax: 13,
  netIncome: 14,
  balanceSheetHeading: 16,
  currentAsset: 17,
  fixedAsset: 18,
  totalAssets: 19,
  currentLiability: 20,
  longTermDebt: 21,
  equity: 22,
  totalLiabEquity: 23,
  cashFlowHeading: 25,
  operatingCF: 26,
  investingCF: 27,
  financingCF: 28,
  netChangeCash: 29,
  kpiHeading: 31,
  grossMargin: 32,
  ebitdaMargin: 33,
  netMargin: 34,
  currentRatio: 35,
  debtToEquity: 36,
  roe: 37,
  roa: 38,
}

export type ProjectionBase = {
  revenue: number
  cogsRatio: number
  opex: number
  depreciation: number
  interest: number
  effectiveTaxRate: number
  currentAsset: number
  fixedAsset: number
  currentLiability: number
  longTermDebt: number
  equity: number
  investingCF: number
  financingCF: number
  workingCapitalRatioCA: number
  workingCapitalRatioCL: number
}

export type ProjectionRawRows = {
  revenue?: number
  opex?: number
  depreciation?: number
  interest?: number
  currentAsset?: number
  fixedAsset?: number
  currentLiability?: number
  longTermDebt?: number
  equity?: number
  investingCF?: number
  financingCF?: number
}

// Computes the same year-by-year figures the sheet-builder will render, purely
// numerically (no ExcelJS/sheet needed), so key metrics and the chart are
// available immediately without waiting for the deferred sheet build.
export function projectFinancials(base: ProjectionBase, cagr: number, inflation: number) {
  const revenue: number[] = []
  const ebitda: number[] = []
  const netIncome: number[] = []
  let runningEquity = base.equity
  for (let y = 0; y < PROJECTION_YEARS; y++) {
    const revenueY = base.revenue * Math.pow(1 + cagr, y)
    const cogsY = revenueY * base.cogsRatio
    const grossY = revenueY - cogsY
    const opexY = base.opex * Math.pow(1 + inflation, y)
    const ebitdaY = grossY - opexY
    const pretaxY = ebitdaY - base.depreciation - base.interest
    const taxY = pretaxY * base.effectiveTaxRate
    const netIncomeY = pretaxY - taxY
    revenue.push(revenueY)
    ebitda.push(ebitdaY)
    netIncome.push(netIncomeY)
    runningEquity = y === 0 ? base.equity : runningEquity + netIncomeY
  }
  return { revenue, ebitda, netIncome, finalEquity: runningEquity }
}

export function buildProjectionScenarioSheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  scenarioLabel: string,
  rawSheetName: string,
  base: ProjectionBase,
  rawRows: ProjectionRawRows,
  cagr: number,
  inflation: number,
) {
  const R = PROJECTION_ROWS
  const sheet = workbook.addWorksheet(sanitizeSheetName(`Projection - ${scenarioLabel}`, usedSheetNames))
  sheet.getColumn(1).width = 28
  for (let c = 2; c <= 6; c++) sheet.getColumn(c).width = 16

  sheet.getCell('A1').value = `5-Year Financial Projection — ${scenarioLabel}`
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:F1')
  sheet.getCell('A2').value = `Assumptions: ${(cagr * 100).toFixed(1)}%/yr revenue growth, ${(inflation * 100).toFixed(1)}%/yr cost inflation`
  sheet.getCell('A2').font = { size: 10, italic: true, color: { argb: 'FF555555' } }
  sheet.mergeCells('A2:F2')

  sheet.getCell(`A${R.header}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
  sheet.getCell(`A${R.header}`).border = THIN_BORDER
  PROJECTION_YEAR_COLS.forEach((col, i) => {
    const cell = sheet.getCell(`${col}${R.header}`)
    cell.value = `Year ${i + 1}`
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.alignment = { horizontal: 'center' }
    cell.border = THIN_BORDER
  })
  sheet.views = [{ state: 'frozen', ySplit: R.header }]

  const labelSpecs: [number, string, boolean?][] = [
    [R.revenue, 'Revenue', true],
    [R.cogs, 'COGS'],
    [R.grossProfit, 'Gross Profit', true],
    [R.opex, 'Operating Expenses'],
    [R.ebitda, 'EBITDA', true],
    [R.depreciation, 'Depreciation'],
    [R.interest, 'Interest Expense'],
    [R.pretax, 'Pre-tax Income'],
    [R.tax, 'Tax'],
    [R.netIncome, 'Net Income', true],
    [R.balanceSheetHeading, 'Balance Sheet'],
    [R.currentAsset, 'Current Assets'],
    [R.fixedAsset, 'Fixed Assets'],
    [R.totalAssets, 'Total Assets', true],
    [R.currentLiability, 'Current Liabilities'],
    [R.longTermDebt, 'Long-Term Debt'],
    [R.equity, 'Equity'],
    [R.totalLiabEquity, 'Total Liabilities & Equity', true],
    [R.cashFlowHeading, 'Cash Flow'],
    [R.operatingCF, 'Operating Cash Flow'],
    [R.investingCF, 'Investing Cash Flow'],
    [R.financingCF, 'Financing Cash Flow'],
    [R.netChangeCash, 'Net Change in Cash', true],
    [R.kpiHeading, 'Key Performance Indicators'],
    [R.grossMargin, 'Gross Margin %'],
    [R.ebitdaMargin, 'EBITDA Margin %'],
    [R.netMargin, 'Net Margin %'],
    [R.currentRatio, 'Current Ratio'],
    [R.debtToEquity, 'Debt-to-Equity (Total Liabilities ÷ Equity)'],
    [R.roe, 'ROE'],
    [R.roa, 'ROA'],
  ]
  const sectionHeadingRows = new Set([R.balanceSheetHeading, R.cashFlowHeading, R.kpiHeading])
  for (const [row, label, bold] of labelSpecs) {
    const cell = sheet.getCell(`A${row}`)
    cell.value = label
    cell.font = { bold: !!bold || sectionHeadingRows.has(row), color: sectionHeadingRows.has(row) ? { argb: NAVY } : undefined }
  }

  function setCell(row: number, col: string, value: MetricValue, format?: MetricFormat) {
    const cell = sheet.getCell(`${col}${row}`)
    setMetricCell(cell, value, format)
    cell.border = THIN_BORDER
  }

  let runningEquity = base.equity
  for (let y = 0; y < PROJECTION_YEARS; y++) {
    const col = PROJECTION_YEAR_COLS[y]
    const prevCol = y > 0 ? PROJECTION_YEAR_COLS[y - 1] : null
    const isYear1 = y === 0

    function flatOrYear1(rawRow: number | undefined, baseValue: number, thisRow: number): MetricValue {
      if (isYear1) return rawRow ? fx(refCell(rawSheetName, rawRow, 'D'), baseValue) : baseValue
      return fx(`=${PROJECTION_YEAR_COLS[0]}${thisRow}`, baseValue)
    }

    const revenueValue = base.revenue * Math.pow(1 + cagr, y)
    setCell(
      R.revenue,
      col,
      isYear1
        ? rawRows.revenue
          ? fx(refCell(rawSheetName, rawRows.revenue, 'D'), base.revenue)
          : base.revenue
        : fx(`=${prevCol}${R.revenue}*(1+${cagr})`, revenueValue),
      'currency',
    )

    const cogsValue = revenueValue * base.cogsRatio
    setCell(R.cogs, col, fx(`=${col}${R.revenue}*${base.cogsRatio}`, cogsValue), 'currency')

    const grossValue = revenueValue - cogsValue
    setCell(R.grossProfit, col, fx(`=${col}${R.revenue}-${col}${R.cogs}`, grossValue), 'currency')

    const opexValue = base.opex * Math.pow(1 + inflation, y)
    setCell(
      R.opex,
      col,
      isYear1
        ? rawRows.opex
          ? fx(refCell(rawSheetName, rawRows.opex, 'D'), base.opex)
          : base.opex
        : fx(`=${prevCol}${R.opex}*(1+${inflation})`, opexValue),
      'currency',
    )

    const ebitdaValue = grossValue - opexValue
    setCell(R.ebitda, col, fx(`=${col}${R.grossProfit}-${col}${R.opex}`, ebitdaValue), 'currency')

    const depValue = base.depreciation
    setCell(R.depreciation, col, flatOrYear1(rawRows.depreciation, base.depreciation, R.depreciation), 'currency')

    const intValue = base.interest
    setCell(R.interest, col, flatOrYear1(rawRows.interest, base.interest, R.interest), 'currency')

    const pretaxValue = ebitdaValue - depValue - intValue
    setCell(R.pretax, col, fx(`=${col}${R.ebitda}-${col}${R.depreciation}-${col}${R.interest}`, pretaxValue), 'currency')

    const taxValue = pretaxValue * base.effectiveTaxRate
    setCell(R.tax, col, fx(`=${col}${R.pretax}*${base.effectiveTaxRate}`, taxValue), 'currency')

    const netIncomeValue = pretaxValue - taxValue
    setCell(R.netIncome, col, fx(`=${col}${R.pretax}-${col}${R.tax}`, netIncomeValue), 'currency')

    const caValue = revenueValue * base.workingCapitalRatioCA
    setCell(
      R.currentAsset,
      col,
      isYear1
        ? rawRows.currentAsset
          ? fx(refCell(rawSheetName, rawRows.currentAsset, 'D'), base.currentAsset)
          : base.currentAsset
        : fx(`=${col}${R.revenue}*${base.workingCapitalRatioCA}`, caValue),
      'currency',
    )

    const faValue = base.fixedAsset
    setCell(R.fixedAsset, col, flatOrYear1(rawRows.fixedAsset, base.fixedAsset, R.fixedAsset), 'currency')

    const totalAssetsValue = caValue + faValue
    setCell(R.totalAssets, col, fx(`=${col}${R.currentAsset}+${col}${R.fixedAsset}`, totalAssetsValue), 'currency')

    const clValue = revenueValue * base.workingCapitalRatioCL
    setCell(
      R.currentLiability,
      col,
      isYear1
        ? rawRows.currentLiability
          ? fx(refCell(rawSheetName, rawRows.currentLiability, 'D'), base.currentLiability)
          : base.currentLiability
        : fx(`=${col}${R.revenue}*${base.workingCapitalRatioCL}`, clValue),
      'currency',
    )

    const ltdValue = base.longTermDebt
    setCell(R.longTermDebt, col, flatOrYear1(rawRows.longTermDebt, base.longTermDebt, R.longTermDebt), 'currency')

    let equityValue: number
    if (isYear1) {
      equityValue = base.equity
      setCell(R.equity, col, rawRows.equity ? fx(refCell(rawSheetName, rawRows.equity, 'D'), base.equity) : base.equity, 'currency')
    } else {
      equityValue = runningEquity + netIncomeValue
      setCell(R.equity, col, fx(`=${prevCol}${R.equity}+${col}${R.netIncome}`, equityValue), 'currency')
    }
    runningEquity = equityValue

    const totalLEValue = clValue + ltdValue + equityValue
    setCell(R.totalLiabEquity, col, fx(`=${col}${R.currentLiability}+${col}${R.longTermDebt}+${col}${R.equity}`, totalLEValue), 'currency')

    const opCFValue = netIncomeValue + depValue
    setCell(R.operatingCF, col, fx(`=${col}${R.netIncome}+${col}${R.depreciation}`, opCFValue), 'currency')

    const invCFValue = base.investingCF
    setCell(R.investingCF, col, flatOrYear1(rawRows.investingCF, base.investingCF, R.investingCF), 'currency')

    const finCFValue = base.financingCF
    setCell(R.financingCF, col, flatOrYear1(rawRows.financingCF, base.financingCF, R.financingCF), 'currency')

    const netChangeValue = opCFValue + invCFValue + finCFValue
    setCell(R.netChangeCash, col, fx(`=${col}${R.operatingCF}+${col}${R.investingCF}+${col}${R.financingCF}`, netChangeValue), 'currency')

    if (revenueValue !== 0) {
      setCell(R.grossMargin, col, fx(`=${col}${R.grossProfit}/${col}${R.revenue}`, grossValue / revenueValue), 'percent')
      setCell(R.ebitdaMargin, col, fx(`=${col}${R.ebitda}/${col}${R.revenue}`, ebitdaValue / revenueValue), 'percent')
      setCell(R.netMargin, col, fx(`=${col}${R.netIncome}/${col}${R.revenue}`, netIncomeValue / revenueValue), 'percent')
    } else {
      setCell(R.grossMargin, col, 'N/A')
      setCell(R.ebitdaMargin, col, 'N/A')
      setCell(R.netMargin, col, 'N/A')
    }
    setCell(R.currentRatio, col, clValue !== 0 ? fx(`=${col}${R.currentAsset}/${col}${R.currentLiability}`, caValue / clValue) : 'N/A', 'ratio')
    if (equityValue !== 0) {
      // Debt-to-equity = Total liabilities / Total shareholders' equity
      setCell(R.debtToEquity, col, fx(`=(${col}${R.currentLiability}+${col}${R.longTermDebt})/${col}${R.equity}`, (clValue + ltdValue) / equityValue), 'ratio')
      setCell(R.roe, col, fx(`=${col}${R.netIncome}/${col}${R.equity}`, netIncomeValue / equityValue), 'percent')
    } else {
      setCell(R.debtToEquity, col, 'N/A')
      setCell(R.roe, col, 'N/A')
    }
    setCell(
      R.roa,
      col,
      totalAssetsValue !== 0 ? fx(`=${col}${R.netIncome}/${col}${R.totalAssets}`, netIncomeValue / totalAssetsValue) : 'N/A',
      'percent',
    )
  }

  return sheet.name
}

export function calculateFinancialProjections(lineItems: LineItem[], growth?: GrowthAssumptions): FinancialProjectionsModel {
  const extraCaveats: string[] = []
  const rawSheetName = 'Year 1 Financials'
  const rawSheet = layoutRawDataSheet(rawSheetName, 'Year 1 (Base) Financials', lineItems, [
    { tag: 'revenue', label: 'Revenue' },
    { tag: 'cogs', label: 'COGS' },
    { tag: 'opex', label: 'Operating Expenses' },
    { tag: 'depreciation', label: 'Depreciation' },
    { tag: 'interest', label: 'Interest' },
    { tag: 'tax', label: 'Tax' },
    { tag: 'current_asset', label: 'Current Assets' },
    { tag: 'fixed_asset', label: 'Fixed Assets' },
    { tag: 'current_liability', label: 'Current Liabilities' },
    { tag: 'long_term_debt', label: 'Long-Term Debt' },
    { tag: 'equity', label: 'Equity' },
    { tag: 'operating_cash_flow_item', label: 'Operating Cash Flow' },
    { tag: 'investing_cash_flow_item', label: 'Investing Cash Flow' },
    { tag: 'financing_cash_flow_item', label: 'Financing Cash Flow' },
  ])

  const revenueBlock = findBlock(rawSheet, 'revenue')
  const cogsBlock = findBlock(rawSheet, 'cogs')
  const opexBlock = findBlock(rawSheet, 'opex')
  const depreciationBlock = findBlock(rawSheet, 'depreciation')
  const interestBlock = findBlock(rawSheet, 'interest')
  const taxBlock = findBlock(rawSheet, 'tax')
  const currentAssetBlock = findBlock(rawSheet, 'current_asset')
  const fixedAssetBlock = findBlock(rawSheet, 'fixed_asset')
  const currentLiabilityBlock = findBlock(rawSheet, 'current_liability')
  const longTermDebtBlock = findBlock(rawSheet, 'long_term_debt')
  const equityBlock = findBlock(rawSheet, 'equity')
  const opCFBlock = findBlock(rawSheet, 'operating_cash_flow_item')
  const invCFBlock = findBlock(rawSheet, 'investing_cash_flow_item')
  const finCFBlock = findBlock(rawSheet, 'financing_cash_flow_item')

  if (!revenueBlock) extraCaveats.push('No line items were classified as revenue; projections below could not be fully computed.')
  if (!equityBlock) extraCaveats.push('No line items were classified as equity; balance sheet and ROE projections below assume a starting equity of zero.')

  const revenue1 = revenueBlock?.subtotalValue ?? 0
  const cogs1 = cogsBlock?.subtotalValue ?? 0
  const opex1 = opexBlock?.subtotalValue ?? 0
  const depreciation1 = depreciationBlock?.subtotalValue ?? 0
  const interest1 = interestBlock?.subtotalValue ?? 0
  const tax1 = taxBlock?.subtotalValue ?? 0

  const pretax1 = revenue1 - cogs1 - opex1 - depreciation1 - interest1
  const effectiveTaxRate = pretax1 !== 0 ? tax1 / pretax1 : 0
  const cogsRatio = revenue1 !== 0 ? cogs1 / revenue1 : 0
  const workingCapitalRatioCA = revenue1 !== 0 ? (currentAssetBlock?.subtotalValue ?? 0) / revenue1 : 0
  const workingCapitalRatioCL = revenue1 !== 0 ? (currentLiabilityBlock?.subtotalValue ?? 0) / revenue1 : 0

  if (revenue1 === 0) {
    extraCaveats.push(
      'Year-1 revenue is zero — cost ratios, working capital scaling, and margin KPIs below default to zero and are not meaningful.',
    )
  }

  const base: ProjectionBase = {
    revenue: revenue1,
    cogsRatio,
    opex: opex1,
    depreciation: depreciation1,
    interest: interest1,
    effectiveTaxRate,
    currentAsset: currentAssetBlock?.subtotalValue ?? 0,
    fixedAsset: fixedAssetBlock?.subtotalValue ?? 0,
    currentLiability: currentLiabilityBlock?.subtotalValue ?? 0,
    longTermDebt: longTermDebtBlock?.subtotalValue ?? 0,
    equity: equityBlock?.subtotalValue ?? 0,
    investingCF: invCFBlock?.subtotalValue ?? 0,
    financingCF: finCFBlock?.subtotalValue ?? 0,
    workingCapitalRatioCA,
    workingCapitalRatioCL,
  }
  const rawRows: ProjectionRawRows = {
    revenue: revenueBlock?.subtotalRow,
    opex: opexBlock?.subtotalRow,
    depreciation: depreciationBlock?.subtotalRow,
    interest: interestBlock?.subtotalRow,
    currentAsset: currentAssetBlock?.subtotalRow,
    fixedAsset: fixedAssetBlock?.subtotalRow,
    currentLiability: currentLiabilityBlock?.subtotalRow,
    longTermDebt: longTermDebtBlock?.subtotalRow,
    equity: equityBlock?.subtotalRow,
    investingCF: invCFBlock?.subtotalRow,
    financingCF: finCFBlock?.subtotalRow,
  }

  const { optimisticCagrOffset, pessimisticCagrOffset } = PLATFORM_ASSUMPTIONS
  const annualInflationRate = growth?.inflation ?? PLATFORM_ASSUMPTIONS.annualInflationRate
  const baseRevenueCagr = growth?.baseCagr ?? PLATFORM_ASSUMPTIONS.baseRevenueCagr
  const scenarios = [
    { label: 'Base Case', cagr: baseRevenueCagr },
    { label: 'Optimistic', cagr: baseRevenueCagr + optimisticCagrOffset },
    { label: 'Pessimistic', cagr: Math.max(baseRevenueCagr - pessimisticCagrOffset, -0.95) },
  ]

  const customSheetBuilders = scenarios.map(
    (scenario) => (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildProjectionScenarioSheet(workbook, usedSheetNames, scenario.label, rawSheetName, base, rawRows, scenario.cagr, annualInflationRate),
  )

  const baseProjection = projectFinancials(base, baseRevenueCagr, annualInflationRate)
  const baseSheetName = 'Projection - Base Case'
  const keyMetrics: KeyMetric[] = [
    { label: 'Year 1 Revenue', cellRef: refCell(baseSheetName, PROJECTION_ROWS.revenue, 'B'), value: baseProjection.revenue[0], format: 'currency' },
    {
      label: 'Year 5 Revenue (Base Case)',
      cellRef: refCell(baseSheetName, PROJECTION_ROWS.revenue, 'F'),
      value: baseProjection.revenue[4],
      format: 'currency',
    },
    {
      label: 'Year 5 Net Income (Base Case)',
      cellRef: refCell(baseSheetName, PROJECTION_ROWS.netIncome, 'F'),
      value: baseProjection.netIncome[4],
      format: 'currency',
    },
    {
      label: 'Year 5 EBITDA Margin (Base Case)',
      cellRef: refCell(baseSheetName, PROJECTION_ROWS.ebitdaMargin, 'F'),
      value: baseProjection.revenue[4] !== 0 ? baseProjection.ebitda[4] / baseProjection.revenue[4] : 0,
      format: 'percent',
    },
  ]

  const chart: ReportChart = {
    type: 'line',
    title: '5-Year Projection (Base Case) — Revenue, EBITDA, Net Income',
    categories: ['Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'],
    series: [
      { name: 'Revenue', values: baseProjection.revenue.map((v) => Math.round(v)) },
      { name: 'EBITDA', values: baseProjection.ebitda.map((v) => Math.round(v)) },
      { name: 'Net Income', values: baseProjection.netIncome.map((v) => Math.round(v)) },
    ],
  }

  extraCaveats.push(
    `Years 2-5 are projected from Year-1 figures using ${growth?.inputs ? 'risk-adjusted' : 'platform-standard'} assumptions: ${(baseRevenueCagr * 100).toFixed(1)}% base-case revenue CAGR (optimistic +${(optimisticCagrOffset * 100).toFixed(0)}pp, pessimistic -${(pessimisticCagrOffset * 100).toFixed(0)}pp) and ${(annualInflationRate * 100).toFixed(1)}% annual operating-expense inflation${growth ? '. ' + growth.rationale : ', applied identically to every upload'}. Depreciation, interest, fixed assets, long-term debt, and investing/financing cash flows are held flat at Year-1 levels; COGS and working capital (current assets/liabilities) scale proportionally with revenue.`,
  )
  extraCaveats.push(
    'The projected balance sheet is a simplified, ratio-driven model, not a fully articulated three-statement forecast: equity rolls forward via retained earnings independently of the asset side, with no cash-balancing mechanism. Total Assets and Total Liabilities & Equity will not reconcile exactly in Years 2-5 — treat the balance sheet rows as directional, not precise.',
  )

  return {
    kind: 'financial-projections',
    rawSheets: [rawSheet],
    calcSheets: [],
    keyMetrics,
    charts: [chart],
    extraCaveats,
    customSheetBuilders,
  }
}
