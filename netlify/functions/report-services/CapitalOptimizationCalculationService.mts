import { type ReportChart, type LineItem } from './types.mts'
import { fx, refCell, type RawDataBlock, layoutRawDataSheet, findBlock, newCalcSheet, addSection, addMetric, type KeyMetric, type ReportModelBase } from './calculationSupport.mts'

export type CapitalOptimizationModel = ReportModelBase & { kind: 'capital-optimization' }

// A reasonable, commonly-cited default target for a healthy debt-to-equity
// ratio, used as the starting point for the restructuring comparison. This is
// an editable input cell in the rendered workbook — users can override it.
export const DEFAULT_TARGET_DEBT_TO_EQUITY = 1.0

export function calculateCapitalOptimization(lineItems: LineItem[]): CapitalOptimizationModel {
  const extraCaveats: string[] = []
  const DATA = 'Capital Structure Data'
  const rawSheet = layoutRawDataSheet(DATA, 'Capital Structure Data', lineItems, [
    { tag: 'short_term_debt', label: 'Short-Term Debt' },
    { tag: 'long_term_debt', label: 'Long-Term Debt' },
    { tag: 'accounts_payable', label: 'Accounts Payable' },
    { tag: 'past_due_payable', label: 'Accounts Payable Past Due (memo, included above)' },
    { tag: 'current_liability', label: 'Other Current Liabilities' },
    { tag: 'other_liability', label: 'Other Non-Current Liabilities' },
    { tag: 'equity', label: 'Equity' },
    { tag: 'current_asset', label: 'Current Assets' },
    { tag: 'fixed_asset', label: 'Fixed Assets' },
    { tag: 'interest_expense', label: 'Interest Expense' },
    { tag: 'ebit', label: 'EBIT' },
  ])

  const shortDebtBlock = findBlock(rawSheet, 'short_term_debt')
  const longDebtBlock = findBlock(rawSheet, 'long_term_debt')
  const equityBlock = findBlock(rawSheet, 'equity')
  const currentAssetBlock = findBlock(rawSheet, 'current_asset')
  const payablesBlock = findBlock(rawSheet, 'accounts_payable')
  const pastDuePayablesBlock = findBlock(rawSheet, 'past_due_payable')
  const currentLiabilityBlock = findBlock(rawSheet, 'current_liability')
  const otherLiabilityBlock = findBlock(rawSheet, 'other_liability')
  const interestBlock = findBlock(rawSheet, 'interest_expense')
  const ebitBlock = findBlock(rawSheet, 'ebit')
  const hasDebt = Boolean(shortDebtBlock || longDebtBlock)
  const hasLiabilities = Boolean(hasDebt || payablesBlock || currentLiabilityBlock || otherLiabilityBlock)

  if (!hasLiabilities || !equityBlock) extraCaveats.push('Liabilities and/or equity line items were not classified; capital structure ratios below could not be fully computed.')

  // Raw-sheet subtotal reference, or 0 when the group is absent.
  const sub = (block: RawDataBlock | undefined) => (block ? fx(refCell(DATA, block.subtotalRow, 'D'), block.subtotalValue) : fx('=0', 0))
  const shortDebt = shortDebtBlock?.subtotalValue ?? 0
  const longDebt = longDebtBlock?.subtotalValue ?? 0
  const debt = shortDebt + longDebt
  const payables = payablesBlock?.subtotalValue ?? 0
  const otherCurrentLiabilities = currentLiabilityBlock?.subtotalValue ?? 0
  const otherLiabilities = otherLiabilityBlock?.subtotalValue ?? 0
  const currentLiabilities = payables + otherCurrentLiabilities + shortDebt
  const totalLiabilities = currentLiabilities + longDebt + otherLiabilities
  const equity = equityBlock?.subtotalValue ?? 0

  const calcSheet = newCalcSheet('Capital Analysis', 'Capital Structure Analysis')
  const keyMetrics: KeyMetric[] = []

  addSection(calcSheet, 'Current Capital Structure')
  const shortDebtRow = addMetric(calcSheet, 'Short-Term Debt', sub(shortDebtBlock))
  const longDebtRow = addMetric(calcSheet, 'Long-Term Debt', sub(longDebtBlock))
  const debtRow = addMetric(calcSheet, 'Total Interest-Bearing Debt', fx(`=C${shortDebtRow}+C${longDebtRow}`, debt))
  const payablesRow = addMetric(calcSheet, 'Accounts Payable', sub(payablesBlock))
  const otherCurrentRow = addMetric(calcSheet, 'Other Current Liabilities', sub(currentLiabilityBlock))
  const currentLiabilitiesRow = addMetric(
    calcSheet,
    'Total Current Liabilities',
    fx(`=C${payablesRow}+C${otherCurrentRow}+C${shortDebtRow}`, currentLiabilities),
  )
  const otherLiabilitiesRow = addMetric(calcSheet, 'Other Non-Current Liabilities', sub(otherLiabilityBlock))
  const totalLiabilitiesRow = addMetric(
    calcSheet,
    'Total Liabilities',
    fx(`=C${currentLiabilitiesRow}+C${longDebtRow}+C${otherLiabilitiesRow}`, totalLiabilities),
    { bold: true },
  )
  const equityRow = addMetric(calcSheet, "Total Shareholders' Equity", sub(equityBlock))
  const totalCapitalRow = addMetric(calcSheet, 'Total Capital (Debt + Equity)', fx(`=C${debtRow}+C${equityRow}`, debt + equity))

  // Debt-to-equity = Total liabilities / Total shareholders' equity
  if (equity !== 0 && hasLiabilities) {
    const deRatioRow = addMetric(calcSheet, 'Debt-to-Equity Ratio (Total Liabilities ÷ Equity)', fx(`=C${totalLiabilitiesRow}/C${equityRow}`, totalLiabilities / equity), {
      format: 'ratio',
    })
    keyMetrics.push({ label: 'Debt-to-Equity Ratio', cellRef: refCell('Capital Analysis', deRatioRow, 'C'), value: totalLiabilities / equity, format: 'ratio' })
  } else {
    addMetric(calcSheet, 'Debt-to-Equity Ratio (Total Liabilities ÷ Equity)', equity === 0 ? 'N/A (no equity data)' : 'N/A (no liability data)')
  }
  if (debt + equity !== 0) {
    addMetric(calcSheet, 'Debt-to-Capital Ratio', fx(`=C${debtRow}/C${totalCapitalRow}`, debt / (debt + equity)), { format: 'percent' })
  } else {
    addMetric(calcSheet, 'Debt-to-Capital Ratio', 'N/A (no capital data)')
  }

  // Current ratio = Current assets / Current liabilities (short-term debt included)
  if (currentAssetBlock && currentLiabilities !== 0) {
    const caRow = addMetric(calcSheet, 'Total Current Assets', sub(currentAssetBlock))
    const currentRatio = currentAssetBlock.subtotalValue / currentLiabilities
    const currentRatioRow = addMetric(calcSheet, 'Current Ratio', fx(`=C${caRow}/C${currentLiabilitiesRow}`, currentRatio), { format: 'ratio' })
    keyMetrics.push({ label: 'Current Ratio', cellRef: refCell('Capital Analysis', currentRatioRow, 'C'), value: currentRatio, format: 'ratio' })
  } else {
    addMetric(calcSheet, 'Current Ratio', 'N/A (insufficient current asset/liability data)')
  }

  // Current accounts payable = (Total AP − Past due AP) / Total AP
  if (payablesBlock && pastDuePayablesBlock && payables !== 0) {
    const pastDueRow = addMetric(calcSheet, 'Accounts Payable Past Due', sub(pastDuePayablesBlock))
    const currentApShare = (payables - pastDuePayablesBlock.subtotalValue) / payables
    const currentApRow = addMetric(calcSheet, 'Current Accounts Payable (share not past due)', fx(`=(C${payablesRow}-C${pastDueRow})/C${payablesRow}`, currentApShare), {
      format: 'percent',
    })
    keyMetrics.push({ label: 'Current Accounts Payable', cellRef: refCell('Capital Analysis', currentApRow, 'C'), value: currentApShare, format: 'percent' })
    if (pastDuePayablesBlock.subtotalValue > payables) extraCaveats.push('Past-due payables exceed total accounts payable — check the AP aging against the balance sheet.')
  } else {
    addMetric(calcSheet, 'Current Accounts Payable (share not past due)', payablesBlock ? 'N/A (no AP aging / past-due payables data)' : 'N/A (no accounts payable data)')
  }

  // Interest coverage = EBIT / Interest expense
  let interestRow: number | undefined
  if (ebitBlock && interestBlock && interestBlock.subtotalValue !== 0) {
    const ebitRow = addMetric(calcSheet, 'EBIT', sub(ebitBlock))
    interestRow = addMetric(calcSheet, 'Interest Expense', sub(interestBlock))
    const coverage = ebitBlock.subtotalValue / interestBlock.subtotalValue
    const coverageRow = addMetric(calcSheet, 'Interest Coverage Ratio (EBIT ÷ Interest)', fx(`=C${ebitRow}/C${interestRow}`, coverage), { format: 'ratio' })
    keyMetrics.push({ label: 'Interest Coverage Ratio', cellRef: refCell('Capital Analysis', coverageRow, 'C'), value: coverage, format: 'ratio' })
  } else {
    addMetric(calcSheet, 'Interest Coverage Ratio (EBIT ÷ Interest)', 'N/A (insufficient EBIT/interest expense data)')
    if (!ebitBlock || !interestBlock) extraCaveats.push('EBIT and/or interest expense were not classified; interest coverage could not be computed.')
  }

  if (equityBlock && equity !== 0 && interestBlock && interestBlock.subtotalValue !== 0 && debt !== 0) {
    addSection(calcSheet, 'Restructuring Comparison — Current vs. Target Structure')
    const intRow = interestRow ?? addMetric(calcSheet, 'Interest Expense', sub(interestBlock))
    const targetRatioRow = addMetric(calcSheet, 'Target Debt-to-Equity Ratio (Total Liabilities ÷ Equity, editable)', DEFAULT_TARGET_DEBT_TO_EQUITY, {
      format: 'ratio',
    })
    const targetLiabilitiesRow = addMetric(calcSheet, 'Target Total Liabilities', fx(`=C${targetRatioRow}*C${equityRow}`, DEFAULT_TARGET_DEBT_TO_EQUITY * equity))
    const reduction = totalLiabilities - DEFAULT_TARGET_DEBT_TO_EQUITY * equity
    const reductionRow = addMetric(calcSheet, 'Liability Reduction Needed', fx(`=C${totalLiabilitiesRow}-C${targetLiabilitiesRow}`, reduction))
    const effectiveRate = interestBlock.subtotalValue / debt
    const effectiveRateRow = addMetric(calcSheet, 'Current Effective Interest Rate (Interest ÷ Debt)', fx(`=C${intRow}/C${debtRow}`, effectiveRate), {
      format: 'percent',
    })
    addMetric(
      calcSheet,
      'Estimated Annual Interest Savings',
      fx(`=MIN(C${reductionRow},C${debtRow})*C${effectiveRateRow}`, Math.min(reduction, debt) * effectiveRate),
      { bold: true, highlight: true },
    )
    extraCaveats.push(
      'Interest savings assume the liability reduction is achieved by repaying interest-bearing debt at the current effective rate, capped at total interest-bearing debt.',
    )
  }

  const chart: ReportChart = {
    type: 'bar',
    title: 'Current Capital Structure — Liabilities vs. Equity',
    categories: ['Interest-Bearing Debt', 'Other Liabilities', 'Equity'],
    series: [{ name: 'Amount', values: [debt, payables + otherCurrentLiabilities + otherLiabilities, equity] }],
  }

  return { kind: 'capital-optimization', rawSheets: [rawSheet], calcSheets: [calcSheet], keyMetrics, charts: [chart], extraCaveats }
}

// ---- Financial Projections calculator ----
//
// Unlike the other calculators, a 5-year, 3-scenario projection doesn't fit
// the generic account-list / label-value sheet shapes, so it renders into a
// dedicated year-columns grid (one sheet per scenario) via customSheetBuilders.
// Row numbers are fixed constants (not a dynamic running counter) so the
// key-metrics/chart figures computed here stay in exact sync with what the
// sheet-builder actually writes, even though the sheets themselves are only
// built later once the real workbook exists.
