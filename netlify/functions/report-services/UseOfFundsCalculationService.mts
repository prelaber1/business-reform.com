import { type ReportChart, type LineItem } from './types.mts'
import { fx, refCell, type RawDataItemRow, layoutRawDataSheet, findBlock, newCalcSheet, addSection, addMetric, type KeyMetric, type ReportModelBase } from './calculationSupport.mts'

export type UseOfFundsModel = ReportModelBase & { kind: 'use-of-funds' }

export function calculateUseOfFunds(lineItems: LineItem[]): UseOfFundsModel {
  const extraCaveats: string[] = []
  const rawSheet = layoutRawDataSheet('Fund Allocation', 'Fund Allocation by Category', lineItems, [
    { tag: 'allocation_category', label: 'Allocation Category' },
  ])
  const block = findBlock(rawSheet, 'allocation_category')
  if (!block) extraCaveats.push('No line items were classified as allocation categories; allocation analysis could not be computed.')
  const total = block?.subtotalValue ?? 0

  const calcSheet = newCalcSheet('Allocation Analysis', 'Allocation Analysis')
  const keyMetrics: KeyMetric[] = []
  const itemRows = block ? (rawSheet.rows.filter((r) => r.kind === 'item') as RawDataItemRow[]) : []

  addSection(calcSheet, 'Allocation % of Total Funding')
  if (block && total !== 0) {
    for (const item of itemRows) {
      addMetric(
        calcSheet,
        item.description,
        fx(`=${refCell('Fund Allocation', item.row, 'D')}/${refCell('Fund Allocation', block.subtotalRow, 'D')}`, item.amount / total),
        { format: 'percent' },
      )
    }
  } else if (block) {
    addMetric(calcSheet, 'Allocation %', 'N/A (zero total funding)')
    extraCaveats.push('Total funding is zero — allocation percentages could not be computed.')
  }

  if (block) {
    keyMetrics.push({ label: 'Total Funding', cellRef: refCell('Fund Allocation', block.subtotalRow, 'D'), value: total, format: 'currency' })
  }

  const chart: ReportChart = {
    type: 'bar',
    title: 'Fund Allocation by Category',
    categories: itemRows.map((i) => i.description),
    series: [{ name: 'Amount', values: itemRows.map((i) => i.amount) }],
  }

  return { kind: 'use-of-funds', rawSheets: [rawSheet], calcSheets: [calcSheet], keyMetrics, charts: [chart], extraCaveats }
}

// ---- Burn Rate calculator ----
