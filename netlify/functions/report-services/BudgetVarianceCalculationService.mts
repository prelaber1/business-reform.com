// ---- Budget vs Actual Variance calculator ----
//
// Pairs each budget line with its actual by label, then computes variance,
// variance %, favorable/unfavorable and a materiality flag per line, plus
// totals and an optional volume/price decomposition for lines that carry
// units. The AI only classifies figures; every number below is computed here
// and rendered as live Excel formulas (thresholds are editable input cells).
import ExcelJS from 'exceljs'
import { NAVY } from './constants.mts'
import { type ReportChart, type LineItem } from './types.mts'
import {
  fx,
  groupItems,
  layoutRawDataSheet,
  refCell,
  newCalcSheet,
  addSection,
  addMetric,
  RESULT_FILL,
  SUBTOTAL_FILL,
  CAVEAT_FILL,
  THIN_BORDER,
  type KeyMetric,
  type ReportModelBase,
} from './calculationSupport.mts'
import { sanitizeSheetName } from './excelSheetHelpers.mts'

export type BudgetVarianceModel = ReportModelBase & { kind: 'budget-vs-actual' }

// Materiality: a variance is flagged when it is both relatively large and
// large enough in absolute terms to matter against the budget's scale.
export const VARIANCE_RULES = {
  materialPct: 0.1, // |variance| >= 10% of the budget line
  materialAbsShareOfBase: 0.005, // ... and >= 0.5% of total budget revenue (or total budget costs if no revenue)
}

const RAW_SHEET = 'Budget vs Actual Data'
const DETAIL_SHEET = 'Variance Detail'
const DECOMP_SHEET = 'Driver Decomposition'
const GROUP_ORDER = [
  { tag: 'budget_revenue', label: 'Budget Revenue' },
  { tag: 'actual_revenue', label: 'Actual Revenue' },
  { tag: 'budget_cost', label: 'Budget Cost' },
  { tag: 'actual_cost', label: 'Actual Cost' },
  { tag: 'budget_units', label: 'Budget Units' },
  { tag: 'actual_units', label: 'Actual Units' },
]

type Kind = 'revenue' | 'cost'
type Pair = {
  key: string
  label: string
  kind: Kind
  budget: number
  actual: number
  budgetRow?: number
  actualRow?: number
  budgetUnits?: number
  actualUnits?: number
  budgetUnitsRow?: number
  actualUnitsRow?: number
}

const normalize = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

// Same label + group appearing twice is summed so each label maps to one row.
function mergeDuplicates(items: LineItem[]): LineItem[] {
  const merged = new Map<string, LineItem>()
  for (const item of items) {
    const key = `${item.group}|${normalize(item.label)}`
    const existing = merged.get(key)
    if (existing) existing.amount += item.amount
    else merged.set(key, { ...item })
  }
  return [...merged.values()]
}

// Pure, so the rules can be tested without a workbook.
export function computeVariance(kind: Kind, budget: number, actual: number, totalBase: number) {
  const variance = actual - budget
  const pct = budget !== 0 ? variance / budget : null
  const profitImpact = kind === 'revenue' ? variance : -variance
  const minAbs = totalBase * VARIANCE_RULES.materialAbsShareOfBase
  const material = Math.abs(variance) >= minAbs && (pct === null ? variance !== 0 : Math.abs(pct) >= VARIANCE_RULES.materialPct)
  const status = variance === 0 ? 'On budget' : profitImpact > 0 ? 'Favorable' : 'Unfavorable'
  return { variance, pct, profitImpact, material, status }
}

// Volume effect = change in units at budget price; price effect = change in
// price at actual units. Together they equal the line variance exactly.
export function decomposeVolumePrice(budget: number, actual: number, budgetUnits: number, actualUnits: number) {
  if (!(budgetUnits > 0) || !(actualUnits > 0)) return null
  const budgetPrice = budget / budgetUnits
  const actualPrice = actual / actualUnits
  return { budgetPrice, actualPrice, volumeEffect: (actualUnits - budgetUnits) * budgetPrice, priceEffect: (actualPrice - budgetPrice) * actualUnits }
}

export function calculateBudgetVariance(lineItems: LineItem[]): BudgetVarianceModel {
  const extraCaveats: string[] = []
  const items = mergeDuplicates(lineItems.filter((li) => GROUP_ORDER.some((g) => g.tag === li.group)))
  const rawSheet = layoutRawDataSheet(RAW_SHEET, 'Budget vs Actual — Source Figures', items, GROUP_ORDER)

  const rowOf = (groupLabel: string, label: string) =>
    rawSheet.rows.find((r) => r.kind === 'item' && r.groupLabel === groupLabel && normalize(r.description) === normalize(label))?.row

  const pairs = new Map<string, Pair>()
  const touch = (item: LineItem, kind: Kind) => {
    const key = `${kind}|${normalize(item.label)}`
    let p = pairs.get(key)
    if (!p) {
      p = { key, label: item.label, kind, budget: 0, actual: 0 }
      pairs.set(key, p)
    }
    return p
  }
  for (const item of groupItems(items, 'budget_revenue')) Object.assign(touch(item, 'revenue'), { budget: item.amount, budgetRow: rowOf('Budget Revenue', item.label) })
  for (const item of groupItems(items, 'actual_revenue')) Object.assign(touch(item, 'revenue'), { actual: item.amount, actualRow: rowOf('Actual Revenue', item.label) })
  for (const item of groupItems(items, 'budget_cost')) Object.assign(touch(item, 'cost'), { budget: item.amount, budgetRow: rowOf('Budget Cost', item.label) })
  for (const item of groupItems(items, 'actual_cost')) Object.assign(touch(item, 'cost'), { actual: item.amount, actualRow: rowOf('Actual Cost', item.label) })

  // Units attach to whichever revenue/cost line shares the label.
  for (const [tag, field, rowField, groupLabel] of [
    ['budget_units', 'budgetUnits', 'budgetUnitsRow', 'Budget Units'],
    ['actual_units', 'actualUnits', 'actualUnitsRow', 'Actual Units'],
  ] as const) {
    for (const item of groupItems(items, tag)) {
      const target = pairs.get(`revenue|${normalize(item.label)}`) ?? pairs.get(`cost|${normalize(item.label)}`)
      if (target) Object.assign(target, { [field]: item.amount, [rowField]: rowOf(groupLabel, item.label) })
      else extraCaveats.push(`Units given for "${item.label}" but no matching budget/actual line; volume/price split skipped for it.`)
    }
  }

  const all = [...pairs.values()]
  const revenue = all.filter((p) => p.kind === 'revenue')
  const costs = all.filter((p) => p.kind === 'cost')
  const unmatched = all.filter((p) => !p.budgetRow || !p.actualRow)
  if (!all.length) extraCaveats.push('No budget/actual line items were classified; the variance analysis below could not be computed.')
  if (unmatched.length) {
    extraCaveats.push(
      `${unmatched.length} line(s) have only a budget or only an actual figure (${unmatched.slice(0, 5).map((p) => p.label).join(', ')}${unmatched.length > 5 ? ', ...' : ''}); the missing side is treated as 0, which shows the full amount as variance. Check for label mismatches between the budget and actual files.`,
    )
  }

  const sum = (ps: Pair[], f: 'budget' | 'actual') => ps.reduce((t, p) => t + p[f], 0)
  const budgetRevenue = sum(revenue, 'budget')
  const budgetCosts = sum(costs, 'budget')
  const totalBase = budgetRevenue > 0 ? budgetRevenue : budgetCosts
  const minAbs = totalBase * VARIANCE_RULES.materialAbsShareOfBase
  const baseLabel = budgetRevenue > 0 ? 'budget revenue' : 'budget costs'

  // ---- Variance Detail sheet layout (computed up front so key metrics can reference it).
  const FIRST = 6 // rows 1-4 title/thresholds, row 5 header
  type DetailRow = { row: number; pair: Pair }
  const detail: DetailRow[] = []
  let r = FIRST
  const revStart = r
  for (const p of revenue) detail.push({ row: r++, pair: p })
  const revTotalRow = r++
  const costStart = r
  for (const p of costs) detail.push({ row: r++, pair: p })
  const costTotalRow = r++
  const resultRow = r++

  const computed = new Map(detail.map((d) => [d.row, computeVariance(d.pair.kind, d.pair.budget, d.pair.actual, totalBase)]))
  const resultBudget = budgetRevenue - budgetCosts
  const resultActual = sum(revenue, 'actual') - sum(costs, 'actual')
  const revVar = sum(revenue, 'actual') - budgetRevenue
  const costVar = sum(costs, 'actual') - budgetCosts
  const resultVar = resultActual - resultBudget
  const materialCount = [...computed.values()].filter((c) => c.material).length

  const metricRef = (row: number, col: string) => refCell(DETAIL_SHEET, row, col)
  const keyMetrics: KeyMetric[] = [
    { label: 'Revenue Variance (Actual − Budget)', cellRef: metricRef(revTotalRow, 'E'), value: revVar, format: 'currency' },
    { label: 'Cost Variance (Actual − Budget)', cellRef: metricRef(costTotalRow, 'E'), value: costVar, format: 'currency' },
    { label: 'Operating Result Variance', cellRef: metricRef(resultRow, 'E'), value: resultVar, format: 'currency' },
    ...(budgetRevenue !== 0 ? [{ label: 'Revenue Variance %', cellRef: metricRef(revTotalRow, 'F'), value: revVar / budgetRevenue, format: 'percent' as const }] : []),
  ]

  // Calc sheet: headline summary referencing the detail sheet.
  const calcSheet = newCalcSheet('Variance Summary', 'Budget vs Actual — Summary')
  addSection(calcSheet, 'Totals')
  addMetric(calcSheet, 'Budget Revenue', fx(`=${metricRef(revTotalRow, 'C')}`, budgetRevenue))
  addMetric(calcSheet, 'Actual Revenue', fx(`=${metricRef(revTotalRow, 'D')}`, sum(revenue, 'actual')))
  addMetric(calcSheet, 'Budget Costs', fx(`=${metricRef(costTotalRow, 'C')}`, budgetCosts))
  addMetric(calcSheet, 'Actual Costs', fx(`=${metricRef(costTotalRow, 'D')}`, sum(costs, 'actual')))
  addMetric(calcSheet, 'Budget Operating Result', fx(`=${metricRef(resultRow, 'C')}`, resultBudget))
  addMetric(calcSheet, 'Actual Operating Result', fx(`=${metricRef(resultRow, 'D')}`, resultActual))
  addMetric(calcSheet, 'Operating Result Variance', fx(`=${metricRef(resultRow, 'E')}`, resultVar), { bold: true, highlight: true })
  addMetric(
    calcSheet,
    'Material variances flagged (count)',
    fx(`=COUNTIF('${DETAIL_SHEET}'!I${revStart}:I${resultRow},"MATERIAL")`, materialCount),
    { format: 'number', bold: true },
  )

  // ---- Charts: profit impact of the largest movers (positive = favorable).
  const movers = detail
    .map((d) => ({ label: d.pair.label, impact: computed.get(d.row)!.profitImpact }))
    .filter((m) => m.impact !== 0)
    .sort((a, b) => Math.abs(b.impact) - Math.abs(a.impact))
    .slice(0, 10)
  const charts: ReportChart[] = []
  if (movers.length) {
    charts.push({
      type: 'bar',
      title: 'Largest Variances — Impact on Operating Result (positive = favorable)',
      categories: movers.map((m) => m.label),
      series: [{ name: 'Impact on operating result', values: movers.map((m) => Math.round(m.impact)) }],
    })
  }
  if (revenue.length + costs.length > 0) {
    charts.push({
      type: 'bar',
      title: 'Budget vs Actual — Revenue, Costs, Operating Result',
      categories: ['Revenue', 'Costs', 'Operating Result'],
      series: [
        { name: 'Budget', values: [Math.round(budgetRevenue), Math.round(budgetCosts), Math.round(resultBudget)] },
        { name: 'Actual', values: [Math.round(sum(revenue, 'actual')), Math.round(sum(costs, 'actual')), Math.round(resultActual)] },
      ],
    })
  }

  // ---- Narrative facts for the commentary pass: computed, so the AI never does the math.
  const money = (n: number) => Math.round(n).toLocaleString('en-US')
  const pctText = (p: number | null) => (p === null ? 'n/a (no budget)' : `${(p * 100).toFixed(1)}%`)
  const narrativeFacts: string[] = [
    `Totals: revenue budget ${money(budgetRevenue)}, actual ${money(sum(revenue, 'actual'))}, variance ${money(revVar)}; costs budget ${money(budgetCosts)}, actual ${money(sum(costs, 'actual'))}, variance ${money(costVar)}; operating result budget ${money(resultBudget)}, actual ${money(resultActual)}, variance ${money(resultVar)}.`,
    `Materiality rule: |variance| >= ${VARIANCE_RULES.materialPct * 100}% of the budget line and >= ${money(minAbs)} (${VARIANCE_RULES.materialAbsShareOfBase * 100}% of ${baseLabel}).`,
    ...detail
      .filter((d) => computed.get(d.row)!.material)
      .sort((a, b) => Math.abs(computed.get(b.row)!.profitImpact) - Math.abs(computed.get(a.row)!.profitImpact))
      .map((d) => {
        const c = computed.get(d.row)!
        const p = d.pair
        const dec = p.budgetUnits !== undefined && p.actualUnits !== undefined ? decomposeVolumePrice(p.budget, p.actual, p.budgetUnits, p.actualUnits) : null
        return `MATERIAL ${p.kind} line "${p.label}": budget ${money(p.budget)}, actual ${money(p.actual)}, variance ${money(c.variance)} (${pctText(c.pct)}), ${c.status}${dec ? `; volume effect ${money(dec.volumeEffect)}, price/rate effect ${money(dec.priceEffect)}` : ''}`
      }),
  ]

  if (materialCount === 0 && all.length) extraCaveats.push('No variance met the materiality thresholds.')
  extraCaveats.push(
    `Variance = Actual − Budget. Favorable means higher revenue or lower cost than budget. A line is flagged MATERIAL when its variance is at least ${VARIANCE_RULES.materialPct * 100}% of budget and at least ${VARIANCE_RULES.materialAbsShareOfBase * 100}% of ${baseLabel}; both thresholds are editable on the "${DETAIL_SHEET}" sheet. Lines are paired by label, so differently named accounts appear as separate rows.`,
  )

  const decomposable = all.filter((p) => p.budgetUnits !== undefined && p.actualUnits !== undefined && decomposeVolumePrice(p.budget, p.actual, p.budgetUnits, p.actualUnits))
  if (!decomposable.length) extraCaveats.push('No units/volume figures were available, so variances are not split into volume and price/rate effects.')

  const customSheetBuilders = [
    (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildDetailSheet(workbook, usedSheetNames, { detail, revStart, revTotalRow, costStart, costTotalRow, resultRow, minAbs, totals: { budgetRevenue, budgetCosts, actualRevenue: sum(revenue, 'actual'), actualCosts: sum(costs, 'actual') }, computed, totalBase, baseLabel }),
    ...(decomposable.length ? [(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) => buildDecompositionSheet(workbook, usedSheetNames, decomposable)] : []),
  ]

  return { kind: 'budget-vs-actual', rawSheets: [rawSheet], calcSheets: [calcSheet], keyMetrics, charts, extraCaveats, narrativeFacts, customSheetBuilders }
}

const MONEY = '#,##0.00" €"'
const header = (cell: ExcelJS.Cell) => {
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
  cell.alignment = { horizontal: 'center', wrapText: true }
  cell.border = THIN_BORDER
}
const raw = (row?: number) => (row ? `'${RAW_SHEET}'!D${row}` : '0')

function buildDetailSheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  s: {
    detail: { row: number; pair: Pair }[]
    revStart: number
    revTotalRow: number
    costStart: number
    costTotalRow: number
    resultRow: number
    minAbs: number
    totals: { budgetRevenue: number; budgetCosts: number; actualRevenue: number; actualCosts: number }
    computed: Map<number, ReturnType<typeof computeVariance>>
    totalBase: number
    baseLabel: string
  },
) {
  const sheet = workbook.addWorksheet(sanitizeSheetName(DETAIL_SHEET, usedSheetNames))
  ;[34, 10, 16, 16, 16, 12, 18, 14, 12].forEach((w, i) => (sheet.getColumn(i + 1).width = w))
  sheet.getCell('A1').value = 'Budget vs Actual — Variance Detail'
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:I1')

  sheet.getCell('A2').value = 'Materiality: minimum variance % of budget line (editable)'
  sheet.getCell('C2').value = VARIANCE_RULES.materialPct
  sheet.getCell('C2').numFmt = '0.0%'
  sheet.getCell('A3').value = `Materiality: minimum absolute variance (${VARIANCE_RULES.materialAbsShareOfBase * 100}% of ${s.baseLabel}, editable)`
  const baseRef = s.baseLabel === 'budget revenue' ? `C${s.revTotalRow}` : `C${s.costTotalRow}`
  sheet.getCell('C3').value = { formula: `=${baseRef}*${VARIANCE_RULES.materialAbsShareOfBase}`, result: s.minAbs }
  sheet.getCell('C3').numFmt = MONEY
  for (const c of ['C2', 'C3']) sheet.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CAVEAT_FILL } }

  ;['Line', 'Type', 'Budget', 'Actual', 'Variance (Actual − Budget)', 'Variance %', 'Impact on Operating Result', 'Favorable / Unfavorable', 'Material?'].forEach((t, i) => {
    sheet.getCell(5, i + 1).value = t
    header(sheet.getCell(5, i + 1))
  })
  sheet.getRow(5).height = 32
  sheet.views = [{ state: 'frozen', ySplit: 5 }]

  for (const { row, pair: p } of s.detail) {
    const c = s.computed.get(row)!
    const sign = p.kind === 'revenue' ? '' : '-'
    sheet.getCell(`A${row}`).value = p.label
    sheet.getCell(`B${row}`).value = p.kind === 'revenue' ? 'Revenue' : 'Cost'
    sheet.getCell(`C${row}`).value = { formula: `=${raw(p.budgetRow)}`, result: p.budget }
    sheet.getCell(`D${row}`).value = { formula: `=${raw(p.actualRow)}`, result: p.actual }
    sheet.getCell(`E${row}`).value = { formula: `=D${row}-C${row}`, result: c.variance }
    sheet.getCell(`F${row}`).value = { formula: `=IF(C${row}=0,"n/a",E${row}/C${row})`, result: c.pct === null ? 'n/a' : c.pct }
    sheet.getCell(`G${row}`).value = { formula: `=${sign}E${row}`, result: c.profitImpact }
    sheet.getCell(`H${row}`).value = {
      formula: `=IF(G${row}>0,"Favorable",IF(G${row}<0,"Unfavorable","On budget"))`,
      result: c.status,
    }
    sheet.getCell(`I${row}`).value = {
      formula: `=IF(ABS(E${row})<$C$3,"",IF(C${row}=0,"MATERIAL",IF(ABS(E${row}/C${row})>=$C$2,"MATERIAL","")))`,
      result: c.material ? 'MATERIAL' : '',
    }
    ;['C', 'D', 'E', 'G'].forEach((col) => (sheet.getCell(`${col}${row}`).numFmt = MONEY))
    sheet.getCell(`F${row}`).numFmt = '0.0%'
    for (let col = 1; col <= 9; col++) sheet.getCell(row, col).border = THIN_BORDER
    sheet.getCell(`I${row}`).alignment = { horizontal: 'center' }
    if (c.material) sheet.getCell(`I${row}`).font = { bold: true, color: { argb: 'FFC00000' } }
  }

  const total = (row: number, label: string, start: number, end: number, budget: number, actual: number, fill: string) => {
    sheet.getCell(`A${row}`).value = label
    sheet.getCell(`C${row}`).value = end >= start ? { formula: `=SUM(C${start}:C${end})`, result: budget } : 0
    sheet.getCell(`D${row}`).value = end >= start ? { formula: `=SUM(D${start}:D${end})`, result: actual } : 0
    sheet.getCell(`E${row}`).value = { formula: `=D${row}-C${row}`, result: actual - budget }
    sheet.getCell(`F${row}`).value = { formula: `=IF(C${row}=0,"n/a",E${row}/C${row})`, result: budget !== 0 ? (actual - budget) / budget : 'n/a' }
    ;['C', 'D', 'E'].forEach((col) => (sheet.getCell(`${col}${row}`).numFmt = MONEY))
    sheet.getCell(`F${row}`).numFmt = '0.0%'
    for (let col = 1; col <= 9; col++) {
      const cell = sheet.getCell(row, col)
      cell.font = { bold: true }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } }
      cell.border = THIN_BORDER
    }
  }
  const t = s.totals
  total(s.revTotalRow, 'Total Revenue', s.revStart, s.revTotalRow - 1, t.budgetRevenue, t.actualRevenue, SUBTOTAL_FILL)
  total(s.costTotalRow, 'Total Costs', s.costStart, s.costTotalRow - 1, t.budgetCosts, t.actualCosts, SUBTOTAL_FILL)

  const rr = s.resultRow
  sheet.getCell(`A${rr}`).value = 'Operating Result (Revenue − Costs)'
  sheet.getCell(`C${rr}`).value = { formula: `=C${s.revTotalRow}-C${s.costTotalRow}`, result: t.budgetRevenue - t.budgetCosts }
  sheet.getCell(`D${rr}`).value = { formula: `=D${s.revTotalRow}-D${s.costTotalRow}`, result: t.actualRevenue - t.actualCosts }
  sheet.getCell(`E${rr}`).value = { formula: `=D${rr}-C${rr}`, result: t.actualRevenue - t.actualCosts - (t.budgetRevenue - t.budgetCosts) }
  sheet.getCell(`F${rr}`).value = {
    formula: `=IF(C${rr}=0,"n/a",E${rr}/ABS(C${rr}))`,
    result: t.budgetRevenue - t.budgetCosts !== 0 ? (t.actualRevenue - t.actualCosts - (t.budgetRevenue - t.budgetCosts)) / Math.abs(t.budgetRevenue - t.budgetCosts) : 'n/a',
  }
  ;['C', 'D', 'E'].forEach((col) => (sheet.getCell(`${col}${rr}`).numFmt = MONEY))
  sheet.getCell(`F${rr}`).numFmt = '0.0%'
  for (let col = 1; col <= 9; col++) {
    const cell = sheet.getCell(rr, col)
    cell.font = { bold: true }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RESULT_FILL } }
    cell.border = THIN_BORDER
  }
}

function buildDecompositionSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, pairs: Pair[]) {
  const sheet = workbook.addWorksheet(sanitizeSheetName(DECOMP_SHEET, usedSheetNames))
  ;[34, 10, 14, 14, 14, 14, 18, 18, 18, 14].forEach((w, i) => (sheet.getColumn(i + 1).width = w))
  sheet.getCell('A1').value = 'Variance Drivers — Volume vs Price/Rate'
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:J1')
  sheet.getCell('A2').value = 'Volume effect = (actual units − budget units) × budget price. Price/rate effect = (actual price − budget price) × actual units. They sum to the line variance.'
  sheet.getCell('A2').font = { italic: true, size: 10 }
  const heads = ['Line', 'Type', 'Budget units', 'Actual units', 'Budget price', 'Actual price', 'Volume effect', 'Price/rate effect', 'Total (check)', 'Ties to variance?']
  heads.forEach((t, i) => {
    sheet.getCell(4, i + 1).value = t
    header(sheet.getCell(4, i + 1))
  })
  sheet.getRow(4).height = 32

  pairs.forEach((p, i) => {
    const row = 5 + i
    const d = decomposeVolumePrice(p.budget, p.actual, p.budgetUnits!, p.actualUnits!)!
    const bu = `'${RAW_SHEET}'!D${p.budgetUnitsRow}`
    const au = `'${RAW_SHEET}'!D${p.actualUnitsRow}`
    sheet.getCell(`A${row}`).value = p.label
    sheet.getCell(`B${row}`).value = p.kind === 'revenue' ? 'Revenue' : 'Cost'
    sheet.getCell(`C${row}`).value = { formula: `=${bu}`, result: p.budgetUnits! }
    sheet.getCell(`D${row}`).value = { formula: `=${au}`, result: p.actualUnits! }
    sheet.getCell(`E${row}`).value = { formula: `=${raw(p.budgetRow)}/C${row}`, result: d.budgetPrice }
    sheet.getCell(`F${row}`).value = { formula: `=${raw(p.actualRow)}/D${row}`, result: d.actualPrice }
    sheet.getCell(`G${row}`).value = { formula: `=(D${row}-C${row})*E${row}`, result: d.volumeEffect }
    sheet.getCell(`H${row}`).value = { formula: `=(F${row}-E${row})*D${row}`, result: d.priceEffect }
    sheet.getCell(`I${row}`).value = { formula: `=G${row}+H${row}`, result: d.volumeEffect + d.priceEffect }
    sheet.getCell(`J${row}`).value = {
      formula: `=IF(ABS(I${row}-(${raw(p.actualRow)}-${raw(p.budgetRow)}))<0.01,"Yes","No")`,
      result: Math.abs(d.volumeEffect + d.priceEffect - (p.actual - p.budget)) < 0.01 ? 'Yes' : 'No',
    }
    ;['E', 'F', 'G', 'H', 'I'].forEach((col) => (sheet.getCell(`${col}${row}`).numFmt = MONEY))
    ;['C', 'D'].forEach((col) => (sheet.getCell(`${col}${row}`).numFmt = '#,##0.##'))
    for (let col = 1; col <= 10; col++) sheet.getCell(row, col).border = THIN_BORDER
  })
}
