import type { LineItem } from './types.mts'
import type { GrowthAssumptions } from './GrowthAssumptionService.mts'
import type { ReportModelBase } from './calculationSupport.mts'
import { calculateBreakEven } from './BreakEvenCalculationService.mts'
import { calculateUseOfFunds } from './UseOfFundsCalculationService.mts'
import { calculateBurnRate } from './BurnRateCalculationService.mts'
import { calculateCashFlowForecast } from './CashFlowProjectionService.mts'
import { calculateCapitalOptimization } from './CapitalOptimizationCalculationService.mts'
import { calculateFpaDigest } from './FpaDigestCalculationService.mts'
import { calculateFinancialProjections } from './FinancialProjectionsCalculationService.mts'

// Maps each offered report type to its calculation service. Adding a report
// means adding a service module and one case here (plus its REPORT_DEFINITIONS
// entry); nothing else in the pipeline is report-specific.
export function calculateReportModel(reportType: string, lineItems: LineItem[], growth?: GrowthAssumptions): ReportModelBase | null {
  switch (reportType) {
    case 'break-even':
      return calculateBreakEven(lineItems)
    case 'use-of-funds':
      return calculateUseOfFunds(lineItems)
    case 'burn-rate':
      return calculateBurnRate(lineItems)
    case 'cash-flow-forecast':
      return calculateCashFlowForecast(lineItems)
    case 'capital-optimization':
      return calculateCapitalOptimization(lineItems)
    case 'financial-projections':
      return calculateFinancialProjections(lineItems, growth)
    case 'fpa-digest':
      return calculateFpaDigest(lineItems)
    default:
      return null
  }
}
