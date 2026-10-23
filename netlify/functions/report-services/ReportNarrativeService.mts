import type { ExtractedDocument } from './DocumentExtractionService.mts'
import type { LineItem, ReportData } from './types.mts'
import type { ReportModelBase } from './calculationSupport.mts'

// Produces the prose, supplementary tables/charts and next steps that wrap the
// deterministic figures, and assembles the final ReportData.
//
// Today the commentary is authored by Claude in the same call that extracts the
// document (see DocumentExtractionService), so this step carries it forward
// unchanged. The verified calculation model is already an input here so that a
// second Claude call, grounded in the computed results, can replace the body of
// this function without touching the orchestrator or any other service.
export async function buildReportNarrative(
  extracted: ExtractedDocument,
  lineItems: LineItem[],
  _model: ReportModelBase | null,
): Promise<ReportData> {
  return {
    ...extracted.narrative,
    lineItems,
    assumptions: extracted.assumptions,
  }
}
