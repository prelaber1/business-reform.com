

export const REPORT_DEFINITIONS: Record<string, { label: string; instructions: string }> = {
  'break-even': {
    label: 'Break-Even Analysis Report',
    instructions: `Produce a comprehensive Break-Even Analysis Report. Include:
- Fixed vs variable cost identification and totals
- Break-even point in units and in revenue
- Contribution margin per unit and as a percentage
- Margin of safety calculation
- A table of revenue vs total cost at increasing volume levels, and a line chart of the same
- Key assumptions and sensitivity analysis
- Executive summary with actionable recommendations

Classify every relevant line item into "lineItems" using one of these group tags: "revenue" (sales/income accounts), "fixed_cost" (costs that do not vary with production/sales volume, e.g. rent, salaries, depreciation, insurance), or "variable_cost" (costs that scale with volume, e.g. materials, direct labor, sales commissions, subcontractor fees tied to project volume). Do NOT calculate break-even point, contribution margin, or margin of safety yourself — provide only the classified raw figures; the analysis engine computes these deterministically from your classifications. Explain any non-obvious classification choices in assumptions.classificationRationale.`,
  },
  'financial-projections': {
    label: '5-Year Financial Projections Report with KPIs',
    instructions: `Produce a 5-Year Financial Projections Report with KPIs. Include:
- Income Statement projections for Years 1-5 (as a table, and a line or bar chart of revenue/EBITDA/net income across the 5 years)
- Balance Sheet snapshot projections for Years 1-5
- Cash Flow Statement projections for Years 1-5
- KPI dashboard with margins, liquidity, leverage, ROE, and ROA (as a table)
- Growth assumptions clearly stated
- Scenario analysis for base, optimistic, and pessimistic cases
- Executive summary with investment highlights and risks

Classify every relevant Year-1 (base/historical) line item into "lineItems" using one of these group tags: "revenue", "cogs", "opex", "depreciation", "interest", "tax" (income statement); "current_asset", "fixed_asset", "current_liability" (all current liabilities, including short-term borrowings), "long_term_debt" (all non-current liabilities: long-term borrowings plus provisions and other non-current liabilities), "equity" (balance sheet, base year only — every liability must fall into "current_liability" or "long_term_debt" so that total liabilities are complete); "operating_cash_flow_item", "investing_cash_flow_item", "financing_cash_flow_item" (cash flow, base year only). Do NOT compute Years 2-5, growth rates, margins, or ratios yourself — the analysis engine projects Years 2-5 for base/optimistic/pessimistic scenarios and computes KPIs deterministically from your Year-1 classification using platform-standard growth and cost-inflation assumptions. Explain any non-obvious classification choices in assumptions.classificationRationale.`,
  },
  'burn-rate': {
    label: 'Burn Rate Report',
    instructions: `Produce a Burn Rate and Runway Analysis Report. Include:
- Gross burn rate and net burn rate
- Current cash balance and projected runway
- A month-by-month burn schedule for the next 12 months as both a table and a line chart
- Largest cash consumption categories (as a bar chart)
- Inflection points where burn rate changes significantly
- Recommendations to extend runway
- Executive summary with critical dates and funding milestones

Classify every relevant line item into "lineItems" using one of these group tags: "cash_inflow" or "cash_outflow" (each tagged with a "period" value like "Month 1".."Month 12" whenever the source data allows dating individual figures to a month; if the source only gives undated totals, omit "period" and note this limitation), and exactly one "starting_cash_balance" item for the current cash position. Do NOT calculate gross/net burn rate or runway yourself — the analysis engine computes these deterministically from your classified, dated figures. Explain any non-obvious classification choices or dating assumptions in assumptions.classificationRationale.`,
  },
  'cash-flow-forecast': {
    label: '13-Week / 12-Month Rolling Cash Flow Forecast Report',
    instructions: `Produce a 13-Week / 12-Month Rolling Cash Flow Forecast Report. Include:
- A week-by-week projection of beginning balance, cash inflows, cash outflows, net cash flow, and ending balance for the next 13 weeks (as a table and a line chart of the ending balance trend)
- A month-by-month projection of the same, rolling forward 12 months (as a table and a line chart of the ending balance trend)
- Identification of the lowest projected cash balance and the week/month it occurs in, to flag potential liquidity shortfalls
- Key recurring inflow and outflow categories driving the forecast (as a bar chart)
- Commentary on covenant or working-capital risk if the balance approaches zero
- Executive summary suitable for a lender or investor liquidity review

Classify every relevant line item into "lineItems" using one of these group tags: "cash_inflow" or "cash_outflow" (each tagged with a "period" value — use "Week 1".."Week 13" if the source data is weekly-dated, e.g. an existing 13-week cash flow worksheet, dated AR/AP aging, or a weekly payroll/sales calendar; use "Month 1".."Month 12" if the source data is only monthly-dated, e.g. a monthly P&L or cash flow statement; use whichever single granularity actually matches the source — omit "period" entirely if the source only gives undated totals, and note this limitation), and exactly one "starting_cash_balance" item for the current cash position. Do NOT calculate run-rates, roll the balance forward, or convert between weekly and monthly figures yourself — the analysis engine computes these deterministically from your classified, dated figures using a fixed weeks-per-month conversion factor. Explain any non-obvious classification choices or dating assumptions in assumptions.classificationRationale.`,
  },
  'capital-optimization': {
    label: 'Capital Optimization & Restructuring Report',
    instructions: `Produce a Capital Optimization and Restructuring Report. Include:
- Current capital structure analysis (debt vs equity mix, as a table and bar chart)
- Asset efficiency metrics
- Underperforming asset identification
- Debt restructuring options with cost-benefit comparison
- Equity optimisation opportunities
- Working capital improvement actions
- Recommended target capital structure
- Implementation roadmap with 30/60/90-day milestones

Classify every relevant line item into "lineItems" using one of these group tags: "short_term_debt" (interest-bearing borrowings due within 12 months, including the current portion of long-term loans and overdrafts), "long_term_debt" (interest-bearing borrowings due after 12 months), "accounts_payable" (trade payables to suppliers, in total), "current_liability" (other non-interest-bearing current liabilities, e.g. taxes, accruals — not trade payables), "other_liability" (non-interest-bearing non-current liabilities, e.g. provisions, deferred tax, employee benefit obligations), "equity", "current_asset", "fixed_asset", "interest_expense", or "ebit" (operating profit before interest and tax). Every liability on the balance sheet must be classified into exactly one of the five liability tags so that total liabilities are complete. If the source includes an accounts payable aging, also add one "past_due_payable" item with the amount of trade payables past their due date — this is a memo item that is already part of "accounts_payable", so it is never counted again in liabilities. Do NOT calculate debt-to-equity, current ratio, interest coverage, or restructuring savings yourself — the analysis engine computes these deterministically from your classified figures. Qualitative content that isn't a computed ratio (underperforming asset identification, working capital actions, the 30/60/90-day roadmap) belongs in "sections" as prose/bullets, not "lineItems". Explain any non-obvious classification choices in assumptions.classificationRationale.`,
  },
  'use-of-funds': {
    label: 'Use of Funds Report',
    instructions: `Produce a Use of Funds Report. Include:
- Total funding amount and sources
- A detailed allocation table (category, amount, percentage of total) and a bar chart of allocation by category
- Justification for each major spending category
- Expected ROI or business impact per allocation
- Milestone-based disbursement schedule
- Risk factors affecting fund utilisation
- Governance and accountability controls
- Executive summary suitable for investors or lenders

Classify every funding allocation category into "lineItems" using the group tag "allocation_category", with "label" as the category name and "amount" as the € allocated to it. Do NOT calculate each category's percentage of total funding yourself — the analysis engine computes this deterministically from your classified figures. Qualitative content (justification per category, ROI narrative, disbursement schedule, risk factors, governance controls) belongs in "sections" as prose/bullets, not "lineItems".`,
  },
  'budget-vs-actual': {
    label: 'Budget vs Actual Variance Report',
    instructions: `Produce a Budget vs Actual Variance Report. Include:
- Headline variance for revenue, costs and operating result (budget, actual, variance, variance %)
- A line-by-line variance table with favorable/unfavorable status
- Flagging of material variances and the likely drivers behind each, using only explanations found in the source (volume, price/rate, timing, one-off, mix, FX); if the source gives no explanation, say the driver is not documented rather than guessing
- A bar chart of the largest variances
- Open questions for budget owners on unexplained material variances
- Executive summary with the key messages for leadership and recommended actions

Classify every relevant line item into "lineItems" using one of these group tags: "budget_revenue", "actual_revenue", "budget_cost", "actual_cost" (the amount for the same period on both sides, costs as positive amounts), and optionally "budget_units" / "actual_units" (volume or quantity for a revenue or cost line, only when the source gives it). Use the identical "label" for a budget line and its matching actual line (and units line) so the engine can pair them; use the same level of detail on both sides and never net a line against another. If the source covers several periods, use the period the user most plausibly wants (the latest closed period) and say so in assumptions.classificationRationale. Do NOT calculate variances, percentages, totals or volume/price effects yourself — the analysis engine computes these deterministically from your classified figures. Explain any non-obvious classification or pairing choices in assumptions.classificationRationale.`,
  },
}
