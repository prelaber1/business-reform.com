import Anthropic from '@anthropic-ai/sdk'
import type { ExtractedDocument } from './DocumentExtractionService.mts'
import { PLATFORM_ASSUMPTIONS } from './calculationSupport.mts'

// Market inputs for the growth adjustment. Claude researches these (web search)
// for the business described in the upload; the CAGR itself is then computed
// deterministically below, in keeping with the "AI supplies figures, code does
// the arithmetic" rule used everywhere else in the pipeline.
export type GrowthInputs = {
  industry: string
  country: string
  industryBeta?: number // levered industry beta
  bondYtm?: number // yield to maturity of the company's bonds, decimal (0.06 = 6%)
  governmentBond10y?: number // 10-year government bond yield of the country, decimal
  isInnovativeHighTech: boolean
  hasProductSales: boolean // false = still developing, no sales yet
  highRnD: boolean
  sources: string[]
}

export type GrowthAssumptions = {
  baseCagr: number
  inflation: number
  adjustments: { label: string; delta: number }[]
  inputs: GrowthInputs | null
  rationale: string
}

// Rules from the 5-year cash flow projections skill. Values the skill leaves
// open (bond-yield and country-risk bands, the no-sales cap) are marked and
// are meant to be tuned.
export const GROWTH_RULES = {
  baseCagr: 0.05, // ~2.0-2.5% inflation, so a good CAGR is about 5%
  inflation: PLATFORM_ASSUMPTIONS.annualInflationRate,
  betaHigh: { above: 1.5, delta: -0.01 },
  betaMid: { above: 1.0, delta: -0.005 },
  highTechWithSalesCap: 0.03, // "starts at 3% at best"
  highTechHighRnDCap: 0.015, // "even lower" — assumed
  floor: 0,
  // Assumed bands (not specified by the skill):
  bondSpreadBands: [{ above: 0.04, delta: -0.01 }, { above: 0.02, delta: -0.005 }],
  countryYieldBands: [{ above: 0.06, delta: -0.01 }, { above: 0.04, delta: -0.005 }],
}

const firstBand = (value: number | undefined, bands: { above: number; delta: number }[]) =>
  value === undefined ? undefined : bands.find((b) => value > b.above)

export function computeAdjustedCagr(inputs: GrowthInputs): GrowthAssumptions {
  const r = GROWTH_RULES
  const adjustments: { label: string; delta: number }[] = []

  const { industryBeta: beta } = inputs
  if (beta !== undefined && beta > r.betaHigh.above) adjustments.push({ label: `Industry beta ${beta.toFixed(2)} > 1.5`, delta: r.betaHigh.delta })
  else if (beta !== undefined && beta > r.betaMid.above) adjustments.push({ label: `Industry beta ${beta.toFixed(2)} between 1 and 1.5`, delta: r.betaMid.delta })

  // Company bond YTM, as spread over the country's 10y government bond when known.
  if (inputs.bondYtm !== undefined) {
    const spread = inputs.bondYtm - (inputs.governmentBond10y ?? r.inflation + 0.01)
    const band = firstBand(spread, r.bondSpreadBands)
    if (band) adjustments.push({ label: `Bond YTM ${(inputs.bondYtm * 100).toFixed(1)}% (spread ${(spread * 100).toFixed(1)}pp)`, delta: band.delta })
  }

  const countryBand = firstBand(inputs.governmentBond10y, r.countryYieldBands)
  if (countryBand) adjustments.push({ label: `${inputs.country} 10y government bond ${(inputs.governmentBond10y! * 100).toFixed(1)}%`, delta: countryBand.delta })

  let cagr = r.baseCagr + adjustments.reduce((sum, a) => sum + a.delta, 0)

  if (inputs.isInnovativeHighTech) {
    const cap = inputs.hasProductSales && !inputs.highRnD ? r.highTechWithSalesCap : r.highTechHighRnDCap
    if (cagr > cap) {
      adjustments.push({ label: `Innovative high-tech cap (${inputs.hasProductSales && !inputs.highRnD ? 'selling product' : 'still in heavy R&D / pre-sales'})`, delta: cap - cagr })
      cagr = cap
    }
  }

  cagr = Math.max(cagr, r.floor)
  return {
    baseCagr: cagr,
    inflation: r.inflation,
    adjustments,
    inputs,
    rationale: `Industry: ${inputs.industry}; country: ${inputs.country}. Starting from ${(r.baseCagr * 100).toFixed(1)}%, ${adjustments.length ? adjustments.map((a) => `${a.label} (${(a.delta * 100).toFixed(1)}pp)`).join('; ') : 'no risk adjustments applied'} => ${(cagr * 100).toFixed(1)}% base-case revenue CAGR.`,
  }
}

export function platformDefaultGrowth(reason: string): GrowthAssumptions {
  return {
    baseCagr: PLATFORM_ASSUMPTIONS.baseRevenueCagr,
    inflation: PLATFORM_ASSUMPTIONS.annualInflationRate,
    adjustments: [],
    inputs: null,
    rationale: `${reason} Platform default ${(PLATFORM_ASSUMPTIONS.baseRevenueCagr * 100).toFixed(1)}% CAGR used.`,
  }
}

const GROWTH_TOOL: Anthropic.Tool = {
  name: 'submit_growth_inputs',
  description: 'Submit the researched market inputs for the business. Do not compute a growth rate.',
  input_schema: {
    type: 'object',
    properties: {
      industry: { type: 'string' },
      country: { type: 'string', description: 'Country of the main operations' },
      industryBeta: { type: 'number', description: 'Current levered industry beta for this industry; omit if not found' },
      bondYtm: { type: 'number', description: 'Yield to maturity of bonds issued by the company, decimal; omit if none' },
      governmentBond10y: { type: 'number', description: 'Current 10-year government bond yield of the country, decimal; omit if not found' },
      isInnovativeHighTech: { type: 'boolean' },
      hasProductSales: { type: 'boolean', description: 'false if the company is still developing and not yet selling' },
      highRnD: { type: 'boolean', description: 'true if R&D expenses are a major cost' },
      sources: { type: 'array', items: { type: 'string' }, description: 'URLs or publisher names for beta and yields' },
    },
    required: ['industry', 'country', 'isInnovativeHighTech', 'hasProductSales', 'highRnD', 'sources'],
  },
}

const optionalNumber = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : undefined)

export function normalizeGrowthInputs(input: unknown): GrowthInputs {
  const raw = (input ?? {}) as Record<string, unknown>
  // Yields must be decimals; treat anything above 1 as a percentage the model forgot to convert.
  const rate = (v: unknown) => {
    const n = optionalNumber(v)
    return n === undefined ? undefined : n > 1 ? n / 100 : n
  }
  return {
    industry: typeof raw.industry === 'string' ? raw.industry : 'Unknown',
    country: typeof raw.country === 'string' ? raw.country : 'Unknown',
    industryBeta: optionalNumber(raw.industryBeta),
    bondYtm: rate(raw.bondYtm),
    governmentBond10y: rate(raw.governmentBond10y),
    isInnovativeHighTech: raw.isInnovativeHighTech === true,
    hasProductSales: raw.hasProductSales !== false,
    highRnD: raw.highRnD === true,
    sources: Array.isArray(raw.sources) ? raw.sources.filter((s): s is string => typeof s === 'string') : [],
  }
}

// Researches industry beta and country/bond yields for the uploaded business
// and returns the risk-adjusted growth assumptions. Any failure (no API
// access, search unavailable, bad response) falls back to the platform default
// so the report still generates.
export async function researchGrowthAssumptions(extracted: ExtractedDocument, filename: string): Promise<GrowthAssumptions> {
  try {
    const context = [
      extracted.narrative.executiveSummary,
      ...extracted.narrative.sections.map((s) => `${s.heading}: ${s.body ?? ''} ${(s.bullets ?? []).join('; ')}`),
      'Line items: ' + (extracted.lineItems as { label?: string }[]).map((li) => li?.label).filter(Boolean).slice(0, 80).join(', '),
    ].join('\n').slice(0, 12000)

    const anthropic = new Anthropic()
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-5-5',
      max_tokens: 4000,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 5 }, GROWTH_TOOL],
      tool_choice: { type: 'auto' },
      messages: [{
        role: 'user',
        content: `From the description of the business in "${filename}" below, identify its industry and country of operations, then use web search to find (1) the current industry beta, (2) the current 10-year government bond yield of that country, and (3) the yield to maturity of any bonds the company has issued. Also judge whether it is an innovative high-tech business, whether it already sells product, and whether R&D is a major cost. Report only what you can source; omit values you cannot find rather than guessing. Finish by calling submit_growth_inputs.\n\n${context}`,
      }],
    })

    const toolUse = message.content.find((p): p is Anthropic.ToolUseBlock => p.type === 'tool_use' && p.name === 'submit_growth_inputs')
    if (!toolUse) return platformDefaultGrowth('Market inputs could not be researched.')
    return computeAdjustedCagr(normalizeGrowthInputs(toolUse.input))
  } catch (error) {
    console.error('Growth assumption research failed', error)
    return platformDefaultGrowth('Market inputs could not be researched.')
  }
}
