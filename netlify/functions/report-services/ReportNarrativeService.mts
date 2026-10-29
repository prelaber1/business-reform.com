import Anthropic from '@anthropic-ai/sdk'
import type { ExtractedDocument } from './DocumentExtractionService.mts'
import type { LineItem, ReportData, ReportSection } from './types.mts'
import type { ReportModelBase } from './calculationSupport.mts'
import type { GrowthAssumptions } from './GrowthAssumptionService.mts'
import type { ValidationIssue } from './FinancialValidationService.mts'
import { buildVarianceCommentary } from './VarianceCommentaryService.mts'

// Everything the review needs, all of it already verified upstream.
export type NarrativeContext = {
  reportLabel: string
  issues?: ValidationIssue[]
  growth?: GrowthAssumptions
}

type Dimension = { name: string; status: 'pass' | 'flag' | 'fail'; headline: string }
type FailureStory = { title: string; whatHappened: string; earlyWarning: string; likelihood: string; damage: string }
type Requirement = { driver: string; required: string; evidence: string; grade: 'in-line' | 'aggressive' | 'stretch' | 'unknown' }

export type NarrativeReview = {
  verdict: 'ready' | 'close' | 'not_yet'
  verdictReason: string
  dimensions: Dimension[]
  fixFirst: string[]
  boardQuestions: { question: string; answered: boolean; note: string }[]
  failureStories: FailureStory[]
  unaskedQuestions: string[]
  mustBeTrue: Requirement[]
  biggestLever: string
}

const REVIEW_TOOL: Anthropic.Tool = {
  name: 'submit_review',
  description: 'Submit the CFO-style review of the computed report. Cite only figures that appear in the provided context.',
  input_schema: {
    type: 'object',
    properties: {
      verdict: { type: 'string', enum: ['ready', 'close', 'not_yet'] },
      verdictReason: { type: 'string' },
      dimensions: {
        type: 'array',
        description: 'Exactly five, in order: Completeness, Internal Consistency, Realism, Benchmark Sanity, Board Defensibility',
        items: {
          type: 'object',
          properties: { name: { type: 'string' }, status: { type: 'string', enum: ['pass', 'flag', 'fail'] }, headline: { type: 'string', description: '8 words or fewer' } },
          required: ['name', 'status', 'headline'],
        },
      },
      fixFirst: { type: 'array', items: { type: 'string' }, description: 'Top 3 issues, each with the one-line fix' },
      boardQuestions: {
        type: 'array',
        description: '2-3 questions a sharp board member or lender will ask',
        items: {
          type: 'object',
          properties: { question: { type: 'string' }, answered: { type: 'boolean' }, note: { type: 'string' } },
          required: ['question', 'answered', 'note'],
        },
      },
      failureStories: {
        type: 'array',
        description: 'The 2-3 most plausible ways this plan fails a year from now, each anchored in a load-bearing assumption',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            whatHappened: { type: 'string' },
            earlyWarning: { type: 'string', description: 'What would be visible 60-90 days in' },
            likelihood: { type: 'string', enum: ['High', 'Medium', 'Low'] },
            damage: { type: 'string', enum: ['High', 'Medium', 'Low'] },
          },
          required: ['title', 'whatHappened', 'earlyWarning', 'likelihood', 'damage'],
        },
      },
      unaskedQuestions: {
        type: 'array',
        items: { type: 'string' },
        description: 'Only the reversibility / base rate / capacity / dependency / opportunity-cost questions the report leaves unanswered, each phrased as a question',
      },
      mustBeTrue: {
        type: 'array',
        description: 'The assumptions the headline outcome implies (growth, capacity/hiring, margin/cash), each graded against what the source data shows',
        items: {
          type: 'object',
          properties: {
            driver: { type: 'string' },
            required: { type: 'string', description: 'What has to hold, with the arithmetic when figures are available' },
            evidence: { type: 'string', description: 'What the source data shows, or "No history provided"' },
            grade: { type: 'string', enum: ['in-line', 'aggressive', 'stretch', 'unknown'] },
          },
          required: ['driver', 'required', 'evidence', 'grade'],
        },
      },
      biggestLever: { type: 'string', description: 'The single assumption that, relaxed slightly, most reduces the stretch' },
    },
    required: ['verdict', 'verdictReason', 'dimensions', 'fixFirst', 'boardQuestions', 'failureStories', 'unaskedQuestions', 'mustBeTrue', 'biggestLever'],
  },
}

const SYSTEM_PROMPT = `You are a seasoned CFO reviewing a financial report before it goes to a board, lender or investor. Apply three fixed frameworks and return them through submit_review.

1. Readiness scorecard. Rate five dimensions in this order. Completeness: is anything missing (revenue build, churn, headcount link, cash view, below-the-line items)? Internal Consistency: do the numbers tie, does spend scale with growth, is timing logical? Realism: growth versus the company's own history, hiring and productivity assumptions, margin direction. Benchmark Sanity: key ratios against defensible ranges. Board Defensibility: the top challenges and whether the report answers them, and whether there is any scenario framing. Verdict: ready (no material holes), close (fix the listed items first), not_yet (material holes). A dimension's status is its worst check.

2. Pre-mortem. Assume a year from now the plan failed. Give the two or three most plausible failure stories, each anchored in a load-bearing assumption and ranked by likelihood and damage, with the early warning sign at 60-90 days. Do not pad with far-fetched risks; two low-risk stories are a valid answer for a safe plan. Then list only the unasked questions (reversibility, base rate, capacity, dependency, opportunity cost) that the report leaves unanswered.

3. What has to be true. Decompose the headline outcome into growth, capacity and margin/cash drivers, state what each must hit with the arithmetic, and grade it against the source data: in-line (within about 1.25x of the company's trend), aggressive (1.25x-2x with a named catalyst), stretch (about 2x or more with no catalyst). If the source gives no history for a driver, grade it unknown and say so; never invent a baseline or substitute industry benchmarks for the company's own history. Name the single biggest lever.

Rules: use only figures in the provided context and quote them exactly; do not recompute or invent numbers. Tie each finding to a specific figure or section. Do not soften real findings and do not manufacture risk. Plain text only, no markdown. Professional English.`

const STATUS_MARK = { pass: '✓', flag: '⚠', fail: '✗' } as const
const VERDICT_TEXT = {
  ready: 'Ready to send up',
  close: 'Close, fix the items below first',
  not_yet: 'Not yet, material holes',
} as const

const asStrings = (v: unknown): string[] => (Array.isArray(v) ? v.filter((s): s is string => typeof s === 'string') : [])
const asObjects = <T,>(v: unknown): T[] => (Array.isArray(v) ? (v.filter((o) => o && typeof o === 'object') as T[]) : [])

// A truncated or malformed tool response must degrade, not crash the report.
export function normalizeReview(input: unknown): NarrativeReview | null {
  const raw = (input ?? {}) as Record<string, unknown>
  if (typeof raw.verdict !== 'string' || !(raw.verdict in VERDICT_TEXT)) return null
  return {
    verdict: raw.verdict as NarrativeReview['verdict'],
    verdictReason: typeof raw.verdictReason === 'string' ? raw.verdictReason : '',
    dimensions: asObjects<Dimension>(raw.dimensions).filter((d) => d.name && d.status in STATUS_MARK),
    fixFirst: asStrings(raw.fixFirst),
    boardQuestions: asObjects<NarrativeReview['boardQuestions'][number]>(raw.boardQuestions),
    failureStories: asObjects<FailureStory>(raw.failureStories),
    unaskedQuestions: asStrings(raw.unaskedQuestions),
    mustBeTrue: asObjects<Requirement>(raw.mustBeTrue),
    biggestLever: typeof raw.biggestLever === 'string' ? raw.biggestLever : '',
  }
}

export function reviewToSections(review: NarrativeReview): ReportSection[] {
  const sections: ReportSection[] = []

  sections.push({
    heading: 'CFO Readiness Scorecard',
    body: `Verdict: ${VERDICT_TEXT[review.verdict]}. ${review.verdictReason}`.trim(),
    bullets: [
      ...review.dimensions.map((d) => `${STATUS_MARK[d.status]} ${d.name}: ${d.headline}`),
      ...review.fixFirst.map((f, i) => `Fix first ${i + 1}: ${f}`),
      ...review.boardQuestions.map((q) => `Board question: "${q.question}" - ${q.answered ? 'answered' : 'no answer yet'}: ${q.note}`),
    ],
  })

  if (review.failureStories.length || review.unaskedQuestions.length) {
    sections.push({
      heading: 'Pre-Mortem: How This Could Fail',
      bullets: [
        ...review.failureStories.map((s) => `${s.title} (likelihood ${s.likelihood}, damage ${s.damage}): ${s.whatHappened} Early warning: ${s.earlyWarning}`),
        ...review.unaskedQuestions.map((q) => `Unanswered: ${q}`),
      ],
    })
  }

  if (review.mustBeTrue.length) {
    sections.push({
      heading: 'What Has To Be True',
      bullets: [
        ...review.mustBeTrue.map((r) => `${r.driver} [${r.grade}]: ${r.required} Source data: ${r.evidence}`),
        ...(review.biggestLever ? [`Biggest lever: ${review.biggestLever}`] : []),
      ],
    })
  }
  return sections
}

const describeModel = (model: ReportModelBase | null) =>
  model
    ? [
        'Computed key metrics:',
        ...model.keyMetrics.map((m) => `- ${m.label}: ${m.value}`),
        ...(model.narrativeFacts ?? []).map((f) => `Fact: ${f}`),
        ...model.extraCaveats.map((c) => `Caveat: ${c}`),
      ].join('\n')
    : 'No computed model for this report type.'

// Second Claude pass. The extraction call already wrote the report's prose;
// this one reviews the computed result and appends the scorecard, pre-mortem
// and what-has-to-be-true sections. If it fails the report ships without them.
export async function buildReportNarrative(
  extracted: ExtractedDocument,
  lineItems: LineItem[],
  model: ReportModelBase | null,
  context?: NarrativeContext,
): Promise<ReportData> {
  const base: ReportData = {
    ...extracted.narrative,
    lineItems,
    assumptions: extracted.assumptions,
  }
  if (!context) return base

  // Variance reports lead with driver commentary, ahead of the generic CFO review.
  if (model?.narrativeFacts?.length) {
    const commentary = await buildVarianceCommentary(extracted, model.narrativeFacts)
    base.sections = [...commentary, ...base.sections]
  }

  const dataQuality = (context.issues ?? []).filter((i) => i.severity === 'warning')
  if (dataQuality.length) {
    base.assumptions = {
      ...base.assumptions,
      dataCaveats: [...base.assumptions.dataCaveats, ...dataQuality.map((i) => `Data check: ${i.message}`)],
    }
  }

  try {
    const anthropic = new Anthropic()
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-5-5',
      max_tokens: 6000,
      system: SYSTEM_PROMPT,
      tools: [REVIEW_TOOL],
      tool_choice: { type: 'auto' },
      messages: [{
        role: 'user',
        content: [
          `Report type: ${context.reportLabel}`,
          `Executive summary drafted from the source file:\n${extracted.narrative.executiveSummary}`,
          `Sections drafted so far:\n${extracted.narrative.sections.map((s) => `${s.heading}: ${s.body ?? ''} ${(s.bullets ?? []).join('; ')}`).join('\n')}`,
          `Source line items (${lineItems.length}):\n${lineItems.slice(0, 150).map((li) => `${li.group} | ${li.label} | ${li.amount}${li.period ? ` | ${li.period}` : ''}`).join('\n')}`,
          describeModel(model),
          context.growth ? `Growth assumption used: ${context.growth.rationale}` : '',
          `Data-quality checks run on the extraction:\n${dataQuality.length ? dataQuality.map((i) => `- ${i.message}`).join('\n') : '- none raised'}`,
          'Call submit_review.',
        ].filter(Boolean).join('\n\n'),
      }],
    })

    const toolUse = message.content.find((p): p is Anthropic.ToolUseBlock => p.type === 'tool_use' && p.name === 'submit_review')
    const review = toolUse ? normalizeReview(toolUse.input) : null
    if (!review) {
      console.warn('Narrative review returned no usable result; report generated without it.')
      return base
    }
    return { ...base, sections: [...base.sections, ...reviewToSections(review)] }
  } catch (error) {
    console.error('Narrative review failed; report generated without it.', error)
    return base
  }
}
