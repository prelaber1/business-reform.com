import Anthropic from '@anthropic-ai/sdk'
import type { ExtractedDocument } from './DocumentExtractionService.mts'
import type { ReportSection } from './types.mts'

// Driver types follow the variance-analysis rubric: decompose, explain, flag.
const DRIVER_TYPES = ['volume', 'price_rate', 'mix', 'timing', 'one_time', 'efficiency', 'other', 'not_documented'] as const
type Driver = { line: string; driverType: (typeof DRIVER_TYPES)[number]; explanation: string; recurring: 'recurring' | 'one-time' | 'unknown'; action: string }
export type VarianceCommentary = { headline: string; drivers: Driver[]; ownerQuestions: string[]; outlook: string }

const COMMENTARY_TOOL: Anthropic.Tool = {
  name: 'submit_variance_commentary',
  description: 'Submit leadership commentary on the computed budget vs actual variances.',
  input_schema: {
    type: 'object',
    properties: {
      headline: { type: 'string', description: 'Two sentences: the net operating result variance and the one or two lines that explain most of it' },
      drivers: {
        type: 'array',
        description: 'One entry per MATERIAL line given in the context, largest impact first',
        items: {
          type: 'object',
          properties: {
            line: { type: 'string', description: 'Exact line label' },
            driverType: { type: 'string', enum: [...DRIVER_TYPES] },
            explanation: { type: 'string', description: 'Why it moved, using the provided volume/price split and the source narrative only. If the source does not say, state that the driver is not documented.' },
            recurring: { type: 'string', enum: ['recurring', 'one-time', 'unknown'] },
            action: { type: 'string', description: 'The one follow-up or decision this variance calls for' },
          },
          required: ['line', 'driverType', 'explanation', 'recurring', 'action'],
        },
      },
      ownerQuestions: { type: 'array', items: { type: 'string' }, description: 'Questions for budget owners on drivers the source leaves unexplained' },
      outlook: { type: 'string', description: 'Whether the variance pattern suggests a reforecast, and which lines it would touch' },
    },
    required: ['headline', 'drivers', 'ownerQuestions', 'outlook'],
  },
}

const SYSTEM_PROMPT = `You are an FP&A lead writing variance commentary for leadership. Rules:
- Every figure is already computed and given to you; quote them exactly and never recompute or invent numbers.
- Explain only the lines flagged MATERIAL. Decompose each into drivers: volume, price/rate, mix, timing, one-time, efficiency. Use the volume and price/rate split when provided.
- Take explanations only from the source narrative given to you. When the source does not explain a variance, set driverType to not_documented and say so; do not guess a cause, and put the question to the budget owner instead.
- Say whether each variance is likely recurring or one-time only if the source supports it; otherwise unknown.
- Favorable and unfavorable are already determined; do not reverse them.
- Plain text only, no markdown. Professional English, concise.`

const asStrings = (v: unknown): string[] => (Array.isArray(v) ? v.filter((s): s is string => typeof s === 'string') : [])

export function normalizeCommentary(input: unknown): VarianceCommentary | null {
  const raw = (input ?? {}) as Record<string, unknown>
  if (typeof raw.headline !== 'string' || !raw.headline) return null
  const drivers = (Array.isArray(raw.drivers) ? raw.drivers : [])
    .filter((d): d is Driver => !!d && typeof d === 'object' && typeof (d as Driver).line === 'string' && typeof (d as Driver).explanation === 'string')
    .map((d) => ({ ...d, driverType: DRIVER_TYPES.includes(d.driverType) ? d.driverType : ('other' as const) }))
  return { headline: raw.headline, drivers, ownerQuestions: asStrings(raw.ownerQuestions), outlook: typeof raw.outlook === 'string' ? raw.outlook : '' }
}

const DRIVER_LABEL: Record<Driver['driverType'], string> = {
  volume: 'Volume',
  price_rate: 'Price/rate',
  mix: 'Mix',
  timing: 'Timing',
  one_time: 'One-time',
  efficiency: 'Efficiency',
  other: 'Other',
  not_documented: 'Driver not documented',
}

export function commentaryToSections(c: VarianceCommentary): ReportSection[] {
  const sections: ReportSection[] = [{ heading: 'Variance Commentary', body: c.headline }]
  if (c.drivers.length) {
    sections.push({
      heading: 'Material Variances and Drivers',
      bullets: c.drivers.map((d) => `${d.line} [${DRIVER_LABEL[d.driverType]}, ${d.recurring}]: ${d.explanation} Action: ${d.action}`),
    })
  }
  if (c.ownerQuestions.length || c.outlook) {
    sections.push({
      heading: 'Open Questions and Outlook',
      bullets: [...c.ownerQuestions.map((q) => `Ask budget owner: ${q}`), ...(c.outlook ? [`Outlook: ${c.outlook}`] : [])],
    })
  }
  return sections
}

// Second Claude call over the computed variances. Any
// failure returns no sections and the report ships with the numbers alone.
export async function buildVarianceCommentary(extracted: ExtractedDocument, facts: string[]): Promise<ReportSection[]> {
  if (!facts.length) return []
  try {
    const anthropic = new Anthropic()
    const message = await anthropic.messages.create({
      model: 'claude-sonnet-5-5',
      max_tokens: 3000,
      system: SYSTEM_PROMPT,
      tools: [COMMENTARY_TOOL],
      tool_choice: { type: 'tool', name: COMMENTARY_TOOL.name },
      messages: [{
        role: 'user',
        content: [
          `Computed variances:\n${facts.join('\n')}`,
          `Source narrative drafted from the uploaded file:\n${extracted.narrative.executiveSummary}\n${extracted.narrative.sections.map((s) => `${s.heading}: ${s.body ?? ''} ${(s.bullets ?? []).join('; ')}`).join('\n')}`,
          'Call submit_variance_commentary.',
        ].join('\n\n'),
      }],
    })
    const toolUse = message.content.find((p): p is Anthropic.ToolUseBlock => p.type === 'tool_use' && p.name === COMMENTARY_TOOL.name)
    const commentary = toolUse ? normalizeCommentary(toolUse.input) : null
    if (!commentary) {
      console.warn('Variance commentary returned no usable result; report generated without it.')
      return []
    }
    return commentaryToSections(commentary)
  } catch (error) {
    console.error('Variance commentary failed; report generated without it.', error)
    return []
  }
}
