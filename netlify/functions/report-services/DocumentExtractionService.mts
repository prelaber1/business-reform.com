import Anthropic from '@anthropic-ai/sdk'
import { REPORT_DEFINITIONS } from './reportDefinitions.mts'
import { REPORT_TOOL } from './extractionSchema.mts'
import { buildMessageContent } from './extractionPrompt.mts'
import type { ReportAssumptions, ReportData } from './types.mts'

export type DocumentExtractionInput = {
  fileBuffer: Buffer
  filename: string
  mimeType: string
  reportType: string
}

// What Claude returned for the uploaded document, normalized to a predictable
// shape but not yet validated. `lineItems` is deliberately `unknown[]`: deciding
// which items are usable is FinancialValidationService's job, not this one's.
export type ExtractedDocument = {
  lineItems: unknown[]
  assumptions: ReportAssumptions
  narrative: Pick<ReportData, 'executiveSummary' | 'sections' | 'tables' | 'charts' | 'nextSteps'>
}

const asArray = <T,>(value: unknown): T[] => (Array.isArray(value) ? (value as T[]) : [])

// Calls Claude once with the uploaded document and converts its tool-use
// response into typed inputs. A truncated generation (e.g. hitting max_tokens
// mid-way) can leave the tool input missing fields entirely even though the
// schema marks them required, so everything is normalized here and a partial
// response degrades gracefully instead of crashing the whole job.
export async function extractDocument(input: DocumentExtractionInput): Promise<ExtractedDocument> {
  const { fileBuffer, filename, mimeType, reportType } = input
  const reportDef = REPORT_DEFINITIONS[reportType]

  const anthropic = new Anthropic()
  const messageContent = await buildMessageContent(fileBuffer, filename, mimeType, reportDef.label, reportDef.instructions)
  const message = await anthropic.messages.create({
    model: 'claude-sonnet-5-5',
    max_tokens: 16000,
    tools: [REPORT_TOOL],
    // Sonnet 5.5 rejects forced tool_choice (type 'tool'/'any'); the prompt
    // already instructs the model to call submit_report.
    tool_choice: { type: 'auto' },
    messages: [{ role: 'user', content: messageContent }],
  })

  if (message.stop_reason === 'max_tokens') {
    console.error('Anthropic response hit max_tokens — the structured report may be truncated/incomplete.')
  }

  const toolUse = message.content.find(
    (part): part is Anthropic.ToolUseBlock => part.type === 'tool_use' && part.name === 'submit_report',
  )
  if (!toolUse) throw new Error('The AI worker returned an unexpected response.')

  return normalizeToolInput(toolUse.input)
}

export function normalizeToolInput(input: unknown): ExtractedDocument {
  const raw = (input ?? {}) as Record<string, unknown>
  const rawAssumptions = raw.assumptions as Partial<ReportAssumptions> | undefined
  return {
    lineItems: asArray(raw.lineItems),
    assumptions: {
      dataCaveats: asArray(rawAssumptions?.dataCaveats),
      excludedItems: asArray(rawAssumptions?.excludedItems),
      classificationRationale: asArray(rawAssumptions?.classificationRationale),
      limitations: asArray(rawAssumptions?.limitations),
    },
    narrative: {
      executiveSummary: typeof raw.executiveSummary === 'string' ? raw.executiveSummary : '',
      sections: asArray(raw.sections),
      tables: asArray(raw.tables),
      charts: asArray(raw.charts),
      nextSteps: asArray(raw.nextSteps),
    },
  }
}
