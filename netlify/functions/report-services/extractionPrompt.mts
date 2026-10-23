import { MAX_TEXT_CHARS } from './constants.mts'
import { isNativeAnthropicType, isSpreadsheetType, isWordType, parseSpreadsheetToText, parseDocxToText } from './documentParsing.mts'

export async function buildMessageContent(fileBuffer: Buffer, filename: string, mimeType: string, reportLabel: string, instructions: string) {
  const prompt = `You are a senior financial analyst and CFO advisor. Use the uploaded file "${filename}" as the source material to produce a ${reportLabel}.

${instructions}

Call the submit_report tool with the fully structured result. Requirements:
- Extract real, concrete figures from the uploaded file wherever present; state assumptions explicitly when data is missing or ambiguous.
- Populate "lineItems" with every individual raw figure from the source document, each tagged with the correct "group" classification for this report type (see the group vocabulary in the instructions above). Do NOT sum, net, or otherwise pre-compute totals, ratios, or derived metrics yourself — list only atomic source figures; all arithmetic is performed downstream in code from your classifications.
- Use "assumptions" to record data caveats, any source items you deliberately excluded and why, your rationale for classifying borderline items into their groups, and named limitations of this analysis.
- Use "tables"/"charts" only for supplementary qualitative or non-deterministic content not covered by "lineItems" (see their descriptions above) — do not duplicate figures that "lineItems" already covers.
- Keep prose in "executiveSummary" and section "body" fields as plain text (no markdown syntax) — formatting is handled by the spreadsheet renderer.
- Write in professional English suitable for C-suite and investor audiences.
- Be concise: keep each table to at most ~15 rows and each section body to at most ~3 short paragraphs, so the full structured response fits comfortably within the output budget.`

  if (isNativeAnthropicType(mimeType)) {
    return [
      {
        type: mimeType.startsWith('image/') ? 'image' : 'document',
        source: {
          type: 'base64',
          media_type: mimeType,
          data: fileBuffer.toString('base64'),
        },
      },
      { type: 'text', text: prompt },
    ] as const
  }

  let extractedText: string
  if (isSpreadsheetType(mimeType)) {
    extractedText = parseSpreadsheetToText(fileBuffer)
  } else if (isWordType(mimeType)) {
    extractedText = await parseDocxToText(fileBuffer)
  } else {
    // Plain text formats (CSV, TXT) are already valid UTF-8.
    extractedText = fileBuffer.toString('utf8').replace(/[^\x20-\x7E\n\r\t]/g, ' ').slice(0, MAX_TEXT_CHARS)
  }

  return [
    {
      type: 'text',
      text: `File contents of "${filename}" extracted as text:\n\n\`\`\`\n${extractedText}\n\`\`\`\n\n${prompt}`,
    },
  ] as const
}
