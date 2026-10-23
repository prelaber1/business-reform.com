// Must run before `sharp` is imported — see the comment inside for why.
import './fonts/setup-fonts.mts'
import Anthropic from '@anthropic-ai/sdk'
import { getStore } from '@netlify/blobs'
import { getUser } from '@netlify/identity'
import type { Config, Context } from '@netlify/functions'
import * as XLSX from 'xlsx'
import mammoth from 'mammoth'
import ExcelJS from 'exceljs'
import sharp from 'sharp'
import { LOGO_PNG_BASE64, LOGO_WIDTH, LOGO_HEIGHT } from './logo-data.mts'
import { AI_MARK, AI_METADATA_DESCRIPTION, AI_METADATA_KEYWORDS, aiMarkLine, applyAiMarkToWorkbook } from './ai-disclosure.mts'
// pptxgenjs's published dual ESM/CJS package confused Netlify's function
// bundler two different ways in a row: a plain `import` resolved to its
// ESM build (which isn't actually marked as a module) and crashed with
// "Cannot use import statement outside a module"; a runtime require() of
// the npm package wasn't recognized by the bundler's dependency scan and
// got left out of the deploy entirely ("Cannot find module"). Vendoring a
// pre-bundled, fully self-contained copy sidesteps both failure modes — a
// relative import of a local file is always reliably included by every
// bundler, no external-module detection involved. Regenerate this file with
// `npx esbuild --bundle --platform=node --format=esm --outfile=netlify/functions/vendor/pptxgenjs.bundle.mjs <entry re-exporting pptxgenjs default>`
// after any pptxgenjs version bump.
import PptxGenJS from './vendor/pptxgenjs.bundle.mjs'
import type PptxGenJSType from 'pptxgenjs'

const UPLOADS_STORE = 'client-uploads'
const REPORTS_STORE = 'client-report-jobs'
const MAX_TEXT_CHARS = 50000

const NAVY = 'FF1F3B5B'
const LIGHT_BAND = 'FFF4F7FB'
const CHART_PALETTE = ['#1f3b5b', '#c9a227', '#1d6b3a', '#b3261e', '#6b46c1', '#2c5380']

type JobStatus = 'processing' | 'complete' | 'failed'

type ReportJob = {
  id: string
  userId: string
  fileKey: string
  fileName: string
  reportType: string
  reportLabel: string
  status: JobStatus
  createdAt: string
  updatedAt: string
  workerToken: string
  downloadKey?: string
  downloadKeyPptx?: string
  error?: string
}

type ReportSection = { heading: string; body?: string; bullets?: string[] }
type ReportTable = { title: string; columns: string[]; rows: (string | number)[][] }
type ReportChartSeries = { name: string; values: number[] }
type ReportChart = { type: 'line' | 'bar'; title: string; categories: string[]; series: ReportChartSeries[] }

// A single atomic figure extracted from the source document, classified with a
// semantic `group` tag (vocabulary varies per report type — see REPORT_DEFINITIONS).
// The AI supplies these raw and unaggregated; all arithmetic on top of them is
// performed deterministically in code (see "Deterministic calculation engine" below)
// and rendered as real cross-referencing Excel formulas, never as AI-computed numbers.
export type LineItem = { label: string; account?: string; group: string; amount: number; period?: string }
export type ReportAssumptions = {
  dataCaveats: string[]
  excludedItems: string[]
  classificationRationale: string[]
  limitations: string[]
}
export type ReportData = {
  executiveSummary: string
  sections: ReportSection[]
  lineItems: LineItem[]
  assumptions: ReportAssumptions
  tables: ReportTable[]
  charts: ReportChart[]
  nextSteps: string[]
}

const REPORT_DEFINITIONS: Record<string, { label: string; instructions: string }> = {
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
}

function json(data: unknown, status = 200) {
  return Response.json(data, { status })
}

function getMimeType(filename: string, fallback?: string) {
  if (fallback) return fallback
  const ext = (filename.split('.').pop() || '').toLowerCase()
  const map: Record<string, string> = {
    pdf: 'application/pdf',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    webp: 'image/webp',
    csv: 'text/plain',
    txt: 'text/plain',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }
  return map[ext] || 'application/octet-stream'
}

function isNativeAnthropicType(mimeType: string) {
  return ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp'].includes(mimeType)
}

function isSpreadsheetType(mimeType: string) {
  return [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ].includes(mimeType)
}

function isWordType(mimeType: string) {
  return [
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ].includes(mimeType)
}

function parseSpreadsheetToText(fileBuffer: Buffer) {
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' })
  const sheets = workbook.SheetNames.map((sheetName) => {
    const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName])
    return `--- Sheet: ${sheetName} ---\n${csv}`
  })
  return sheets.join('\n\n').slice(0, MAX_TEXT_CHARS)
}

async function parseDocxToText(fileBuffer: Buffer) {
  const result = await mammoth.extractRawText({ buffer: fileBuffer })
  return result.value.slice(0, MAX_TEXT_CHARS)
}

function getReportDownloadFilename(reportLabel: string, extension: 'xlsx' | 'pptx' = 'xlsx') {
  const safeName = reportLabel.replace(/[^\w.\- ]+/g, '_').trim() || 'Generated Report'
  return `${safeName}.${extension}`
}

// ── Chart rendering (hand-rolled SVG -> PNG, no headless browser / canvas needed) ──

function escapeXml(value: string) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function formatAxisNumber(n: number) {
  const abs = Math.abs(n)
  if (abs >= 1_000_000) return `${+(n / 1_000_000).toFixed(1)}M`
  if (abs >= 1_000) return `${+(n / 1_000).toFixed(1)}K`
  return `${Math.round(n * 100) / 100}`
}

function buildChartSvg(chart: ReportChart) {
  const width = 900
  const height = 420
  const marginLeft = 70
  const marginRight = 30
  const marginTop = 50
  const marginBottom = 70
  const plotWidth = width - marginLeft - marginRight
  const plotHeight = height - marginTop - marginBottom

  const allValues = chart.series.flatMap((s) => s.values)
  const maxVal = Math.max(0, ...allValues)
  const minVal = Math.min(0, ...allValues)
  const range = maxVal - minVal || 1

  const gridLines = 5
  const yTicks = Array.from({ length: gridLines + 1 }, (_, i) => minVal + (range * i) / gridLines)
  const yScale = (v: number) => marginTop + plotHeight - ((v - minVal) / range) * plotHeight

  let gridSvg = ''
  let yLabelsSvg = ''
  for (const tick of yTicks) {
    const y = yScale(tick)
    gridSvg += `<line x1="${marginLeft}" y1="${y}" x2="${width - marginRight}" y2="${y}" stroke="#e3e9f0" stroke-width="1" />`
    yLabelsSvg += `<text x="${marginLeft - 10}" y="${y + 4}" font-size="11" fill="#5a6b7d" text-anchor="end" font-family="DejaVu Sans, sans-serif">${formatAxisNumber(tick)}</text>`
  }

  const categoryCount = chart.categories.length
  const xStep = categoryCount > 1 ? plotWidth / (categoryCount - 1) : plotWidth
  let xLabelsSvg = ''
  chart.categories.forEach((cat, i) => {
    const x = categoryCount > 1 ? marginLeft + i * xStep : marginLeft + plotWidth / 2
    xLabelsSvg += `<text x="${x}" y="${height - marginBottom + 20}" font-size="11" fill="#5a6b7d" text-anchor="middle" font-family="DejaVu Sans, sans-serif">${escapeXml(cat)}</text>`
  })

  let seriesSvg = ''
  let legendSvg = ''
  const legendItemWidth = Math.min(160, plotWidth / Math.max(1, chart.series.length))
  const legendStartX = marginLeft + Math.max(0, (plotWidth - legendItemWidth * chart.series.length) / 2)

  if (chart.type === 'line') {
    chart.series.forEach((s, si) => {
      const color = CHART_PALETTE[si % CHART_PALETTE.length]
      const points = s.values
        .map((v, i) => {
          const x = categoryCount > 1 ? marginLeft + i * xStep : marginLeft + plotWidth / 2
          return `${x},${yScale(v)}`
        })
        .join(' ')
      seriesSvg += `<polyline points="${points}" fill="none" stroke="${color}" stroke-width="2.5" />`
      s.values.forEach((v, i) => {
        const x = categoryCount > 1 ? marginLeft + i * xStep : marginLeft + plotWidth / 2
        seriesSvg += `<circle cx="${x}" cy="${yScale(v)}" r="3.5" fill="${color}" />`
      })
      legendSvg += `<rect x="${legendStartX + si * legendItemWidth}" y="${height - 22}" width="12" height="12" fill="${color}" /><text x="${legendStartX + si * legendItemWidth + 18}" y="${height - 12}" font-size="11" fill="#2b2b2b" font-family="DejaVu Sans, sans-serif">${escapeXml(s.name)}</text>`
    })
  } else {
    const groupWidth = categoryCount > 0 ? plotWidth / categoryCount : plotWidth
    const barGap = 4
    const barWidth = Math.max(4, (groupWidth - barGap * (chart.series.length + 1)) / Math.max(1, chart.series.length))
    chart.categories.forEach((_cat, ci) => {
      const groupX = marginLeft + ci * groupWidth
      chart.series.forEach((s, si) => {
        const color = CHART_PALETTE[si % CHART_PALETTE.length]
        const v = s.values[ci] ?? 0
        const y0 = yScale(0)
        const y1 = yScale(v)
        const barX = groupX + barGap + si * (barWidth + barGap)
        seriesSvg += `<rect x="${barX}" y="${Math.min(y0, y1)}" width="${barWidth}" height="${Math.abs(y1 - y0)}" fill="${color}" />`
      })
    })
    chart.series.forEach((s, si) => {
      const color = CHART_PALETTE[si % CHART_PALETTE.length]
      legendSvg += `<rect x="${legendStartX + si * legendItemWidth}" y="${height - 22}" width="12" height="12" fill="${color}" /><text x="${legendStartX + si * legendItemWidth + 18}" y="${height - 12}" font-size="11" fill="#2b2b2b" font-family="DejaVu Sans, sans-serif">${escapeXml(s.name)}</text>`
    })
  }

  return `<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">
<rect width="${width}" height="${height}" fill="#ffffff" />
<text x="${width / 2}" y="26" font-size="16" font-weight="bold" fill="#1f3b5b" text-anchor="middle" font-family="DejaVu Sans, sans-serif">${escapeXml(chart.title)}</text>
${gridSvg}
<line x1="${marginLeft}" y1="${marginTop}" x2="${marginLeft}" y2="${height - marginBottom}" stroke="#b9c8d8" stroke-width="1" />
<line x1="${marginLeft}" y1="${height - marginBottom}" x2="${width - marginRight}" y2="${height - marginBottom}" stroke="#b9c8d8" stroke-width="1" />
${yLabelsSvg}
${xLabelsSvg}
${seriesSvg}
${legendSvg}
</svg>`
}

async function renderChartPng(chart: ReportChart) {
  const svg = buildChartSvg(chart)
  return sharp(Buffer.from(svg)).png().toBuffer()
}

// Renders every chart's PNG exactly once, shared across the Excel workbook
// (Summary sheet + dedicated Charts sheet) and the PowerPoint deck, instead
// of re-rendering the same SVG->PNG conversion for each consumer.
async function renderChartPngCache(charts: ReportChart[]): Promise<Map<ReportChart, Buffer>> {
  const cache = new Map<ReportChart, Buffer>()
  for (const chart of charts) {
    try {
      cache.set(chart, await renderChartPng(chart))
    } catch (error) {
      console.error('Chart render failed, skipping', chart.title, error)
    }
  }
  return cache
}

// ── AI report structuring ──────────────────────────────────────────────────

const REPORT_TOOL: Anthropic.Tool = {
  name: 'submit_report',
  description: 'Submit the fully structured financial report content for rendering into a formatted Excel workbook.',
  input_schema: {
    type: 'object',
    properties: {
      executiveSummary: {
        type: 'string',
        description: 'A 2-4 paragraph plain-prose executive summary (no markdown syntax, no headings).',
      },
      sections: {
        type: 'array',
        description: 'Analytical sections of the report, each with a heading and a prose body and/or bullet points.',
        items: {
          type: 'object',
          properties: {
            heading: { type: 'string' },
            body: { type: 'string', description: 'Plain prose paragraph(s), no markdown syntax.' },
            bullets: { type: 'array', items: { type: 'string' } },
          },
          required: ['heading'],
        },
      },
      lineItems: {
        type: 'array',
        description: 'Every individual raw financial line item extracted from the source document, classified with a semantic "group" tag (the allowed group values for this report type are listed in the instructions above). List ONLY atomic figures exactly as they appear in the source — do NOT pre-aggregate, subtotal, net, or compute derived ratios/metrics here. Downstream code computes all totals and derived metrics deterministically from these classifications.',
        items: {
          type: 'object',
          properties: {
            label: { type: 'string', description: 'Human-readable description, e.g. the account description from the source document.' },
            account: { type: 'string', description: 'Source account code if present, e.g. "60", "61.00". Omit if not applicable.' },
            group: { type: 'string', description: 'Semantic classification tag — use only the allowed values listed in the instructions for this report type.' },
            amount: { type: 'number' },
            period: { type: 'string', description: 'Time period label if the source spans multiple periods, e.g. "Year 1", "Month 3". Omit for single-period figures.' },
          },
          required: ['label', 'group', 'amount'],
        },
      },
      assumptions: {
        type: 'object',
        description: 'Structured notes on data quality, exclusions, and classification decisions, rendered on a dedicated Assumptions & Notes sheet.',
        properties: {
          dataCaveats: { type: 'array', items: { type: 'string' }, description: 'Caveats about the source data itself: period covered, completeness, currency, estimates used.' },
          excludedItems: { type: 'array', items: { type: 'string' }, description: 'Line items seen in the source but deliberately excluded from this analysis, and why (e.g. prior-year equity carryforwards, non-operating items).' },
          classificationRationale: { type: 'array', items: { type: 'string' }, description: 'Why specific accounts/items were tagged into their groups, especially for borderline or non-obvious cases.' },
          limitations: { type: 'array', items: { type: 'string' }, description: 'Named limitations of this analysis or model.' },
        },
        required: ['dataCaveats', 'excludedItems', 'classificationRationale', 'limitations'],
      },
      tables: {
        type: 'array',
        description: 'Supplementary structured tables for qualitative or non-deterministic data that does not belong in "lineItems" (e.g. a competitor comparison, a risk-factor matrix, a qualitative scoring table). Do not use this for figures that "lineItems" already covers — those are computed deterministically downstream, not tabulated here.',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            columns: { type: 'array', items: { type: 'string' } },
            rows: { type: 'array', items: { type: 'array', items: { type: ['string', 'number'] } } },
          },
          required: ['title', 'columns', 'rows'],
        },
      },
      charts: {
        type: 'array',
        description:
          'Supplementary charts for qualitative or non-deterministic comparisons only. Charts that visualize a quantity derivable from "lineItems" (e.g. cost vs. volume, burn schedule, allocation breakdown, capital structure, 5-year trends) are generated automatically from the computed model — do not duplicate those here. Use "line" for trends over time and "bar" for category comparisons.',
        items: {
          type: 'object',
          properties: {
            type: { type: 'string', enum: ['line', 'bar'] },
            title: { type: 'string' },
            categories: { type: 'array', items: { type: 'string' }, description: 'X-axis labels, e.g. month or year names.' },
            series: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  name: { type: 'string' },
                  values: { type: 'array', items: { type: 'number' } },
                },
                required: ['name', 'values'],
              },
            },
          },
          required: ['type', 'title', 'categories', 'series'],
        },
      },
      nextSteps: {
        type: 'array',
        items: { type: 'string' },
        description: 'Concrete, actionable next steps.',
      },
    },
    required: ['executiveSummary', 'sections', 'lineItems', 'assumptions', 'tables', 'charts', 'nextSteps'],
  },
}

async function buildMessageContent(fileBuffer: Buffer, filename: string, mimeType: string, reportLabel: string, instructions: string) {
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

// ── Deterministic calculation engine (no AI involved) ───────────────────────
//
// The AI's only job (see REPORT_TOOL / REPORT_DEFINITIONS above) is to extract
// and classify raw `lineItems` into semantic groups. Every derived metric
// below — subtotals, ratios, break-even points, KPIs — is computed here in
// plain TypeScript and rendered as a real cross-sheet Excel formula with a
// cached numeric result, so the workbook is a genuine live spreadsheet (edit a
// raw figure, everything downstream recalculates) instead of static,
// unverifiable AI-computed numbers.

const RESULT_FILL = 'FFC6E0B4'
const CAVEAT_FILL = 'FFFCE4D6'
const SUBTOTAL_FILL = 'FFDDEBF7'
const BORDER_COLOR = 'FFB7B7B7'
const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: BORDER_COLOR } },
  left: { style: 'thin', color: { argb: BORDER_COLOR } },
  bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
  right: { style: 'thin', color: { argb: BORDER_COLOR } },
}

type Formula = { formula: string; result: number }
// 'ratio' = a plain multiple (e.g. current ratio 1.85), shown as "1.85x";
// 'number' = a plain count such as months of runway.
type MetricFormat = 'currency' | 'percent' | 'ratio' | 'number'
function fx(formula: string, result: number): Formula {
  return { formula, result }
}
// A metric cell is a live formula, a plain editable numeric input (e.g. a
// target ratio or growth-rate assumption the user may want to tweak), or a
// plain-text placeholder (e.g. "N/A (no revenue data)") when the underlying
// figure genuinely can't be computed — never a broken formula / #DIV0! left
// for the user to puzzle over.
type MetricValue = Formula | string | number

function sumGroup(items: LineItem[]): number {
  return items.reduce((total, item) => total + (Number.isFinite(item.amount) ? item.amount : 0), 0)
}

function groupItems(lineItems: LineItem[], tag: string): LineItem[] {
  return lineItems.filter((item) => item.group === tag && Number.isFinite(item.amount))
}

function refCell(sheetName: string, row: number, col: string): string {
  return `'${sheetName}'!${col}${row}`
}

// ---- Raw-data sheet model: account-level line items grouped into contiguous
// blocks (one per classification tag), each followed by a SUM() subtotal row.
// Fixed column layout: A=Account, B=Description, C=Group, D=Amount.

type RawDataItemRow = { row: number; kind: 'item'; account?: string; description: string; groupLabel: string; amount: number }
type RawDataTotalRow = { row: number; kind: 'subtotal' | 'total'; label: string; value: MetricValue }
type RawDataRow = RawDataItemRow | RawDataTotalRow
type RawDataBlock = { tag: string; label: string; subtotalRow: number; subtotalValue: number }
type RawDataSheetModel = { sheetName: string; title: string; rows: RawDataRow[]; blocks: RawDataBlock[]; lastRow: number }

const RAW_DATA_FIRST_DATA_ROW = 4 // row 1 = title, row 2 = blank, row 3 = column headers

function layoutRawDataSheet(
  sheetName: string,
  title: string,
  lineItems: LineItem[],
  groupOrder: { tag: string; label: string }[],
): RawDataSheetModel {
  const rows: RawDataRow[] = []
  const blocks: RawDataBlock[] = []
  let row = RAW_DATA_FIRST_DATA_ROW

  for (const { tag, label } of groupOrder) {
    const items = groupItems(lineItems, tag)
    if (items.length === 0) continue
    const blockStart = row
    for (const item of items) {
      rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: label, amount: item.amount })
      row += 1
    }
    const blockEnd = row - 1
    const subtotalValue = sumGroup(items)
    rows.push({ row, kind: 'subtotal', label: `Total ${label}`, value: fx(`=SUM(D${blockStart}:D${blockEnd})`, subtotalValue) })
    blocks.push({ tag, label, subtotalRow: row, subtotalValue })
    row += 1
  }

  return { sheetName, title, rows, blocks, lastRow: row - 1 }
}

function findBlock(model: RawDataSheetModel, tag: string) {
  return model.blocks.find((b) => b.tag === tag)
}

// Appends a derived row (e.g. "Total Operating Expenses", "Net Result") below
// the last block/row on a raw-data sheet, referencing other rows on that same
// sheet. Returns the row number so calculators can reference it in turn.
function addTotalRow(model: RawDataSheetModel, label: string, value: MetricValue): number {
  const row = model.lastRow + 1
  model.rows.push({ row, kind: 'total', label, value })
  model.lastRow = row
  return row
}

// ---- Calculations sheet model: label/value metric rows grouped into
// optionally-headed sections (e.g. "Scenario A", "Scenario B"). Fixed column
// layout: A=spacer, B=label, C=value.

type CalcRow =
  | { row: number; kind: 'sectionHeading'; label: string }
  | { row: number; kind: 'columnHeader' }
  | { row: number; kind: 'metric'; label: string; value: MetricValue; bold?: boolean; highlight?: boolean; format?: MetricFormat }

type CalcSheetModel = { sheetName: string; title: string; rows: CalcRow[]; lastRow: number }

function newCalcSheet(sheetName: string, title: string): CalcSheetModel {
  return { sheetName, title, rows: [], lastRow: 2 }
}

function addSection(model: CalcSheetModel, heading: string) {
  model.lastRow += 1
  model.rows.push({ row: model.lastRow, kind: 'sectionHeading', label: heading })
  model.lastRow += 1
  model.rows.push({ row: model.lastRow, kind: 'columnHeader' })
}

function addMetric(
  model: CalcSheetModel,
  label: string,
  value: MetricValue,
  opts: { bold?: boolean; highlight?: boolean; format?: MetricFormat } = {},
): number {
  model.lastRow += 1
  model.rows.push({ row: model.lastRow, kind: 'metric', label, value, format: 'currency', ...opts })
  return model.lastRow
}

// ---- Break-Even calculator ----

type KeyMetric = { label: string; cellRef: string; value: number; format: MetricFormat }
type ReportModelBase = {
  rawSheets: RawDataSheetModel[]
  calcSheets: CalcSheetModel[]
  keyMetrics: KeyMetric[]
  charts: ReportChart[]
  extraCaveats: string[]
  // Escape hatch for calculators whose output doesn't fit the generic
  // account-list (RawDataSheetModel) / label-value (CalcSheetModel) shapes —
  // e.g. a multi-year, multi-column projection grid. Each builder is invoked
  // once the real workbook/usedSheetNames exist, after the raw/calc sheets.
  customSheetBuilders?: ((workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) => void)[]
}

// ── Platform-controlled financial assumptions ───────────────────────────────
// These are NOT user-facing and NOT AI-supplied — they're fixed inputs the
// analysis engine applies deterministically to every report of the relevant
// type, regardless of what's in the uploaded file. Update the values here
// (and redeploy) whenever the macro outlook or benchmark data changes; the
// frontend, the AI prompt, and the client experience are completely
// unaffected — the only visible difference is the numbers in the report.
const PLATFORM_ASSUMPTIONS = {
  // Applied in the Financial Projections calculator.
  annualInflationRate: 0.025, // 2.5%/year — escalates operating expenses across all scenarios
  baseRevenueCagr: 0.08, // 8%/year — base-case revenue growth
  optimisticCagrOffset: 0.05, // optimistic scenario = base + 5 percentage points
  pessimisticCagrOffset: 0.05, // pessimistic scenario = base - 5 percentage points
  // Applied in the Cash Flow Forecast calculator to convert a run-rate from
  // one granularity to the other when the source data only supports one.
  weeksPerMonth: 4.33, // 52 weeks / 12 months
}

type BreakEvenModel = ReportModelBase & { kind: 'break-even' }

function buildBreakEvenChart(revenue: number, fixed: number, variable: number): ReportChart {
  const base = revenue > 0 ? revenue : fixed + variable || 1
  const points = [0, 0.5, 1, 1.5, 2]
  const categories = points.map((p) => `${Math.round(p * 100)}%`)
  return {
    type: 'line',
    title: 'Revenue vs. Total Cost at Increasing Volume',
    categories,
    series: [
      { name: 'Revenue', values: points.map((p) => Math.round(base * p)) },
      { name: 'Total Cost', values: points.map((p) => Math.round(fixed + variable * p)) },
    ],
  }
}

export function calculateBreakEven(lineItems: LineItem[]): BreakEvenModel {
  const extraCaveats: string[] = []
  const rawSheet = layoutRawDataSheet('P&L', 'Reconstructed Income Statement', lineItems, [
    { tag: 'revenue', label: 'Revenue' },
    { tag: 'fixed_cost', label: 'Fixed Costs' },
    { tag: 'variable_cost', label: 'Variable Costs' },
  ])

  const revenueBlock = findBlock(rawSheet, 'revenue')
  const fixedBlock = findBlock(rawSheet, 'fixed_cost')
  const variableBlock = findBlock(rawSheet, 'variable_cost')
  if (!revenueBlock) extraCaveats.push('No line items were classified as revenue; break-even metrics below could not be fully computed.')
  if (!fixedBlock && !variableBlock) extraCaveats.push('No cost line items were classified as fixed or variable; break-even metrics below could not be fully computed.')

  const revenue = revenueBlock?.subtotalValue ?? 0
  const fixed = fixedBlock?.subtotalValue ?? 0
  const variable = variableBlock?.subtotalValue ?? 0

  const opexValue: MetricValue =
    fixedBlock || variableBlock
      ? fx(
          `=${fixedBlock ? `D${fixedBlock.subtotalRow}` : '0'}+${variableBlock ? `D${variableBlock.subtotalRow}` : '0'}`,
          fixed + variable,
        )
      : 'N/A (no cost data)'
  const opexRow = addTotalRow(rawSheet, 'Total Operating Expenses', opexValue)
  const netResultValue: MetricValue = revenueBlock
    ? fx(`=D${revenueBlock.subtotalRow}-D${opexRow}`, revenue - (fixed + variable))
    : 'N/A (no revenue data)'
  addTotalRow(rawSheet, 'Net Result', netResultValue)

  const calcSheet = newCalcSheet('Break-Even', 'Break-Even Calculation')
  const keyMetrics: KeyMetric[] = []

  addSection(calcSheet, 'Scenario A — Realistic (fixed vs. variable as classified)')
  const revRowA = addMetric(
    calcSheet,
    'Revenue',
    revenueBlock ? fx(refCell('P&L', revenueBlock.subtotalRow, 'D'), revenue) : 'N/A (no revenue data)',
  )
  const fixedRowA = addMetric(calcSheet, 'Fixed Costs', fixedBlock ? fx(refCell('P&L', fixedBlock.subtotalRow, 'D'), fixed) : fx('=0', 0))
  const varRowA = addMetric(
    calcSheet,
    'Variable Costs',
    variableBlock ? fx(refCell('P&L', variableBlock.subtotalRow, 'D'), variable) : fx('=0', 0),
  )

  if (revenueBlock && revenue !== 0) {
    const cm = revenue - variable
    const cmRow = addMetric(calcSheet, 'Contribution Margin (Revenue − Variable Costs)', fx(`=C${revRowA}-C${varRowA}`, cm))
    const cmRatio = cm / revenue
    const cmRatioRow = addMetric(calcSheet, 'Contribution Margin Ratio', fx(`=C${cmRow}/C${revRowA}`, cmRatio), { format: 'percent' })
    if (cmRatio !== 0) {
      const be = fixed / cmRatio
      const beRow = addMetric(calcSheet, 'Break-Even Revenue', fx(`=C${fixedRowA}/C${cmRatioRow}`, be), { bold: true, highlight: true })
      const mos = revenue - be
      const mosRow = addMetric(calcSheet, 'Margin of Safety (€)', fx(`=C${revRowA}-C${beRow}`, mos))
      addMetric(calcSheet, 'Margin of Safety (%)', fx(`=C${mosRow}/C${revRowA}`, mos / revenue), { format: 'percent' })
      keyMetrics.push({ label: 'Break-Even Revenue (Scenario A)', cellRef: refCell('Break-Even', beRow, 'C'), value: be, format: 'currency' })
      keyMetrics.push({
        label: 'Margin of Safety % (Scenario A)',
        cellRef: refCell('Break-Even', mosRow + 1, 'C'),
        value: mos / revenue,
        format: 'percent',
      })
    } else {
      addMetric(calcSheet, 'Break-Even Revenue', 'N/A (zero contribution margin)', { bold: true })
      extraCaveats.push('Contribution margin ratio is zero — break-even revenue could not be computed for Scenario A.')
    }
  } else {
    addMetric(calcSheet, 'Break-Even Revenue', 'N/A (no revenue data)', { bold: true })
  }

  addSection(calcSheet, 'Scenario B — Conservative (all operating costs treated as fixed)')
  const fixedRowB = addMetric(
    calcSheet,
    'Fixed Costs (= Total Operating Expenses)',
    fx(refCell('P&L', opexRow, 'D'), fixed + variable),
  )
  const beRowB = addMetric(calcSheet, 'Break-Even Revenue', fx(`=C${fixedRowB}`, fixed + variable), { bold: true, highlight: true })
  if (revenueBlock && revenue !== 0) {
    const mosBRow = addMetric(calcSheet, 'Margin of Safety (€)', fx(`=C${revRowA}-C${beRowB}`, revenue - (fixed + variable)))
    addMetric(calcSheet, 'Margin of Safety (%)', fx(`=C${mosBRow}/C${revRowA}`, (revenue - (fixed + variable)) / revenue), {
      format: 'percent',
    })
  } else {
    addMetric(calcSheet, 'Margin of Safety (€)', 'N/A (no revenue data)')
  }

  const headlineMetrics: KeyMetric[] = [
    { label: 'Total Operating Expenses', cellRef: refCell('P&L', opexRow, 'D'), value: fixed + variable, format: 'currency' },
  ]
  if (revenueBlock) {
    headlineMetrics.unshift({
      label: 'Total Revenue',
      cellRef: refCell('P&L', revenueBlock.subtotalRow, 'D'),
      value: revenue,
      format: 'currency',
    })
  }

  return {
    kind: 'break-even',
    rawSheets: [rawSheet],
    calcSheets: [calcSheet],
    keyMetrics: [...headlineMetrics, ...keyMetrics],
    charts: [buildBreakEvenChart(revenue, fixed, variable)],
    extraCaveats,
  }
}

// ---- Use of Funds calculator ----

type UseOfFundsModel = ReportModelBase & { kind: 'use-of-funds' }

function calculateUseOfFunds(lineItems: LineItem[]): UseOfFundsModel {
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

type BurnRateModel = ReportModelBase & { kind: 'burn-rate' }

function periodSortKey(period: string | undefined): number {
  const match = period?.match(/(\d+)/)
  return match ? parseInt(match[1], 10) : 0
}

function layoutBurnRateSheet(lineItems: LineItem[]) {
  const inflowItems = groupItems(lineItems, 'cash_inflow')
  const outflowItems = groupItems(lineItems, 'cash_outflow')
  const startingItems = groupItems(lineItems, 'starting_cash_balance')

  const periodsSet = new Set<string>()
  for (const item of [...inflowItems, ...outflowItems]) periodsSet.add(item.period || 'All Periods')
  const periods = Array.from(periodsSet).sort((a, b) => periodSortKey(a) - periodSortKey(b))

  const rows: RawDataRow[] = []
  const blocks: RawDataBlock[] = []
  let row = RAW_DATA_FIRST_DATA_ROW
  let startingBalanceRow: number | null = null
  const startingBalanceValue = sumGroup(startingItems)

  if (startingItems.length) {
    const item = startingItems[0]
    rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: 'Starting Cash Balance', amount: item.amount })
    startingBalanceRow = row
    row += 1
  }

  const inflowRowByPeriod: Record<string, number> = {}
  const outflowRowByPeriod: Record<string, number> = {}
  const inflowValueByPeriod: Record<string, number> = {}
  const outflowValueByPeriod: Record<string, number> = {}

  for (const period of periods) {
    const periodInflow = inflowItems.filter((i) => (i.period || 'All Periods') === period)
    const periodOutflow = outflowItems.filter((i) => (i.period || 'All Periods') === period)

    if (periodInflow.length) {
      const start = row
      for (const item of periodInflow) {
        rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: `Inflow — ${period}`, amount: item.amount })
        row += 1
      }
      const subtotalValue = sumGroup(periodInflow)
      rows.push({ row, kind: 'subtotal', label: `Total Inflow — ${period}`, value: fx(`=SUM(D${start}:D${row - 1})`, subtotalValue) })
      inflowRowByPeriod[period] = row
      inflowValueByPeriod[period] = subtotalValue
      blocks.push({ tag: `cash_inflow:${period}`, label: `Inflow — ${period}`, subtotalRow: row, subtotalValue })
      row += 1
    }
    if (periodOutflow.length) {
      const start = row
      for (const item of periodOutflow) {
        rows.push({ row, kind: 'item', account: item.account, description: item.label, groupLabel: `Outflow — ${period}`, amount: item.amount })
        row += 1
      }
      const subtotalValue = sumGroup(periodOutflow)
      rows.push({ row, kind: 'subtotal', label: `Total Outflow — ${period}`, value: fx(`=SUM(D${start}:D${row - 1})`, subtotalValue) })
      outflowRowByPeriod[period] = row
      outflowValueByPeriod[period] = subtotalValue
      blocks.push({ tag: `cash_outflow:${period}`, label: `Outflow — ${period}`, subtotalRow: row, subtotalValue })
      row += 1
    }
  }

  const model: RawDataSheetModel = { sheetName: 'Cash Flow Detail', title: 'Cash Flow Detail', rows, blocks, lastRow: row - 1 }
  return { model, periods, inflowRowByPeriod, outflowRowByPeriod, inflowValueByPeriod, outflowValueByPeriod, startingBalanceRow, startingBalanceValue, outflowItems }
}

function calculateBurnRate(lineItems: LineItem[]): BurnRateModel {
  const extraCaveats: string[] = []
  const {
    model: rawSheet,
    periods,
    inflowRowByPeriod,
    outflowRowByPeriod,
    inflowValueByPeriod,
    outflowValueByPeriod,
    startingBalanceRow,
    startingBalanceValue,
    outflowItems,
  } = layoutBurnRateSheet(lineItems)

  if (!periods.length) extraCaveats.push('No cash inflow/outflow line items were classified; burn rate metrics below could not be computed.')
  if (periods.length === 1 && periods[0] === 'All Periods') {
    extraCaveats.push(
      'Source data did not distinguish individual months — the figures below treat the classified data as a single representative period rather than a true 12-month average.',
    )
  }
  if (startingBalanceRow === null) extraCaveats.push('No starting cash balance was identified; runway could not be computed.')

  const calcSheet = newCalcSheet('Burn Analysis', 'Burn Rate & Runway Analysis')
  const keyMetrics: KeyMetric[] = []

  addSection(calcSheet, 'Burn Rate & Runway')
  const outflowRefs = periods.filter((p) => outflowRowByPeriod[p]).map((p) => refCell('Cash Flow Detail', outflowRowByPeriod[p], 'D'))
  const outflowValues = periods.filter((p) => outflowRowByPeriod[p]).map((p) => outflowValueByPeriod[p])
  const inflowRefs = periods.filter((p) => inflowRowByPeriod[p]).map((p) => refCell('Cash Flow Detail', inflowRowByPeriod[p], 'D'))
  const inflowValues = periods.filter((p) => inflowRowByPeriod[p]).map((p) => inflowValueByPeriod[p])

  const grossBurnValue = outflowValues.length ? outflowValues.reduce((a, b) => a + b, 0) / outflowValues.length : 0
  const grossBurnRow = addMetric(
    calcSheet,
    'Gross Burn Rate (monthly operating expenses, average per period)',
    outflowRefs.length ? fx(`=AVERAGE(${outflowRefs.join(',')})`, grossBurnValue) : 'N/A (no outflow data)',
  )
  const grossInflowValue = inflowValues.length ? inflowValues.reduce((a, b) => a + b, 0) / inflowValues.length : 0
  const grossInflowRow = addMetric(
    calcSheet,
    'Gross Inflow (average per period)',
    inflowRefs.length ? fx(`=AVERAGE(${inflowRefs.join(',')})`, grossInflowValue) : 'N/A (no inflow data)',
  )
  const netBurnValue = grossBurnValue - grossInflowValue
  const netBurnRow = addMetric(calcSheet, 'Net Burn (average per period)', fx(`=C${grossBurnRow}-C${grossInflowRow}`, netBurnValue), {
    bold: true,
  })

  let currentBalanceRow: number | null = null
  if (startingBalanceRow !== null) {
    currentBalanceRow = addMetric(
      calcSheet,
      'Current Cash Balance',
      fx(refCell('Cash Flow Detail', startingBalanceRow, 'D'), startingBalanceValue),
    )
    keyMetrics.push({ label: 'Current Cash Balance', cellRef: refCell('Burn Analysis', currentBalanceRow, 'C'), value: startingBalanceValue, format: 'currency' })
  }

  if (currentBalanceRow !== null && netBurnValue > 0) {
    const runway = startingBalanceValue / netBurnValue
    const runwayRow = addMetric(calcSheet, 'Runway (periods)', fx(`=C${currentBalanceRow}/C${netBurnRow}`, runway), {
      bold: true,
      highlight: true,
      format: 'number',
    })
    keyMetrics.push({ label: 'Runway (periods)', cellRef: refCell('Burn Analysis', runwayRow, 'C'), value: runway, format: 'number' })
  } else if (currentBalanceRow !== null) {
    addMetric(calcSheet, 'Runway (periods)', 'N/A — cash flow break-even or positive', { bold: true, highlight: true })
  } else {
    addMetric(calcSheet, 'Runway (periods)', 'N/A (no starting cash balance)', { bold: true })
  }

  // Gross runway = company cash / monthly operating expenses: how long cash
  // lasts if no cash comes in at all.
  if (currentBalanceRow !== null && grossBurnValue > 0) {
    const grossRunway = startingBalanceValue / grossBurnValue
    const grossRunwayRow = addMetric(calcSheet, 'Gross Runway (cash ÷ monthly operating expenses, periods)', fx(`=C${currentBalanceRow}/C${grossBurnRow}`, grossRunway), {
      format: 'number',
    })
    keyMetrics.push({ label: 'Gross Runway (periods, no inflows)', cellRef: refCell('Burn Analysis', grossRunwayRow, 'C'), value: grossRunway, format: 'number' })
  } else {
    addMetric(calcSheet, 'Gross Runway (cash ÷ monthly operating expenses, periods)', currentBalanceRow === null ? 'N/A (no starting cash balance)' : 'N/A (no outflow data)')
  }

  keyMetrics.push({ label: 'Gross Burn Rate (average per period)', cellRef: refCell('Burn Analysis', grossBurnRow, 'C'), value: grossBurnValue, format: 'currency' })
  keyMetrics.push({ label: 'Net Burn (average per period)', cellRef: refCell('Burn Analysis', netBurnRow, 'C'), value: netBurnValue, format: 'currency' })

  addSection(calcSheet, 'Period-by-Period Cash Position')
  let prevRef: string | null = startingBalanceRow !== null ? refCell('Cash Flow Detail', startingBalanceRow, 'D') : null
  let runningBalance = startingBalanceValue
  const balanceSeries: number[] = []
  for (const period of periods) {
    const inflowVal = inflowValueByPeriod[period] ?? 0
    const outflowVal = outflowValueByPeriod[period] ?? 0
    runningBalance = runningBalance + inflowVal - outflowVal
    const inflowRef = inflowRowByPeriod[period] ? refCell('Cash Flow Detail', inflowRowByPeriod[period], 'D') : null
    const outflowRef = outflowRowByPeriod[period] ? refCell('Cash Flow Detail', outflowRowByPeriod[period], 'D') : null
    let formulaStr = prevRef ? `=${prevRef}` : '=0'
    if (inflowRef) formulaStr += `+${inflowRef}`
    if (outflowRef) formulaStr += `-${outflowRef}`
    const balRow = addMetric(calcSheet, `Cash Balance — ${period}`, fx(formulaStr, runningBalance))
    prevRef = `C${balRow}`
    balanceSeries.push(runningBalance)
  }

  const categoryTotals = new Map<string, number>()
  for (const item of outflowItems) categoryTotals.set(item.label, (categoryTotals.get(item.label) ?? 0) + item.amount)
  const topCategories = [...categoryTotals.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8)

  const charts: ReportChart[] = []
  if (periods.length) {
    charts.push({
      type: 'line',
      title: 'Period-by-Period Cash Balance',
      categories: periods,
      series: [{ name: 'Cash Balance', values: balanceSeries }],
    })
  }
  if (topCategories.length) {
    charts.push({
      type: 'bar',
      title: 'Largest Cash Outflow Categories',
      categories: topCategories.map(([label]) => label),
      series: [{ name: 'Total Outflow', values: topCategories.map(([, value]) => value) }],
    })
  }

  return { kind: 'burn-rate', rawSheets: [rawSheet], calcSheets: [calcSheet], keyMetrics, charts, extraCaveats }
}

// ---- Cash Flow Forecast calculator (13-week / 12-month rolling) ----
//
// Reuses the burn-rate raw sheet layout (same "cash_inflow"/"cash_outflow"/
// "starting_cash_balance" group vocabulary, period-dated), then rolls the
// balance forward across two fixed horizons. Whichever granularity the
// source data actually supports (weekly or monthly periods) is used as the
// baseline run-rate; the other horizon is derived from it via the
// platform-controlled weeks-per-month conversion factor, exactly like the
// Financial Projections calculator derives Years 2-5 from Year-1 data.

type CashFlowForecastModel = ReportModelBase & { kind: 'cash-flow-forecast' }

const CF_WEEK_HORIZON = 13
const CF_MONTH_HORIZON = 12
const CF_GRID_ROWS = { header: 4, beginning: 5, inflow: 6, outflow: 7, net: 8, ending: 9 }

function cfColLetter(indexFromB: number): string {
  // indexFromB is 0-based; 0 -> B, 1 -> C, ... (stays within single letters for our horizons)
  return String.fromCharCode('B'.charCodeAt(0) + indexFromB)
}

function averageOf(values: number[]): number {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : 0
}

function classifyPeriodGranularity(period: string): 'week' | 'month' | 'unlabeled' {
  if (/week/i.test(period)) return 'week'
  if (/month/i.test(period)) return 'month'
  return 'unlabeled'
}

function buildCashFlowGridSheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  title: string,
  periodLabelPrefix: string,
  horizon: number,
  startingBalance: MetricValue,
  startingBalanceValue: number,
  periodInflow: (index: number) => MetricValue,
  periodInflowValue: (index: number) => number,
  periodOutflow: (index: number) => MetricValue,
  periodOutflowValue: (index: number) => number,
) {
  const R = CF_GRID_ROWS
  const sheet = workbook.addWorksheet(sanitizeSheetName(title, usedSheetNames))
  sheet.getColumn(1).width = 22
  for (let i = 0; i < horizon; i++) sheet.getColumn(2 + i).width = 15

  sheet.getCell('A1').value = title
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells(`A1:${cfColLetter(horizon - 1)}1`)

  sheet.getCell(`A${R.header}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
  sheet.getCell(`A${R.header}`).border = THIN_BORDER
  for (let i = 0; i < horizon; i++) {
    const col = cfColLetter(i)
    const cell = sheet.getCell(`${col}${R.header}`)
    cell.value = `${periodLabelPrefix} ${i + 1}`
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.alignment = { horizontal: 'center' }
    cell.border = THIN_BORDER
  }
  sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: R.header }]

  const labelSpecs: [number, string, boolean?][] = [
    [R.beginning, 'Beginning Balance'],
    [R.inflow, 'Cash Inflows'],
    [R.outflow, 'Cash Outflows'],
    [R.net, 'Net Cash Flow', true],
    [R.ending, 'Ending Balance', true],
  ]
  for (const [row, label, bold] of labelSpecs) {
    const cell = sheet.getCell(`A${row}`)
    cell.value = label
    cell.font = { bold: !!bold }
  }

  function setCell(row: number, col: string, value: MetricValue) {
    const cell = sheet.getCell(`${col}${row}`)
    setMetricCell(cell, value, 'currency')
    cell.border = THIN_BORDER
  }

  let runningBalance = startingBalanceValue
  const endingSeries: number[] = []
  for (let i = 0; i < horizon; i++) {
    const col = cfColLetter(i)
    const prevCol = i > 0 ? cfColLetter(i - 1) : null

    setCell(R.beginning, col, i === 0 ? startingBalance : fx(`=${prevCol}${R.ending}`, runningBalance))

    const inflowVal = periodInflowValue(i)
    setCell(R.inflow, col, periodInflow(i))
    const outflowVal = periodOutflowValue(i)
    setCell(R.outflow, col, periodOutflow(i))

    const netVal = inflowVal - outflowVal
    setCell(R.net, col, fx(`=${col}${R.inflow}-${col}${R.outflow}`, netVal))

    const beginningVal = runningBalance
    runningBalance = beginningVal + netVal
    setCell(R.ending, col, fx(`=${col}${R.beginning}+${col}${R.net}`, runningBalance))

    endingSeries.push(runningBalance)
  }

  return { sheetName: sheet.name, endingSeries }
}

function calculateCashFlowForecast(lineItems: LineItem[]): CashFlowForecastModel {
  const extraCaveats: string[] = []
  const {
    model: rawSheet,
    periods,
    inflowRowByPeriod,
    outflowRowByPeriod,
    inflowValueByPeriod,
    outflowValueByPeriod,
    startingBalanceRow,
    startingBalanceValue,
  } = layoutBurnRateSheet(lineItems)

  if (startingBalanceRow === null) extraCaveats.push('No starting cash balance was identified; the forecast below assumes a starting balance of zero.')

  const weekPeriods = periods.filter((p) => classifyPeriodGranularity(p) === 'week').sort((a, b) => periodSortKey(a) - periodSortKey(b))
  const monthPeriods = periods.filter((p) => classifyPeriodGranularity(p) === 'month').sort((a, b) => periodSortKey(a) - periodSortKey(b))
  const unlabeledPeriods = periods.filter((p) => classifyPeriodGranularity(p) === 'unlabeled')

  const { weeksPerMonth } = PLATFORM_ASSUMPTIONS
  const rawSheetName = rawSheet.sheetName

  const assumptionsSheet = newCalcSheet('Run-Rate Assumptions', 'Baseline Run-Rate Assumptions')
  addSection(assumptionsSheet, 'Recurring Cash Flow Run-Rate')
  const weeksPerMonthRow = addMetric(assumptionsSheet, 'Weeks per Month (conversion factor, editable)', weeksPerMonth)

  let weeklyInflowRow: number, weeklyOutflowRow: number, monthlyInflowRow: number, monthlyOutflowRow: number
  let weeklyInflowValue: number, weeklyOutflowValue: number, monthlyInflowValue: number, monthlyOutflowValue: number

  if (monthPeriods.length) {
    const inflowRefs = monthPeriods.filter((p) => inflowRowByPeriod[p]).map((p) => refCell(rawSheetName, inflowRowByPeriod[p], 'D'))
    monthlyInflowValue = averageOf(monthPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    monthlyInflowRow = addMetric(assumptionsSheet, 'Monthly Inflow Run-Rate', inflowRefs.length ? fx(`=AVERAGE(${inflowRefs.join(',')})`, monthlyInflowValue) : 'N/A (no inflow data)')
    const outflowRefs = monthPeriods.filter((p) => outflowRowByPeriod[p]).map((p) => refCell(rawSheetName, outflowRowByPeriod[p], 'D'))
    monthlyOutflowValue = averageOf(monthPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    monthlyOutflowRow = addMetric(assumptionsSheet, 'Monthly Outflow Run-Rate', outflowRefs.length ? fx(`=AVERAGE(${outflowRefs.join(',')})`, monthlyOutflowValue) : 'N/A (no outflow data)')
  } else if (weekPeriods.length) {
    const weeklyInflowForConversion = averageOf(weekPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    const weeklyOutflowForConversion = averageOf(weekPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    monthlyInflowValue = weeklyInflowForConversion * weeksPerMonth
    monthlyInflowRow = addMetric(assumptionsSheet, 'Monthly Inflow Run-Rate (derived from weekly data)', monthlyInflowValue)
    monthlyOutflowValue = weeklyOutflowForConversion * weeksPerMonth
    monthlyOutflowRow = addMetric(assumptionsSheet, 'Monthly Outflow Run-Rate (derived from weekly data)', monthlyOutflowValue)
    extraCaveats.push('Source data was only weekly-dated — the 12-month forecast derives its monthly run-rate from the weekly figures using a fixed 4.33 weeks/month conversion factor.')
  } else {
    monthlyInflowValue = averageOf(unlabeledPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    monthlyInflowRow = addMetric(assumptionsSheet, 'Monthly Inflow Run-Rate (assumed from undated total)', monthlyInflowValue)
    monthlyOutflowValue = averageOf(unlabeledPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    monthlyOutflowRow = addMetric(assumptionsSheet, 'Monthly Outflow Run-Rate (assumed from undated total)', monthlyOutflowValue)
    if (unlabeledPeriods.length) extraCaveats.push('Source data did not distinguish individual weeks or months — the undated total is treated as a single representative month, and both horizons below are extrapolated flat from it.')
  }

  if (weekPeriods.length) {
    const inflowRefs = weekPeriods.filter((p) => inflowRowByPeriod[p]).map((p) => refCell(rawSheetName, inflowRowByPeriod[p], 'D'))
    weeklyInflowValue = averageOf(weekPeriods.map((p) => inflowValueByPeriod[p] ?? 0))
    weeklyInflowRow = addMetric(assumptionsSheet, 'Weekly Inflow Run-Rate', inflowRefs.length ? fx(`=AVERAGE(${inflowRefs.join(',')})`, weeklyInflowValue) : 'N/A (no inflow data)')
    const outflowRefs = weekPeriods.filter((p) => outflowRowByPeriod[p]).map((p) => refCell(rawSheetName, outflowRowByPeriod[p], 'D'))
    weeklyOutflowValue = averageOf(weekPeriods.map((p) => outflowValueByPeriod[p] ?? 0))
    weeklyOutflowRow = addMetric(assumptionsSheet, 'Weekly Outflow Run-Rate', outflowRefs.length ? fx(`=AVERAGE(${outflowRefs.join(',')})`, weeklyOutflowValue) : 'N/A (no outflow data)')
  } else {
    weeklyInflowValue = monthlyInflowValue / weeksPerMonth
    weeklyInflowRow = addMetric(assumptionsSheet, 'Weekly Inflow Run-Rate (derived from monthly data)', fx(`=C${monthlyInflowRow}/C${weeksPerMonthRow}`, weeklyInflowValue))
    weeklyOutflowValue = monthlyOutflowValue / weeksPerMonth
    weeklyOutflowRow = addMetric(assumptionsSheet, 'Weekly Outflow Run-Rate (derived from monthly data)', fx(`=C${monthlyOutflowRow}/C${weeksPerMonthRow}`, weeklyOutflowValue))
    if (monthPeriods.length) extraCaveats.push('Source data was only monthly-dated — the 13-week forecast derives its weekly run-rate from the monthly figures using a fixed 4.33 weeks/month conversion factor.')
  }

  const assumptionsSheetName = assumptionsSheet.sheetName

  const startingBalanceRef: MetricValue =
    startingBalanceRow !== null ? fx(refCell(rawSheetName, startingBalanceRow, 'D'), startingBalanceValue) : fx('=0', 0)

  function makeAccessors(
    periodsForHorizon: string[],
    rowByPeriod: Record<string, number>,
    valueByPeriod: Record<string, number>,
    flatRunRateRow: number,
    flatRunRateValue: number,
    labelPrefix: string,
  ) {
    return {
      value: (index: number) => valueByPeriod[`${labelPrefix} ${index + 1}`] ?? flatRunRateValue,
      cell: (index: number): MetricValue => {
        const label = `${labelPrefix} ${index + 1}`
        if (rowByPeriod[label]) return fx(refCell(rawSheetName, rowByPeriod[label], 'D'), valueByPeriod[label])
        return fx(refCell(assumptionsSheetName, flatRunRateRow, 'C'), flatRunRateValue)
      },
    }
  }

  const weekInflowAccessor = makeAccessors(weekPeriods, inflowRowByPeriod, inflowValueByPeriod, weeklyInflowRow, weeklyInflowValue, 'Week')
  const weekOutflowAccessor = makeAccessors(weekPeriods, outflowRowByPeriod, outflowValueByPeriod, weeklyOutflowRow, weeklyOutflowValue, 'Week')
  const monthInflowAccessor = makeAccessors(monthPeriods, inflowRowByPeriod, inflowValueByPeriod, monthlyInflowRow, monthlyInflowValue, 'Month')
  const monthOutflowAccessor = makeAccessors(monthPeriods, outflowRowByPeriod, outflowValueByPeriod, monthlyOutflowRow, monthlyOutflowValue, 'Month')

  const customSheetBuilders = [
    (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildCashFlowGridSheet(
        workbook,
        usedSheetNames,
        '13-Week Cash Flow Forecast',
        'Week',
        CF_WEEK_HORIZON,
        startingBalanceRef,
        startingBalanceValue,
        weekInflowAccessor.cell,
        weekInflowAccessor.value,
        weekOutflowAccessor.cell,
        weekOutflowAccessor.value,
      ),
    (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildCashFlowGridSheet(
        workbook,
        usedSheetNames,
        '12-Month Cash Flow Forecast',
        'Month',
        CF_MONTH_HORIZON,
        startingBalanceRef,
        startingBalanceValue,
        monthInflowAccessor.cell,
        monthInflowAccessor.value,
        monthOutflowAccessor.cell,
        monthOutflowAccessor.value,
      ),
  ]

  // Precompute the same ending-balance series numerically (independent of the
  // deferred sheet builders above) so key metrics and charts are available
  // immediately, mirroring the Financial Projections calculator's approach.
  function projectEndingSeries(horizon: number, inflowValue: (i: number) => number, outflowValue: (i: number) => number) {
    let balance = startingBalanceValue
    const series: number[] = []
    for (let i = 0; i < horizon; i++) {
      balance = balance + inflowValue(i) - outflowValue(i)
      series.push(balance)
    }
    return series
  }

  const weekSeries = projectEndingSeries(CF_WEEK_HORIZON, weekInflowAccessor.value, weekOutflowAccessor.value)
  const monthSeries = projectEndingSeries(CF_MONTH_HORIZON, monthInflowAccessor.value, monthOutflowAccessor.value)

  const minWeekBalance = Math.min(...weekSeries)
  const minWeekIndex = weekSeries.indexOf(minWeekBalance)
  const minMonthBalance = Math.min(...monthSeries)
  const minMonthIndex = monthSeries.indexOf(minMonthBalance)

  const keyMetrics: KeyMetric[] = [
    {
      label: 'Starting Cash Balance',
      cellRef: refCell('13-Week Cash Flow Forecast', CF_GRID_ROWS.beginning, cfColLetter(0)),
      value: startingBalanceValue,
      format: 'currency',
    },
    {
      label: `Lowest Projected Balance — 13-Week (Week ${minWeekIndex + 1})`,
      cellRef: refCell('13-Week Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(minWeekIndex)),
      value: minWeekBalance,
      format: 'currency',
    },
    {
      label: 'Ending Balance — Week 13',
      cellRef: refCell('13-Week Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(CF_WEEK_HORIZON - 1)),
      value: weekSeries[CF_WEEK_HORIZON - 1],
      format: 'currency',
    },
    {
      label: `Lowest Projected Balance — 12-Month (Month ${minMonthIndex + 1})`,
      cellRef: refCell('12-Month Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(minMonthIndex)),
      value: minMonthBalance,
      format: 'currency',
    },
    {
      label: 'Ending Balance — Month 12',
      cellRef: refCell('12-Month Cash Flow Forecast', CF_GRID_ROWS.ending, cfColLetter(CF_MONTH_HORIZON - 1)),
      value: monthSeries[CF_MONTH_HORIZON - 1],
      format: 'currency',
    },
  ]

  if (minWeekBalance < 0 || minMonthBalance < 0) {
    extraCaveats.push('The projected cash balance goes negative within the forecast horizon under the current run-rate assumptions — treat this as a liquidity risk flag, not a certainty, since it is driven by a flat extrapolated run-rate wherever period-specific data was not available.')
  }

  const charts: ReportChart[] = [
    {
      type: 'line',
      title: '13-Week Rolling Cash Balance',
      categories: Array.from({ length: CF_WEEK_HORIZON }, (_, i) => `Week ${i + 1}`),
      series: [{ name: 'Ending Balance', values: weekSeries.map((v) => Math.round(v)) }],
    },
    {
      type: 'line',
      title: '12-Month Rolling Cash Balance',
      categories: Array.from({ length: CF_MONTH_HORIZON }, (_, i) => `Month ${i + 1}`),
      series: [{ name: 'Ending Balance', values: monthSeries.map((v) => Math.round(v)) }],
    },
  ]

  return {
    kind: 'cash-flow-forecast',
    rawSheets: [rawSheet],
    calcSheets: [assumptionsSheet],
    keyMetrics,
    charts,
    extraCaveats,
    customSheetBuilders,
  }
}

// ---- Capital Optimization calculator ----

type CapitalOptimizationModel = ReportModelBase & { kind: 'capital-optimization' }

// A reasonable, commonly-cited default target for a healthy debt-to-equity
// ratio, used as the starting point for the restructuring comparison. This is
// an editable input cell in the rendered workbook — users can override it.
const DEFAULT_TARGET_DEBT_TO_EQUITY = 1.0

function calculateCapitalOptimization(lineItems: LineItem[]): CapitalOptimizationModel {
  const extraCaveats: string[] = []
  const DATA = 'Capital Structure Data'
  const rawSheet = layoutRawDataSheet(DATA, 'Capital Structure Data', lineItems, [
    { tag: 'short_term_debt', label: 'Short-Term Debt' },
    { tag: 'long_term_debt', label: 'Long-Term Debt' },
    { tag: 'accounts_payable', label: 'Accounts Payable' },
    { tag: 'past_due_payable', label: 'Accounts Payable Past Due (memo, included above)' },
    { tag: 'current_liability', label: 'Other Current Liabilities' },
    { tag: 'other_liability', label: 'Other Non-Current Liabilities' },
    { tag: 'equity', label: 'Equity' },
    { tag: 'current_asset', label: 'Current Assets' },
    { tag: 'fixed_asset', label: 'Fixed Assets' },
    { tag: 'interest_expense', label: 'Interest Expense' },
    { tag: 'ebit', label: 'EBIT' },
  ])

  const shortDebtBlock = findBlock(rawSheet, 'short_term_debt')
  const longDebtBlock = findBlock(rawSheet, 'long_term_debt')
  const equityBlock = findBlock(rawSheet, 'equity')
  const currentAssetBlock = findBlock(rawSheet, 'current_asset')
  const payablesBlock = findBlock(rawSheet, 'accounts_payable')
  const pastDuePayablesBlock = findBlock(rawSheet, 'past_due_payable')
  const currentLiabilityBlock = findBlock(rawSheet, 'current_liability')
  const otherLiabilityBlock = findBlock(rawSheet, 'other_liability')
  const interestBlock = findBlock(rawSheet, 'interest_expense')
  const ebitBlock = findBlock(rawSheet, 'ebit')
  const hasDebt = Boolean(shortDebtBlock || longDebtBlock)
  const hasLiabilities = Boolean(hasDebt || payablesBlock || currentLiabilityBlock || otherLiabilityBlock)

  if (!hasLiabilities || !equityBlock) extraCaveats.push('Liabilities and/or equity line items were not classified; capital structure ratios below could not be fully computed.')

  // Raw-sheet subtotal reference, or 0 when the group is absent.
  const sub = (block: RawDataBlock | undefined) => (block ? fx(refCell(DATA, block.subtotalRow, 'D'), block.subtotalValue) : fx('=0', 0))
  const shortDebt = shortDebtBlock?.subtotalValue ?? 0
  const longDebt = longDebtBlock?.subtotalValue ?? 0
  const debt = shortDebt + longDebt
  const payables = payablesBlock?.subtotalValue ?? 0
  const otherCurrentLiabilities = currentLiabilityBlock?.subtotalValue ?? 0
  const otherLiabilities = otherLiabilityBlock?.subtotalValue ?? 0
  const currentLiabilities = payables + otherCurrentLiabilities + shortDebt
  const totalLiabilities = currentLiabilities + longDebt + otherLiabilities
  const equity = equityBlock?.subtotalValue ?? 0

  const calcSheet = newCalcSheet('Capital Analysis', 'Capital Structure Analysis')
  const keyMetrics: KeyMetric[] = []

  addSection(calcSheet, 'Current Capital Structure')
  const shortDebtRow = addMetric(calcSheet, 'Short-Term Debt', sub(shortDebtBlock))
  const longDebtRow = addMetric(calcSheet, 'Long-Term Debt', sub(longDebtBlock))
  const debtRow = addMetric(calcSheet, 'Total Interest-Bearing Debt', fx(`=C${shortDebtRow}+C${longDebtRow}`, debt))
  const payablesRow = addMetric(calcSheet, 'Accounts Payable', sub(payablesBlock))
  const otherCurrentRow = addMetric(calcSheet, 'Other Current Liabilities', sub(currentLiabilityBlock))
  const currentLiabilitiesRow = addMetric(
    calcSheet,
    'Total Current Liabilities',
    fx(`=C${payablesRow}+C${otherCurrentRow}+C${shortDebtRow}`, currentLiabilities),
  )
  const otherLiabilitiesRow = addMetric(calcSheet, 'Other Non-Current Liabilities', sub(otherLiabilityBlock))
  const totalLiabilitiesRow = addMetric(
    calcSheet,
    'Total Liabilities',
    fx(`=C${currentLiabilitiesRow}+C${longDebtRow}+C${otherLiabilitiesRow}`, totalLiabilities),
    { bold: true },
  )
  const equityRow = addMetric(calcSheet, "Total Shareholders' Equity", sub(equityBlock))
  const totalCapitalRow = addMetric(calcSheet, 'Total Capital (Debt + Equity)', fx(`=C${debtRow}+C${equityRow}`, debt + equity))

  // Debt-to-equity = Total liabilities / Total shareholders' equity
  if (equity !== 0 && hasLiabilities) {
    const deRatioRow = addMetric(calcSheet, 'Debt-to-Equity Ratio (Total Liabilities ÷ Equity)', fx(`=C${totalLiabilitiesRow}/C${equityRow}`, totalLiabilities / equity), {
      format: 'ratio',
    })
    keyMetrics.push({ label: 'Debt-to-Equity Ratio', cellRef: refCell('Capital Analysis', deRatioRow, 'C'), value: totalLiabilities / equity, format: 'ratio' })
  } else {
    addMetric(calcSheet, 'Debt-to-Equity Ratio (Total Liabilities ÷ Equity)', equity === 0 ? 'N/A (no equity data)' : 'N/A (no liability data)')
  }
  if (debt + equity !== 0) {
    addMetric(calcSheet, 'Debt-to-Capital Ratio', fx(`=C${debtRow}/C${totalCapitalRow}`, debt / (debt + equity)), { format: 'percent' })
  } else {
    addMetric(calcSheet, 'Debt-to-Capital Ratio', 'N/A (no capital data)')
  }

  // Current ratio = Current assets / Current liabilities (short-term debt included)
  if (currentAssetBlock && currentLiabilities !== 0) {
    const caRow = addMetric(calcSheet, 'Total Current Assets', sub(currentAssetBlock))
    const currentRatio = currentAssetBlock.subtotalValue / currentLiabilities
    const currentRatioRow = addMetric(calcSheet, 'Current Ratio', fx(`=C${caRow}/C${currentLiabilitiesRow}`, currentRatio), { format: 'ratio' })
    keyMetrics.push({ label: 'Current Ratio', cellRef: refCell('Capital Analysis', currentRatioRow, 'C'), value: currentRatio, format: 'ratio' })
  } else {
    addMetric(calcSheet, 'Current Ratio', 'N/A (insufficient current asset/liability data)')
  }

  // Current accounts payable = (Total AP − Past due AP) / Total AP
  if (payablesBlock && pastDuePayablesBlock && payables !== 0) {
    const pastDueRow = addMetric(calcSheet, 'Accounts Payable Past Due', sub(pastDuePayablesBlock))
    const currentApShare = (payables - pastDuePayablesBlock.subtotalValue) / payables
    const currentApRow = addMetric(calcSheet, 'Current Accounts Payable (share not past due)', fx(`=(C${payablesRow}-C${pastDueRow})/C${payablesRow}`, currentApShare), {
      format: 'percent',
    })
    keyMetrics.push({ label: 'Current Accounts Payable', cellRef: refCell('Capital Analysis', currentApRow, 'C'), value: currentApShare, format: 'percent' })
    if (pastDuePayablesBlock.subtotalValue > payables) extraCaveats.push('Past-due payables exceed total accounts payable — check the AP aging against the balance sheet.')
  } else {
    addMetric(calcSheet, 'Current Accounts Payable (share not past due)', payablesBlock ? 'N/A (no AP aging / past-due payables data)' : 'N/A (no accounts payable data)')
  }

  // Interest coverage = EBIT / Interest expense
  let interestRow: number | undefined
  if (ebitBlock && interestBlock && interestBlock.subtotalValue !== 0) {
    const ebitRow = addMetric(calcSheet, 'EBIT', sub(ebitBlock))
    interestRow = addMetric(calcSheet, 'Interest Expense', sub(interestBlock))
    const coverage = ebitBlock.subtotalValue / interestBlock.subtotalValue
    const coverageRow = addMetric(calcSheet, 'Interest Coverage Ratio (EBIT ÷ Interest)', fx(`=C${ebitRow}/C${interestRow}`, coverage), { format: 'ratio' })
    keyMetrics.push({ label: 'Interest Coverage Ratio', cellRef: refCell('Capital Analysis', coverageRow, 'C'), value: coverage, format: 'ratio' })
  } else {
    addMetric(calcSheet, 'Interest Coverage Ratio (EBIT ÷ Interest)', 'N/A (insufficient EBIT/interest expense data)')
    if (!ebitBlock || !interestBlock) extraCaveats.push('EBIT and/or interest expense were not classified; interest coverage could not be computed.')
  }

  if (equityBlock && equity !== 0 && interestBlock && interestBlock.subtotalValue !== 0 && debt !== 0) {
    addSection(calcSheet, 'Restructuring Comparison — Current vs. Target Structure')
    const intRow = interestRow ?? addMetric(calcSheet, 'Interest Expense', sub(interestBlock))
    const targetRatioRow = addMetric(calcSheet, 'Target Debt-to-Equity Ratio (Total Liabilities ÷ Equity, editable)', DEFAULT_TARGET_DEBT_TO_EQUITY, {
      format: 'ratio',
    })
    const targetLiabilitiesRow = addMetric(calcSheet, 'Target Total Liabilities', fx(`=C${targetRatioRow}*C${equityRow}`, DEFAULT_TARGET_DEBT_TO_EQUITY * equity))
    const reduction = totalLiabilities - DEFAULT_TARGET_DEBT_TO_EQUITY * equity
    const reductionRow = addMetric(calcSheet, 'Liability Reduction Needed', fx(`=C${totalLiabilitiesRow}-C${targetLiabilitiesRow}`, reduction))
    const effectiveRate = interestBlock.subtotalValue / debt
    const effectiveRateRow = addMetric(calcSheet, 'Current Effective Interest Rate (Interest ÷ Debt)', fx(`=C${intRow}/C${debtRow}`, effectiveRate), {
      format: 'percent',
    })
    addMetric(
      calcSheet,
      'Estimated Annual Interest Savings',
      fx(`=MIN(C${reductionRow},C${debtRow})*C${effectiveRateRow}`, Math.min(reduction, debt) * effectiveRate),
      { bold: true, highlight: true },
    )
    extraCaveats.push(
      'Interest savings assume the liability reduction is achieved by repaying interest-bearing debt at the current effective rate, capped at total interest-bearing debt.',
    )
  }

  const chart: ReportChart = {
    type: 'bar',
    title: 'Current Capital Structure — Liabilities vs. Equity',
    categories: ['Interest-Bearing Debt', 'Other Liabilities', 'Equity'],
    series: [{ name: 'Amount', values: [debt, payables + otherCurrentLiabilities + otherLiabilities, equity] }],
  }

  return { kind: 'capital-optimization', rawSheets: [rawSheet], calcSheets: [calcSheet], keyMetrics, charts: [chart], extraCaveats }
}

// ---- Financial Projections calculator ----
//
// Unlike the other calculators, a 5-year, 3-scenario projection doesn't fit
// the generic account-list / label-value sheet shapes, so it renders into a
// dedicated year-columns grid (one sheet per scenario) via customSheetBuilders.
// Row numbers are fixed constants (not a dynamic running counter) so the
// key-metrics/chart figures computed here stay in exact sync with what the
// sheet-builder actually writes, even though the sheets themselves are only
// built later once the real workbook exists.

type FinancialProjectionsModel = ReportModelBase & { kind: 'financial-projections' }

const PROJECTION_YEARS = 5
const PROJECTION_YEAR_COLS = ['B', 'C', 'D', 'E', 'F']
const PROJECTION_ROWS = {
  header: 4,
  revenue: 5,
  cogs: 6,
  grossProfit: 7,
  opex: 8,
  ebitda: 9,
  depreciation: 10,
  interest: 11,
  pretax: 12,
  tax: 13,
  netIncome: 14,
  balanceSheetHeading: 16,
  currentAsset: 17,
  fixedAsset: 18,
  totalAssets: 19,
  currentLiability: 20,
  longTermDebt: 21,
  equity: 22,
  totalLiabEquity: 23,
  cashFlowHeading: 25,
  operatingCF: 26,
  investingCF: 27,
  financingCF: 28,
  netChangeCash: 29,
  kpiHeading: 31,
  grossMargin: 32,
  ebitdaMargin: 33,
  netMargin: 34,
  currentRatio: 35,
  debtToEquity: 36,
  roe: 37,
  roa: 38,
}

type ProjectionBase = {
  revenue: number
  cogsRatio: number
  opex: number
  depreciation: number
  interest: number
  effectiveTaxRate: number
  currentAsset: number
  fixedAsset: number
  currentLiability: number
  longTermDebt: number
  equity: number
  investingCF: number
  financingCF: number
  workingCapitalRatioCA: number
  workingCapitalRatioCL: number
}

type ProjectionRawRows = {
  revenue?: number
  opex?: number
  depreciation?: number
  interest?: number
  currentAsset?: number
  fixedAsset?: number
  currentLiability?: number
  longTermDebt?: number
  equity?: number
  investingCF?: number
  financingCF?: number
}

// Computes the same year-by-year figures the sheet-builder will render, purely
// numerically (no ExcelJS/sheet needed), so key metrics and the chart are
// available immediately without waiting for the deferred sheet build.
function projectFinancials(base: ProjectionBase, cagr: number, inflation: number) {
  const revenue: number[] = []
  const ebitda: number[] = []
  const netIncome: number[] = []
  let runningEquity = base.equity
  for (let y = 0; y < PROJECTION_YEARS; y++) {
    const revenueY = base.revenue * Math.pow(1 + cagr, y)
    const cogsY = revenueY * base.cogsRatio
    const grossY = revenueY - cogsY
    const opexY = base.opex * Math.pow(1 + inflation, y)
    const ebitdaY = grossY - opexY
    const pretaxY = ebitdaY - base.depreciation - base.interest
    const taxY = pretaxY * base.effectiveTaxRate
    const netIncomeY = pretaxY - taxY
    revenue.push(revenueY)
    ebitda.push(ebitdaY)
    netIncome.push(netIncomeY)
    runningEquity = y === 0 ? base.equity : runningEquity + netIncomeY
  }
  return { revenue, ebitda, netIncome, finalEquity: runningEquity }
}

function buildProjectionScenarioSheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  scenarioLabel: string,
  rawSheetName: string,
  base: ProjectionBase,
  rawRows: ProjectionRawRows,
  cagr: number,
  inflation: number,
) {
  const R = PROJECTION_ROWS
  const sheet = workbook.addWorksheet(sanitizeSheetName(`Projection - ${scenarioLabel}`, usedSheetNames))
  sheet.getColumn(1).width = 28
  for (let c = 2; c <= 6; c++) sheet.getColumn(c).width = 16

  sheet.getCell('A1').value = `5-Year Financial Projection — ${scenarioLabel}`
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:F1')
  sheet.getCell('A2').value = `Assumptions: ${(cagr * 100).toFixed(1)}%/yr revenue growth, ${(inflation * 100).toFixed(1)}%/yr cost inflation`
  sheet.getCell('A2').font = { size: 10, italic: true, color: { argb: 'FF555555' } }
  sheet.mergeCells('A2:F2')

  sheet.getCell(`A${R.header}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
  sheet.getCell(`A${R.header}`).border = THIN_BORDER
  PROJECTION_YEAR_COLS.forEach((col, i) => {
    const cell = sheet.getCell(`${col}${R.header}`)
    cell.value = `Year ${i + 1}`
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.alignment = { horizontal: 'center' }
    cell.border = THIN_BORDER
  })
  sheet.views = [{ state: 'frozen', ySplit: R.header }]

  const labelSpecs: [number, string, boolean?][] = [
    [R.revenue, 'Revenue', true],
    [R.cogs, 'COGS'],
    [R.grossProfit, 'Gross Profit', true],
    [R.opex, 'Operating Expenses'],
    [R.ebitda, 'EBITDA', true],
    [R.depreciation, 'Depreciation'],
    [R.interest, 'Interest Expense'],
    [R.pretax, 'Pre-tax Income'],
    [R.tax, 'Tax'],
    [R.netIncome, 'Net Income', true],
    [R.balanceSheetHeading, 'Balance Sheet'],
    [R.currentAsset, 'Current Assets'],
    [R.fixedAsset, 'Fixed Assets'],
    [R.totalAssets, 'Total Assets', true],
    [R.currentLiability, 'Current Liabilities'],
    [R.longTermDebt, 'Long-Term Debt'],
    [R.equity, 'Equity'],
    [R.totalLiabEquity, 'Total Liabilities & Equity', true],
    [R.cashFlowHeading, 'Cash Flow'],
    [R.operatingCF, 'Operating Cash Flow'],
    [R.investingCF, 'Investing Cash Flow'],
    [R.financingCF, 'Financing Cash Flow'],
    [R.netChangeCash, 'Net Change in Cash', true],
    [R.kpiHeading, 'Key Performance Indicators'],
    [R.grossMargin, 'Gross Margin %'],
    [R.ebitdaMargin, 'EBITDA Margin %'],
    [R.netMargin, 'Net Margin %'],
    [R.currentRatio, 'Current Ratio'],
    [R.debtToEquity, 'Debt-to-Equity (Total Liabilities ÷ Equity)'],
    [R.roe, 'ROE'],
    [R.roa, 'ROA'],
  ]
  const sectionHeadingRows = new Set([R.balanceSheetHeading, R.cashFlowHeading, R.kpiHeading])
  for (const [row, label, bold] of labelSpecs) {
    const cell = sheet.getCell(`A${row}`)
    cell.value = label
    cell.font = { bold: !!bold || sectionHeadingRows.has(row), color: sectionHeadingRows.has(row) ? { argb: NAVY } : undefined }
  }

  function setCell(row: number, col: string, value: MetricValue, format?: MetricFormat) {
    const cell = sheet.getCell(`${col}${row}`)
    setMetricCell(cell, value, format)
    cell.border = THIN_BORDER
  }

  let runningEquity = base.equity
  for (let y = 0; y < PROJECTION_YEARS; y++) {
    const col = PROJECTION_YEAR_COLS[y]
    const prevCol = y > 0 ? PROJECTION_YEAR_COLS[y - 1] : null
    const isYear1 = y === 0

    function flatOrYear1(rawRow: number | undefined, baseValue: number, thisRow: number): MetricValue {
      if (isYear1) return rawRow ? fx(refCell(rawSheetName, rawRow, 'D'), baseValue) : baseValue
      return fx(`=${PROJECTION_YEAR_COLS[0]}${thisRow}`, baseValue)
    }

    const revenueValue = base.revenue * Math.pow(1 + cagr, y)
    setCell(
      R.revenue,
      col,
      isYear1
        ? rawRows.revenue
          ? fx(refCell(rawSheetName, rawRows.revenue, 'D'), base.revenue)
          : base.revenue
        : fx(`=${prevCol}${R.revenue}*(1+${cagr})`, revenueValue),
      'currency',
    )

    const cogsValue = revenueValue * base.cogsRatio
    setCell(R.cogs, col, fx(`=${col}${R.revenue}*${base.cogsRatio}`, cogsValue), 'currency')

    const grossValue = revenueValue - cogsValue
    setCell(R.grossProfit, col, fx(`=${col}${R.revenue}-${col}${R.cogs}`, grossValue), 'currency')

    const opexValue = base.opex * Math.pow(1 + inflation, y)
    setCell(
      R.opex,
      col,
      isYear1
        ? rawRows.opex
          ? fx(refCell(rawSheetName, rawRows.opex, 'D'), base.opex)
          : base.opex
        : fx(`=${prevCol}${R.opex}*(1+${inflation})`, opexValue),
      'currency',
    )

    const ebitdaValue = grossValue - opexValue
    setCell(R.ebitda, col, fx(`=${col}${R.grossProfit}-${col}${R.opex}`, ebitdaValue), 'currency')

    const depValue = base.depreciation
    setCell(R.depreciation, col, flatOrYear1(rawRows.depreciation, base.depreciation, R.depreciation), 'currency')

    const intValue = base.interest
    setCell(R.interest, col, flatOrYear1(rawRows.interest, base.interest, R.interest), 'currency')

    const pretaxValue = ebitdaValue - depValue - intValue
    setCell(R.pretax, col, fx(`=${col}${R.ebitda}-${col}${R.depreciation}-${col}${R.interest}`, pretaxValue), 'currency')

    const taxValue = pretaxValue * base.effectiveTaxRate
    setCell(R.tax, col, fx(`=${col}${R.pretax}*${base.effectiveTaxRate}`, taxValue), 'currency')

    const netIncomeValue = pretaxValue - taxValue
    setCell(R.netIncome, col, fx(`=${col}${R.pretax}-${col}${R.tax}`, netIncomeValue), 'currency')

    const caValue = revenueValue * base.workingCapitalRatioCA
    setCell(
      R.currentAsset,
      col,
      isYear1
        ? rawRows.currentAsset
          ? fx(refCell(rawSheetName, rawRows.currentAsset, 'D'), base.currentAsset)
          : base.currentAsset
        : fx(`=${col}${R.revenue}*${base.workingCapitalRatioCA}`, caValue),
      'currency',
    )

    const faValue = base.fixedAsset
    setCell(R.fixedAsset, col, flatOrYear1(rawRows.fixedAsset, base.fixedAsset, R.fixedAsset), 'currency')

    const totalAssetsValue = caValue + faValue
    setCell(R.totalAssets, col, fx(`=${col}${R.currentAsset}+${col}${R.fixedAsset}`, totalAssetsValue), 'currency')

    const clValue = revenueValue * base.workingCapitalRatioCL
    setCell(
      R.currentLiability,
      col,
      isYear1
        ? rawRows.currentLiability
          ? fx(refCell(rawSheetName, rawRows.currentLiability, 'D'), base.currentLiability)
          : base.currentLiability
        : fx(`=${col}${R.revenue}*${base.workingCapitalRatioCL}`, clValue),
      'currency',
    )

    const ltdValue = base.longTermDebt
    setCell(R.longTermDebt, col, flatOrYear1(rawRows.longTermDebt, base.longTermDebt, R.longTermDebt), 'currency')

    let equityValue: number
    if (isYear1) {
      equityValue = base.equity
      setCell(R.equity, col, rawRows.equity ? fx(refCell(rawSheetName, rawRows.equity, 'D'), base.equity) : base.equity, 'currency')
    } else {
      equityValue = runningEquity + netIncomeValue
      setCell(R.equity, col, fx(`=${prevCol}${R.equity}+${col}${R.netIncome}`, equityValue), 'currency')
    }
    runningEquity = equityValue

    const totalLEValue = clValue + ltdValue + equityValue
    setCell(R.totalLiabEquity, col, fx(`=${col}${R.currentLiability}+${col}${R.longTermDebt}+${col}${R.equity}`, totalLEValue), 'currency')

    const opCFValue = netIncomeValue + depValue
    setCell(R.operatingCF, col, fx(`=${col}${R.netIncome}+${col}${R.depreciation}`, opCFValue), 'currency')

    const invCFValue = base.investingCF
    setCell(R.investingCF, col, flatOrYear1(rawRows.investingCF, base.investingCF, R.investingCF), 'currency')

    const finCFValue = base.financingCF
    setCell(R.financingCF, col, flatOrYear1(rawRows.financingCF, base.financingCF, R.financingCF), 'currency')

    const netChangeValue = opCFValue + invCFValue + finCFValue
    setCell(R.netChangeCash, col, fx(`=${col}${R.operatingCF}+${col}${R.investingCF}+${col}${R.financingCF}`, netChangeValue), 'currency')

    if (revenueValue !== 0) {
      setCell(R.grossMargin, col, fx(`=${col}${R.grossProfit}/${col}${R.revenue}`, grossValue / revenueValue), 'percent')
      setCell(R.ebitdaMargin, col, fx(`=${col}${R.ebitda}/${col}${R.revenue}`, ebitdaValue / revenueValue), 'percent')
      setCell(R.netMargin, col, fx(`=${col}${R.netIncome}/${col}${R.revenue}`, netIncomeValue / revenueValue), 'percent')
    } else {
      setCell(R.grossMargin, col, 'N/A')
      setCell(R.ebitdaMargin, col, 'N/A')
      setCell(R.netMargin, col, 'N/A')
    }
    setCell(R.currentRatio, col, clValue !== 0 ? fx(`=${col}${R.currentAsset}/${col}${R.currentLiability}`, caValue / clValue) : 'N/A', 'ratio')
    if (equityValue !== 0) {
      // Debt-to-equity = Total liabilities / Total shareholders' equity
      setCell(R.debtToEquity, col, fx(`=(${col}${R.currentLiability}+${col}${R.longTermDebt})/${col}${R.equity}`, (clValue + ltdValue) / equityValue), 'ratio')
      setCell(R.roe, col, fx(`=${col}${R.netIncome}/${col}${R.equity}`, netIncomeValue / equityValue), 'percent')
    } else {
      setCell(R.debtToEquity, col, 'N/A')
      setCell(R.roe, col, 'N/A')
    }
    setCell(
      R.roa,
      col,
      totalAssetsValue !== 0 ? fx(`=${col}${R.netIncome}/${col}${R.totalAssets}`, netIncomeValue / totalAssetsValue) : 'N/A',
      'percent',
    )
  }

  return sheet.name
}

function calculateFinancialProjections(lineItems: LineItem[]): FinancialProjectionsModel {
  const extraCaveats: string[] = []
  const rawSheetName = 'Year 1 Financials'
  const rawSheet = layoutRawDataSheet(rawSheetName, 'Year 1 (Base) Financials', lineItems, [
    { tag: 'revenue', label: 'Revenue' },
    { tag: 'cogs', label: 'COGS' },
    { tag: 'opex', label: 'Operating Expenses' },
    { tag: 'depreciation', label: 'Depreciation' },
    { tag: 'interest', label: 'Interest' },
    { tag: 'tax', label: 'Tax' },
    { tag: 'current_asset', label: 'Current Assets' },
    { tag: 'fixed_asset', label: 'Fixed Assets' },
    { tag: 'current_liability', label: 'Current Liabilities' },
    { tag: 'long_term_debt', label: 'Long-Term Debt' },
    { tag: 'equity', label: 'Equity' },
    { tag: 'operating_cash_flow_item', label: 'Operating Cash Flow' },
    { tag: 'investing_cash_flow_item', label: 'Investing Cash Flow' },
    { tag: 'financing_cash_flow_item', label: 'Financing Cash Flow' },
  ])

  const revenueBlock = findBlock(rawSheet, 'revenue')
  const cogsBlock = findBlock(rawSheet, 'cogs')
  const opexBlock = findBlock(rawSheet, 'opex')
  const depreciationBlock = findBlock(rawSheet, 'depreciation')
  const interestBlock = findBlock(rawSheet, 'interest')
  const taxBlock = findBlock(rawSheet, 'tax')
  const currentAssetBlock = findBlock(rawSheet, 'current_asset')
  const fixedAssetBlock = findBlock(rawSheet, 'fixed_asset')
  const currentLiabilityBlock = findBlock(rawSheet, 'current_liability')
  const longTermDebtBlock = findBlock(rawSheet, 'long_term_debt')
  const equityBlock = findBlock(rawSheet, 'equity')
  const opCFBlock = findBlock(rawSheet, 'operating_cash_flow_item')
  const invCFBlock = findBlock(rawSheet, 'investing_cash_flow_item')
  const finCFBlock = findBlock(rawSheet, 'financing_cash_flow_item')

  if (!revenueBlock) extraCaveats.push('No line items were classified as revenue; projections below could not be fully computed.')
  if (!equityBlock) extraCaveats.push('No line items were classified as equity; balance sheet and ROE projections below assume a starting equity of zero.')

  const revenue1 = revenueBlock?.subtotalValue ?? 0
  const cogs1 = cogsBlock?.subtotalValue ?? 0
  const opex1 = opexBlock?.subtotalValue ?? 0
  const depreciation1 = depreciationBlock?.subtotalValue ?? 0
  const interest1 = interestBlock?.subtotalValue ?? 0
  const tax1 = taxBlock?.subtotalValue ?? 0

  const pretax1 = revenue1 - cogs1 - opex1 - depreciation1 - interest1
  const effectiveTaxRate = pretax1 !== 0 ? tax1 / pretax1 : 0
  const cogsRatio = revenue1 !== 0 ? cogs1 / revenue1 : 0
  const workingCapitalRatioCA = revenue1 !== 0 ? (currentAssetBlock?.subtotalValue ?? 0) / revenue1 : 0
  const workingCapitalRatioCL = revenue1 !== 0 ? (currentLiabilityBlock?.subtotalValue ?? 0) / revenue1 : 0

  if (revenue1 === 0) {
    extraCaveats.push(
      'Year-1 revenue is zero — cost ratios, working capital scaling, and margin KPIs below default to zero and are not meaningful.',
    )
  }

  const base: ProjectionBase = {
    revenue: revenue1,
    cogsRatio,
    opex: opex1,
    depreciation: depreciation1,
    interest: interest1,
    effectiveTaxRate,
    currentAsset: currentAssetBlock?.subtotalValue ?? 0,
    fixedAsset: fixedAssetBlock?.subtotalValue ?? 0,
    currentLiability: currentLiabilityBlock?.subtotalValue ?? 0,
    longTermDebt: longTermDebtBlock?.subtotalValue ?? 0,
    equity: equityBlock?.subtotalValue ?? 0,
    investingCF: invCFBlock?.subtotalValue ?? 0,
    financingCF: finCFBlock?.subtotalValue ?? 0,
    workingCapitalRatioCA,
    workingCapitalRatioCL,
  }
  const rawRows: ProjectionRawRows = {
    revenue: revenueBlock?.subtotalRow,
    opex: opexBlock?.subtotalRow,
    depreciation: depreciationBlock?.subtotalRow,
    interest: interestBlock?.subtotalRow,
    currentAsset: currentAssetBlock?.subtotalRow,
    fixedAsset: fixedAssetBlock?.subtotalRow,
    currentLiability: currentLiabilityBlock?.subtotalRow,
    longTermDebt: longTermDebtBlock?.subtotalRow,
    equity: equityBlock?.subtotalRow,
    investingCF: invCFBlock?.subtotalRow,
    financingCF: finCFBlock?.subtotalRow,
  }

  const { annualInflationRate, baseRevenueCagr, optimisticCagrOffset, pessimisticCagrOffset } = PLATFORM_ASSUMPTIONS
  const scenarios = [
    { label: 'Base Case', cagr: baseRevenueCagr },
    { label: 'Optimistic', cagr: baseRevenueCagr + optimisticCagrOffset },
    { label: 'Pessimistic', cagr: Math.max(baseRevenueCagr - pessimisticCagrOffset, -0.95) },
  ]

  const customSheetBuilders = scenarios.map(
    (scenario) => (workbook: ExcelJS.Workbook, usedSheetNames: Set<string>) =>
      buildProjectionScenarioSheet(workbook, usedSheetNames, scenario.label, rawSheetName, base, rawRows, scenario.cagr, annualInflationRate),
  )

  const baseProjection = projectFinancials(base, baseRevenueCagr, annualInflationRate)
  const baseSheetName = 'Projection - Base Case'
  const keyMetrics: KeyMetric[] = [
    { label: 'Year 1 Revenue', cellRef: refCell(baseSheetName, PROJECTION_ROWS.revenue, 'B'), value: baseProjection.revenue[0], format: 'currency' },
    {
      label: 'Year 5 Revenue (Base Case)',
      cellRef: refCell(baseSheetName, PROJECTION_ROWS.revenue, 'F'),
      value: baseProjection.revenue[4],
      format: 'currency',
    },
    {
      label: 'Year 5 Net Income (Base Case)',
      cellRef: refCell(baseSheetName, PROJECTION_ROWS.netIncome, 'F'),
      value: baseProjection.netIncome[4],
      format: 'currency',
    },
    {
      label: 'Year 5 EBITDA Margin (Base Case)',
      cellRef: refCell(baseSheetName, PROJECTION_ROWS.ebitdaMargin, 'F'),
      value: baseProjection.revenue[4] !== 0 ? baseProjection.ebitda[4] / baseProjection.revenue[4] : 0,
      format: 'percent',
    },
  ]

  const chart: ReportChart = {
    type: 'line',
    title: '5-Year Projection (Base Case) — Revenue, EBITDA, Net Income',
    categories: ['Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'],
    series: [
      { name: 'Revenue', values: baseProjection.revenue.map((v) => Math.round(v)) },
      { name: 'EBITDA', values: baseProjection.ebitda.map((v) => Math.round(v)) },
      { name: 'Net Income', values: baseProjection.netIncome.map((v) => Math.round(v)) },
    ],
  }

  extraCaveats.push(
    `Years 2-5 are projected from Year-1 figures using platform-standard assumptions: ${(baseRevenueCagr * 100).toFixed(1)}% base-case revenue CAGR (optimistic +${(optimisticCagrOffset * 100).toFixed(0)}pp, pessimistic -${(pessimisticCagrOffset * 100).toFixed(0)}pp) and ${(annualInflationRate * 100).toFixed(1)}% annual operating-expense inflation, applied identically to every upload. Depreciation, interest, fixed assets, long-term debt, and investing/financing cash flows are held flat at Year-1 levels; COGS and working capital (current assets/liabilities) scale proportionally with revenue.`,
  )
  extraCaveats.push(
    'The projected balance sheet is a simplified, ratio-driven model, not a fully articulated three-statement forecast: equity rolls forward via retained earnings independently of the asset side, with no cash-balancing mechanism. Total Assets and Total Liabilities & Equity will not reconcile exactly in Years 2-5 — treat the balance sheet rows as directional, not precise.',
  )

  return {
    kind: 'financial-projections',
    rawSheets: [rawSheet],
    calcSheets: [],
    keyMetrics,
    charts: [chart],
    extraCaveats,
    customSheetBuilders,
  }
}

export function calculateReportModel(reportType: string, lineItems: LineItem[]): ReportModelBase | null {
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
      return calculateFinancialProjections(lineItems)
    default:
      return null
  }
}

// ── Excel workbook rendering ────────────────────────────────────────────────

function addWrappedParagraphs(sheet: ExcelJS.Worksheet, text: string, startRow: number) {
  let row = startRow
  for (const para of text.split('\n').map((p) => p.trim()).filter(Boolean)) {
    const cell = sheet.getCell(`A${row}`)
    cell.value = para
    cell.alignment = { wrapText: true, vertical: 'top' }
    sheet.getRow(row).height = Math.max(20, Math.ceil(para.length / 105) * 15)
    row += 1
  }
  return row
}

function sanitizeSheetName(name: string, used: Set<string>) {
  let base = name.replace(/[\\/*?:[\]]/g, ' ').trim().slice(0, 31) || 'Sheet'
  let candidate = base
  let suffix = 1
  while (used.has(candidate)) {
    suffix += 1
    candidate = `${base.slice(0, 28)} ${suffix}`
  }
  used.add(candidate)
  return candidate
}

// Renders executive summary / sections / next steps / charts into an existing
// worksheet starting at `startRow`. Shared by both the legacy fully-AI-driven
// workbook and the new deterministic workbook's Summary sheet.
// Draws each chart's title + cached PNG image into `sheet`, stacked
// vertically starting at `startRow`. Shared by the narrative Summary sheet
// (contextual, interleaved with prose) and the dedicated Charts sheet
// (all charts collected together).
function addChartImagesToSheet(
  workbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  charts: ReportChart[],
  chartPngCache: Map<ReportChart, Buffer>,
  startRow: number,
) {
  let row = startRow
  for (const chart of charts) {
    const pngBuffer = chartPngCache.get(chart)
    if (!pngBuffer) continue
    const imageId = workbook.addImage({ buffer: pngBuffer, extension: 'png' })
    sheet.addImage(imageId, { tl: { col: 0, row }, ext: { width: 640, height: 299 } })
    row += 17 // reserve rows so later content doesn't overlap the floating image
  }
  return row
}

async function renderNarrativeSummary(
  workbook: ExcelJS.Workbook,
  summary: ExcelJS.Worksheet,
  reportData: ReportData,
  startRow: number,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  let row = startRow
  summary.getCell(`A${row}`).value = 'Executive Summary'
  summary.getCell(`A${row}`).font = { size: 14, bold: true, color: { argb: NAVY } }
  row += 1
  row = addWrappedParagraphs(summary, reportData.executiveSummary, row)
  row += 1

  for (const section of reportData.sections) {
    const headingCell = summary.getCell(`A${row}`)
    headingCell.value = section.heading
    headingCell.font = { size: 13, bold: true, color: { argb: NAVY } }
    headingCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BAND } }
    row += 1

    if (section.body) row = addWrappedParagraphs(summary, section.body, row)

    if (section.bullets?.length) {
      for (const bullet of section.bullets) {
        const cell = summary.getCell(`A${row}`)
        cell.value = `•  ${bullet}`
        cell.alignment = { wrapText: true, vertical: 'top' }
        row += 1
      }
    }
    row += 1
  }

  if (reportData.nextSteps?.length) {
    summary.getCell(`A${row}`).value = 'Next Steps'
    summary.getCell(`A${row}`).font = { size: 13, bold: true, color: { argb: NAVY } }
    row += 1
    reportData.nextSteps.forEach((step, i) => {
      const cell = summary.getCell(`A${row}`)
      cell.value = `${i + 1}. ${step}`
      cell.alignment = { wrapText: true, vertical: 'top' }
      row += 1
    })
    row += 1
  }

  row = addChartImagesToSheet(workbook, summary, reportData.charts, chartPngCache, row)

  return row
}

// A dedicated tab collecting every chart from the report (deterministic
// model charts + any supplementary AI-authored charts) in one place, so the
// final-outcome visuals can be reviewed or presented without hunting through
// the narrative Summary sheet.
function buildChartsSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, charts: ReportChart[], chartPngCache: Map<ReportChart, Buffer>) {
  if (!charts.length) return
  const sheet = workbook.addWorksheet(sanitizeSheetName('Charts', usedSheetNames))
  sheet.getColumn(1).width = 14

  sheet.getCell('A1').value = 'Charts'
  sheet.getCell('A1').font = { size: 16, bold: true, color: { argb: NAVY } }

  let row = 3
  for (const chart of charts) {
    if (!chartPngCache.has(chart)) continue
    sheet.getCell(`A${row}`).value = chart.title
    sheet.getCell(`A${row}`).font = { size: 12, bold: true, color: { argb: NAVY } }
    row += 1
    row = addChartImagesToSheet(workbook, sheet, [chart], chartPngCache, row)
    row += 2
  }
}

// One sheet per supplementary AI-authored table — the escape hatch for
// qualitative/non-deterministic content that doesn't fit the line-item model.
function renderTableSheets(workbook: ExcelJS.Workbook, tables: ReportTable[], usedSheetNames: Set<string>) {
  for (const table of tables) {
    if (!table.columns?.length) continue
    if (!Array.isArray(table.rows)) table.rows = []
    const sheet = workbook.addWorksheet(sanitizeSheetName(table.title || 'Data', usedSheetNames))

    sheet.addRow(table.columns)
    sheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    })
    sheet.views = [{ state: 'frozen', ySplit: 1 }]
    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: table.columns.length } }

    for (const dataRow of table.rows) sheet.addRow(dataRow)

    table.columns.forEach((col, ci) => {
      const headerLower = col.toLowerCase()
      const maxContentLength = Math.max(col.length, ...table.rows.map((r) => String(r[ci] ?? '').length), 1)
      const column = sheet.getColumn(ci + 1)
      column.width = Math.min(40, Math.max(12, maxContentLength + 3))

      const isNumericColumn = table.rows.length > 0 && table.rows.every((r) => typeof r[ci] === 'number')
      if (isNumericColumn) {
        if (/%|percent|margin|rate|ratio/.test(headerLower)) {
          column.numFmt = '0.0%'
        } else if (/€|\$|revenue|cost|profit|cash|amount|ebitda|price|budget|expense|income|funding|burn/.test(headerLower)) {
          column.numFmt = '#,##0.00'
        } else {
          column.numFmt = '#,##0'
        }

        if (table.rows.length > 1) {
          const colLetter = column.letter
          sheet.addConditionalFormatting({
            ref: `${colLetter}2:${colLetter}${table.rows.length + 1}`,
            rules: [
              {
                type: 'dataBar',
                priority: 1,
                gradient: true,
                cfvo: [{ type: 'min' }, { type: 'max' }],
              },
            ],
          })
        }
      }
    })

    for (let r = 2; r <= table.rows.length + 1; r += 2) {
      sheet.getRow(r).eachCell({ includeEmpty: true }, (cell) => {
        if (!cell.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BAND } }
      })
    }
  }
}

// @netlify/blobs' set() only accepts string | ArrayBuffer | Blob — a Node
// Buffer is a Uint8Array view, not itself an ArrayBuffer, so slice it out.
function bufferToArrayBuffer(buffer: Buffer): ArrayBuffer {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer
}

function workbookToArrayBuffer(buffer: ExcelJS.Buffer): ArrayBuffer {
  return bufferToArrayBuffer(buffer as unknown as Buffer)
}

// ── Branding: cover page + a discrete logo mark on every other sheet ───────

const LOGO_ASPECT = LOGO_HEIGHT / LOGO_WIDTH

// A dedicated title/cover sheet, always added first. Carries the logo at a
// modest, tasteful size — the "prominent" cover treatment is still restrained
// (this platform's reports are meant to look like they belong on a CFO's
// desk, not a pitch deck), consistent with the small corner mark used on
// every other sheet.
function buildCoverSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, logoImageId: number, reportLabel: string, sourceFileName: string) {
  const sheet = workbook.addWorksheet(sanitizeSheetName('Cover', usedSheetNames))
  sheet.getColumn(1).width = 100

  const logoWidth = 200
  const logoHeight = Math.round(logoWidth * LOGO_ASPECT)
  sheet.addImage(logoImageId, { tl: { col: 0.15, row: 1.3 }, ext: { width: logoWidth, height: logoHeight } })

  const titleRow = 9
  sheet.getCell(`A${titleRow}`).value = reportLabel
  sheet.getCell(`A${titleRow}`).font = { size: 22, bold: true, color: { argb: NAVY } }

  sheet.getCell(`A${titleRow + 2}`).value = `Source file: ${sourceFileName}`
  sheet.getCell(`A${titleRow + 2}`).font = { size: 12, color: { argb: 'FF555555' } }

  sheet.getCell(`A${titleRow + 3}`).value = `Generated: ${new Date().toLocaleString()}`
  sheet.getCell(`A${titleRow + 3}`).font = { size: 12, color: { argb: 'FF555555' } }

  sheet.getCell(`A${titleRow + 5}`).value = 'Prepared by the Business Reform AI CFO platform.'
  sheet.getCell(`A${titleRow + 5}`).font = { size: 11, italic: true, color: { argb: 'FF8A98A8' } }

  sheet.getCell(`A${titleRow + 7}`).value = aiMarkLine('en')
  sheet.getCell(`A${titleRow + 7}`).font = { size: 10, bold: true, color: { argb: NAVY } }
}

// Stamps a small logo mark in the top-right corner of every sheet in the
// workbook except the Cover sheet (which already carries the logo at a more
// prominent size) — run once, after all sheets exist, so it doesn't need to
// be threaded through every individual sheet-builder function.
function addDiscreteLogoToAllSheets(workbook: ExcelJS.Workbook, logoImageId: number) {
  const logoWidth = 60
  const logoHeight = Math.round(logoWidth * LOGO_ASPECT)
  for (const sheet of workbook.worksheets) {
    if (sheet.name === 'Cover') continue
    const anchorCol = Math.max(sheet.columnCount || 1, 1)
    sheet.addImage(logoImageId, { tl: { col: anchorCol, row: 0.1 }, ext: { width: logoWidth, height: logoHeight } })
  }
}

async function buildExcelWorkbook(reportData: ReportData, reportLabel: string, sourceFileName: string, chartPngCache: Map<ReportChart, Buffer>) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Business Reform AI CFO'
  workbook.created = new Date()

  const usedSheetNames = new Set<string>()
  const logoImageId = workbook.addImage({ base64: LOGO_PNG_BASE64, extension: 'png' })
  buildCoverSheet(workbook, usedSheetNames, logoImageId, reportLabel, sourceFileName)

  // --- Summary sheet: narrative + embedded chart images ---
  const summary = workbook.addWorksheet(sanitizeSheetName('Summary', usedSheetNames))
  summary.getColumn(1).width = 100

  summary.getCell('A1').value = reportLabel
  summary.getCell('A1').font = { size: 18, bold: true, color: { argb: NAVY } }

  summary.getCell('A2').value = `Source file: ${sourceFileName}  |  Generated: ${new Date().toLocaleString()}`
  summary.getCell('A2').font = { size: 10, italic: true, color: { argb: 'FF555555' } }

  await renderNarrativeSummary(workbook, summary, reportData, 4, chartPngCache)
  buildChartsSheet(workbook, usedSheetNames, reportData.charts, chartPngCache)

  // --- One sheet per data table, properly formatted ---
  renderTableSheets(workbook, reportData.tables, usedSheetNames)

  addDiscreteLogoToAllSheets(workbook, logoImageId)
  applyAiMarkToWorkbook(workbook, 'en')

  const buffer = await workbook.xlsx.writeBuffer()
  return workbookToArrayBuffer(buffer)
}

// ── Deterministic Excel workbook rendering ──────────────────────────────────
//
// Renders a ReportModelBase (raw-data sheets + calculations sheets, all
// formula-driven) plus the narrative/assumptions content, into a workbook
// whose sheet order is always: Summary → raw-data sheet(s) → calculations
// sheet(s) → any supplementary AI tables/charts (the escape hatch) →
// Assumptions & Notes last.

function numFmtFor(format?: MetricFormat) {
  if (format === 'percent') return '0.00%'
  if (format === 'ratio') return '0.00"x"'
  if (format === 'number') return '#,##0.0'
  return '#,##0.00" €"'
}

function setMetricCell(cell: ExcelJS.Cell, value: MetricValue, format?: MetricFormat) {
  if (typeof value === 'string') {
    cell.value = value
    cell.font = { italic: true, color: { argb: 'FF884400' } }
    return
  }
  if (typeof value === 'number') {
    cell.value = value
    cell.numFmt = numFmtFor(format)
    return
  }
  cell.value = { formula: value.formula.replace(/^=/, ''), result: value.result }
  cell.numFmt = numFmtFor(format)
}

function buildKeyMetricsTable(summary: ExcelJS.Worksheet, startRow: number, keyMetrics: KeyMetric[]) {
  let row = startRow
  summary.getCell(`A${row}`).value = 'Key Metrics'
  summary.getCell(`A${row}`).font = { size: 13, bold: true, color: { argb: NAVY } }
  row += 1

  const headerRow = row
  summary.getCell(`A${headerRow}`).value = 'Metric'
  summary.getCell(`B${headerRow}`).value = 'Value'
  ;['A', 'B'].forEach((col) => {
    const cell = summary.getCell(`${col}${headerRow}`)
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.border = THIN_BORDER
    cell.alignment = { horizontal: 'center' }
  })
  row += 1

  for (const metric of keyMetrics) {
    const labelCell = summary.getCell(`A${row}`)
    labelCell.value = metric.label
    labelCell.border = THIN_BORDER
    const valueCell = summary.getCell(`B${row}`)
    setMetricCell(valueCell, fx(metric.cellRef, metric.value), metric.format)
    valueCell.border = THIN_BORDER
    row += 1
  }
  return row + 1
}

function buildCaveatCallout(summary: ExcelJS.Worksheet, startRow: number, assumptions: ReportAssumptions) {
  const caveat = assumptions.dataCaveats[0] || assumptions.limitations[0]
  if (!caveat) return startRow
  const row = startRow
  summary.mergeCells(`A${row}:B${row + 2}`)
  const cell = summary.getCell(`A${row}`)
  cell.value = `Caveat: ${caveat} See the Assumptions & Notes sheet for full detail.`
  cell.font = { italic: true, size: 10, color: { argb: 'FF663C00' } }
  cell.alignment = { wrapText: true, vertical: 'top' }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: CAVEAT_FILL } }
  return row + 4
}

async function buildDeterministicSummarySheet(
  workbook: ExcelJS.Workbook,
  usedSheetNames: Set<string>,
  reportData: ReportData,
  model: ReportModelBase,
  reportLabel: string,
  sourceFileName: string,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  const summary = workbook.addWorksheet(sanitizeSheetName('Summary', usedSheetNames))
  summary.getColumn(1).width = 55
  summary.getColumn(2).width = 45

  summary.getCell('A1').value = reportLabel
  summary.getCell('A1').font = { size: 18, bold: true, color: { argb: NAVY } }
  summary.mergeCells('A1:B1')

  summary.getCell('A2').value = `Source file: ${sourceFileName}  |  Generated: ${new Date().toLocaleString()}`
  summary.getCell('A2').font = { size: 10, italic: true, color: { argb: 'FF555555' } }
  summary.mergeCells('A2:B2')

  let row = 4
  if (model.keyMetrics.length) row = buildKeyMetricsTable(summary, row, model.keyMetrics)
  row = buildCaveatCallout(summary, row, reportData.assumptions)
  await renderNarrativeSummary(workbook, summary, { ...reportData, charts: [...model.charts, ...reportData.charts] }, row, chartPngCache)
}

function buildRawDataSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, model: RawDataSheetModel) {
  const sheet = workbook.addWorksheet(sanitizeSheetName(model.sheetName, usedSheetNames))
  sheet.getColumn(1).width = 12
  sheet.getColumn(2).width = 42
  sheet.getColumn(3).width = 18
  sheet.getColumn(4).width = 18

  sheet.getCell('A1').value = model.title
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:D1')

  const headerRow = 3
  ;['Account', 'Description', 'Group', 'Amount'].forEach((label, i) => {
    const cell = sheet.getCell(headerRow, i + 1)
    cell.value = label
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    cell.alignment = { horizontal: 'center' }
    cell.border = THIN_BORDER
  })
  sheet.views = [{ state: 'frozen', ySplit: headerRow }]

  for (const dataRow of model.rows) {
    if (dataRow.kind === 'item') {
      sheet.getCell(`A${dataRow.row}`).value = dataRow.account || ''
      sheet.getCell(`B${dataRow.row}`).value = dataRow.description
      sheet.getCell(`C${dataRow.row}`).value = dataRow.groupLabel
      const amountCell = sheet.getCell(`D${dataRow.row}`)
      amountCell.value = dataRow.amount
      amountCell.numFmt = '#,##0.00" €"'
      ;['A', 'B', 'C', 'D'].forEach((col) => (sheet.getCell(`${col}${dataRow.row}`).border = THIN_BORDER))
    } else {
      sheet.mergeCells(`A${dataRow.row}:C${dataRow.row}`)
      const labelCell = sheet.getCell(`A${dataRow.row}`)
      labelCell.value = dataRow.label
      labelCell.font = { bold: true }
      const valueCell = sheet.getCell(`D${dataRow.row}`)
      setMetricCell(valueCell, dataRow.value)
      const fillColor = dataRow.kind === 'total' ? RESULT_FILL : SUBTOTAL_FILL
      ;['A', 'B', 'C', 'D'].forEach((col) => {
        const cell = sheet.getCell(`${col}${dataRow.row}`)
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } }
        cell.border = THIN_BORDER
      })
    }
  }
}

function buildCalculationsSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, model: CalcSheetModel) {
  const sheet = workbook.addWorksheet(sanitizeSheetName(model.sheetName, usedSheetNames))
  sheet.getColumn(1).width = 4
  sheet.getColumn(2).width = 48
  sheet.getColumn(3).width = 22

  sheet.getCell('A1').value = model.title
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }
  sheet.mergeCells('A1:C1')

  for (const dataRow of model.rows) {
    if (dataRow.kind === 'sectionHeading') {
      sheet.mergeCells(`B${dataRow.row}:C${dataRow.row}`)
      const cell = sheet.getCell(`B${dataRow.row}`)
      cell.value = dataRow.label
      cell.font = { size: 12, bold: true, color: { argb: NAVY } }
    } else if (dataRow.kind === 'columnHeader') {
      ;['B', 'C'].forEach((col, i) => {
        const cell = sheet.getCell(`${col}${dataRow.row}`)
        cell.value = i === 0 ? 'Metric' : 'Value'
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
        cell.alignment = { horizontal: 'center' }
        cell.border = THIN_BORDER
      })
    } else {
      const labelCell = sheet.getCell(`B${dataRow.row}`)
      labelCell.value = dataRow.label
      labelCell.font = { bold: !!dataRow.bold }
      labelCell.border = THIN_BORDER
      const valueCell = sheet.getCell(`C${dataRow.row}`)
      setMetricCell(valueCell, dataRow.value, dataRow.format)
      if (dataRow.bold) valueCell.font = { ...valueCell.font, bold: true }
      valueCell.border = THIN_BORDER
      if (dataRow.highlight) {
        ;['B', 'C'].forEach((col) => {
          sheet.getCell(`${col}${dataRow.row}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RESULT_FILL } }
        })
      }
    }
  }
}

function buildAssumptionsSheet(workbook: ExcelJS.Workbook, usedSheetNames: Set<string>, assumptions: ReportAssumptions, extraCaveats: string[]) {
  const sheet = workbook.addWorksheet(sanitizeSheetName('Assumptions & Notes', usedSheetNames))
  sheet.getColumn(1).width = 100

  sheet.getCell('A1').value = 'Assumptions, Methodology & Limitations'
  sheet.getCell('A1').font = { size: 14, bold: true, color: { argb: NAVY } }

  let row = 3
  const sections: { heading: string; items: string[] }[] = [
    { heading: 'Data Caveats', items: assumptions.dataCaveats || [] },
    { heading: 'Excluded Items', items: assumptions.excludedItems || [] },
    { heading: 'Classification Rationale', items: assumptions.classificationRationale || [] },
    { heading: 'Limitations', items: [...(assumptions.limitations || []), ...extraCaveats] },
  ]

  for (const section of sections) {
    if (!section.items.length) continue
    const headingCell = sheet.getCell(`A${row}`)
    headingCell.value = section.heading
    headingCell.font = { size: 12, bold: true, color: { argb: NAVY } }
    row += 1
    for (const item of section.items) {
      const cell = sheet.getCell(`A${row}`)
      cell.value = `•  ${item}`
      cell.alignment = { wrapText: true, vertical: 'top' }
      sheet.getRow(row).height = Math.max(20, Math.ceil(item.length / 105) * 15)
      row += 1
    }
    row += 1
  }
}

export async function buildDeterministicExcelWorkbook(
  reportData: ReportData,
  model: ReportModelBase,
  reportLabel: string,
  sourceFileName: string,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Business Reform AI CFO'
  workbook.created = new Date()
  workbook.calcProperties.fullCalcOnLoad = true

  const usedSheetNames = new Set<string>()
  const allCharts = [...model.charts, ...reportData.charts]
  const logoImageId = workbook.addImage({ base64: LOGO_PNG_BASE64, extension: 'png' })
  buildCoverSheet(workbook, usedSheetNames, logoImageId, reportLabel, sourceFileName)

  await buildDeterministicSummarySheet(workbook, usedSheetNames, reportData, model, reportLabel, sourceFileName, chartPngCache)
  for (const rawSheet of model.rawSheets) buildRawDataSheet(workbook, usedSheetNames, rawSheet)
  for (const calcSheet of model.calcSheets) buildCalculationsSheet(workbook, usedSheetNames, calcSheet)
  for (const buildCustomSheet of model.customSheetBuilders || []) buildCustomSheet(workbook, usedSheetNames)
  buildChartsSheet(workbook, usedSheetNames, allCharts, chartPngCache)
  renderTableSheets(workbook, reportData.tables, usedSheetNames)
  buildAssumptionsSheet(workbook, usedSheetNames, reportData.assumptions, model.extraCaveats)

  addDiscreteLogoToAllSheets(workbook, logoImageId)
  applyAiMarkToWorkbook(workbook, 'en')

  const buffer = await workbook.xlsx.writeBuffer()
  return workbookToArrayBuffer(buffer)
}

// ── PowerPoint deck rendering ────────────────────────────────────────────────
//
// A slide-form companion to the Excel workbook: title, executive summary, key
// metrics table, one slide per chart (reusing the same cached PNGs as the
// Excel Charts sheet), and next steps — generated from the same ReportData /
// ReportModelBase the Excel workbook is built from, so the two artifacts
// never disagree.

const PPTX_NAVY = '1F3B5B'
const PPTX_NAVY_DEEP = '15293D'
const PPTX_INK = '2B2B2B'
const PPTX_MUTED = 'D9E2EC'
const PPTX_GOLD = 'C9A227'
const PPTX_CARD_BG = 'F4F7FB'
const PPTX_CARD_BORDER = 'E3E9F0'
const PPTX_LABEL_GREY = '6B7785'
const PPTX_SLIDE_W = 13.33
const PPTX_SLIDE_H = 7.5
const PPTX_MARGIN = 0.55

function formatMetricForSlide(value: number, format?: MetricFormat): string {
  if (format === 'percent') return `${(value * 100).toFixed(2)}%`
  if (format === 'ratio') return `${value.toFixed(2)}x`
  if (format === 'number') return value.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })
  return `${value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`
}

// A consistent masthead for every content slide: a short gold tick mark, the
// slide title, and a thin rule underneath — gives the deck a designed,
// templated feel instead of a bare heading floating on white.
function addContentSlideHeader(slide: PptxGenJSType.Slide, title: string) {
  slide.addShape('rect', { x: PPTX_MARGIN, y: 0.42, w: 0.5, h: 0.06, fill: { color: PPTX_GOLD }, line: { type: 'none' } })
  slide.addText(title, {
    x: PPTX_MARGIN,
    y: 0.54,
    w: PPTX_SLIDE_W - PPTX_MARGIN * 2,
    h: 0.65,
    fontSize: 24,
    bold: true,
    color: PPTX_NAVY,
    fontFace: 'Arial',
  })
  slide.addShape('line', {
    x: PPTX_MARGIN,
    y: 1.28,
    w: PPTX_SLIDE_W - PPTX_MARGIN * 2,
    h: 0,
    line: { color: PPTX_CARD_BORDER, width: 1 },
  })
}

export async function buildPowerPointPresentation(
  reportData: ReportData,
  model: ReportModelBase | null,
  reportLabel: string,
  sourceFileName: string,
  chartPngCache: Map<ReportChart, Buffer>,
) {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'Business Reform AI CFO'
  pptx.title = reportLabel
  pptx.subject = `${AI_MARK.en} — ${AI_METADATA_DESCRIPTION} Keywords: ${AI_METADATA_KEYWORDS}`

  // LAYOUT_WIDE is 13.33 x 7.5in. A small logo mark, top-right, on every
  // slide — via a slide master for content slides, and placed directly on
  // the title slide (which uses its own navy background, not the master).
  const logoDataUri = `image/png;base64,${LOGO_PNG_BASE64}`
  const logoW = 0.85
  const logoH = logoW * LOGO_ASPECT
  const logoX = PPTX_SLIDE_W - logoW - 0.35
  const logoY = 0.25

  // "AI generated" mark (EU AI Act, Art. 50), bottom-left of every slide.
  const aiMarkText = aiMarkLine('en')
  const aiMarkOptions = { x: 0.35, y: PPTX_SLIDE_H - 0.4, w: 7, h: 0.28, fontSize: 9, fontFace: 'Arial', valign: 'middle' as const }

  pptx.defineSlideMaster({
    title: 'CONTENT_MASTER',
    background: { color: 'FFFFFF' },
    objects: [
      { image: { data: logoDataUri, x: logoX, y: logoY, w: logoW, h: logoH } },
      { text: { text: aiMarkText, options: { ...aiMarkOptions, color: PPTX_LABEL_GREY } } },
    ],
  })

  // --- Title / cover slide: navy field with a soft layered circle accent
  // bleeding off the bottom-right corner, a gold kicker + rule around the
  // title, for a designed opening rather than a flat color block. ---
  const titleSlide = pptx.addSlide()
  titleSlide.background = { color: PPTX_NAVY }
  titleSlide.addShape('ellipse', { x: 9.0, y: 4.6, w: 6.2, h: 6.2, fill: { color: PPTX_NAVY_DEEP }, line: { type: 'none' } })
  titleSlide.addShape('ellipse', { x: 10.6, y: 5.7, w: 3.6, h: 3.6, fill: { color: PPTX_GOLD, transparency: 82 }, line: { type: 'none' } })

  // The logo's navy elements would be invisible directly on this slide's
  // navy background, so give it a small white backing chip for contrast.
  titleSlide.addShape('rect', {
    x: logoX - 0.12,
    y: logoY - 0.1,
    w: logoW + 0.24,
    h: logoH + 0.2,
    fill: { color: 'FFFFFF' },
    line: { type: 'none' },
  })
  titleSlide.addImage({ data: logoDataUri, x: logoX, y: logoY, w: logoW, h: logoH })
  titleSlide.addText(aiMarkText, { ...aiMarkOptions, color: 'FFFFFF' })

  titleSlide.addShape('rect', { x: 0.6, y: 1.65, w: 0.5, h: 0.055, fill: { color: PPTX_GOLD }, line: { type: 'none' } })
  titleSlide.addText('AI CFO FINANCIAL REPORT', {
    x: 0.6,
    y: 1.78,
    w: 9,
    h: 0.35,
    fontSize: 12,
    bold: true,
    color: PPTX_GOLD,
    charSpacing: 2,
    fontFace: 'Arial',
  })
  titleSlide.addText(reportLabel, { x: 0.6, y: 2.25, w: 12.13, h: 1.4, fontSize: 32, bold: true, color: 'FFFFFF', fontFace: 'Arial' })
  titleSlide.addShape('rect', { x: 0.6, y: 3.72, w: 1.1, h: 0.04, fill: { color: PPTX_GOLD }, line: { type: 'none' } })
  titleSlide.addText(`Source file: ${sourceFileName}`, { x: 0.6, y: 3.95, w: 12.13, h: 0.4, fontSize: 14, color: PPTX_MUTED, fontFace: 'Arial' })
  titleSlide.addText(`Generated: ${new Date().toLocaleString()}`, { x: 0.6, y: 4.35, w: 12.13, h: 0.4, fontSize: 14, color: PPTX_MUTED, fontFace: 'Arial' })

  // --- Executive Summary: prose set inside a card with a gold rule down the
  // left edge, like a pull-quote — reads as an editorial callout rather than
  // a plain paragraph dropped onto a slide. ---
  if (reportData.executiveSummary) {
    const slide = pptx.addSlide({ masterName: 'CONTENT_MASTER' })
    addContentSlideHeader(slide, 'Executive Summary')
    const cardX = PPTX_MARGIN
    const cardY = 1.55
    const cardW = PPTX_SLIDE_W - PPTX_MARGIN * 2
    const cardH = PPTX_SLIDE_H - cardY - 0.5
    slide.addShape('roundRect', {
      x: cardX,
      y: cardY,
      w: cardW,
      h: cardH,
      rectRadius: 0.08,
      fill: { color: PPTX_CARD_BG },
      line: { color: PPTX_CARD_BORDER, width: 1 },
    })
    slide.addShape('rect', { x: cardX, y: cardY, w: 0.07, h: cardH, fill: { color: PPTX_GOLD }, line: { type: 'none' } })
    slide.addText(reportData.executiveSummary, {
      x: cardX + 0.4,
      y: cardY + 0.3,
      w: cardW - 0.75,
      h: cardH - 0.6,
      fontSize: 15,
      color: PPTX_INK,
      valign: 'top',
      lineSpacingMultiple: 1.35,
      fontFace: 'Arial',
    })
  }

  // --- Key Metrics: a grid of metric cards (big number + label) instead of
  // an Excel-style two-column table — the standard "stat cards" pattern for
  // presentation decks. ---
  if (model?.keyMetrics.length) {
    const slide = pptx.addSlide({ masterName: 'CONTENT_MASTER' })
    addContentSlideHeader(slide, 'Key Metrics')

    const metrics = model.keyMetrics
    const columns = Math.min(4, metrics.length) || 1
    const rows = Math.ceil(metrics.length / columns)
    const gridX = PPTX_MARGIN
    const gridY = 1.55
    const gridW = PPTX_SLIDE_W - PPTX_MARGIN * 2
    const gridH = PPTX_SLIDE_H - gridY - 0.5
    const gap = 0.3
    const cardW = (gridW - gap * (columns - 1)) / columns
    const cardH = Math.min(1.9, (gridH - gap * (rows - 1)) / rows)

    metrics.forEach((metric, i) => {
      const col = i % columns
      const row = Math.floor(i / columns)
      const cardX = gridX + col * (cardW + gap)
      const cardY = gridY + row * (cardH + gap)

      slide.addShape('roundRect', {
        x: cardX,
        y: cardY,
        w: cardW,
        h: cardH,
        rectRadius: 0.06,
        fill: { color: 'FFFFFF' },
        line: { color: PPTX_CARD_BORDER, width: 1 },
        shadow: { type: 'outer', color: '1F3B5B', opacity: 0.15, blur: 6, offset: 2, angle: 90 },
      })
      slide.addShape('rect', { x: cardX, y: cardY, w: cardW, h: 0.07, fill: { color: PPTX_GOLD }, line: { type: 'none' } })
      slide.addText(formatMetricForSlide(metric.value, metric.format), {
        x: cardX + 0.08,
        y: cardY + 0.28,
        w: cardW - 0.16,
        h: cardH * 0.48,
        fontSize: cardW < 2.6 ? 19 : 23,
        bold: true,
        color: PPTX_NAVY,
        align: 'center',
        valign: 'middle',
        fontFace: 'Arial',
      })
      slide.addText(metric.label, {
        x: cardX + 0.1,
        y: cardY + cardH * 0.6,
        w: cardW - 0.2,
        h: cardH * 0.38,
        fontSize: 11,
        color: PPTX_LABEL_GREY,
        align: 'center',
        valign: 'top',
        fontFace: 'Arial',
      })
    })
  }

  // --- Charts: framed in a card so the image reads as a placed figure
  // rather than floating loose on the slide. ---
  const allCharts = [...(model?.charts ?? []), ...reportData.charts]
  for (const chart of allCharts) {
    const pngBuffer = chartPngCache.get(chart)
    if (!pngBuffer) continue
    const slide = pptx.addSlide({ masterName: 'CONTENT_MASTER' })
    addContentSlideHeader(slide, chart.title)

    const cardX = 0.9
    const cardY = 1.5
    const cardW = PPTX_SLIDE_W - cardX * 2
    const cardH = PPTX_SLIDE_H - cardY - 0.5
    slide.addShape('roundRect', {
      x: cardX,
      y: cardY,
      w: cardW,
      h: cardH,
      rectRadius: 0.08,
      fill: { color: 'FFFFFF' },
      line: { color: PPTX_CARD_BORDER, width: 1 },
    })
    // Chart PNGs are rendered at 900x420 (see buildChartSvg) — a 2.14:1 ratio.
    const imgH = cardH - 0.5
    const imgW = imgH * (900 / 420)
    slide.addImage({
      data: `image/png;base64,${pngBuffer.toString('base64')}`,
      x: cardX + (cardW - imgW) / 2,
      y: cardY + (cardH - imgH) / 2,
      w: imgW,
      h: imgH,
    })
  }

  // --- Next Steps: numbered gold badges beside each step instead of plain
  // "1. text" lines — reads as a roadmap rather than a bullet list. ---
  if (reportData.nextSteps?.length) {
    const slide = pptx.addSlide({ masterName: 'CONTENT_MASTER' })
    addContentSlideHeader(slide, 'Next Steps')

    const steps = reportData.nextSteps
    const areaY = 1.6
    const areaH = PPTX_SLIDE_H - areaY - 0.5
    const rowH = Math.min(1.15, areaH / steps.length)
    const badgeSize = 0.5

    steps.forEach((step, i) => {
      const rowY = areaY + i * rowH
      const badgeY = rowY + (rowH - badgeSize) / 2
      slide.addShape('ellipse', { x: PPTX_MARGIN, y: badgeY, w: badgeSize, h: badgeSize, fill: { color: PPTX_GOLD }, line: { type: 'none' } })
      slide.addText(String(i + 1), {
        x: PPTX_MARGIN,
        y: badgeY,
        w: badgeSize,
        h: badgeSize,
        fontSize: 16,
        bold: true,
        color: 'FFFFFF',
        align: 'center',
        valign: 'middle',
        fontFace: 'Arial',
      })
      slide.addText(step, {
        x: PPTX_MARGIN + badgeSize + 0.3,
        y: rowY,
        w: PPTX_SLIDE_W - PPTX_MARGIN * 2 - badgeSize - 0.3,
        h: rowH,
        fontSize: 14,
        color: PPTX_INK,
        valign: 'middle',
        lineSpacingMultiple: 1.2,
        fontFace: 'Arial',
      })
    })
  }

  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as unknown as Buffer
  return bufferToArrayBuffer(buffer)
}

// ── Job storage & orchestration ─────────────────────────────────────────────

export async function saveJob(job: ReportJob) {
  const store = getStore({ name: REPORTS_STORE, consistency: 'strong' })
  await store.setJSON(`${job.userId}/${job.id}.json`, job)
}

export async function getJob(userId: string, jobId: string) {
  const store = getStore({ name: REPORTS_STORE, consistency: 'strong' })
  return (await store.get(`${userId}/${jobId}.json`, { type: 'json' })) as ReportJob | null
}

export async function processReport(job: ReportJob) {
  const reportsStore = getStore({ name: REPORTS_STORE, consistency: 'strong' })
  const uploadsStore = getStore(UPLOADS_STORE)

  try {
    const upload = await uploadsStore.getWithMetadata(job.fileKey, { type: 'arrayBuffer' })
    if (!upload) throw new Error('The uploaded file could not be found.')

    const meta = upload.metadata as Record<string, string>
    const filename = meta.originalName || job.fileName
    const mimeType = getMimeType(filename, meta.contentType)
    const fileBuffer = Buffer.from(upload.data as ArrayBuffer)
    const reportDef = REPORT_DEFINITIONS[job.reportType]

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

    // A truncated generation (e.g. hitting max_tokens mid-way) can leave the
    // tool input missing fields entirely, even though the schema marks them
    // required — normalize so a partial response degrades gracefully instead
    // of crashing the whole job.
    const rawInput = toolUse.input as Partial<ReportData>
    const rawAssumptions = rawInput.assumptions as Partial<ReportAssumptions> | undefined
    const reportData: ReportData = {
      executiveSummary: typeof rawInput.executiveSummary === 'string' ? rawInput.executiveSummary : '',
      sections: Array.isArray(rawInput.sections) ? rawInput.sections : [],
      lineItems: Array.isArray(rawInput.lineItems)
        ? rawInput.lineItems.filter(
            (li): li is LineItem =>
              !!li && typeof li.label === 'string' && typeof li.group === 'string' && Number.isFinite(li.amount),
          )
        : [],
      assumptions: {
        dataCaveats: Array.isArray(rawAssumptions?.dataCaveats) ? rawAssumptions.dataCaveats : [],
        excludedItems: Array.isArray(rawAssumptions?.excludedItems) ? rawAssumptions.excludedItems : [],
        classificationRationale: Array.isArray(rawAssumptions?.classificationRationale) ? rawAssumptions.classificationRationale : [],
        limitations: Array.isArray(rawAssumptions?.limitations) ? rawAssumptions.limitations : [],
      },
      tables: Array.isArray(rawInput.tables) ? rawInput.tables : [],
      charts: Array.isArray(rawInput.charts) ? rawInput.charts : [],
      nextSteps: Array.isArray(rawInput.nextSteps) ? rawInput.nextSteps : [],
    }

    if (!reportData.executiveSummary && !reportData.sections.length && !reportData.tables.length && !reportData.lineItems.length) {
      throw new Error('The AI worker returned an empty or incomplete report.')
    }

    const downloadKey = `${job.userId}/${job.id}.xlsx`
    const model = calculateReportModel(job.reportType, reportData.lineItems)
    const allCharts = [...(model?.charts ?? []), ...reportData.charts]
    const chartPngCache = await renderChartPngCache(allCharts)
    const workbookBuffer = model
      ? await buildDeterministicExcelWorkbook(reportData, model, reportDef.label, filename, chartPngCache)
      : await buildExcelWorkbook(reportData, reportDef.label, filename, chartPngCache)
    await reportsStore.set(downloadKey, workbookBuffer)

    // The PowerPoint deck is a companion artifact — if it fails to generate,
    // the Excel report (the primary deliverable) still completes normally,
    // just without a PowerPoint download link.
    let downloadKeyPptx: string | undefined
    try {
      const pptxBuffer = await buildPowerPointPresentation(reportData, model, reportDef.label, filename, chartPngCache)
      downloadKeyPptx = `${job.userId}/${job.id}.pptx`
      await reportsStore.set(downloadKeyPptx, pptxBuffer)
    } catch (error) {
      console.error('PowerPoint generation failed, continuing with Excel-only report', error)
    }

    await saveJob({
      ...job,
      status: 'complete',
      fileName: filename,
      updatedAt: new Date().toISOString(),
      downloadKey,
      downloadKeyPptx,
    })
  } catch (error) {
    console.error('Report worker failed', error)
    await saveJob({
      ...job,
      status: 'failed',
      updatedAt: new Date().toISOString(),
      error: 'Report generation failed. Please try again.',
    })
  }
}

async function startBackgroundReportWorker(req: Request, job: ReportJob) {
  // report-worker-background.mts declares a custom config.path, which in
  // Netlify Functions v2 replaces its default /.netlify/functions/<name>
  // route rather than adding to it — so it must be invoked at that path.
  const workerUrl = new URL('/api/portal/report-worker-background', req.url)
  const response = await fetch(workerUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      jobId: job.id,
      userId: job.userId,
      workerToken: job.workerToken,
    }),
  })

  if (!response.ok) {
    const bodyText = await response.text().catch(() => '')
    console.error('report-worker-background invocation failed', {
      url: workerUrl.toString(),
      status: response.status,
      statusText: response.statusText,
      body: bodyText,
    })
    await saveJob({
      ...job,
      status: 'failed',
      updatedAt: new Date().toISOString(),
      error: 'Report generation could not start. Please try again.',
    })
    throw new Error('Report generation could not start. Please try again.')
  }
}

const ALLOWED_REPORT_EMAILS = new Set([
  'a.romosios@gmail.com',
  'a_romosios@yahoo.gr',
'anna.eftychidou@gmail.com',
  'themis.eftychidoy@gmail.com',
'vasilis.tsolis@cognitivplus.com',
'gianniszetas@gmail.com',
'angelos.p.eftychidis@gmail.com',
])

async function handlePost(req: Request, _context: Context, userId: string, userEmail: string | undefined) {
  if (!userEmail || !ALLOWED_REPORT_EMAILS.has(userEmail.toLowerCase())) {
    return json({ error: 'Not authorized user request. Please contact site administrators' }, 403)
  }

  let body: { fileKey?: string; reportType?: string }
  try {
    body = await req.json()
  } catch {
    return json({ error: 'Invalid JSON body.' }, 400)
  }

  const fileKey = body.fileKey
  const reportType = body.reportType
  if (!fileKey || !fileKey.startsWith(`${userId}/`)) {
    return json({ error: 'Choose one of your uploaded files first.' }, 400)
  }
  if (!reportType || !REPORT_DEFINITIONS[reportType]) {
    return json({ error: 'Choose a valid report type.' }, 400)
  }

  const uploadsStore = getStore(UPLOADS_STORE)
  const uploadMeta = await uploadsStore.getMetadata(fileKey)
  if (!uploadMeta) return json({ error: 'The selected file was not found.' }, 404)

  const meta = uploadMeta.metadata as Record<string, string>
  const now = new Date().toISOString()
  const job: ReportJob = {
    id: crypto.randomUUID(),
    userId,
    fileKey,
    fileName: meta.originalName || fileKey.slice(`${userId}/`.length),
    reportType,
    reportLabel: REPORT_DEFINITIONS[reportType].label,
    status: 'processing',
    createdAt: now,
    updatedAt: now,
    workerToken: crypto.randomUUID(),
  }

  await saveJob(job)
  try {
    await startBackgroundReportWorker(req, job)
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : 'Report generation could not start. Please try again.' }, 502)
  }

  return json({ jobId: job.id, status: job.status, reportLabel: job.reportLabel }, 202)
}

async function handleGet(req: Request, userId: string) {
  const url = new URL(req.url)
  const jobId = url.searchParams.get('jobId')
  if (!jobId) return json({ error: 'Missing jobId.' }, 400)

  const job = await getJob(userId, jobId)
  if (!job) return json({ error: 'Report job not found.' }, 404)

  if (url.searchParams.get('download') === '1') {
    const wantsPptx = url.searchParams.get('format') === 'pptx'
    const downloadKey = wantsPptx ? job.downloadKeyPptx : job.downloadKey
    if (job.status !== 'complete' || !downloadKey) {
      return json({ error: 'Report is not ready yet.' }, 409)
    }
    const reportsStore = getStore({ name: REPORTS_STORE, consistency: 'strong' })
    const report = await reportsStore.get(downloadKey, { type: 'arrayBuffer' })
    if (!report) return json({ error: 'Report file not found.' }, 404)

    const filename = getReportDownloadFilename(job.reportLabel, wantsPptx ? 'pptx' : 'xlsx')
    return new Response(report, {
      headers: {
        'Content-Type': wantsPptx
          ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`,
      },
    })
  }

  return json({
    jobId: job.id,
    status: job.status,
    reportLabel: job.reportLabel,
    fileName: job.fileName,
    downloadFileName: getReportDownloadFilename(job.reportLabel),
    downloadFileNamePptx: job.downloadKeyPptx ? getReportDownloadFilename(job.reportLabel, 'pptx') : null,
    error: job.error,
    downloadUrl: job.status === 'complete' && job.downloadKey ? `/api/portal/generate-report?jobId=${encodeURIComponent(job.id)}&download=1` : null,
    downloadUrlPptx:
      job.status === 'complete' && job.downloadKeyPptx
        ? `/api/portal/generate-report?jobId=${encodeURIComponent(job.id)}&download=1&format=pptx`
        : null,
  })
}

export default async (req: Request, context: Context) => {
  const user = await getUser()
  if (!user) return json({ error: 'You must be logged in.' }, 401)

  if (req.method === 'POST') return handlePost(req, context, user.id, user.email)
  if (req.method === 'GET') return handleGet(req, user.id)
  return json({ error: 'Method not allowed.' }, 405)
}

export const config: Config = {
  path: '/api/portal/generate-report',
  method: ['POST', 'GET'],
}
