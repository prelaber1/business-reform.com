import Anthropic from '@anthropic-ai/sdk'

export const REPORT_TOOL: Anthropic.Tool = {
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
