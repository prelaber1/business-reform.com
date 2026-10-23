import type PptxGenJSType from 'pptxgenjs'
import PptxGenJS from '../vendor/pptxgenjs.bundle.mjs'
import { LOGO_PNG_BASE64 } from '../logo-data.mts'
import { AI_MARK, AI_METADATA_DESCRIPTION, AI_METADATA_KEYWORDS, aiMarkLine } from '../ai-disclosure.mts'
import { type ReportChart, type ReportData } from './types.mts'
import { buildChartSvg } from './chartRendering.mts'
import { type MetricFormat, type ReportModelBase } from './calculationSupport.mts'
import { bufferToArrayBuffer } from './excelSheetHelpers.mts'
import { LOGO_ASPECT } from './ExcelReportBuilder.mts'

export const PPTX_NAVY = '1F3B5B'
export const PPTX_NAVY_DEEP = '15293D'
export const PPTX_INK = '2B2B2B'
export const PPTX_MUTED = 'D9E2EC'
export const PPTX_GOLD = 'C9A227'
export const PPTX_CARD_BG = 'F4F7FB'
export const PPTX_CARD_BORDER = 'E3E9F0'
export const PPTX_LABEL_GREY = '6B7785'
export const PPTX_SLIDE_W = 13.33
export const PPTX_SLIDE_H = 7.5
export const PPTX_MARGIN = 0.55

export function formatMetricForSlide(value: number, format?: MetricFormat): string {
  if (format === 'percent') return `${(value * 100).toFixed(2)}%`
  if (format === 'ratio') return `${value.toFixed(2)}x`
  if (format === 'number') return value.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })
  return `${value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`
}

// A consistent masthead for every content slide: a short gold tick mark, the
// slide title, and a thin rule underneath — gives the deck a designed,
// templated feel instead of a bare heading floating on white.
export function addContentSlideHeader(slide: PptxGenJSType.Slide, title: string) {
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
