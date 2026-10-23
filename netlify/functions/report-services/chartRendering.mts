// Must run before `sharp` is imported — see the comment inside for why.
import '../fonts/setup-fonts.mts'
import sharp from 'sharp'
import { CHART_PALETTE } from './constants.mts'
import { type ReportChart } from './types.mts'

export function escapeXml(value: string) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

export function formatAxisNumber(n: number) {
  const abs = Math.abs(n)
  if (abs >= 1_000_000) return `${+(n / 1_000_000).toFixed(1)}M`
  if (abs >= 1_000) return `${+(n / 1_000).toFixed(1)}K`
  return `${Math.round(n * 100) / 100}`
}

export function buildChartSvg(chart: ReportChart) {
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

export async function renderChartPng(chart: ReportChart) {
  const svg = buildChartSvg(chart)
  return sharp(Buffer.from(svg)).png().toBuffer()
}

// Renders every chart's PNG exactly once, shared across the Excel workbook
// (Summary sheet + dedicated Charts sheet) and the PowerPoint deck, instead
// of re-rendering the same SVG->PNG conversion for each consumer.
export async function renderChartPngCache(charts: ReportChart[]): Promise<Map<ReportChart, Buffer>> {
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
