// AWS Lambda (which Netlify Functions run on) ships with zero fonts
// installed, so sharp/libvips' SVG->PNG text rendering silently produces
// missing-glyph "tofu" boxes for every character in generated chart images —
// not just non-Latin text, since there is no font available at all, not
// even for plain digits.
//
// This writes a bundled font + a fontconfig config pointing at it into /tmp
// (the only writable directory in the Lambda runtime) as an immediate
// top-level side effect, and points fontconfig at it via env vars. Runs
// once per cold start. Imported before `sharp` itself in
// generate-report.mts specifically so these env vars are in place before
// libvips/fontconfig can initialize with no fonts found.
import { writeFileSync, mkdirSync, existsSync } from 'node:fs'
import { DEJAVU_SANS_REGULAR_BASE64, DEJAVU_SANS_BOLD_BASE64 } from './dejavu-sans-base64.mts'

const FONT_DIR = '/tmp/business-reform-fonts'
const CACHE_DIR = '/tmp/business-reform-fontconfig-cache'
const FONTS_CONF_PATH = `${FONT_DIR}/fonts.conf`

mkdirSync(FONT_DIR, { recursive: true })
mkdirSync(CACHE_DIR, { recursive: true })

if (!existsSync(`${FONT_DIR}/DejaVuSans.ttf`)) {
  writeFileSync(`${FONT_DIR}/DejaVuSans.ttf`, Buffer.from(DEJAVU_SANS_REGULAR_BASE64, 'base64'))
}
if (!existsSync(`${FONT_DIR}/DejaVuSans-Bold.ttf`)) {
  writeFileSync(`${FONT_DIR}/DejaVuSans-Bold.ttf`, Buffer.from(DEJAVU_SANS_BOLD_BASE64, 'base64'))
}
if (!existsSync(FONTS_CONF_PATH)) {
  writeFileSync(
    FONTS_CONF_PATH,
    `<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "fonts.dtd">\n<fontconfig>\n  <dir>${FONT_DIR}</dir>\n  <cachedir>${CACHE_DIR}</cachedir>\n</fontconfig>\n`,
    'utf8',
  )
}

process.env.FONTCONFIG_FILE = FONTS_CONF_PATH
process.env.FONTCONFIG_PATH = FONT_DIR
