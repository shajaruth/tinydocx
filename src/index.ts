export interface TextOptions {
  align?: 'left' | 'center' | 'right'
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strikethrough?: boolean
  code?: boolean
  color?: string
  font?: string
  size?: number
}

export interface TableOptions {
  colWidths?: number[]
}

export interface ImageOptions {
  width: number
  height: number
}

export interface TextRun {
  text: string
  opts?: TextOptions
  link?: string
}

export interface DocContext {
  heading(str: string, level: 1 | 2 | 3): void
  paragraph(str: string, opts?: TextOptions): void
  text(str: string, size: number, opts?: TextOptions): void
  lineBreak(): void
  horizontalRule(): void
  list(items: string[], ordered?: boolean): void
  table(rows: string[][], opts?: TableOptions): void
  link(text: string, url: string, opts?: TextOptions): void
  image(data: Uint8Array, opts: ImageOptions): void
  pageNumber(): void
}

export interface DOCXBuilder {
  content(fn: (ctx: DocContext) => void): DOCXBuilder
  header(fn: (ctx: DocContext) => void): DOCXBuilder
  footer(fn: (ctx: DocContext) => void): DOCXBuilder
  build(): Uint8Array
}

export interface ODTBuilder {
  content(fn: (ctx: DocContext) => void): ODTBuilder
  build(): Uint8Array
}

interface ListItem {
  runs: TextRun[]
  children?: { items: ListItem[]; ordered: boolean }
}

type DocElement =
  | { type: 'heading'; text: string; level: 1 | 2 | 3 }
  | { type: 'paragraph'; text: string; opts?: TextOptions }
  | { type: 'richParagraph'; runs: TextRun[]; align?: 'left' | 'center' | 'right' }
  | { type: 'text'; text: string; size: number; opts?: TextOptions }
  | { type: 'lineBreak' }
  | { type: 'horizontalRule' }
  | { type: 'list'; items: string[]; ordered: boolean }
  | { type: 'richList'; items: ListItem[]; ordered: boolean }
  | { type: 'table'; rows: string[][]; opts?: TableOptions }
  | { type: 'richTable'; rows: TextRun[][][]; opts?: TableOptions }
  | { type: 'link'; text: string; url: string; opts?: TextOptions; rId: string }
  | { type: 'image'; data: Uint8Array; opts: ImageOptions; rId: string; docPrId: number }
  | { type: 'blockquote'; runs: TextRun[] }
  | { type: 'codeBlock'; code: string; language?: string }
  | { type: 'pageNumber' }

type InlineToken =
  | { tag: 'text'; content: string }
  | { tag: 'bold'; children: InlineToken[] }
  | { tag: 'italic'; children: InlineToken[] }
  | { tag: 'code'; content: string }
  | { tag: 'strike'; children: InlineToken[] }
  | { tag: 'link'; children: InlineToken[]; href: string }
  | { tag: 'image'; alt: string; src: string }

type BlockToken =
  | { tag: 'h'; level: 1 | 2 | 3; content: InlineToken[] }
  | { tag: 'p'; content: InlineToken[] }
  | { tag: 'pre'; code: string; lang?: string }
  | { tag: 'quote'; blocks: BlockToken[] }
  | { tag: 'hr' }
  | { tag: 'ul'; items: ListNodeItem[] }
  | { tag: 'ol'; items: ListNodeItem[] }
  | { tag: 'table'; head: InlineToken[][]; body: InlineToken[][][] }

type ListNodeItem = {
  content: InlineToken[]
  nested?: { ordered: boolean; items: ListNodeItem[] }
}

type Result<T> = { ok: true; val: T; rest: string } | { ok: false }
type Parser<T> = (s: string) => Result<T>

const ok = <T>(val: T, rest: string): Result<T> => ({ ok: true, val, rest })
const fail = <T>(): Result<T> => ({ ok: false })

const lit = (match: string): Parser<string> => s =>
  s.startsWith(match) ? ok(match, s.slice(match.length)) : fail()

const re = (pattern: RegExp): Parser<string> => s => {
  const m = s.match(new RegExp('^' + pattern.source))
  return m ? ok(m[0], s.slice(m[0].length)) : fail()
}

const map = <A, B>(p: Parser<A>, f: (a: A) => B): Parser<B> => s => {
  const r = p(s)
  return r.ok ? ok(f(r.val), r.rest) : fail()
}

const seq = <A, B>(p1: Parser<A>, p2: Parser<B>): Parser<[A, B]> => s => {
  const r1 = p1(s)
  if (!r1.ok) return fail()
  const r2 = p2(r1.rest)
  return r2.ok ? ok([r1.val, r2.val], r2.rest) : fail()
}

const seq3 = <A, B, C>(p1: Parser<A>, p2: Parser<B>, p3: Parser<C>): Parser<[A, B, C]> => s => {
  const r1 = p1(s)
  if (!r1.ok) return fail()
  const r2 = p2(r1.rest)
  if (!r2.ok) return fail()
  const r3 = p3(r2.rest)
  return r3.ok ? ok([r1.val, r2.val, r3.val], r3.rest) : fail()
}

const alt = <T>(...ps: Parser<T>[]): Parser<T> => s => {
  for (const p of ps) {
    const r = p(s)
    if (r.ok) return r
  }
  return fail()
}

const many = <T>(p: Parser<T>): Parser<T[]> => s => {
  const results: T[] = []
  let rest = s
  while (true) {
    const r = p(rest)
    if (!r.ok) break
    results.push(r.val)
    rest = r.rest
  }
  return ok(results, rest)
}

const lazy = <T>(f: () => Parser<T>): Parser<T> => s => f()(s)

const between = <T>(open: string, close: string, p: Parser<T>): Parser<T> => s => {
  const r1 = lit(open)(s)
  if (!r1.ok) return fail()
  const r2 = p(r1.rest)
  if (!r2.ok) return fail()
  const r3 = lit(close)(r2.rest)
  return r3.ok ? ok(r2.val, r3.rest) : fail()
}

const untilStr = (delim: string): Parser<string> => s => {
  const idx = s.indexOf(delim)
  return idx > 0 ? ok(s.slice(0, idx), s.slice(idx)) : fail()
}

const takeWhile = (pred: (c: string) => boolean): Parser<string> => s => {
  let i = 0
  while (i < s.length && pred(s[i])) i++
  return i > 0 ? ok(s.slice(0, i), s.slice(i)) : fail()
}

const anyChar: Parser<string> = s => s.length > 0 ? ok(s[0], s.slice(1)) : fail()

const inlineSpecials = new Set(['*', '_', '`', '~', '[', '!', '\\'])

const escapeP: Parser<InlineToken> = map(
  seq(lit('\\'), anyChar),
  ([, c]) => ({ tag: 'text', content: c })
)

const codeP: Parser<InlineToken> = map(
  between('`', '`', takeWhile(c => c !== '`')),
  content => ({ tag: 'code', content })
)

const inlineTokens: Parser<InlineToken[]> = lazy(() => many(inlineElem))

const boldP: Parser<InlineToken> = alt(
  map(between('**', '**', lazy(() => untilInline('**'))), children => ({ tag: 'bold', children })),
  map(between('__', '__', lazy(() => untilInline('__'))), children => ({ tag: 'bold', children }))
)

const strikeP: Parser<InlineToken> = map(
  between('~~', '~~', lazy(() => untilInline('~~'))),
  children => ({ tag: 'strike', children })
)

const italicP: Parser<InlineToken> = s => {
  if (s[0] !== '*' && s[0] !== '_') return fail()
  if (s[1] === s[0]) return fail()
  const marker = s[0]
  const end = s.indexOf(marker, 1)
  if (end <= 1) return fail()
  const inner = s.slice(1, end)
  const children = parseInlineTokens(inner)
  return ok({ tag: 'italic', children }, s.slice(end + 1))
}

const imageP: Parser<InlineToken> = s => {
  if (!s.startsWith('![')) return fail()
  const altEnd = s.indexOf(']', 2)
  if (altEnd === -1 || s[altEnd + 1] !== '(') return fail()
  const srcEnd = s.indexOf(')', altEnd + 2)
  if (srcEnd === -1) return fail()
  return ok(
    { tag: 'image', alt: s.slice(2, altEnd), src: s.slice(altEnd + 2, srcEnd) },
    s.slice(srcEnd + 1)
  )
}

const linkP: Parser<InlineToken> = s => {
  if (s[0] !== '[') return fail()
  const textEnd = s.indexOf(']', 1)
  if (textEnd === -1 || s[textEnd + 1] !== '(') return fail()
  const hrefEnd = s.indexOf(')', textEnd + 2)
  if (hrefEnd === -1) return fail()
  const children = parseInlineTokens(s.slice(1, textEnd))
  return ok(
    { tag: 'link', children, href: s.slice(textEnd + 2, hrefEnd) },
    s.slice(hrefEnd + 1)
  )
}

const plainTextP: Parser<InlineToken> = map(
  takeWhile(c => !inlineSpecials.has(c)),
  content => ({ tag: 'text', content })
)

const singleCharP: Parser<InlineToken> = map(anyChar, content => ({ tag: 'text', content }))

const inlineElem: Parser<InlineToken> = alt(
  escapeP, codeP, boldP, strikeP, italicP, imageP, linkP, plainTextP, singleCharP
)

const untilInline = (delim: string): Parser<InlineToken[]> => s => {
  const tokens: InlineToken[] = []
  let rest = s
  while (rest.length > 0 && !rest.startsWith(delim)) {
    const r = inlineElem(rest)
    if (!r.ok) break
    tokens.push(r.val)
    rest = r.rest
  }
  return ok(tokens, rest)
}

const parseInlineTokens = (s: string): InlineToken[] => {
  const r = inlineTokens(s)
  return r.ok ? r.val : [{ tag: 'text', content: s }]
}

const mergeTextTokens = (tokens: InlineToken[]): InlineToken[] =>
  tokens.reduce<InlineToken[]>((acc, t) => {
    const last = acc[acc.length - 1]
    if (t.tag === 'text' && last?.tag === 'text') {
      acc[acc.length - 1] = { tag: 'text', content: last.content + t.content }
    } else {
      acc.push(t)
    }
    return acc
  }, [])

const tokensToRuns = (tokens: InlineToken[], inherited: TextOptions = {}): TextRun[] =>
  mergeTextTokens(tokens).flatMap((t): TextRun[] => {
    switch (t.tag) {
      case 'text': return [{ text: t.content, opts: Object.keys(inherited).length ? { ...inherited } : undefined }]
      case 'bold': return tokensToRuns(t.children, { ...inherited, bold: true })
      case 'italic': return tokensToRuns(t.children, { ...inherited, italic: true })
      case 'strike': return tokensToRuns(t.children, { ...inherited, strikethrough: true })
      case 'code': return [{ text: t.content, opts: { ...inherited, code: true } }]
      case 'link': return tokensToRuns(t.children, { ...inherited, color: '#0563C1', underline: true }).map(r => ({ ...r, link: t.href }))
      case 'image': return []
    }
  })

const tokensToPlainText = (tokens: InlineToken[]): string =>
  tokens.map(t => {
    switch (t.tag) {
      case 'text': case 'code': return t.content
      case 'bold': case 'italic': case 'strike': return tokensToPlainText(t.children)
      case 'link': return tokensToPlainText(t.children)
      case 'image': return t.alt
    }
  }).join('')

const parseBlockLines = (lines: string[]): BlockToken[] => {
  const blocks: BlockToken[] = []
  let i = 0

  while (i < lines.length) {
    const line = lines[i]

    if (line.trim() === '') { i++; continue }

    if (line.startsWith('```')) {
      const lang = line.slice(3).trim() || undefined
      const codeLines: string[] = []
      i++
      while (i < lines.length && !lines[i].startsWith('```')) {
        codeLines.push(lines[i])
        i++
      }
      blocks.push({ tag: 'pre', code: codeLines.join('\n'), lang })
      i++
      continue
    }

    if (/^#{1,3}\s/.test(line)) {
      const m = line.match(/^(#{1,3})\s+(.*)$/)!
      const level = Math.min(m[1].length, 3) as 1 | 2 | 3
      blocks.push({ tag: 'h', level, content: parseInlineTokens(m[2]) })
      i++
      continue
    }

    if (/^(-{3,}|\*{3,}|_{3,})$/.test(line.trim())) {
      blocks.push({ tag: 'hr' })
      i++
      continue
    }

    if (line.startsWith('> ')) {
      const quoteLines: string[] = []
      while (i < lines.length && lines[i].startsWith('> ')) {
        quoteLines.push(lines[i].slice(2))
        i++
      }
      blocks.push({ tag: 'quote', blocks: parseBlockLines(quoteLines) })
      continue
    }

    if (/^\|.+\|$/.test(line)) {
      const parseRow = (r: string) => r.slice(1, -1).split('|').map(c => parseInlineTokens(c.trim()))
      const head = parseRow(line)
      i++
      if (i < lines.length && /^\|[-:| ]+\|$/.test(lines[i])) i++
      const body: InlineToken[][][] = []
      while (i < lines.length && /^\|.+\|$/.test(lines[i])) {
        body.push(parseRow(lines[i]))
        i++
      }
      blocks.push({ tag: 'table', head, body })
      continue
    }

    if (/^(\s*)[-*]\s/.test(line) || /^(\s*)\d+\.\s/.test(line)) {
      const parseList = (startIdx: number, indent: number): { items: ListNodeItem[]; nextIdx: number } => {
        const items: ListNodeItem[] = []
        let idx = startIdx
        const firstLine = lines[idx]
        const ordered = /^\s*\d+\./.test(firstLine)
        const marker = ordered ? /^\s*\d+\.\s/ : /^\s*[-*]\s/

        while (idx < lines.length) {
          const l = lines[idx]
          const spaces = l.match(/^(\s*)/)?.[1].length ?? 0
          if (spaces < indent && idx !== startIdx) break
          if (spaces === indent && marker.test(l)) {
            const text = l.replace(/^\s*[-*]\s|^\s*\d+\.\s/, '')
            const content = parseInlineTokens(text)
            idx++
            let nested: ListNodeItem['nested'] = undefined
            if (idx < lines.length) {
              const nextSpaces = lines[idx].match(/^(\s*)/)?.[1].length ?? 0
              if (nextSpaces > indent && (/^\s*[-*]\s/.test(lines[idx]) || /^\s*\d+\.\s/.test(lines[idx]))) {
                const sub = parseList(idx, nextSpaces)
                nested = { ordered: /^\s*\d+\./.test(lines[idx]), items: sub.items }
                idx = sub.nextIdx
              }
            }
            items.push({ content, nested })
          } else if (spaces > indent) {
            idx++
          } else {
            break
          }
        }
        return { items, nextIdx: idx }
      }

      const indent = line.match(/^(\s*)/)?.[1].length ?? 0
      const result = parseList(i, indent)
      const ordered = /^\s*\d+\./.test(line)
      blocks.push(ordered ? { tag: 'ol', items: result.items } : { tag: 'ul', items: result.items })
      i = result.nextIdx
      continue
    }

    blocks.push({ tag: 'p', content: parseInlineTokens(line) })
    i++
  }

  return blocks
}

const listNodeToListItem = (node: ListNodeItem): ListItem => ({
  runs: tokensToRuns(node.content),
  children: node.nested ? { ordered: node.nested.ordered, items: node.nested.items.map(listNodeToListItem) } : undefined
})

const blockToElement = (b: BlockToken): DocElement[] => {
  switch (b.tag) {
    case 'h': return [{ type: 'heading', text: tokensToPlainText(b.content), level: b.level }]
    case 'p': return [{ type: 'richParagraph', runs: tokensToRuns(b.content) }]
    case 'pre': return [{ type: 'codeBlock', code: b.code, language: b.lang }]
    case 'hr': return [{ type: 'horizontalRule' }]
    case 'quote': return [{ type: 'blockquote', runs: b.blocks.flatMap(bb => bb.tag === 'p' ? tokensToRuns(bb.content) : []) }]
    case 'ul': return [{ type: 'richList', ordered: false, items: b.items.map(listNodeToListItem) }]
    case 'ol': return [{ type: 'richList', ordered: true, items: b.items.map(listNodeToListItem) }]
    case 'table': return [{ type: 'richTable', rows: [b.head.map(h => tokensToRuns(h)), ...b.body.map(r => r.map(c => tokensToRuns(c)))] }]
  }
}

const parseMarkdownAST = (md: string): BlockToken[] => parseBlockLines(md.replace(/\r\n/g, '\n').split('\n'))

const astToElements = (blocks: BlockToken[]): DocElement[] => blocks.flatMap(blockToElement)

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i]
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
    }
  }
  return (crc ^ 0xffffffff) >>> 0
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/**
 * Centralized XML namespace registry.
 * All namespace URIs are defined here to ensure consistency across the codebase
 * and prevent typos from causing invalid XML.
 */
const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
} as const

function createZip(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const entries: { name: Uint8Array; data: Uint8Array; crc: number; offset: number }[] = []
  const parts: Uint8Array[] = []
  let offset = 0

  for (const file of files) {
    const nameBytes = new TextEncoder().encode(file.name)
    const crc = crc32(file.data)
    const header = new Uint8Array(30 + nameBytes.length)
    const view = new DataView(header.buffer)
    view.setUint32(0, 0x04034b50, true)
    view.setUint16(4, 20, true)
    view.setUint16(6, 0, true)
    view.setUint16(8, 0, true)
    view.setUint16(10, 0, true)
    view.setUint16(12, 0, true)
    view.setUint32(14, crc, true)
    view.setUint32(18, file.data.length, true)
    view.setUint32(22, file.data.length, true)
    view.setUint16(26, nameBytes.length, true)
    view.setUint16(28, 0, true)
    header.set(nameBytes, 30)
    entries.push({ name: nameBytes, data: file.data, crc, offset })
    parts.push(header, file.data)
    offset += header.length + file.data.length
  }

  const centralDirOffset = offset
  for (const entry of entries) {
    const central = new Uint8Array(46 + entry.name.length)
    const view = new DataView(central.buffer)
    view.setUint32(0, 0x02014b50, true)
    view.setUint16(4, 20, true)
    view.setUint16(6, 20, true)
    view.setUint16(8, 0, true)
    view.setUint16(10, 0, true)
    view.setUint16(12, 0, true)
    view.setUint16(14, 0, true)
    view.setUint32(16, entry.crc, true)
    view.setUint32(20, entry.data.length, true)
    view.setUint32(24, entry.data.length, true)
    view.setUint16(28, entry.name.length, true)
    view.setUint16(30, 0, true)
    view.setUint16(32, 0, true)
    view.setUint16(34, 0, true)
    view.setUint16(36, 0, true)
    view.setUint32(38, 0, true)
    view.setUint32(42, entry.offset, true)
    central.set(entry.name, 46)
    parts.push(central)
    offset += central.length
  }

  const endRecord = new Uint8Array(22)
  const endView = new DataView(endRecord.buffer)
  endView.setUint32(0, 0x06054b50, true)
  endView.setUint16(4, 0, true)
  endView.setUint16(6, 0, true)
  endView.setUint16(8, entries.length, true)
  endView.setUint16(10, entries.length, true)
  endView.setUint32(12, offset - centralDirOffset, true)
  endView.setUint32(16, centralDirOffset, true)
  endView.setUint16(20, 0, true)
  parts.push(endRecord)

  const totalLength = parts.reduce((sum, p) => sum + p.length, 0)
  const result = new Uint8Array(totalLength)
  let pos = 0
  for (const part of parts) {
    result.set(part, pos)
    pos += part.length
  }
  return result
}

interface BuildContext {
  elements: DocElement[]
  hyperlinks: { url: string; rId: string }[]
  images: { data: Uint8Array; rId: string; ext: string }[]
  nextRId: number
  nextDocPrId: { value: number }
}

function createContext(ctx: BuildContext): DocContext {
  return {
    heading(str: string, level: 1 | 2 | 3) {
      ctx.elements.push({ type: 'heading', text: str, level })
    },
    paragraph(str: string, opts?: TextOptions) {
      ctx.elements.push({ type: 'paragraph', text: str, opts })
    },
    text(str: string, size: number, opts?: TextOptions) {
      ctx.elements.push({ type: 'text', text: str, size, opts })
    },
    lineBreak() {
      ctx.elements.push({ type: 'lineBreak' })
    },
    horizontalRule() {
      ctx.elements.push({ type: 'horizontalRule' })
    },
    list(items: string[], ordered = false) {
      ctx.elements.push({ type: 'list', items, ordered })
    },
    table(rows: string[][], opts?: TableOptions) {
      ctx.elements.push({ type: 'table', rows, opts })
    },
    link(text: string, url: string, opts?: TextOptions) {
      const rId = `rId${ctx.nextRId++}`
      ctx.hyperlinks.push({ url, rId })
      ctx.elements.push({ type: 'link', text, url, opts, rId })
    },
    image(data: Uint8Array, opts: ImageOptions) {
      const rId = `rId${ctx.nextRId++}`
      const docPrId = ctx.nextDocPrId.value++
      const ext = detectImageType(data)
      ctx.images.push({ data, rId, ext })
      ctx.elements.push({ type: 'image', data, opts, rId, docPrId })
    },
    pageNumber() {
      ctx.elements.push({ type: 'pageNumber' })
    }
  }
}

function detectImageType(data: Uint8Array): string {
  if (data[0] === 0x89 && data[1] === 0x50) return 'png'
  if (data[0] === 0xff && data[1] === 0xd8) return 'jpeg'
  if (data[0] === 0x47 && data[1] === 0x49) return 'gif'
  if (data[0] === 0x52 && data[1] === 0x49 && data[2] === 0x46 && data[3] === 0x46 &&
      data[8] === 0x57 && data[9] === 0x45 && data[10] === 0x42 && data[11] === 0x50) return 'webp'
  return 'png'
}

function buildDocxRunProps(opts?: TextOptions, size?: number): string {
  let rPr = ''
  const font = opts?.code ? 'Courier New' : opts?.font
  if (font) rPr += `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`
  const sz = size || opts?.size
  if (sz) rPr += `<w:sz w:val="${sz * 2}"/><w:szCs w:val="${sz * 2}"/>`
  if (opts?.bold) rPr += '<w:b/>'
  if (opts?.italic) rPr += '<w:i/>'
  if (opts?.underline) rPr += '<w:u w:val="single"/>'
  if (opts?.strikethrough) rPr += '<w:strike/>'
  if (opts?.code) rPr += '<w:shd w:val="clear" w:color="auto" w:fill="E8E8E8"/>'
  if (opts?.color) rPr += `<w:color w:val="${opts.color.replace('#', '')}"/>`
  return rPr ? `<w:rPr>${rPr}</w:rPr>` : ''
}

function getAlignment(align?: string): string {
  return align === 'center' ? 'center' : align === 'right' ? 'right' : align === 'justify' ? 'both' : 'left'
}

const parseInline = (text: string): TextRun[] => tokensToRuns(parseInlineTokens(text))

function buildDocxBody(elements: DocElement[]): string {
  const SIZES = { 1: 48, 2: 36, 3: 28 }
  let body = ''

  for (const el of elements) {
    if (el.type === 'heading') {
      const sz = SIZES[el.level]
      body += `<w:p><w:pPr><w:pStyle w:val="Heading${el.level}"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t>${escapeXml(el.text)}</w:t></w:r></w:p>`
    } else if (el.type === 'paragraph') {
      const jc = getAlignment(el.opts?.align)
      const rPr = buildDocxRunProps(el.opts)
      body += `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr><w:r>${rPr}<w:t>${escapeXml(el.text)}</w:t></w:r></w:p>`
    } else if (el.type === 'text') {
      const jc = getAlignment(el.opts?.align)
      const rPr = buildDocxRunProps(el.opts, el.size)
      body += `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr><w:r>${rPr}<w:t>${escapeXml(el.text)}</w:t></w:r></w:p>`
    } else if (el.type === 'lineBreak') {
      body += '<w:p/>'
    } else if (el.type === 'horizontalRule') {
      body += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'
    } else if (el.type === 'list') {
      const numId = el.ordered ? 2 : 1
      for (const item of el.items) {
        body += `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t>${escapeXml(item)}</w:t></w:r></w:p>`
      }
    } else if (el.type === 'table') {
      body += '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
      body += '<w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>'
      body += '<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>'
      body += '<w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/>'
      body += '</w:tblBorders></w:tblPr>'
      if (el.opts?.colWidths) {
        body += '<w:tblGrid>'
        for (const w of el.opts.colWidths) body += `<w:gridCol w:w="${w}"/>`
        body += '</w:tblGrid>'
      }
      for (const row of el.rows) {
        body += '<w:tr>'
        for (let i = 0; i < row.length; i++) {
          body += '<w:tc>'
          if (el.opts?.colWidths?.[i]) {
            body += `<w:tcPr><w:tcW w:w="${el.opts.colWidths[i]}" w:type="dxa"/></w:tcPr>`
          }
          body += `<w:p><w:r><w:t>${escapeXml(row[i])}</w:t></w:r></w:p></w:tc>`
        }
        body += '</w:tr>'
      }
      body += '</w:tbl>'
    } else if (el.type === 'link') {
      const rPr = buildDocxRunProps({ ...el.opts, color: el.opts?.color || '0563C1', underline: true })
      body += `<w:p><w:hyperlink r:id="${el.rId}"><w:r>${rPr}<w:t>${escapeXml(el.text)}</w:t></w:r></w:hyperlink></w:p>`
    } else if (el.type === 'image') {
      const cx = Math.round(el.opts.width * 914400)
      const cy = Math.round(el.opts.height * 914400)
      body += `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${NS.wp}"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${el.docPrId}" name="Image${el.docPrId}"/><a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${NS.pic}"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${el.docPrId}" name="Image${el.docPrId}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${el.rId}" xmlns:r="${NS.r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    } else if (el.type === 'pageNumber') {
      body += '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    } else if (el.type === 'richParagraph') {
      const jc = getAlignment(el.align)
      body += `<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr>`
      for (const run of el.runs) {
        const rPr = buildDocxRunProps(run.opts)
        body += `<w:r>${rPr}<w:t>${escapeXml(run.text)}</w:t></w:r>`
      }
      body += '</w:p>'
    } else if (el.type === 'richList') {
      const buildListItems = (items: ListItem[], ordered: boolean, level: number) => {
        const numId = ordered ? 2 : 1
        for (const item of items) {
          body += `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`
          for (const run of item.runs) {
            const rPr = buildDocxRunProps(run.opts)
            body += `<w:r>${rPr}<w:t>${escapeXml(run.text)}</w:t></w:r>`
          }
          body += '</w:p>'
          if (item.children) {
            buildListItems(item.children.items, item.children.ordered, level + 1)
          }
        }
      }
      buildListItems(el.items, el.ordered, 0)
    } else if (el.type === 'richTable') {
      body += '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
      body += '<w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>'
      body += '<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>'
      body += '<w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/>'
      body += '</w:tblBorders></w:tblPr>'
      if (el.opts?.colWidths) {
        body += '<w:tblGrid>'
        for (const w of el.opts.colWidths) body += `<w:gridCol w:w="${w}"/>`
        body += '</w:tblGrid>'
      }
      for (const row of el.rows) {
        body += '<w:tr>'
        for (let i = 0; i < row.length; i++) {
          body += '<w:tc>'
          if (el.opts?.colWidths?.[i]) {
            body += `<w:tcPr><w:tcW w:w="${el.opts.colWidths[i]}" w:type="dxa"/></w:tcPr>`
          }
          body += '<w:p>'
          for (const run of row[i]) {
            const rPr = buildDocxRunProps(run.opts)
            body += `<w:r>${rPr}<w:t>${escapeXml(run.text)}</w:t></w:r>`
          }
          body += '</w:p></w:tc>'
        }
        body += '</w:tr>'
      }
      body += '</w:tbl>'
    } else if (el.type === 'blockquote') {
      body += '<w:p><w:pPr><w:ind w:left="720"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="4" w:color="CCCCCC"/></w:pBdr></w:pPr>'
      for (const run of el.runs) {
        const rPr = buildDocxRunProps({ ...run.opts, color: run.opts?.color || '#666666' })
        body += `<w:r>${rPr}<w:t>${escapeXml(run.text)}</w:t></w:r>`
      }
      body += '</w:p>'
    } else if (el.type === 'codeBlock') {
      const lines = el.code.split('\n')
      for (const line of lines) {
        body += '<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr>'
        body += `<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`
      }
    }
  }
  return body
}

function generateDocxDocument(elements: DocElement[], headerRId?: string, footerRId?: string): string {
  const body = buildDocxBody(elements)

  let sectPr = '<w:sectPr>'
  if (headerRId) sectPr += `<w:headerReference w:type="default" r:id="${headerRId}"/>`
  if (footerRId) sectPr += `<w:footerReference w:type="default" r:id="${footerRId}"/>`
  sectPr += '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}">
<w:body>
${body}
${sectPr}
</w:body>
</w:document>`
}

function generateDocxHeader(elements: DocElement[], hasLinksOrImages: boolean): string {
  const body = buildDocxBody(elements)
  const rAttr = hasLinksOrImages ? ` xmlns:r="${NS.r}"` : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${NS.w}"${rAttr}>
${body}
</w:hdr>`
}

function generateDocxFooter(elements: DocElement[], hasLinksOrImages: boolean): string {
  const body = buildDocxBody(elements)
  const rAttr = hasLinksOrImages ? ` xmlns:r="${NS.r}"` : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${NS.w}"${rAttr}>
${body}
</w:ftr>`
}

function generateDocxStyles(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="48"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="Heading 3"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>`
}

function generateDocxNumbering(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS.w}">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`
}

function generateDocxContentTypes(hasLists: boolean, hasHeader: boolean, hasFooter: boolean, imageExts: string[]): string {
  let types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS.ct}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>`

  for (const ext of Array.from(new Set(imageExts))) {
    const mime = ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/png'
    types += `<Default Extension="${ext}" ContentType="${mime}"/>`
  }

  types += `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`

  if (hasLists) {
    types += `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
  }
  if (hasHeader) {
    types += `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`
  }
  if (hasFooter) {
    types += `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`
  }

  types += `</Types>`
  return types
}

function generateDocxRels(
  hasLists: boolean,
  hasHeader: boolean,
  hasFooter: boolean,
  hyperlinks: { url: string; rId: string }[],
  images: { rId: string; ext: string }[],
  imageOffset: number
): string {
  let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS.rel}">
<Relationship Id="rId1" Type="${NS.r}/styles" Target="styles.xml"/>`

  let nextId = 2
  if (hasLists) {
    rels += `<Relationship Id="rId${nextId++}" Type="${NS.r}/numbering" Target="numbering.xml"/>`
  }
  if (hasHeader) {
    rels += `<Relationship Id="rIdHeader" Type="${NS.r}/header" Target="header1.xml"/>`
  }
  if (hasFooter) {
    rels += `<Relationship Id="rIdFooter" Type="${NS.r}/footer" Target="footer1.xml"/>`
  }

  for (const link of hyperlinks) {
    rels += `<Relationship Id="${link.rId}" Type="${NS.r}/hyperlink" Target="${escapeXml(link.url)}" TargetMode="External"/>`
  }

  for (let i = 0; i < images.length; i++) {
    rels += `<Relationship Id="${images[i].rId}" Type="${NS.r}/image" Target="media/image${imageOffset + i + 1}.${images[i].ext}"/>`
  }

  rels += `</Relationships>`
  return rels
}

function generatePartRels(
  hyperlinks: { url: string; rId: string }[],
  images: { rId: string; ext: string }[],
  imageOffset: number
): string {
  let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS.rel}">`

  for (const link of hyperlinks) {
    rels += `<Relationship Id="${link.rId}" Type="${NS.r}/hyperlink" Target="${escapeXml(link.url)}" TargetMode="External"/>`
  }

  for (let i = 0; i < images.length; i++) {
    rels += `<Relationship Id="${images[i].rId}" Type="${NS.r}/image" Target="media/image${imageOffset + i + 1}.${images[i].ext}"/>`
  }

  rels += `</Relationships>`
  return rels
}

const DOCX_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS.rel}">
<Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/>
</Relationships>`

/**
 * Create a new DOCX document
 */
export function docx(): DOCXBuilder {
  const docPrIdCounter = { value: 1 }
  const mainCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 10, nextDocPrId: docPrIdCounter }
  const headerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 100, nextDocPrId: docPrIdCounter }
  const footerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 200, nextDocPrId: docPrIdCounter }
  let hasHeader = false
  let hasFooter = false

  const builder: DOCXBuilder = {
    content(fn) {
      fn(createContext(mainCtx))
      return builder
    },
    header(fn) {
      hasHeader = true
      fn(createContext(headerCtx))
      return builder
    },
    footer(fn) {
      hasFooter = true
      fn(createContext(footerCtx))
      return builder
    },
    build() {
      const enc = new TextEncoder()
      const hasLists = mainCtx.elements.some(el => el.type === 'list')
      const allImages = [...mainCtx.images, ...headerCtx.images, ...footerCtx.images]
      const imageExts = allImages.map(img => img.ext)

      const mainImageOffset = 0
      const headerImageOffset = mainCtx.images.length
      const footerImageOffset = mainCtx.images.length + headerCtx.images.length

      const files: { name: string; data: Uint8Array }[] = [
        { name: '[Content_Types].xml', data: enc.encode(generateDocxContentTypes(hasLists, hasHeader, hasFooter, imageExts)) },
        { name: '_rels/.rels', data: enc.encode(DOCX_RELS) },
        { name: 'word/_rels/document.xml.rels', data: enc.encode(generateDocxRels(hasLists, hasHeader, hasFooter, mainCtx.hyperlinks, mainCtx.images, mainImageOffset)) },
        { name: 'word/document.xml', data: enc.encode(generateDocxDocument(mainCtx.elements, hasHeader ? 'rIdHeader' : undefined, hasFooter ? 'rIdFooter' : undefined)) },
        { name: 'word/styles.xml', data: enc.encode(generateDocxStyles()) }
      ]

      if (hasLists) {
        files.push({ name: 'word/numbering.xml', data: enc.encode(generateDocxNumbering()) })
      }

      if (hasHeader) {
        const hasLinksOrImages = headerCtx.hyperlinks.length > 0 || headerCtx.images.length > 0
        files.push({ name: 'word/header1.xml', data: enc.encode(generateDocxHeader(headerCtx.elements, hasLinksOrImages)) })
        if (hasLinksOrImages) {
          files.push({ name: 'word/_rels/header1.xml.rels', data: enc.encode(generatePartRels(headerCtx.hyperlinks, headerCtx.images, headerImageOffset)) })
        }
      }

      if (hasFooter) {
        const hasLinksOrImages = footerCtx.hyperlinks.length > 0 || footerCtx.images.length > 0
        files.push({ name: 'word/footer1.xml', data: enc.encode(generateDocxFooter(footerCtx.elements, hasLinksOrImages)) })
        if (hasLinksOrImages) {
          files.push({ name: 'word/_rels/footer1.xml.rels', data: enc.encode(generatePartRels(footerCtx.hyperlinks, footerCtx.images, footerImageOffset)) })
        }
      }

      for (let i = 0; i < allImages.length; i++) {
        files.push({ name: `word/media/image${i + 1}.${allImages[i].ext}`, data: allImages[i].data })
      }

      return createZip(files)
    }
  }
  return builder
}

const ODT_MIMETYPE = 'application/vnd.oasis.opendocument.text'

const ODT_MANIFEST = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="${NS.manifest}" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`

const ODT_STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="${NS.office}" xmlns:style="${NS.style}" xmlns:fo="${NS.fo}" office:version="1.2">
<office:styles>
<style:style style:name="Standard" style:family="paragraph"/>
<style:style style:name="Heading1" style:family="paragraph"><style:text-properties fo:font-size="24pt" fo:font-weight="bold"/></style:style>
<style:style style:name="Heading2" style:family="paragraph"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>
<style:style style:name="Heading3" style:family="paragraph"><style:text-properties fo:font-size="14pt" fo:font-weight="bold"/></style:style>
</office:styles>
</office:document-styles>`

function buildOdtStyle(name: string, opts?: TextOptions, size?: number): string {
  let textProps = ''
  const font = opts?.code ? 'Courier New' : opts?.font
  if (font) textProps += ` style:font-name="${font}"`
  const sz = size || opts?.size
  if (sz) textProps += ` fo:font-size="${sz}pt"`
  if (opts?.bold) textProps += ' fo:font-weight="bold"'
  if (opts?.italic) textProps += ' fo:font-style="italic"'
  if (opts?.underline) textProps += ' style:text-underline-style="solid"'
  if (opts?.strikethrough) textProps += ' style:text-line-through-style="solid"'
  if (opts?.code) textProps += ' fo:background-color="#E8E8E8"'
  if (opts?.color) textProps += ` fo:color="${opts.color}"`
  const pProps = opts?.align ? `<style:paragraph-properties fo:text-align="${opts.align}"/>` : ''
  return `<style:style style:name="${name}" style:family="paragraph">${pProps}<style:text-properties${textProps}/></style:style>`
}

function buildOdtTextStyle(name: string, opts?: TextOptions): string {
  let textProps = ''
  const font = opts?.code ? 'Courier New' : opts?.font
  if (font) textProps += ` style:font-name="${font}"`
  if (opts?.size) textProps += ` fo:font-size="${opts.size}pt"`
  if (opts?.bold) textProps += ' fo:font-weight="bold"'
  if (opts?.italic) textProps += ' fo:font-style="italic"'
  if (opts?.underline) textProps += ' style:text-underline-style="solid"'
  if (opts?.strikethrough) textProps += ' style:text-line-through-style="solid"'
  if (opts?.code) textProps += ' fo:background-color="#E8E8E8"'
  if (opts?.color) textProps += ` fo:color="${opts.color}"`
  return `<style:style style:name="${name}" style:family="text"><style:text-properties${textProps}/></style:style>`
}

function generateOdtContent(elements: DocElement[]): string {
  let body = ''
  let styleCount = 0
  const styles: string[] = []

  for (const el of elements) {
    if (el.type === 'heading') {
      body += `<text:h text:style-name="Heading${el.level}" text:outline-level="${el.level}">${escapeXml(el.text)}</text:h>`
    } else if (el.type === 'paragraph') {
      let styleName = 'Standard'
      if (el.opts?.bold || el.opts?.italic || el.opts?.color || el.opts?.align || el.opts?.font || el.opts?.underline) {
        styleName = `P${++styleCount}`
        styles.push(buildOdtStyle(styleName, el.opts))
      }
      body += `<text:p text:style-name="${styleName}">${escapeXml(el.text)}</text:p>`
    } else if (el.type === 'text') {
      const styleName = `P${++styleCount}`
      styles.push(buildOdtStyle(styleName, el.opts, el.size))
      body += `<text:p text:style-name="${styleName}">${escapeXml(el.text)}</text:p>`
    } else if (el.type === 'lineBreak') {
      body += '<text:p text:style-name="Standard"/>'
    } else if (el.type === 'horizontalRule') {
      body += '<text:p text:style-name="Standard">────────────────────────────────────────</text:p>'
    } else if (el.type === 'list') {
      body += `<text:list text:style-name="${el.ordered ? 'Numbering_20_1' : 'List_20_1'}">`
      for (const item of el.items) {
        body += `<text:list-item><text:p text:style-name="Standard">${escapeXml(item)}</text:p></text:list-item>`
      }
      body += '</text:list>'
    } else if (el.type === 'table') {
      body += `<table:table xmlns:table="${NS.table}">`
      for (const row of el.rows) {
        body += '<table:table-row>'
        for (const cell of row) {
          body += `<table:table-cell><text:p>${escapeXml(cell)}</text:p></table:table-cell>`
        }
        body += '</table:table-row>'
      }
      body += '</table:table>'
    } else if (el.type === 'link') {
      body += `<text:p><text:a xlink:href="${escapeXml(el.url)}" xmlns:xlink="${NS.xlink}">${escapeXml(el.text)}</text:a></text:p>`
    } else if (el.type === 'richParagraph') {
      body += `<text:p text:style-name="Standard">`
      for (const run of el.runs) {
        if (run.opts?.bold || run.opts?.italic || run.opts?.code || run.opts?.underline || run.opts?.strikethrough || run.opts?.color) {
          const styleName = `T${++styleCount}`
          styles.push(buildOdtTextStyle(styleName, run.opts))
          body += `<text:span text:style-name="${styleName}">${escapeXml(run.text)}</text:span>`
        } else {
          body += escapeXml(run.text)
        }
      }
      body += '</text:p>'
    } else if (el.type === 'richList') {
      const buildOdtList = (items: ListItem[], ordered: boolean) => {
        body += `<text:list text:style-name="${ordered ? 'Numbering_20_1' : 'List_20_1'}">`
        for (const item of items) {
          body += '<text:list-item><text:p text:style-name="Standard">'
          for (const run of item.runs) {
            if (run.opts?.bold || run.opts?.italic || run.opts?.code || run.opts?.underline || run.opts?.strikethrough || run.opts?.color) {
              const styleName = `T${++styleCount}`
              styles.push(buildOdtTextStyle(styleName, run.opts))
              body += `<text:span text:style-name="${styleName}">${escapeXml(run.text)}</text:span>`
            } else {
              body += escapeXml(run.text)
            }
          }
          body += '</text:p>'
          if (item.children) buildOdtList(item.children.items, item.children.ordered)
          body += '</text:list-item>'
        }
        body += '</text:list>'
      }
      buildOdtList(el.items, el.ordered)
    } else if (el.type === 'richTable') {
      body += `<table:table xmlns:table="${NS.table}">`
      for (const row of el.rows) {
        body += '<table:table-row>'
        for (const cell of row) {
          body += '<table:table-cell><text:p>'
          for (const run of cell) {
            if (run.opts?.bold || run.opts?.italic || run.opts?.code || run.opts?.underline || run.opts?.strikethrough || run.opts?.color) {
              const styleName = `T${++styleCount}`
              styles.push(buildOdtTextStyle(styleName, run.opts))
              body += `<text:span text:style-name="${styleName}">${escapeXml(run.text)}</text:span>`
            } else {
              body += escapeXml(run.text)
            }
          }
          body += '</text:p></table:table-cell>'
        }
        body += '</table:table-row>'
      }
      body += '</table:table>'
    } else if (el.type === 'blockquote') {
      body += '<text:p text:style-name="Standard">'
      for (const run of el.runs) {
        body += escapeXml(run.text)
      }
      body += '</text:p>'
    } else if (el.type === 'codeBlock') {
      const codeStyleName = `P${++styleCount}`
      styles.push(`<style:style style:name="${codeStyleName}" style:family="paragraph"><style:text-properties style:font-name="Courier New" fo:background-color="#F5F5F5"/></style:style>`)
      for (const line of el.code.split('\n')) {
        body += `<text:p text:style-name="${codeStyleName}">${escapeXml(line)}</text:p>`
      }
    }
  }

  const autoStyles = styles.length > 0 ? `<office:automatic-styles>${styles.join('')}</office:automatic-styles>` : ''
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="${NS.office}" xmlns:text="${NS.text}" xmlns:style="${NS.style}" xmlns:fo="${NS.fo}" office:version="1.2">
${autoStyles}
<office:body>
<office:text>
${body}
</office:text>
</office:body>
</office:document-content>`
}

/**
 * Create a new ODT document
 */
export function odt(): ODTBuilder {
  const ctx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 1, nextDocPrId: { value: 1 } }
  const builder: ODTBuilder = {
    content(fn) {
      fn(createContext(ctx))
      return builder
    },
    build() {
      const enc = new TextEncoder()
      return createZip([
        { name: 'mimetype', data: enc.encode(ODT_MIMETYPE) },
        { name: 'META-INF/manifest.xml', data: enc.encode(ODT_MANIFEST) },
        { name: 'content.xml', data: enc.encode(generateOdtContent(ctx.elements)) },
        { name: 'styles.xml', data: enc.encode(ODT_STYLES) }
      ])
    }
  }
  return builder
}

const parseMarkdownToElements = (md: string): DocElement[] => astToElements(parseMarkdownAST(md))

export function markdownToDocx(md: string): Uint8Array {
  const docPrIdCounter = { value: 1 }
  const mainCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 10, nextDocPrId: docPrIdCounter }
  mainCtx.elements.push(...parseMarkdownToElements(md))
  const enc = new TextEncoder()
  const hasLists = mainCtx.elements.some(el => el.type === 'list' || el.type === 'richList')
  const files: { name: string; data: Uint8Array }[] = [
    { name: '[Content_Types].xml', data: enc.encode(generateDocxContentTypes(hasLists, false, false, [])) },
    { name: '_rels/.rels', data: enc.encode(DOCX_RELS) },
    { name: 'word/_rels/document.xml.rels', data: enc.encode(generateDocxRels(hasLists, false, false, [], [], 0)) },
    { name: 'word/document.xml', data: enc.encode(generateDocxDocument(mainCtx.elements, undefined, undefined)) },
    { name: 'word/styles.xml', data: enc.encode(generateDocxStyles()) }
  ]
  if (hasLists) files.push({ name: 'word/numbering.xml', data: enc.encode(generateDocxNumbering()) })
  return createZip(files)
}

export function markdownToOdt(md: string): Uint8Array {
  const elements = parseMarkdownToElements(md)
  const enc = new TextEncoder()
  return createZip([
    { name: 'mimetype', data: enc.encode(ODT_MIMETYPE) },
    { name: 'META-INF/manifest.xml', data: enc.encode(ODT_MANIFEST) },
    { name: 'content.xml', data: enc.encode(generateOdtContent(elements)) },
    { name: 'styles.xml', data: enc.encode(ODT_STYLES) }
  ])
}

export default docx
