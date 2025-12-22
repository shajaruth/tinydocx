import { describe, expect, test } from 'bun:test'
import { docx, odt, markdownToDocx, markdownToOdt } from './index'

describe('docx', () => {
  test('creates valid ZIP archive', () => {
    const doc = docx()
    doc.content(() => {})
    const bytes = doc.build()
    expect(bytes[0]).toBe(0x50)
    expect(bytes[1]).toBe(0x4b)
  })

  test('returns Uint8Array', () => {
    const doc = docx()
    doc.content(() => {})
    expect(doc.build()).toBeInstanceOf(Uint8Array)
  })

  test('includes required DOCX files', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Hello'))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('[Content_Types].xml')
    expect(str).toContain('word/document.xml')
    expect(str).toContain('_rels/.rels')
  })

  test('renders heading', () => {
    const doc = docx()
    doc.content((ctx) => ctx.heading('Test Heading', 1))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Test Heading')
    expect(str).toContain('Heading1')
  })

  test('renders paragraph', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Test paragraph'))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Test paragraph')
  })

  test('renders text with size', () => {
    const doc = docx()
    doc.content((ctx) => ctx.text('Large text', 24))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Large text')
    expect(str).toContain('w:sz w:val="48"')
  })

  test('applies bold formatting', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Bold text', { bold: true }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('<w:b/>')
  })

  test('applies italic formatting', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Italic text', { italic: true }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('<w:i/>')
  })

  test('applies color', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Red text', { color: '#ff0000' }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('w:color w:val="ff0000"')
  })

  test('applies alignment', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Centered', { align: 'center' }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('w:jc w:val="center"')
  })

  test('renders horizontal rule', () => {
    const doc = docx()
    doc.content((ctx) => ctx.horizontalRule())
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('w:pBdr')
  })

  test('escapes XML special characters', () => {
    const doc = docx()
    doc.content((ctx) => ctx.paragraph('Test <tag> & "quotes"'))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('&lt;tag&gt;')
    expect(str).toContain('&amp;')
    expect(str).toContain('&quot;')
  })

  test('supports method chaining', () => {
    const doc = docx()
    const result = doc.content((ctx) => ctx.paragraph('Test'))
    expect(result).toBe(doc)
  })
})

describe('odt', () => {
  test('creates valid ZIP archive', () => {
    const doc = odt()
    doc.content(() => {})
    const bytes = doc.build()
    expect(bytes[0]).toBe(0x50)
    expect(bytes[1]).toBe(0x4b)
  })

  test('returns Uint8Array', () => {
    const doc = odt()
    doc.content(() => {})
    expect(doc.build()).toBeInstanceOf(Uint8Array)
  })

  test('includes mimetype as first file', () => {
    const doc = odt()
    doc.content(() => {})
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('mimetype')
    expect(str).toContain('application/vnd.oasis.opendocument.text')
  })

  test('includes required ODT files', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Hello'))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('content.xml')
    expect(str).toContain('styles.xml')
    expect(str).toContain('META-INF/manifest.xml')
  })

  test('renders heading', () => {
    const doc = odt()
    doc.content((ctx) => ctx.heading('Test Heading', 1))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Test Heading')
    expect(str).toContain('Heading1')
  })

  test('renders paragraph', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Test paragraph'))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Test paragraph')
  })

  test('renders text with size', () => {
    const doc = odt()
    doc.content((ctx) => ctx.text('Large text', 24))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Large text')
    expect(str).toContain('fo:font-size="24pt"')
  })

  test('applies bold formatting', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Bold text', { bold: true }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('fo:font-weight="bold"')
  })

  test('applies italic formatting', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Italic text', { italic: true }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('fo:font-style="italic"')
  })

  test('applies color', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Red text', { color: '#ff0000' }))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('fo:color="#ff0000"')
  })

  test('escapes XML special characters', () => {
    const doc = odt()
    doc.content((ctx) => ctx.paragraph('Test <tag> & "quotes"'))
    const bytes = doc.build()
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('&lt;tag&gt;')
    expect(str).toContain('&amp;')
    expect(str).toContain('&quot;')
  })

  test('supports method chaining', () => {
    const doc = odt()
    const result = doc.content((ctx) => ctx.paragraph('Test'))
    expect(result).toBe(doc)
  })
})

describe('markdownToDocx', () => {
  test('returns Uint8Array', () => {
    const bytes = markdownToDocx('# Hello')
    expect(bytes).toBeInstanceOf(Uint8Array)
  })

  test('converts headers', () => {
    const bytes = markdownToDocx('# H1\n## H2\n### H3')
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('H1')
    expect(str).toContain('H2')
    expect(str).toContain('H3')
  })

  test('converts bullet lists', () => {
    const bytes = markdownToDocx('- Item 1\n- Item 2')
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Item 1')
    expect(str).toContain('Item 2')
  })

  test('converts numbered lists', () => {
    const bytes = markdownToDocx('1. First\n2. Second')
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('1. First')
    expect(str).toContain('2. Second')
  })

  test('converts horizontal rules', () => {
    const bytes = markdownToDocx('---')
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('w:pBdr')
  })

  test('handles empty input', () => {
    const bytes = markdownToDocx('')
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(bytes.length).toBeGreaterThan(0)
  })
})

describe('markdownToOdt', () => {
  test('returns Uint8Array', () => {
    const bytes = markdownToOdt('# Hello')
    expect(bytes).toBeInstanceOf(Uint8Array)
  })

  test('converts headers', () => {
    const bytes = markdownToOdt('# H1\n## H2\n### H3')
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('H1')
    expect(str).toContain('H2')
    expect(str).toContain('H3')
  })

  test('converts bullet lists', () => {
    const bytes = markdownToOdt('- Item 1\n- Item 2')
    const str = new TextDecoder().decode(bytes)
    expect(str).toContain('Item 1')
    expect(str).toContain('Item 2')
  })

  test('handles empty input', () => {
    const bytes = markdownToOdt('')
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(bytes.length).toBeGreaterThan(0)
  })
})
