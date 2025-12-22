export interface TextOptions {
  align?: 'left' | 'center' | 'right'
  bold?: boolean
  italic?: boolean
  color?: string
}

export interface DocContext {
  heading(str: string, level: 1 | 2 | 3): void
  paragraph(str: string, opts?: TextOptions): void
  text(str: string, size: number, opts?: TextOptions): void
  lineBreak(): void
  horizontalRule(): void
}

export interface DOCXBuilder {
  content(fn: (ctx: DocContext) => void): DOCXBuilder
  build(): Uint8Array
}

export interface ODTBuilder {
  content(fn: (ctx: DocContext) => void): ODTBuilder
  build(): Uint8Array
}

type DocElement =
  | { type: 'heading'; text: string; level: 1 | 2 | 3 }
  | { type: 'paragraph'; text: string; opts?: TextOptions }
  | { type: 'text'; text: string; size: number; opts?: TextOptions }
  | { type: 'lineBreak' }
  | { type: 'horizontalRule' }

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

function createContext(elements: DocElement[]): DocContext {
  return {
    heading(str: string, level: 1 | 2 | 3) {
      elements.push({ type: 'heading', text: str, level })
    },
    paragraph(str: string, opts?: TextOptions) {
      elements.push({ type: 'paragraph', text: str, opts })
    },
    text(str: string, size: number, opts?: TextOptions) {
      elements.push({ type: 'text', text: str, size, opts })
    },
    lineBreak() {
      elements.push({ type: 'lineBreak' })
    },
    horizontalRule() {
      elements.push({ type: 'horizontalRule' })
    }
  }
}

function parseMarkdown(ctx: DocContext, md: string): void {
  for (const raw of md.split('\n')) {
    const line = raw.trimEnd()
    if (/^#{1,3}\s/.test(line)) {
      const level = line.match(/^#+/)![0].length as 1 | 2 | 3
      ctx.heading(line.slice(level + 1), level)
    } else if (/^[-*]\s/.test(line)) {
      ctx.paragraph('  • ' + line.slice(2))
    } else if (/^\d+\.\s/.test(line)) {
      ctx.paragraph('  ' + line)
    } else if (/^(-{3,}|\*{3,}|_{3,})$/.test(line)) {
      ctx.horizontalRule()
    } else if (line.trim() === '') {
      ctx.lineBreak()
    } else {
      ctx.paragraph(line)
    }
  }
}

function buildDocxRunProps(opts?: TextOptions, size?: number): string {
  let rPr = ''
  if (size) rPr += `<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/>`
  if (opts?.bold) rPr += '<w:b/>'
  if (opts?.italic) rPr += '<w:i/>'
  if (opts?.color) rPr += `<w:color w:val="${opts.color.replace('#', '')}"/>`
  return rPr ? `<w:rPr>${rPr}</w:rPr>` : ''
}

function getAlignment(align?: string): string {
  return align === 'center' ? 'center' : align === 'right' ? 'right' : 'left'
}

function generateDocxDocument(elements: DocElement[]): string {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
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
      body += '<w:p><w:r><w:br/></w:r></w:p>'
    } else if (el.type === 'horizontalRule') {
      body += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'
    }
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}">
<w:body>
${body}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body>
</w:document>`
}

const DOCX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

const DOCX_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

const DOCX_DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`

/**
 * Create a new DOCX document
 */
export function docx(): DOCXBuilder {
  const elements: DocElement[] = []
  const builder: DOCXBuilder = {
    content(fn) {
      fn(createContext(elements))
      return builder
    },
    build() {
      const enc = new TextEncoder()
      return createZip([
        { name: '[Content_Types].xml', data: enc.encode(DOCX_CONTENT_TYPES) },
        { name: '_rels/.rels', data: enc.encode(DOCX_RELS) },
        { name: 'word/_rels/document.xml.rels', data: enc.encode(DOCX_DOC_RELS) },
        { name: 'word/document.xml', data: enc.encode(generateDocxDocument(elements)) }
      ])
    }
  }
  return builder
}

const ODT_MIMETYPE = 'application/vnd.oasis.opendocument.text'

const ODT_MANIFEST = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`

const ODT_STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
<office:styles>
<style:style style:name="Standard" style:family="paragraph"/>
<style:style style:name="Heading1" style:family="paragraph"><style:text-properties fo:font-size="24pt" fo:font-weight="bold"/></style:style>
<style:style style:name="Heading2" style:family="paragraph"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>
<style:style style:name="Heading3" style:family="paragraph"><style:text-properties fo:font-size="14pt" fo:font-weight="bold"/></style:style>
</office:styles>
</office:document-styles>`

function buildOdtStyle(name: string, opts?: TextOptions, size?: number): string {
  let textProps = ''
  if (size) textProps += ` fo:font-size="${size}pt"`
  if (opts?.bold) textProps += ' fo:font-weight="bold"'
  if (opts?.italic) textProps += ' fo:font-style="italic"'
  if (opts?.color) textProps += ` fo:color="${opts.color}"`
  const pProps = opts?.align ? `<style:paragraph-properties fo:text-align="${opts.align}"/>` : ''
  return `<style:style style:name="${name}" style:family="paragraph">${pProps}<style:text-properties${textProps}/></style:style>`
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
      if (el.opts?.bold || el.opts?.italic || el.opts?.color || el.opts?.align) {
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
    }
  }

  const autoStyles = styles.length > 0 ? `<office:automatic-styles>${styles.join('')}</office:automatic-styles>` : ''
  return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">
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
  const elements: DocElement[] = []
  const builder: ODTBuilder = {
    content(fn) {
      fn(createContext(elements))
      return builder
    },
    build() {
      const enc = new TextEncoder()
      return createZip([
        { name: 'mimetype', data: enc.encode(ODT_MIMETYPE) },
        { name: 'META-INF/manifest.xml', data: enc.encode(ODT_MANIFEST) },
        { name: 'content.xml', data: enc.encode(generateOdtContent(elements)) },
        { name: 'styles.xml', data: enc.encode(ODT_STYLES) }
      ])
    }
  }
  return builder
}

/**
 * Convert markdown to DOCX
 */
export function markdownToDocx(md: string): Uint8Array {
  const doc = docx()
  doc.content((ctx) => parseMarkdown(ctx, md))
  return doc.build()
}

/**
 * Convert markdown to ODT
 */
export function markdownToOdt(md: string): Uint8Array {
  const doc = odt()
  doc.content((ctx) => parseMarkdown(ctx, md))
  return doc.build()
}

export default docx
