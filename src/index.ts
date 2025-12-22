export interface TextOptions {
  align?: 'left' | 'center' | 'right'
  bold?: boolean
  italic?: boolean
  underline?: boolean
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

type DocElement =
  | { type: 'heading'; text: string; level: 1 | 2 | 3 }
  | { type: 'paragraph'; text: string; opts?: TextOptions }
  | { type: 'text'; text: string; size: number; opts?: TextOptions }
  | { type: 'lineBreak' }
  | { type: 'horizontalRule' }
  | { type: 'list'; items: string[]; ordered: boolean }
  | { type: 'table'; rows: string[][]; opts?: TableOptions }
  | { type: 'link'; text: string; url: string; opts?: TextOptions; rId: string }
  | { type: 'image'; data: Uint8Array; opts: ImageOptions; rId: string }
  | { type: 'pageNumber' }

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

interface BuildContext {
  elements: DocElement[]
  hyperlinks: { url: string; rId: string }[]
  images: { data: Uint8Array; rId: string; ext: string }[]
  nextRId: number
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
      const ext = detectImageType(data)
      ctx.images.push({ data, rId, ext })
      ctx.elements.push({ type: 'image', data, opts, rId })
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
  return 'png'
}

function buildDocxRunProps(opts?: TextOptions, size?: number): string {
  let rPr = ''
  if (opts?.font) rPr += `<w:rFonts w:ascii="${opts.font}" w:hAnsi="${opts.font}"/>`
  const sz = size || opts?.size
  if (sz) rPr += `<w:sz w:val="${sz * 2}"/><w:szCs w:val="${sz * 2}"/>`
  if (opts?.bold) rPr += '<w:b/>'
  if (opts?.italic) rPr += '<w:i/>'
  if (opts?.underline) rPr += '<w:u w:val="single"/>'
  if (opts?.color) rPr += `<w:color w:val="${opts.color.replace('#', '')}"/>`
  return rPr ? `<w:rPr>${rPr}</w:rPr>` : ''
}

function getAlignment(align?: string): string {
  return align === 'center' ? 'center' : align === 'right' ? 'right' : align === 'justify' ? 'both' : 'left'
}

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
      body += `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="Image"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${el.rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    } else if (el.type === 'pageNumber') {
      body += '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    }
  }
  return body
}

function generateDocxDocument(elements: DocElement[], headerRId?: string, footerRId?: string): string {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const body = buildDocxBody(elements)

  let sectPr = '<w:sectPr>'
  if (headerRId) sectPr += `<w:headerReference w:type="default" r:id="${headerRId}"/>`
  if (footerRId) sectPr += `<w:footerReference w:type="default" r:id="${footerRId}"/>`
  sectPr += '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>'

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="${R}">
<w:body>
${body}
${sectPr}
</w:body>
</w:document>`
}

function generateDocxHeader(elements: DocElement[]): string {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const body = buildDocxBody(elements)
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W}">
${body}
</w:hdr>`
}

function generateDocxFooter(elements: DocElement[]): string {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const body = buildDocxBody(elements)
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W}">
${body}
</w:ftr>`
}

function generateDocxStyles(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="48"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="Heading 3"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>`
}

function generateDocxNumbering(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`
}

function generateDocxContentTypes(hasLists: boolean, hasHeader: boolean, hasFooter: boolean, imageExts: string[]): string {
  let types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
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
  images: { rId: string; ext: string }[]
): string {
  const REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
  const OFFREL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

  let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL}">
<Relationship Id="rId1" Type="${OFFREL}/styles" Target="styles.xml"/>`

  let nextId = 2
  if (hasLists) {
    rels += `<Relationship Id="rId${nextId++}" Type="${OFFREL}/numbering" Target="numbering.xml"/>`
  }
  if (hasHeader) {
    rels += `<Relationship Id="rIdHeader" Type="${OFFREL}/header" Target="header1.xml"/>`
  }
  if (hasFooter) {
    rels += `<Relationship Id="rIdFooter" Type="${OFFREL}/footer" Target="footer1.xml"/>`
  }

  for (const link of hyperlinks) {
    rels += `<Relationship Id="${link.rId}" Type="${OFFREL}/hyperlink" Target="${escapeXml(link.url)}" TargetMode="External"/>`
  }

  for (let i = 0; i < images.length; i++) {
    rels += `<Relationship Id="${images[i].rId}" Type="${OFFREL}/image" Target="media/image${i + 1}.${images[i].ext}"/>`
  }

  rels += `</Relationships>`
  return rels
}

const DOCX_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

/**
 * Create a new DOCX document
 */
export function docx(): DOCXBuilder {
  const mainCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 10 }
  const headerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 100 }
  const footerCtx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 200 }
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
      const allHyperlinks = [...mainCtx.hyperlinks, ...headerCtx.hyperlinks, ...footerCtx.hyperlinks]
      const allImages = [...mainCtx.images, ...headerCtx.images, ...footerCtx.images]
      const imageExts = allImages.map(img => img.ext)

      const files: { name: string; data: Uint8Array }[] = [
        { name: '[Content_Types].xml', data: enc.encode(generateDocxContentTypes(hasLists, hasHeader, hasFooter, imageExts)) },
        { name: '_rels/.rels', data: enc.encode(DOCX_RELS) },
        { name: 'word/_rels/document.xml.rels', data: enc.encode(generateDocxRels(hasLists, hasHeader, hasFooter, allHyperlinks, allImages)) },
        { name: 'word/document.xml', data: enc.encode(generateDocxDocument(mainCtx.elements, hasHeader ? 'rIdHeader' : undefined, hasFooter ? 'rIdFooter' : undefined)) },
        { name: 'word/styles.xml', data: enc.encode(generateDocxStyles()) }
      ]

      if (hasLists) {
        files.push({ name: 'word/numbering.xml', data: enc.encode(generateDocxNumbering()) })
      }
      if (hasHeader) {
        files.push({ name: 'word/header1.xml', data: enc.encode(generateDocxHeader(headerCtx.elements)) })
      }
      if (hasFooter) {
        files.push({ name: 'word/footer1.xml', data: enc.encode(generateDocxFooter(footerCtx.elements)) })
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
  if (opts?.font) textProps += ` style:font-name="${opts.font}"`
  const sz = size || opts?.size
  if (sz) textProps += ` fo:font-size="${sz}pt"`
  if (opts?.bold) textProps += ' fo:font-weight="bold"'
  if (opts?.italic) textProps += ' fo:font-style="italic"'
  if (opts?.underline) textProps += ' style:text-underline-style="solid"'
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
      body += '<table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">'
      for (const row of el.rows) {
        body += '<table:table-row>'
        for (const cell of row) {
          body += `<table:table-cell><text:p>${escapeXml(cell)}</text:p></table:table-cell>`
        }
        body += '</table:table-row>'
      }
      body += '</table:table>'
    } else if (el.type === 'link') {
      body += `<text:p><text:a xlink:href="${escapeXml(el.url)}" xmlns:xlink="http://www.w3.org/1999/xlink">${escapeXml(el.text)}</text:a></text:p>`
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
  const ctx: BuildContext = { elements: [], hyperlinks: [], images: [], nextRId: 1 }
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

function parseMarkdown(ctx: DocContext, md: string): void {
  const lines = md.split('\n')
  let i = 0
  while (i < lines.length) {
    const line = lines[i].trimEnd()
    if (/^#{1,3}\s/.test(line)) {
      const level = line.match(/^#+/)![0].length as 1 | 2 | 3
      ctx.heading(line.slice(level + 1), level)
    } else if (/^[-*]\s/.test(line)) {
      const items: string[] = []
      while (i < lines.length && /^[-*]\s/.test(lines[i])) {
        items.push(lines[i].slice(2))
        i++
      }
      ctx.list(items, false)
      continue
    } else if (/^\d+\.\s/.test(line)) {
      const items: string[] = []
      while (i < lines.length && /^\d+\.\s/.test(lines[i])) {
        items.push(lines[i].replace(/^\d+\.\s/, ''))
        i++
      }
      ctx.list(items, true)
      continue
    } else if (/^(-{3,}|\*{3,}|_{3,})$/.test(line)) {
      ctx.horizontalRule()
    } else if (line.trim() === '') {
      ctx.lineBreak()
    } else {
      ctx.paragraph(line)
    }
    i++
  }
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
