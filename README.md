# tinydocx

Minimal DOCX/ODT creation library. **<400 LOC, zero dependencies, makes real documents.**

```
npm install tinydocx
```

---

## Why tinydocx?

|  | tinydocx | docx |
| --- | --- | --- |
| **Size** | ~4 KB | ~180 KB |
| **Dependencies** | 0 | 5 |

**~45x smaller.** We removed custom fonts, images, headers/footers, tables, and advanced features. What's left is the 90% use case: **put text in a document.**

### Build with it

Invoices, receipts, reports, contracts, letters, simple exports

### Features

| Feature | Description |
| --- | --- |
| **Text** | Any size, bold/italic, hex colors |
| **Paragraphs** | Alignment (left/center/right) |
| **Headings** | H1, H2, H3 with appropriate sizing |
| **Markdown** | Convert markdown to DOCX/ODT |
| **ODT** | OpenDocument format support |

### Not included

Images, headers/footers, page numbers, tables, custom fonts, lists, hyperlinks

Need those? Use [docx](https://github.com/dolanmiu/docx).

---

## Quick start

```typescript
import { docx } from 'tinydocx'
import { writeFileSync } from 'fs'

const doc = docx()
doc.content((ctx) => {
  ctx.heading('Hello World', 1)
  ctx.paragraph('This is a paragraph.')
  ctx.paragraph('Bold text', { bold: true })
  ctx.paragraph('Centered', { align: 'center' })
  ctx.text('Custom size (18pt)', 18)
})

writeFileSync('output.docx', doc.build())
```

---

## API

```typescript
import { docx, odt, markdownToDocx, markdownToOdt } from 'tinydocx'

// Create document
const doc = docx()  // or odt()

// Add content
doc.content((ctx) => {
  ctx.heading(str, level)          // level: 1, 2, or 3
  ctx.paragraph(str, opts?)        // simple paragraph
  ctx.text(str, size, opts?)       // text with font size (points)
  ctx.lineBreak()                  // empty line
  ctx.horizontalRule()             // horizontal rule
})

// Build
doc.build()                        // returns Uint8Array
```

### TextOptions

```typescript
{
  align?: 'left' | 'center' | 'right'
  bold?: boolean
  italic?: boolean
  color?: string   // hex color (e.g., '#FF0000')
}
```

### Markdown conversion

```typescript
import { markdownToDocx, markdownToOdt } from 'tinydocx'

const md = `
# Hello World

This is a paragraph.

- Item 1
- Item 2

---

1. First
2. Second
`

writeFileSync('output.docx', markdownToDocx(md))
writeFileSync('output.odt', markdownToOdt(md))
```

Supported markdown: `# ## ###` headings, `- *` bullet lists, `1.` numbered lists, `---` rules, paragraphs

---

## Full example

```typescript
import { docx } from 'tinydocx'
import { writeFileSync } from 'fs'

const doc = docx()
doc.content((ctx) => {
  ctx.heading('INVOICE', 1)
  ctx.text('#INV-2025-001', 10, { color: '#666666' })
  ctx.lineBreak()

  ctx.paragraph('Acme Corporation', { bold: true })
  ctx.paragraph('123 Business Street')
  ctx.paragraph('New York, NY 10001', { color: '#666666' })
  ctx.lineBreak()

  ctx.horizontalRule()
  ctx.lineBreak()

  ctx.paragraph('Website Development - $5,000.00')
  ctx.paragraph('Hosting (Annual) - $200.00')
  ctx.paragraph('Maintenance Package - $1,800.00')
  ctx.lineBreak()

  ctx.paragraph('Total Due: $7,000.00', { bold: true, align: 'right' })
  ctx.lineBreak()

  ctx.paragraph('Thank you for your business!', { italic: true, align: 'center' })
})

writeFileSync('invoice.docx', doc.build())
```

---

## How it works

A .docx file is a ZIP archive containing XML files. tinydocx generates the minimal required structure:

```
[Content_Types].xml    # MIME type declarations
_rels/.rels            # Package relationships
word/document.xml      # Your actual content
word/_rels/document.xml.rels
```

ODT files have a similar structure:

```
mimetype               # MIME type (must be first)
META-INF/manifest.xml  # File manifest
content.xml            # Your actual content
styles.xml             # Style definitions
```

The library includes a minimal ZIP implementation (~60 LOC) and templates the XML directly. No compression (STORE method) keeps the code simple.

---

## License

MIT
