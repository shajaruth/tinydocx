# tinydocx

Minimal DOCX/ODT creation library. Zero dependencies.

## Install

```bash
npm install tinydocx
```

## Usage

### DOCX

```javascript
import { docx } from 'tinydocx'
import { writeFileSync } from 'fs'

const doc = docx()
doc.content((ctx) => {
  ctx.heading('Invoice', 1)
  ctx.paragraph('Thank you for your business!')
  ctx.text('Total: $100.00', 14, { bold: true })
})

writeFileSync('output.docx', doc.build())
```

### ODT

```javascript
import { odt } from 'tinydocx'
import { writeFileSync } from 'fs'

const doc = odt()
doc.content((ctx) => {
  ctx.heading('Report', 1)
  ctx.paragraph('This is the content.')
})

writeFileSync('output.odt', doc.build())
```

### Markdown

```javascript
import { markdownToDocx, markdownToOdt } from 'tinydocx'
import { writeFileSync } from 'fs'

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

## API

### `docx()` / `odt()`

Create a new document builder.

```typescript
const doc = docx() // or odt()
```

### `.content(fn)`

Add content using a callback with a context object.

```typescript
doc.content((ctx) => {
  ctx.heading('Title', 1)        // level: 1, 2, or 3
  ctx.paragraph('Text')          // simple paragraph
  ctx.text('Text', 12)           // text with font size
  ctx.lineBreak()                // empty line
  ctx.horizontalRule()           // horizontal rule
})
```

### Text Options

```typescript
interface TextOptions {
  align?: 'left' | 'center' | 'right'
  bold?: boolean
  italic?: boolean
  color?: string  // hex color like '#ff0000'
}

ctx.paragraph('Centered bold', { align: 'center', bold: true })
ctx.text('Red italic', 14, { color: '#ff0000', italic: true })
```

### `.build()`

Generate the document as a `Uint8Array`.

```typescript
const bytes = doc.build()
```

### `markdownToDocx(md)` / `markdownToOdt(md)`

Convert markdown string to document bytes.

Supported markdown:
- `# H1`, `## H2`, `### H3` - headings
- `- item` or `* item` - bullet lists
- `1. item` - numbered lists
- `---` or `***` - horizontal rules
- Plain text - paragraphs

## License

MIT
