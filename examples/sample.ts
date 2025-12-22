import { docx, odt } from '../src/index'
import { writeFileSync, readFileSync } from 'fs'

const catImage = readFileSync(new URL('./cat.jpg', import.meta.url))

const doc = docx()

doc.header((ctx) => {
  ctx.paragraph('URGENT - PLEASE HELP!', { bold: true, color: '#c0392b', align: 'center' })
})

doc.footer((ctx) => {
  ctx.paragraph('Generated with tinydocx', { italic: true, color: '#95a5a6', align: 'center' })
})

doc.content((ctx) => {
  ctx.heading('LOST CAT', 1)
  ctx.paragraph('REWARD: $500', { bold: true, color: '#27ae60', align: 'center', size: 24 })
  ctx.lineBreak()

  ctx.image(catImage, { width: 3, height: 2.25 })
  ctx.lineBreak()

  ctx.heading('Description', 2)
  ctx.table([
    ['Name', 'Whiskers'],
    ['Breed', 'Domestic Shorthair'],
    ['Color', 'Orange Tabby'],
    ['Age', '3 years old'],
    ['Weight', '~10 lbs'],
    ['Microchipped', 'Yes']
  ], { colWidths: [3000, 5000] })
  ctx.lineBreak()

  ctx.heading('Last Seen', 2)
  ctx.paragraph('Date: December 20, 2025', { bold: true })
  ctx.paragraph('Location: Oak Street & Maple Avenue')
  ctx.paragraph('Time: Around 6:00 PM')
  ctx.lineBreak()

  ctx.heading('Identifying Features', 2)
  ctx.list([
    'White patch on chest',
    'Small notch in left ear',
    'Wearing blue collar with bell',
    'Very friendly, responds to "Whiskers"'
  ])
  ctx.lineBreak()

  ctx.horizontalRule()
  ctx.lineBreak()

  ctx.heading('Contact Information', 2)
  ctx.paragraph('If found, please contact immediately:', { bold: true })
  ctx.lineBreak()
  ctx.paragraph('Phone: (555) 123-4567', { size: 16 })
  ctx.paragraph('Email: findwhiskers@email.com', { size: 16 })
  ctx.link('Submit sighting online', 'https://findwhiskers.com/report')
  ctx.lineBreak()

  ctx.horizontalRule()
  ctx.lineBreak()

  ctx.paragraph('Whiskers is an indoor cat and may be scared. Please approach gently.', { italic: true, align: 'center' })
  ctx.paragraph('Thank you for your help bringing our beloved cat home!', { bold: true, align: 'center', color: '#2980b9' })
})

writeFileSync('sample.docx', doc.build())
console.log('Created sample.docx')

const odtDoc = odt()
odtDoc.content((ctx) => {
  ctx.heading('LOST CAT', 1)
  ctx.paragraph('REWARD: $500', { bold: true, color: '#27ae60', align: 'center' })
  ctx.lineBreak()

  ctx.heading('Description', 2)
  ctx.table([
    ['Name', 'Whiskers'],
    ['Breed', 'Domestic Shorthair'],
    ['Color', 'Orange Tabby'],
    ['Age', '3 years old']
  ])
  ctx.lineBreak()

  ctx.heading('Last Seen', 2)
  ctx.paragraph('Date: December 20, 2025', { bold: true })
  ctx.paragraph('Location: Oak Street & Maple Avenue')
  ctx.lineBreak()

  ctx.heading('Identifying Features', 2)
  ctx.list([
    'White patch on chest',
    'Small notch in left ear',
    'Wearing blue collar with bell',
    'Very friendly'
  ])
  ctx.lineBreak()

  ctx.horizontalRule()

  ctx.heading('Contact', 2)
  ctx.paragraph('Phone: (555) 123-4567')
  ctx.link('Submit sighting online', 'https://findwhiskers.com/report')
  ctx.lineBreak()

  ctx.paragraph('Thank you for your help!', { bold: true, align: 'center' })
})

writeFileSync('sample.odt', odtDoc.build())
console.log('Created sample.odt')
