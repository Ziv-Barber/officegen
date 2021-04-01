var async = require('async')
var officegen = require('../')

var fs = require('fs')
var path = require('path')

var pptx = officegen('pptx')

var outDir = path.join(__dirname, '../tmp/')

var slide
var pObj

pptx.on('finalize', function (written) {
  console.log(
    'Finish to create a PowerPoint file.\nTotal bytes created: ' +
      written +
      '\n'
  )

  // clear the temporatory files
})

pptx.on('error', function (err) {
  console.log(err)
})

pptx.setDocTitle('Sample PPTX Document')

function generateExampleSlides(callback) {
  // do the rest things here
  // console.log('finalize')

  // A way to change the view:
  pptx.view.restoredLeft = 15620 // This is also the default so it's the same as not changing it.
  pptx.view.restoredTop = 54442 // The default is 94660

  // Let's create a new slide:
  slide = pptx.makeNewSlide()

  slide.name = 'The first slide!'

  // Change the background color:
  slide.back = '000000'

  // Declare the default color to use on this slide:
  slide.color = 'ffffff'

  // Basic way to add text string:
  slide.addText('Created using Officegen version ' + officegen.version)
  slide.addText('Fast position', 0, 20)
  slide.addText('Full line', 0, 40, '100%', 20)

  // Add text box with multi colors and fonts:
  slide.addText(
    [
      { text: 'Hello ', options: { font_size: 56 } },
      {
        text: 'World!',
        options: {
          font_size: 56,
          font_face: 'Arial',
          color: 'ffff00'
        }
      }
    ],
    {
      cx: '75%',
      cy: 66,
      y: 150,
      bodyProp: {
        normAutofit: 92500
      }
    }
  )
  // Please note that you can pass object as the text parameter to addText.

  // For a single text just pass a text string to addText:
  slide.addText('Office generator', {
    y: 66,
    x: 'c',
    cx: '50%',
    cy: '1inch',
    font_size: 48,
    color: '0000ff',
    bodyProp: {
      normAutofit: 92500
    }
  })

  pObj = slide.addText('Two\nlines', {
    y: 100,
    x: 10,
    cx: '70%',
    font_face: 'Wide Latin',
    font_size: 54,
    color: 'cc0000',
    bold: true,
    underline: true
  })
  pObj.options.y += 150

  // 2nd slide:
  slide = pptx.makeNewSlide()

  // For every color property (including the back color property) you can pass object instead of the color string:
  slide.back = { type: 'solid', color: '004400' }
  pObj = slide.addText('Office generator', {
    y: 'c',
    x: 0,
    cx: '100%',
    cy: '2cm',
    font_size: 48,
    align: 'center',
    color: { type: 'solid', color: '008800' }
  })
  pObj.setShadowEffect('outerShadow', { bottom: true, right: true })

  slide = pptx.makeNewSlide()

  pObj = slide.addText('Office generator', {
    y: 'c',
    x: 0,
    cx: '100%',
    cy: '2cm',
	font_face: 'Rubik',
	pitch_family: 2,
	charset: -79,
    font_size: 74,
    align: 'center'
  })

  slide = pptx.makeNewSlide()

  slide.show = false
  slide.addText('Red line', 'ff0000')
  slide.addShape(pptx.shapes.OVAL, {
    fill: { type: 'solid', color: 'ff0000', alpha: 50 },
    line: 'ffff00',
    y: 50,
    x: 50
  })
  slide.addText('Red box 1', {
    color: 'ffffff',
    fill: 'ff0000',
    line: 'ffff00',
    line_size: 5,
    y: 100,
    rotate: 45
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '0000ff',
    y: 150,
    x: 150,
    cy: 0,
    cx: 300
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '0000ff',
    y: 150,
    x: 150,
    cy: 100,
    cx: 0
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '0000ff',
    y: 249,
    x: 150,
    cy: 0,
    cx: 300
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '0000ff',
    y: 150,
    x: 449,
    cy: 100,
    cx: 0
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '000088',
    y: 150,
    x: 150,
    cy: 100,
    cx: 300
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '000088',
    y: 150,
    x: 150,
    cy: 100,
    cx: 300
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '000088',
    y: 170,
    x: 150,
    cy: 100,
    cx: 300,
    line_head: 'triangle'
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '000088',
    y: 190,
    x: 150,
    cy: 100,
    cx: 300,
    line_tail: 'triangle'
  })
  slide.addShape(pptx.shapes.LINE, {
    line: '000088',
    y: 210,
    x: 150,
    cy: 100,
    cx: 300,
    line_head: 'stealth',
    line_tail: 'stealth'
  })
  pObj = slide.addShape(pptx.shapes.LINE)
  pObj.options.line = '008888'
  pObj.options.y = 210
  pObj.options.x = 150
  pObj.options.cy = 100
  pObj.options.cx = 300
  pObj.options.line_head = 'stealth'
  pObj.options.line_tail = 'stealth'
  pObj.options.flip_vertical = true
  slide.addText('Red box 2', {
    color: 'ffffff',
    fill: 'ff0000',
    line: 'ffff00',
    y: 350,
    x: 200,
    shape: pptx.shapes.ROUNDED_RECTANGLE,
    indentLevel: 1
  })

  slide = pptx.makeNewSlide()

  slide.addImage(path.resolve(__dirname, 'images_for_examples/image1.png'), {
    y: 'c',
    x: 'c'
  })

  slide = pptx.makeNewSlide()

  slide.addImage(path.resolve(__dirname, 'images_for_examples/image2.jpg'), {
    y: 0,
    x: 0,
    cy: '100%',
    cx: '100%'
  })

  slide = pptx.makeNewSlide()
  slide.addImage(path.resolve(__dirname, 'images_for_examples/image3.png'), {
    y: 'c',
    x: 'c'
  })

  slide = pptx.makeNewSlide()

  slide.addImage(path.resolve(__dirname, 'images_for_examples/image2.jpg'), {
    y: 0,
    x: 0,
    cy: '100%',
    cx: '100%'
  })

  slide = pptx.makeNewSlide()

  slide.addImage(path.resolve(__dirname, 'images_for_examples/image2.jpg'), {
    y: 0,
    x: 0,
    cy: '100%',
    cx: '100%'
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_001.png'), {
    y: 10,
    x: 10
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_002.png'), {
    y: 10,
    x: 110
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_001.png'), {
    y: 110,
    x: 10
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_001.png'), {
    y: 110,
    x: 110
  })

  slide = pptx.makeNewSlide()

  slide.addImage(path.resolve(__dirname, 'images_for_examples/image2.jpg'), {
    y: 0,
    x: 0,
    cy: '100%',
    cx: '100%'
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_001.png'), {
    y: 10,
    x: 10
  })
  slide.addImage(
    path.resolve(__dirname, 'images_for_examples/sword_002.png'),
    110,
    10
  )
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_003.png'), {
    y: 10,
    x: 210
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_004.png'), {
    y: 110,
    x: 10
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_001.png'), {
    y: 110,
    x: 110
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_003.png'), {
    y: 110,
    x: 210
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_002.png'), {
    y: 210,
    x: 10
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_004.png'), {
    y: 210,
    x: 110
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_004.png'), {
    y: 210,
    x: 210
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_004.png'), {
    y: '310',
    x: 10
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_002.png'), {
    y: 310,
    x: 110
  })
  slide.addImage(path.resolve(__dirname, 'images_for_examples/sword_003.png'), {
    y: 310,
    x: 210
  })

  callback()
}

function generateTables(callback) {
  slide = pptx.makeNewSlide()

  // Table with consistent formats
  var rows = []
  var columnWidths = []
  for (var i = 0; i < 12; i++) {
    var row = []
    for (var j = 0; j < 5; j++) {
      row.push('[' + i + ',' + j + ']')
    } // End of for loop.

    rows.push(row)
    columnWidths.push(300 * 1000 + Math.round(Math.random() * 800 * 1000))
  } // End of for loop.

  slide.addTable(rows, {
    font_size: 9,
    font_face: 'Comic Sans MS',
    columnWidths: columnWidths
  })

  // Table with various formats for cells:
  var headerRow = [
    {
      val: 'Region',
      opts: {
        bold: 1
      }
    },
    {
      val: 'Abr.',
      opts: {
        bold: 1
      }
    },
    {
      val: 'Pop.',
      opts: {
        bold: 1
      }
    },
    {
      val: 'Sq. Km.',
      opts: {
        bold: 1
      }
    }
  ]

  var dataRows = [
    {
      val: 'Midwest',
      opts: {
        font_face: 'Verdana',
        align: 'l'
      }
    },
    {
      val: 'MW',
      opts: {
        font_face: 'Verdana',
        align: 'l'
      }
    },
    {
      val: 2000000,
      opts: {
        font_face: 'Verdana',
        align: 'r',
        bold: 1,
        // font_color: 'ffffff',
        fill_color: '00a65a'
      }
    },
    {
      val: 45,
      opts: {
        font_face: 'Verdana',
        align: 'r',
        bold: 1,
        fill_color: 'cccccc'
      }
    }
  ]

  var columnDefinition = [4286250, 952500, 952500, 952500]

  slide.addTable([headerRow, dataRows], {
    font_size: 10,
    font_face: 'Arial',
    columnWidths: columnDefinition
  })

  callback()
}

function finalize() {
  var out = fs.createWriteStream(path.join(outDir, 'example.pptx'))

  out.on('error', function (err) {
    console.log(err)
  })

  pptx.generate(out)
}

async.series(
  [
    generateTables,
    generateExampleSlides // inherited from original project
  ],
  finalize
)
