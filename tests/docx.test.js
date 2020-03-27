//
// officegen: docx tests
//
// Please put here all the docx tests.
//
// Copyright (c) 2013 Ziv Barber;
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//

var assert = require('assert')
var officegen = require('../')
var fs = require('fs')
var path = require('path')

var dirImages = path.join(__dirname, '../examples/')
var outDir = path.join(__dirname, '../tmp/')

// Common error method:
var onError = function (err) {
  console.log(err)
  assert(false)
}

describe('DOCX generator', function () {
  this.timeout(2000)
  this.slow(2000)

  before(function (done) {
    fs.mkdir(outDir, 0o777, function (err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a document with text and styles', function (done) {
    var docx = officegen('docx')
    docx.on('error', onError)

    var pObj = docx.createP()

    pObj.addText('Simple')
    pObj.addText(' with color', { color: '000088' })
    pObj.addText(' and back color.', { color: '00ffff', back: '000088' })

    pObj = docx.createP()

    pObj.addText('Bold + underline', { bold: true, underline: true })

    pObj = docx.createP({ align: 'center' })

    pObj.addText('Center this text.')

    pObj = docx.createP()
    pObj.options.align = 'right'

    pObj.addText('Align this text to the right.')

    pObj = docx.createP()

    pObj.addText('Those two lines are in the same paragraph,')
    pObj.addLineBreak()
    pObj.addText('but they are separated by a line break.')

    docx.putPageBreak()

    pObj = docx.createP()

    pObj.addText('Fonts face only.', { font_face: 'Arial' })
    pObj.addText(' Fonts face and size.', { font_face: 'Arial', font_size: 40 })

    docx.putPageBreak()

    pObj = docx.createListOfNumbers()

    pObj.addText('Option 1')

    pObj = docx.createListOfNumbers()

    pObj.addText('Option 2')

    var outFilename = 'test-doc-1.docx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    docx.generate(out)
    out.on('close', function () {
      done()
    })
  })

  it('can handle text without spaces', function (done) {
    var docx = officegen('docx')
    docx.on('error', onError)

    var pObj = docx.createP()
    pObj.addText('Hello,World')

    var outFilename = 'test-doc-3.docx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    docx.generate(out)
    out.on('close', function () {
      done()
    })
  })

  it('creates a document with images', function (done) {
    var docx = officegen('docx')
    docx.on('error', onError)

    var pObj = docx.createP()

    pObj = docx.createP()

    pObj.addImage(path.resolve(dirImages, 'images_for_examples/image3.png'))

    docx.putPageBreak()

    pObj = docx.createP()

    pObj.addImage(path.resolve(dirImages, 'images_for_examples/image1.png'))

    pObj = docx.createP()

    pObj.addImage(path.resolve(dirImages, 'images_for_examples/sword_001.png'))
    pObj.addImage(path.resolve(dirImages, 'images_for_examples/sword_002.png'))
    pObj.addImage(path.resolve(dirImages, 'images_for_examples/sword_003.png'))
    pObj.addText('... some text here ...', { font_face: 'Arial' })
    pObj.addImage(path.resolve(dirImages, 'images_for_examples/sword_004.png'))

    pObj = docx.createP()

    pObj.addImage(path.resolve(dirImages, 'images_for_examples/image1.png'))

    var outFilename = 'test-doc-2.docx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    out.on('error', onError)

    docx.generate(out)
    out.on('close', function () {
      done()
    })
  })

  it('can handle right-to-left', function (done) {
    var docx = officegen('docx')
    docx.on('error', onError)

    var pObj = docx.createP({ rtl: true })
    pObj.addText('نص ذو اتجاه من اليمين')

    pObj = docx.createP({ rtl: true })
    pObj.addText('ישור לימין')

    var table = [
      [
        {
          val: 'من اليمين',
          opts: {
            b: true,
            sz: '48',
            shd: {
              fill: '7F7F7F',
              themeFill: 'text1',
              themeFillTint: '80'
            },
            fontFamily: 'Avenir Book',
            rtl: true
          }
        },
        {
          val: 'Title Line 1\r\n\r\nTitle Line 2',
          opts: {
            b: true,
            color: 'A00000',
            align: 'right',
            shd: {
              fill: '92CDDC',
              themeFill: 'text1',
              themeFillTint: '80'
            }
          }
        }
      ]
    ]

    var tableStyle = {
      borders: true,
      rtl: true
    }

    docx.createTable(table, tableStyle)

    table = [
      [
        {
          val: 'מימין',
          opts: {
            b: true,
            sz: '48',
            shd: {
              fill: '7F7F7F',
              themeFill: 'text1',
              themeFillTint: '80'
            },
            fontFamily: 'Avenir Book',
            rtl: true
          }
        },
        {
          val: ['Title Line 1', '', 'Title Line 2'],
          opts: {
            b: true,
            color: 'A00000',
            align: 'right',
            shd: {
              fill: '92CDDC',
              themeFill: 'text1',
              themeFillTint: '80'
            }
          }
        }
      ]
    ]

    docx.createTable(table, tableStyle)

    var outFilename = 'test-doc-rtl.docx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    out.on('error', onError)
    docx.generate(out)
    out.on('close', function () {
      done()
    })
  })
})
