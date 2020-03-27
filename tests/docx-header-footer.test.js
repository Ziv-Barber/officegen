//
// officegen: docx header-footer plugin tests
//
// Please put here all the docx header-footer plugin tests.
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

var outDir = path.join(__dirname, '../tmp/')

// var pluginToTest = require ( '../lib/docxplg-headfoot' )

// Common error method:
var onError = function (err) {
  console.log(err)
  assert(false)
}

describe('DOCX generator with header and footer', function () {
  this.timeout(1000)

  before(function (done) {
    fs.mkdir(outDir, 0o777, function (err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a document with header and footer', function (done) {
    var docx = officegen({
      type: 'docx',
      extraPlugs: [
        // pluginToTest // The 'docxplg-headfoot' plugin.
      ]
    })
    docx.on('error', onError)

    // Add a header:
    var header = docx.getHeader().createP()
    header.addText('This is the header')

    var pObj = docx.createP()

    pObj.addText('Click me please!', { hyperlink: 'testBM' })

    pObj = docx.createP()

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

    pObj.startBookmark('testBM')

    pObj.addText('Those two lines are in the same paragraph,')
    pObj.addLineBreak()
    pObj.addText('but they are separated by a line break.')

    pObj.endBookmark()

    docx.putPageBreak()

    pObj = docx.createP()

    pObj.addText('Fonts face only.', { font_face: 'Arial' })
    pObj.addText(' Fonts face and size.', { font_face: 'Arial', font_size: 40 })

    var outFilename = 'test-docx-header-foooter-1.docx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    docx.generate(out)
    out.on('close', function () {
      done()
    })
  })
})
