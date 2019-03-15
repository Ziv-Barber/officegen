//
// officegen: pptx layouts plugin tests
//
// Please put here all the pptx layouts plugin tests.
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

// var pluginLayouts = require('../lib/pptxplg-layouts')

var outDir = path.join(__dirname, '../tmp/')

// Common error method
var onError = function(err) {
  console.log(err)
  assert(false)
}

describe('PPTX Layouts plugin', function() {
  this.slow(1000)

  before(function(done) {
    fs.mkdir(outDir, 0o777, function(err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a presentation with the title layout', function(done) {
    var slide
    var pptx = officegen({
      type: 'pptx',
      extraPlugs: [
        // pluginLayouts // The 'pptxplg-layouts' plugin.
      ]
    })
    pptx.on('error', onError)

    pptx.setDocTitle('Testing Layouts')

    //
    // Slide #1:
    //

    slide = pptx.makeNewSlide({
      useLayout: 'title'
    })

    //
    // Slide #2:
    //

    slide = pptx.makeNewSlide({
      useLayout: 'title'
    })

    slide.setTitle('The title')
    slide.setSubTitle('Another text')

    // Add a speaker note:
    slide.setSpeakerNote(
      'This is a speaker note! Using the new setSpeakerNote feature of the slide API.'
    )

    //
    // Slide #3:
    //

    slide = pptx.makeNewSlide({
      useLayout: 'title'
    })

    slide.setTitle([
      { text: 'Hello ', options: { font_size: 56 } },
      {
        text: 'World!',
        options: { font_size: 56, font_face: 'Arial', color: 'ffff00' }
      }
    ])
    slide.setSubTitle('Another text')

    //
    // Slide #4:
    //

    slide = pptx.makeTitleSlide()

    //
    // Slide #5:
    //

    slide = pptx.makeTitleSlide('The title of this slide', 'Sub title')

    //
    // Slide #6:
    //

    slide = pptx.makeTitleSlide(
      [
        { text: 'Hello ', options: { font_size: 56 } },
        {
          text: 'World!',
          options: { font_size: 56, font_face: 'Arial', color: 'ffff00' }
        }
      ],
      'Sub title'
    )

    //
    // Slide #7:
    //

    slide = pptx.makeObjSlide('The title of slide 7', [
      { text: '', options: { listType: 'dot' } },
      { text: 'Some ', options: { font_size: 56 } },
      {
        text: 'data',
        options: { font_size: 56, font_face: 'Arial', color: 'ff8800' }
      }
    ])

    slide.useLayout.isDate = false
    slide.setFooter('Message in the footer')

    //
    // Slide #8:
    //

    slide = pptx.makeSecHeadSlide('The title of slide 8', 'Sub title')

    //
    // Generate the pptx file:
    //

    var outFilename = 'test-ppt-layouts-1.pptx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    pptx.generate(out)
    out.on('close', function() {
      done()
    })
  })
})
