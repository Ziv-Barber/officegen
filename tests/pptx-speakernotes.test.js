//
// officegen: pptx speakernotes plugin tests
//
// Please put here all the pptx speakernotes plugin tests.
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

// Common error method:
var onError = function (err) {
  console.log(err)
  assert(false)
}

describe('PPTX Speaker Notes plugin', function () {
  this.slow(2000)

  before(function (done) {
    fs.mkdir(outDir, 0o777, function (err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a presentation with one speaker note', function (done) {
    var slide
    var pptx = officegen('pptx')
    pptx.on('error', onError)

    pptx.setDocTitle('Testing Speaker Notes')

    //
    // Slide #1:
    //

    slide = pptx.makeNewSlide()

    slide.name = 'Title to the slide?'

    // Change the background color:
    slide.back = '000000'

    // Declare the default color to use on this slide:
    slide.color = 'ffffff'

    // Add some text:
    slide.addText(
      'Created using Officegen version ' + officegen.version,
      0,
      0,
      '80%',
      20
    )

    //
    // Slide #2:
    //

    slide = pptx.makeNewSlide()

    // Change the background color:
    slide.back = '22c0e7'

    // Declare the default color to use on this slide:
    slide.color = '000000'

    // Add some text:
    slide.addText('Just another one', 0, 0, '100%', 20)

    //
    // Slide #3:
    //

    slide = pptx.makeNewSlide()

    // Change the background color:
    slide.back = 'e7bb22'

    // Declare the default color to use on this slide:
    slide.color = '000000'

    // Add some text:
    slide.addText('This one has a speaker note', 0, 0, '100%', 20)

    // Add a speaker note:
    slide.setSpeakerNote(
      'This is a speaker note! Using the new setSpeakerNote feature of the slide API.'
    )

    //
    // Slide #4:
    //

    slide = pptx.makeNewSlide()

    // Change the background color:
    slide.back = 'e7db19'

    // Declare the default color to use on this slide:
    slide.color = '000000'

    // Add some text:
    slide.addText('This one has a speaker note with 2 lines', 0, 0, '100%', 20)

    // Add a speaker note:
    // slide.setSpeakerNote ( 'This is a speaker note!\nUsing the new setSpeakerNote feature of the slide API.' )
    slide.setSpeakerNote('This is a speaker note!')
    slide.setSpeakerNote(
      'Using the new setSpeakerNote feature of the slide API.',
      true
    )

    //
    // Slide #5:
    //

    slide = pptx.makeNewSlide()

    // Change the background color:
    slide.back = '2fe722'

    // Declare the default color to use on this slide:
    slide.color = '000000'

    // Add some text:
    slide.addText('Just another slide', 0, 0, '100%', 20)

    //
    // Generate the pptx file:
    //

    var outFilename = 'test-ppt-notes-1.pptx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    pptx.generate(out)
    out.on('close', function () {
      done()
    })
  })
})
