//
// officegen: pptx charts tests
//
// Please put here all the pptx charts tests.
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
var chartsData = require('../test_files/charts-data.js')
var path = require('path')

var outDir = path.join(__dirname, '../tmp/')

// Common error method:
var onError = function (err) {
  console.log(err)
  assert(false)
}

describe('PPTX generator - charts', function () {
  this.slow(2000)

  before(function (done) {
    fs.mkdir(outDir, 0o777, function (err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a slides with charts', function (done) {
    var pptx = officegen({ type: 'pptx', tempDir: outDir })
    pptx.on('error', onError)

    pptx.setDocTitle('Sample PPTX Document')
    var slide = pptx.makeNewSlide()

    var rows = []
    for (var i = 0; i < 12; i++) {
      var row = []
      for (var j = 0; j < 5; j++) {
        row.push('[' + i + ',' + j + ']')
      }
      rows.push(row)
    } // End of for loop.

    slide.addTable(rows, {})

    var outFilename = 'test-ppt-table-1.pptx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    pptx.generate(out)
    out.on('close', function () {
      done()
    })
  })

  chartsData.forEach(function (chartInfo, chartIdx) {
    it(
      'creates a presentation with charts >>' + chartInfo.renderType,
      function (done) {
        var officegen = require('../')
        var pptx = officegen({ type: 'pptx', tempDir: outDir })
        pptx.on('error', onError)

        pptx.setDocTitle('Sample PPTX Document')
        var slide = pptx.makeNewSlide()
        slide.name = 'OfficeChart slide'
        slide.back = 'ffffff'

        slide.addChart(
          chartInfo,
          function () {
            var outFilename =
              'test-ppt-chart-' +
              chartIdx +
              '-' +
              chartInfo.renderType +
              '.pptx'
            var out = fs.createWriteStream(path.join(outDir, outFilename))
            pptx.generate(out)
            out.on('close', function () {
              done()
            })
          },
          onError
        )
      }
    )
  })
})
