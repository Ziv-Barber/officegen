//
// officegen: xlsx basic tests
//
// Please put here all the xlsx basic tests.
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

describe('XLSX generator', function () {
  before(function (done) {
    fs.mkdir(outDir, 0o777, function (err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a spreadsheet with text and numbers', function (done) {
    var xlsx = officegen('xlsx')
    xlsx.on('error', onError)

    var sheet = xlsx.makeNewSheet()
    sheet.name = 'Excel Test'

    sheet.setColumnWidth('A', 16.5)
    sheet.setColumnWidth('E', 10.5)
    // sheet.setColumnCenter('C') // NOT working yet!!!!

    // The direct option - two-dimensional array:
    sheet.data[0] = []
    sheet.data[0][0] = 1
    sheet.data[1] = []
    sheet.data[1][3] = 'abc'
    sheet.data[1][4] = 'More'
    sheet.data[1][5] = 'Text'
    sheet.data[1][6] = 'Here'
    sheet.data[2] = []
    sheet.data[2][5] = 'abc'
    sheet.data[2][6] = 900
    sheet.data[6] = []
    sheet.data[6][2] = 1972

    // Using setCell:
    sheet.setCell('E7', 340)
    sheet.setCell('I1', -3)
    sheet.setCell('I2', 31.12)
    sheet.setCell('G102', 'Hello World!')

    var outFilename = 'test-xls-1.xlsx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    xlsx.generate(out)
    out.on('close', function () {
      done()
    })
  })
})
