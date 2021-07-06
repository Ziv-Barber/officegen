//
// officegen: docx table.
//
// Please refer to README.md for this module's documentations.
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

module.exports = {
  // assume passed in an array of row objects
  getTable: function (rows, tblOpts) {
    tblOpts = tblOpts || {}

    var self = this

    return self._getBase(
      rows.map(function (row) {
        return self._getRow(
          row.map(function (cell) {
            cell = cell || {}
            if (
              typeof cell === 'string' ||
              typeof cell === 'number' ||
              Array.isArray(cell)
            ) {
              var val = cell
              cell = {
                val: val
              }
            }

            return self._getCell(cell.val, cell.opts, tblOpts)
          }),
          tblOpts
        )
      }),
      self._getColSpecs(rows, tblOpts),
      tblOpts
    )
  },

  _getBase: function (rowSpecs, colSpecs, opts) {
    var baseTable = {
      'w:tbl': {
        'w:tblPr': {
          'w:tblStyle': {
            '@w:val': 'TableGrid'
          },
          'w:tblW': {
            '@w:w': opts.tableWidth || '0',
            '@w:type': opts.tableWidthType || 'auto'
          },
          'w:tblInd': {
            '@w:w': opts.indent || '0',
            '@w:type': 'dxa'
          },
          'w:tblLook': {
            '@w:val': '05EO'
          }
        },
        'w:tblGrid': colSpecs,
        '#text': rowSpecs
      }
    }
    if (opts.fixedLayout) {
      baseTable['w:tbl']['w:tblLayout']['@w:type'] = 'fixed'
    }
    if (opts.borders) {
      const defaultSize = 4
      baseTable['w:tbl']['w:tblPr']['w:tblBorders'] = {
        'w:top': {
          '@w:val': 'single',
          '@w:sz': opts.borderSize || defaultSize,
          '@w:space': '0',
          '@w:color': '000000'
        },
        'w:bottom': {
          '@w:val': 'single',
          '@w:sz': opts.borderSize || defaultSize,
          '@w:space': '0',
          '@w:color': '000000'
        },
        'w:left': {
          '@w:val': 'single',
          '@w:sz': opts.borderSize || defaultSize,
          '@w:space': '0',
          '@w:color': '000000'
        },
        'w:right': {
          '@w:val': 'single',
          '@w:sz': opts.borderSize || defaultSize,
          '@w:space': '0',
          '@w:color': '000000'
        },
        'w:insideH': {
          '@w:val': 'single',
          '@w:sz': opts.borderSize || defaultSize,
          '@w:space': '0',
          '@w:color': '000000'
        },
        'w:insideV': {
          '@w:val': 'single',
          '@w:sz': opts.borderSize || defaultSize,
          '@w:space': '0',
          '@w:color': '000000'
        }
      }
    }
    if (opts.borderStyle || opts.boderStyle) {
      baseTable['w:tbl']['w:tblPr']['w:tblBorders'] =
        opts.borderStyle || opts.boderStyle
    }
    if (opts.rtl) {
      baseTable['w:tbl']['w:tblPr']['w:bidiVisual'] = {}
    }
    return baseTable
  },

  _getColSpecs: function (cols, opts) {
    var self = this
    if (opts.columns) {
      return opts.columns.map(function (col) {
        return self._tblGrid(col.width)
      })
    }
    return cols[0].map(function (col) {
      return self._tblGrid(col.opts.cellColWidth || opts.tableColWidth)
    })
  },

  // TODO
  _tblGrid: function (width) {
    return {
      'w:gridCol': {
        '@w:w': width || '1'
      }
    }
  },

  _getRow: function (cells, opts) {
    return {
      'w:tr': {
        '@w:rsidR': '00995B51',
        '@w:rsidTr': '007F1D13',
        '#text': cells // populate this with an array of table cell objects
      }
    }
  },

  _getCellP : function (text, opts, tblOpts, style) {

    let p = {
      '@w:rsidR': '00995B51',
      '@w:rsidRPr': '00722E63',
      '@w:rsidRDefault': '00995B51',
      'w:pPr': {
        'w:rPr': {
          'w:lang': {
            '@w:val': 'en-AU'
          }
        }
      },
      'w:r': {
        '@w:rsidRPr': '00722E63',
        'w:rPr': {
          'w:lang': {
            '@w:val': 'en-AU'
          }
        },
        '#text': text
      }
    }
    if (style) {
      p['w:pPr']['w:pStyle'] = {
        '@w:val': style
      };
    }
    return p;
  },

  _getCell: function (val, opts, tblOpts) {
    opts = opts || {}
    // var b = {};

    // if (opts.b) {
    //   b = {
    //     "w:tc": {
    //       "w:p": {
    //         "w:r": {
    //           "w:rPr": {
    //             "w:b": {}
    //           }
    //         }
    //       }
    //     }
    //   }
    // }

    var splitLines = [];

    // handle lines as elements in array
    if (Array.isArray(val)) {
      splitLines = val
    } else if (typeof val === 'string' && val.includes('\r\n')) {
      // handle line breaks in cell text
      splitLines = val.split(/\r?\n/)
    } else {
      splitLines[0] = val
    }
    var multiLineBreakObj = [{ 'w:t': splitLines[0] }]
    for (var i = 1; i < splitLines.length; i++) {
      multiLineBreakObj.push({ 'w:br': '' })
      multiLineBreakObj.push({ 'w:t': splitLines[i] })
    }

    var cellObj = {
      'w:tc': {
        'w:tcPr': {
          'w:gridSpan': {
            '@w:val': opts.gridSpan || '1'
          },
          'w:vAlign': {
            '@w:val': opts.vAlign || 'top'
          }
        }
      }
    }

    if (opts.multiP) {
      let paras = [];
      val.forEach((p) => {
        paras.push({'w:p': this._getCellP({'w:t': p.text}, opts, tblOpts, p.style)});
      });
      cellObj['w:tc']['#text'] = paras;
    } else {
      cellObj['w:tc']['w:p'] = this._getCellP(multiLineBreakObj, opts, tblOpts, opts.pStyle);
    }
    
    if (opts.cellColWidth || tblOpts.tableColWidth) {
      cellObj['w:tc']['w:tcPr']['w:tcW'] = {
        '@w:w': opts.cellColWidth || tblOpts.tableColWidth || '0',
        '@w:type': 'dxa'
      }
    }

    if (opts.shd) {
      cellObj['w:tc']['w:tcPr']['w:shd'] = {
        '@w:val': 'clear',
        '@w:color': 'auto',
        '@w:fill': opts.shd.fill || '',
        '@w:themeFill': opts.shd.themeFill || '',
        '@w:themeFillTint': opts.shd.themeFillTint || ''
      }
    }

   /* if (!opts.b) {
      if (opts.multiP) {
        cellObj['w:tc']['w:p'].forEach((p, index) => {
          delete cellObj['w:tc']['w:p'][index]['w:r']['w:rPr']['w:b']
        })
      } else delete cellObj['w:tc']['w:p']['w:r']['w:rPr']['w:b']
    }
    if (tblOpts.rtl) {
      if (opts.multiP) {
        cellObj['w:tc']['w:p'].forEach((p, index) => {
          cellObj['w:tc']['w:p'][index]['w:pPr']['w:bidi'] = {'@w:val': '1'}
        })
      } else cellObj['w:tc']['w:p']['w:pPr']['w:bidi'] = { '@w:val': '1' }
    }
    if (opts.rtl) {
      if (opts.multiP) {
        cellObj['w:tc']['w:p'].forEach((p, index) => {
          cellObj['w:tc']['w:p'][index]['w:r']['w:rPr']['w:rtl'] = {'@w:val': '1'}
        })
      } else cellObj['w:tc']['w:p']['w:r']['w:rPr']['w:rtl'] = { '@w:val': '1' }
    }
    if (opts.vMerge) {
      cellObj['w:tc']['w:tcPr']['w:vMerge'] = { '@w:val': opts.vMerge }
    }*/
    return cellObj
  }
}
