//
// officegen: tables
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

var EMU = 914400 // OfficeXML measures in English Metric Units

module.exports = {
  // assume passed in an array of row objects
  getTable: function (rows, options) {
    options = options || {}
    options.tabstyle = options.tabstyle
      ? options.tabstyle
      : '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'
    if (options.columnWidth === undefined) {
      options.columnWidth = (8 / rows[0].length) * EMU
    }
    var self = this

    return self._getBase(
      rows.map(function (row, row_idx) {
        return self._getRow(
          row.map(function (val, idx) {
            var cellVal = val
            var cellOptions = options
            if (typeof val === 'object') {
              // Cell-specific formatting passed in, override table options
              cellOptions = Object.prototype.hasOwnProperty.call(val, 'opts')
                ? val.opts
                : options
              cellVal = Object.prototype.hasOwnProperty.call(val, 'val')
                ? val.val
                : val
            }

            return self._getCell(cellVal, cellOptions, idx, row_idx)
          }),
          options
        )
      }),
      self._getColSpecs(rows, options),
      options
    )
  },

  _getBase: function (rowSpecs, colSpecs, options) {
    return {
      'p:graphicFrame': {
        'p:nvGraphicFramePr': {
          'p:cNvPr': {
            '@id': '6',
            '@name': 'Table 5'
          },
          'p:cNvGraphicFramePr': {
            'a:graphicFrameLocks': {
              '@noGrp': '1'
            }
          },
          'p:nvPr': {
            'p:extLst': {
              'p:ext': {
                '@uri': '{D42A27DB-BD31-4B8C-83A1-F6EECF244321}',
                'p14:modId': {
                  '@xmlns:p14':
                    'http://schemas.microsoft.com/office/powerpoint/2010/main',
                  '@val': '1579011935'
                }
              }
            }
          }
        },
        'p:xfrm': {
          'a:off': {
            '@x': options.x || '1524000',
            '@y': options.y || '1397000'
          },
          'a:ext': {
            '@cx': options.cx || '6096000',
            '@cy': options.cy || '1483360'
          }
        },
        'a:graphic': {
          'a:graphicData': {
            '@uri': 'http://schemas.openxmlformats.org/drawingml/2006/table',
            'a:tbl': {
              'a:tblPr': {
                '@firstRow': '1',
                '@bandRow': '1',
                'a:tableStyleId': options.tabstyle // "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
              },
              'a:tblGrid': {
                '#text': colSpecs
              },
              '#text': rowSpecs // replace this with  an array of table row objects
            }
          }
        }
      }
    }
  },

  _getColSpecs: function (rows, options) {
    var self = this
    return rows[0].map(function (val, idx) {
      return self._tblGrid(idx, options)
    })
  },

  _tblGrid: function (idx, options) {
    return {
      'a:gridCol': {
        '@w': options.columnWidths
          ? options.columnWidths[idx]
          : options.columnWidth || '0' // || "2048000"
      }
    }
  },

  _getRow: function (cells, options) {
    return {
      'a:tr': {
        '@h': options.rowHeight || '0', // || "370840",
        '#text': cells // populate this with an array of table cell objects
      }
    }
  },

  _getCell: function (val, options, idx, row_idx) {
    var font_size = options.font_size || 14
    var font_face = options.font_face || 'Times New Roman'
    var cellObject = {
      'a:tc': {
        'a:txBody': {
          'a:bodyPr': {},
          'a:lstStyle': {},

          'a:p': {
            'a:pPr': {
              '@algn': options.align
                ? options.align[idx]
                  ? options.align[idx]
                  : options.align
                : 'ctr'
            },
            'a:r': {
              'a:rPr': {
                '@lang': 'en-US',
                '@sz': '' + font_size * 100,
                '@dirty': '0',
                '@smtClean': '0',
                '@b': options.bold
                  ? options.bold[row_idx]
                    ? options.bold[row_idx][idx]
                      ? options.bold[row_idx][idx]
                      : options.bold[row_idx]
                    : options.bold
                  : '0',
                '@i': options.italics
                  ? options.italics[row_idx]
                    ? options.italics[row_idx][idx]
                      ? options.italics[row_idx][idx]
                      : options.italics[row_idx]
                    : options.italics
                  : '0',
                'a:latin': {
                  '@typeface': font_face
                },
                'a:cs': {
                  '@typeface': font_face
                }
              },
              'a:t': val // this is the cell value
            },
            'a:endParaRPr': {
              '@lang': 'en-US',
              '@sz': '' + font_size * 100,
              '@dirty': '0',
              'a:latin': {
                '@typeface': font_face
              },
              'a:cs': {
                '@typeface': font_face
              }
            }
          }
        },
        'a:tcPr': {}
      }
    }

    if (Object.prototype.hasOwnProperty.call(options, 'fill_color')) {
      // Apply background fill to table cell
      cellObject['a:tc']['a:tcPr']['a:solidFill'] = {
        'a:srgbClr': {
          '@val': options.fill_color
        }
      }
    }

    if (Object.prototype.hasOwnProperty.call(options, 'font_color')) {
      // Apply color to text run
      cellObject['a:tc']['a:txBody']['a:p']['a:r']['a:rPr']['a:solidFill'] = {
        'a:srgbClr': {
          '@val': options.font_color
        }
      }
    }

    return cellObject
  }
}
