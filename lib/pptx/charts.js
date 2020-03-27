//
// officegen: pptx charts
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
  null: function (options) {
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:date1904': { '@val': '1' },
        'c:chart': {}
      }
    }
  },

  bar: function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:chart': {
          'c:plotArea': {
            'c:layout': {},
            'c:barChart': {
              'c:barDir': { '@val': 'bar' },
              'c:grouping': { '@val': 'clustered' },
              '#text': [
                { 'c:axId': { '@val': '64451712' } },
                { 'c:axId': { '@val': '64453248' } }
              ]
            },
            'c:catAx': {
              'c:axId': { '@val': '64451712' },
              'c:scaling': {
                'c:orientation': {
                  '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                }
              },
              'c:axPos': { '@val': 'l' },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64453248' },
              'c:crosses': { '@val': 'autoZero' },
              'c:auto': { '@val': '1' },
              'c:lblAlgn': { '@val': 'ctr' },
              'c:lblOffset': { '@val': '100' }
            },
            'c:valAx': {
              'c:axId': { '@val': '64453248' },
              'c:scaling': {
                'c:orientation': { '@val': 'minMax' }
              },
              'c:axPos': { '@val': 'b' },
              //              "c:majorGridlines": {},
              'c:numFmt': {
                '@formatCode': 'General',
                '@sourceLinked': '1'
              },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64451712' },
              'c:crosses': {
                '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
              },
              'c:crossBetween': { '@val': 'between' }
            }
          },
          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  column: function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:date1904': { '@val': '1' },
        'c:chart': {
          'c:plotArea': {
            'c:layout': {},
            'c:barChart': {
              'c:barDir': { '@val': 'col' },
              'c:grouping': { '@val': 'clustered' },
              'c:overlap': { '@val': options.overlap || '0' },
              'c:gapWidth': { '@val': options.gapWidth || '150' },

              '#text': [
                { 'c:axId': { '@val': '64451712' } },
                { 'c:axId': { '@val': '64453248' } }
              ]
            },
            'c:catAx': {
              'c:axId': { '@val': '64451712' },
              'c:scaling': {
                'c:orientation': {
                  '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                }
              },
              'c:axPos': { '@val': 'l' },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64453248' },
              'c:crosses': { '@val': 'autoZero' },
              'c:auto': { '@val': '1' },
              'c:lblAlgn': { '@val': 'ctr' },
              'c:lblOffset': { '@val': '100' }
            },
            'c:valAx': {
              'c:axId': { '@val': '64453248' },
              'c:scaling': {
                'c:orientation': { '@val': 'minMax' }
              },
              'c:axPos': { '@val': 'b' },
              //              "c:majorGridlines": {},
              'c:numFmt': {
                '@formatCode': 'General',
                '@sourceLinked': '1'
              },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64451712' },
              'c:crosses': {
                '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
              },
              'c:crossBetween': { '@val': 'between' }
            }
          },
          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  'stacked-column': function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:date1904': { '@val': '1' },
        'c:chart': {
          'c:plotArea': {
            'c:layout': {},
            'c:barChart': {
              'c:barDir': { '@val': 'col' },
              'c:grouping': { '@val': 'stacked' },
              'c:overlap': { '@val': options.overlap || '100' },
              'c:gapWidth': { '@val': options.gapWidth || '25' },

              '#text': [
                { 'c:axId': { '@val': '64451712' } },
                { 'c:axId': { '@val': '64453248' } }
              ]
            },
            'c:catAx': {
              'c:axId': { '@val': '64451712' },
              'c:scaling': {
                'c:orientation': {
                  '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                }
              },
              'c:axPos': { '@val': 'l' },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64453248' },
              'c:crosses': { '@val': 'autoZero' },
              'c:auto': { '@val': '1' },
              'c:lblAlgn': { '@val': 'ctr' },
              'c:lblOffset': { '@val': '100' }
            },
            'c:valAx': {
              'c:axId': { '@val': '64453248' },
              'c:scaling': {
                'c:orientation': { '@val': 'minMax' }
              },
              'c:axPos': { '@val': 'b' },
              //              "c:majorGridlines": {},
              'c:numFmt': {
                '@formatCode': 'General',
                '@sourceLinked': '1'
              },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64451712' },
              'c:crosses': {
                '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
              },
              'c:crossBetween': { '@val': 'between' }
            }
          },
          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  'group-bar': function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:date1904': { '@val': '1' },
        'c:chart': {
          'c:plotArea': {
            'c:layout': {},
            'c:barChart': {
              'c:barDir': { '@val': 'bar' },
              'c:grouping': { '@val': 'stacked' },
              'c:overlap': { '@val': options.overlap || '100' },
              'c:gapWidth': { '@val': options.gapWidth || '150' },
              '#text': [
                { 'c:axId': { '@val': '64451712' } },
                { 'c:axId': { '@val': '64453248' } }
              ]
            },
            'c:catAx': {
              'c:axId': { '@val': '64451712' },
              'c:scaling': {
                'c:orientation': {
                  '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                }
              },
              'c:axPos': { '@val': 'l' },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64453248' },
              'c:crosses': { '@val': 'autoZero' },
              'c:auto': { '@val': '1' },
              'c:lblAlgn': { '@val': 'ctr' },
              'c:lblOffset': { '@val': '100' }
            },
            'c:valAx': {
              'c:axId': { '@val': '64453248' },
              'c:scaling': {
                'c:orientation': { '@val': 'minMax' }
              },
              'c:axPos': { '@val': 'b' },
              //              "c:majorGridlines": {},
              'c:numFmt': {
                '@formatCode': 'General',
                '@sourceLinked': '1'
              },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64451712' },
              'c:crosses': {
                '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
              },
              'c:crossBetween': { '@val': 'between' }
            }
          },
          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  pie: function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:chart': {
          'c:title': {
            'c:layout': {}
          },
          'c:plotArea': {
            'c:layout': {},
            'c:pieChart': {
              'c:varyColors': { '@val': '1' },
              'c:firstSliceAng': { '@val': '0' },
              '#text': []
            }
          },

          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  line: function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:chart': {
          'c:plotArea': {
            'c:layout': {},
            'c:lineChart': {
              'c:grouping': { '@val': 'standard' },
              '#text': [
                { 'c:axId': { '@val': '64451712' } },
                { 'c:axId': { '@val': '64453248' } }
              ]
            },
            'c:catAx': {
              'c:axId': { '@val': '64451712' },
              'c:scaling': {
                'c:orientation': {
                  '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                }
              },
              'c:axPos': { '@val': 'l' },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64453248' },
              'c:crosses': { '@val': 'autoZero' },
              'c:auto': { '@val': '1' },
              'c:lblAlgn': { '@val': 'ctr' },
              'c:lblOffset': { '@val': '100' }
            },
            'c:valAx': {
              'c:axId': { '@val': '64453248' },
              'c:scaling': {
                'c:orientation': { '@val': 'minMax' }
              },
              'c:axPos': { '@val': 'b' },
              //              "c:majorGridlines": {},
              'c:numFmt': {
                '@formatCode': 'General',
                '@sourceLinked': '1'
              },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64451712' },
              'c:crosses': {
                '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
              },
              'c:crossBetween': { '@val': 'between' }
            }
          },
          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  area: function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:chart': {
          'c:plotArea': {
            'c:layout': {},
            'c:areaChart': {
              'c:grouping': { '@val': 'standard' },
              '#text': [
                { 'c:axId': { '@val': '64451712' } },
                { 'c:axId': { '@val': '64453248' } }
              ]
            },
            'c:catAx': {
              'c:axId': { '@val': '64451712' },
              'c:scaling': {
                'c:orientation': {
                  '@val': options.catAxisReverseOrder ? 'maxMin' : 'minMax'
                }
              },
              'c:axPos': { '@val': 'l' },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64453248' },
              'c:crosses': { '@val': 'autoZero' },
              'c:auto': { '@val': '1' },
              'c:lblAlgn': { '@val': 'ctr' },
              'c:lblOffset': { '@val': '100' }
            },
            'c:valAx': {
              'c:axId': { '@val': '64453248' },
              'c:scaling': {
                'c:orientation': { '@val': 'minMax' }
              },
              'c:axPos': { '@val': 'b' },
              //              "c:majorGridlines": {},
              'c:numFmt': {
                '@formatCode': 'General',
                '@sourceLinked': '1'
              },
              'c:tickLblPos': { '@val': 'nextTo' },
              'c:crossAx': { '@val': '64451712' },
              'c:crosses': {
                '@val': options.valAxisCrossAtMaxCategory ? 'max' : 'autoZero'
              },
              'c:crossBetween': { '@val': 'between' }
            }
          },
          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  },
  doughnut: function (options) {
    options = options || {}
    return {
      'c:chartSpace': {
        '@xmlns:c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        '@xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@xmlns:r':
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'c:lang': { '@val': 'en-US' },
        'c:chart': {
          'c:title': {
            'c:layout': {}
          },
          'c:plotArea': {
            'c:layout': {},
            'c:doughnutChart': {
              'c:varyColors': { '@val': '1' },
              'c:firstSliceAng': { '@val': '0' },
              'c:holeSize': { '@val': '75' },
              '#text': []
            }
          },

          'c:legend': {
            'c:legendPos': { '@val': 'r' },
            'c:layout': {}
          },
          'c:plotVisOnly': { '@val': '1' }
        },
        'c:txPr': {
          'a:bodyPr': {},
          'a:lstStyle': {},
          'a:p': {
            'a:pPr': {
              'a:defRPr': { '@sz': '1800' }
            },
            'a:endParaRPr': { '@lang': 'en-US' }
          }
        },
        'c:externalData': { '@r:id': 'rId1' }
      }
    }
  }
}
