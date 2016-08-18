
module.exports = {

  // assume passed in an array of row objects
  getTable: function(rows, tblOpts) {
    var tblOpts = tblOpts || {};

    var self = this;

    return self._getBase(
        rows.map(function(row) {
          return self._getRow(
              row.map(function(cell) {
                return self._getCell((cell.val || cell), cell.opts, tblOpts);
              }),
              tblOpts
          );
        }),
        self._getColSpecs(rows, tblOpts),
        tblOpts
    )
  },

  _getBase: function (rowSpecs, colSpecs, opts) {
    var self = this;

    return {
      "w:tbl": {
        "w:tblPr": {
          "w:tblStyle": {
            "@w:val": "a3"
          },
          "w:tblW": {
            "@w:w": "0",
            "@w:type": "auto"
          },
          "w:tblLook": {
            "@w:val": "04A0",
            "@w:firstRow": "1",
            "@w:lastRow": "0",
            "@w:firstColumn": "1",
            "@w:lastColumn": "0",
            "@w:noHBand": "0",
            "@w:noVBand": "1"
          }
        },
        "w:tblGrid": {
          "#list": colSpecs
        },
        "#list": [rowSpecs]
      }
    }
  },

  _getColSpecs: function(cols, opts) {
    var self = this;
    return cols[0].map(function(val,idx) {
      return self._tblGrid(opts);
    })
  },

  // TODO 
  _tblGrid: function(opts) {
    return {
      "w:gridCol": {
        "@w:w": opts.tableColWidth || "0"
      }
    };
  },


  _getRow: function (cells, opts) {
    return {
      "w:tr": {
        "@w:rsidR": "00995B51",
        "@w:rsidTr": "007F1D13",
        "#list": [cells] // populate this with an array of table cell objects
      }
    }
  },

  _getCell: function (val, opts, tblOpts) {
    opts = opts || {};
    // var b = {};
    var content =[{"w:p": {
      "@w:rsidR": "00995B51",
      "@w:rsidRPr": "00722E63",
      "@w:rsidRDefault": "00995B51",
      "w:pPr": {
        "w:jc": {
          "@w:val": opts.align || tblOpts.tableAlign || "center"
        },
      },
      "w:r": {
        "@w:rsidRPr": "00722E63",
        "w:rPr": {
          "w:rFonts": {
            "@w:ascii": opts.fontFamily || tblOpts.tableFontFamily || "宋体",
            "@w:hAnsi": opts.fontFamily || tblOpts.tableFontFamily || "宋体"
          },
          "w:color": {
            "@w:val": opts.color || tblOpts.tableColor || "000"
          },
          "w:b": {},
          "w:sz": {
            "@w:val": opts.sz || tblOpts.sz || "24"
          },
          "w:szCs": {
            "@w:val": opts.sz || tblOpts.sz || "24"
          }
        },
        "w:t": val
      }
    }}]
    if(typeof val != 'string' && typeof val != 'number') {
      if (Array.isArray(val)) {
        content = this._getComplexCell(val, tblOpts);
      }
    } else {
      if (!opts.b) {
        delete content[0]["w:p"]["w:r"]["w:rPr"]["w:b"];
      }
    }
    var cellObj = {
      "w:tc": {
        "w:tcPr": {
          "w:tcW": {
            "@w:w": opts.cellColWidth || tblOpts.tableColWidth || "0",
            "@w:type": "dxa"
          },
          "w:shd": {
            "@w:val": "clear",
            "@w:color": "auto",
            "@w:fill":  opts.shd && opts.shd.fill || "",
            "@w:themeFill": opts.shd && opts.shd.themeFill || "",
            "@w:themeFillTint": opts.shd && opts.shd.themeFillTint || ""
          }
        },
        "#list": [content]

      }
    }
    return cellObj;
  },
  _getComplexCell: function (cellVal, tblOpts) {
    var content = cellVal.map(function(cell) {
      /*
       * TODO: add more options. like adding tables , pictures etc
       */
      if(cell.type === 'text' || cell.type === 'number') {
        var opts = cell.opts || {};
        var rowObject = [];
        if(cell.inline) {
          cell.values.forEach(function(row) {
            var tempRow = {
              "w:r": {
                "@w:rsidRPr": "00722E63",
                "w:rPr": {
                  "w:rFonts": {
                    "@w:ascii": row.opts.fontFamily || tblOpts.tableFontFamily || "宋体",
                    "@w:hAnsi": row.opts.fontFamily || tblOpts.tableFontFamily || "宋体"
                  },
                  "w:color": {
                    "@w:val": row.opts.color || tblOpts.tableColor || "000"
                  },
                  "w:b": {},
                  "w:sz": {
                    "@w:val": row.opts.sz || tblOpts.sz || "16"
                  },
                  "w:szCs": {
                    "@w:val": row.opts.sz || tblOpts.sz || "16"
                  }
                },
                "w:t": row.val
              }
            }
            if(!row.opts.b) {
              delete tempRow["w:r"]["w:rPr"]["w:b"];
            }
            rowObject.push( tempRow);
          })
        } else {
          tempRow =  {
            "w:r": {
              "@w:rsidRPr": "00722E63",
              "w:rPr": {
                "w:rFonts": {
                  "@w:ascii": opts.fontFamily || tblOpts.tableFontFamily || "宋体",
                  "@w:hAnsi": opts.fontFamily || tblOpts.tableFontFamily || "宋体"
                },
                "w:color": {
                  "@w:val": opts.color || tblOpts.tableColor || "000"
                },
                "w:b": {},
                "w:sz": {
                  "@w:val": opts.sz || tblOpts.sz || "16"
                },
                "w:szCs": {
                  "@w:val": opts.sz || tblOpts.sz || "16"
                }
              },
              "w:t": cell.val
            }
          }
          if(!opts.b) {
            delete tempRow["w:r"]["w:rPr"]["w:b"];
          }
          rowObject.push(tempRow);
        }
        return [
          {"w:p": {
            "@w:rsidR": "00995B51",
            "@w:rsidRPr": "00722E63",
            "@w:rsidRDefault": "00995B51",
            "w:pPr": {
              "w:jc": {
                "@w:val": opts.align || tblOpts.tableAlign || "center"
              },
            },
            "#list": [rowObject]
          }}];
      }
    })
    return content;
  }
}