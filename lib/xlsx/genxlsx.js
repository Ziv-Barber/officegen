//
// officegen: All the code to generate XLSX files.
//
// Please refer to README.md for this module's documentations.
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

/**
 * Basicgen plugin to create xlsx files (Microsoft Excel).
 */

var baseobj = require('../core/index.js')
var msdoc = require('../msdoc/msofficegen.js')

var docplugman = require('../core/docplug')

// Officegen xlsx plugins:
// BMK_XLSX_PLUG:

/**
 * Extend officegen object with XLSX support.
 *
 * This method extending the given officegen object to create XLSX document.
 *
 * @param {object} genobj The object to extend.
 * @param {string} new_type The type of object to create.
 * @param {object} options The object's options.
 * @param {object} gen_private Access to the internals of this object.
 * @param {object} type_info Additional information about this type.
 * @constructor
 * @name makeXlsx
 */
function makeXlsx(genobj, new_type, options, gen_private, type_info) {
  /**
   * Prepare the default data.
   * @param {object} docpluginman Access to the document plugins manager.
   */
  function setDefaultDocValues(docpluginman) {
    // var pptxData = docpluginman.getDataStorage()
    // Please put any setting that API can override here:
  }

  /**
   * Create the shared string resource.
   *
   * This resource holding all the text strings of any Excel document.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeXlsSharedStrings(data) {
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
      genobj.generate_data.total_strings +
      '" uniqueCount="' +
      genobj.generate_data.shared_strings.length +
      '">'

    for (
      var i = 0, total_size = genobj.generate_data.shared_strings.length;
      i < total_size;
      i++
    ) {
      outString +=
        '<si><t>' +
        gen_private.plugs.type.msoffice.escapeText(
          genobj.generate_data.shared_strings[i]
        ) +
        '</t></si>'
    } // Endif.

    return outString + '</sst>'
  }

  /**
   * Prepare everything to generate XLSX files.
   *
   * This method working on all the Excel cells to find out information needed by the generator engine.
   */
  function cbPrepareXlsxToGenerate() {
    genobj.generate_data = {}
    genobj.generate_data.shared_strings = []
    genobj.lookup_strings = {}
    genobj.generate_data.total_strings = 0
    genobj.generate_data.cell_strings = []

    // Tell all the features (plugins) that we are about to generate a new document zip:
    gen_private.features.type.xlsx.emitEvent('beforeGen', genobj)

    // Allow some plugins to do more stuff after all the plugins added their data:
    gen_private.features.type.xlsx.emitEvent('beforeGenFinal', genobj)

    // Create the share strings data:
    for (
      var i = 0, total_size = gen_private.pages.length;
      i < total_size;
      i++
    ) {
      if (gen_private.pages[i]) {
        for (
          var rowId = 0, total_size_y = gen_private.pages[i].sheet.data.length;
          rowId < total_size_y;
          rowId++
        ) {
          if (gen_private.pages[i].sheet.data[rowId]) {
            for (
              var columnId = 0,
                total_size_x = gen_private.pages[i].sheet.data[rowId].length;
              columnId < total_size_x;
              columnId++
            ) {
              if (
                typeof gen_private.pages[i].sheet.data[rowId][columnId] !==
                'undefined'
              ) {
                switch (
                  typeof gen_private.pages[i].sheet.data[rowId][columnId]
                ) {
                  case 'string':
                    genobj.generate_data.total_strings++

                    if (!genobj.generate_data.cell_strings[i]) {
                      genobj.generate_data.cell_strings[i] = []
                    } // Endif.

                    if (!genobj.generate_data.cell_strings[i][rowId]) {
                      genobj.generate_data.cell_strings[i][rowId] = []
                    } // Endif.

                    var shared_str =
                      gen_private.pages[i].sheet.data[rowId][columnId]

                    if (shared_str in genobj.lookup_strings) {
                      genobj.generate_data.cell_strings[i][rowId][columnId] =
                        genobj.lookup_strings[shared_str]
                    } else {
                      var shared_str_position =
                        genobj.generate_data.shared_strings.length
                      genobj.generate_data.cell_strings[i][rowId][
                        columnId
                      ] = shared_str_position
                      genobj.lookup_strings[shared_str] = shared_str_position
                      genobj.generate_data.shared_strings[
                        shared_str_position
                      ] = shared_str
                    } // Endif.
                    break
                } // End of switch.
              } // Endif.
            } // End of for loop.
          } // Endif.
        } // End of for loop.
      } // Endif.
    } // End of for loop.

    if (genobj.generate_data.total_strings) {
      gen_private.plugs.intAddAnyResourceToParse(
        'xl\\sharedStrings.xml',
        'buffer',
        null,
        cbMakeXlsSharedStrings,
        false
      )
      gen_private.type.msoffice.files_list.push({
        name: '/xl/sharedStrings.xml',
        type:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
        clear: 'generate'
      })

      gen_private.type.msoffice.rels_app.push({
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
        target: 'sharedStrings.xml',
        clear: 'generate'
      })
    } // Endif.
  }

  /**
   * ???.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeXlsStyles(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf applyAlignment="1" borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'
    )
  }

  /**
   * ???.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeXlsApp(data) {
    var pagesCount = gen_private.pages.length
    var userName =
      genobj.options.author || genobj.options.creator || 'officegen'
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>' +
      pagesCount +
      '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' +
      pagesCount +
      '" baseType="lpstr">'

    for (
      var i = 0, total_size = gen_private.pages.length;
      i < total_size;
      i++
    ) {
      outString += '<vt:lpstr>Sheet' + (i + 1) + '</vt:lpstr>'
    } // End of for loop.

    outString +=
      '</vt:vector></TitlesOfParts><Company>' +
      userName +
      '</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>'
    return outString
  }

  /**
   * ???.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeXlsWorkbook(data) {
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4507"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="19095" windowHeight="7485"/></bookViews><sheets>'

    for (
      var i = 0, total_size = gen_private.pages.length;
      i < total_size;
      i++
    ) {
      var sheetName = gen_private.pages[i].sheet.name || 'Sheet' + (i + 1)
      var rId = gen_private.pages[i].relId
      outString +=
        '<sheet name="' +
        sheetName +
        '" sheetId="' +
        (i + 1) +
        '" r:id="rId' +
        rId +
        '"/>'
    } // End of for loop.

    outString += '</sheets><calcPr calcId="125725"/></workbook>'
    return outString
  }

  /**
   * Translate from the Excel displayed row name into index number.
   *
   * @param {string} cell_string Either the cell displayed position or the row displayed position.
   * @param {boolean} ret_also_column ???.
   * @return The cell's row Id.
   */
  function cbCellToNumber(cell_string, ret_also_column) {
    var cellNumber = 0
    var cellIndex = 0
    var cellMax = cell_string.length
    var rowId = 0

    // Converted from C++ (from DuckWriteC++):
    while (cellIndex < cellMax) {
      var curChar = cell_string.charCodeAt(cellIndex)
      if (curChar >= 0x30 && curChar <= 0x39) {
        rowId = parseInt(cell_string.slice(cellIndex), 10)
        rowId = rowId > 0 ? rowId - 1 : 0
        break
      } else if (curChar >= 0x41 && curChar <= 0x5a) {
        if (cellIndex > 0) {
          cellNumber++
          cellNumber *= 0x5b - 0x41
        } // Endif.

        cellNumber += curChar - 0x41
      } else if (curChar >= 0x61 && curChar <= 0x7a) {
        if (cellIndex > 0) {
          cellNumber++
          cellNumber *= 0x5b - 0x41
        } // Endif.

        cellNumber += curChar - 0x61
      } // Endif.

      cellIndex++
    } // End of while loop.

    if (ret_also_column) {
      return { row: rowId, column: cellNumber }
    } // Endif.

    return cellNumber
  }

  /**
   * ???.
   *
   * @param {object} cell_number ???.
   * @return ???.
   */
  function cbNumberToCell(cell_number) {
    var outCell = ''
    var curCell = cell_number

    while (curCell >= 0) {
      outCell = String.fromCharCode((curCell % (0x5b - 0x41)) + 0x41) + outCell
      if (curCell >= 0x5b - 0x41) {
        curCell = Math.floor(curCell / (0x5b - 0x41)) - 1
      } else {
        break
      }
    } // End of while loop.

    return outCell
  }

  /**
   * ???.
   *
   * @param {object} data The main sheet object.
   * @return Text string.
   */
  function cbMakeXlsSheet(data) {
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    var maxX = 0
    var maxY = 0
    var curColMax
    var rowId
    var columnId
    var colsFound = 0
    var total_size_y
    var total_size_x

    // Find the maximum cells area:
    maxY = data.sheet.data.length ? data.sheet.data.length - 1 : 0
    for (
      rowId = 0, total_size_y = data.sheet.data.length;
      rowId < total_size_y;
      rowId++
    ) {
      if (data.sheet.data[rowId]) {
        curColMax = data.sheet.data[rowId].length
          ? data.sheet.data[rowId].length - 1
          : 0
        maxX = maxX < curColMax ? curColMax : maxX
      } // Endif.
    } // End of for loop.

    outString +=
      '<dimension ref="A1:' +
      cbNumberToCell(maxX) +
      '' +
      (maxY + 1) +
      '"/><sheetViews>'
    outString += '<sheetView tabSelected="1" workbookViewId="0"/>'
    // outString += '<selection activeCell="A1" sqref="A1"/>'
    outString += '</sheetViews><sheetFormatPr defaultRowHeight="15"/>'

    if (data.sheet.width) {
      data.sheet.width.forEach(function (value, indexId) {
        if (typeof value === 'object') {
          var outAttr = ''

          /* eslint-disable no-self-compare */
          if (typeof value.width === 'number' && value.width === value.width) {
            outAttr = ' width="' + value.width + '" customWidth="1"'
          } // Endif.

          /* eslint-disable no-self-compare */
          if (
            typeof value.styleCode === 'number' &&
            value.styleCode === value.styleCode
          ) {
            outAttr = ' style="' + value.styleCode + '"'
          } // Endif.

          if (!colsFound) {
            outString += '<cols>'
          } // Endif.

          outString +=
            '<col min="' +
            (value.colId + 1) +
            '" max="' +
            (value.colId + 1) +
            '"' +
            outAttr +
            '/>'
          colsFound++

          // Support for old code, not recommended:
        } else if (typeof value === 'number') {
          if (!colsFound) {
            outString += '<cols>'
          } // Endif.

          outString +=
            '<col min="' +
            (indexId + 1) +
            '" max="' +
            (indexId + 1) +
            '" width="' +
            value +
            '" customWidth="1"/>'
          colsFound++
        } // Endif.
      })

      if (colsFound) {
        outString += '</cols>'
      } // Endif.
    } // Endif.

    outString += '<sheetData>'

    for (
      rowId = 0, total_size_y = data.sheet.data.length;
      rowId < total_size_y;
      rowId++
    ) {
      if (data.sheet.data[rowId]) {
        // Patch by arnesten <notifications@github.com>: Automatically support line breaks if used in cell + calculates row height:
        var rowLines = 1
        data.sheet.data[rowId].forEach(function (cellData) {
          if (typeof cellData === 'string') {
            var candidate = cellData.split('\n').length
            rowLines = Math.max(rowLines, candidate)
          }
        })
        outString +=
          '<row r="' +
          (rowId + 1) +
          '" spans="1:' +
          data.sheet.data[rowId].length +
          '" ht="' +
          rowLines * 15 +
          '">'
        // End of patch.

        for (
          columnId = 0, total_size_x = data.sheet.data[rowId].length;
          columnId < total_size_x;
          columnId++
        ) {
          var cellData = data.sheet.data[rowId][columnId]
          if (typeof cellData !== 'undefined') {
            var isString = ''
            var cellOutData = '0'

            switch (typeof cellData) {
              case 'number':
                cellOutData = cellData
                break

              case 'string':
                cellOutData =
                  genobj.generate_data.cell_strings[data.id][rowId][columnId]
                if (cellData.indexOf('\n') >= 0) {
                  isString = ' s="1" t="s"'
                } else {
                  isString = ' t="s"'
                }
                break
            } // End of switch.

            outString +=
              '<c r="' +
              cbNumberToCell(columnId) +
              '' +
              (rowId + 1) +
              '"' +
              isString +
              '><v>' +
              cellOutData +
              '</v></c>'
          } // Endif.
        } // End of for loop.

        outString += '</row>'
      } // Endif.
    } // End of for loop.

    outString += '</sheetData>'
    if (data.sheet.mergeCells && data.sheet.mergeCells.length > 0) {
      outString += '<mergeCells count="'+data.sheet.mergeCells.length+'">'
      for (var mergeIndex = 0; mergeIndex < data.sheet.mergeCells.length; mergeIndex++) {
        outString += '<mergeCell ref="'+data.sheet.mergeCells[mergeIndex]+'"/>'
      }
      outString += '</mergeCells>'
    }
    outString += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'
    return outString
  }

  // Prepare genobj for MS-Office:
  msdoc.makemsdoc(genobj, new_type, options, gen_private, type_info)
  gen_private.plugs.type.msoffice.makeOfficeGenerator('xl', 'workbook', {})

  gen_private.features.page_name = 'sheets' // This document type must have pages.

  // On each generate we'll prepare the shared strings list:
  genobj.on('beforeGen', cbPrepareXlsxToGenerate)

  // Create the plugins manager:
  var plugsmanObj = new docplugman(
    genobj,
    gen_private,
    'xlsx',
    setDefaultDocValues
  )

  // We'll register now any officegen internal plugin that we want to always use for Excel based documents:
  // BMK_XLSX_PLUG:

  // Dynamic loading of additional plugins requested by the user:
  if (
    options.extraPlugs &&
    typeof options.extraPlugs === 'object' &&
    options.extraPlugs.forEach
  ) {
    options.extraPlugs.forEach(function (value) {
      var newPlug

      if (value) {
        if (typeof value === 'function') {
          // You already loaded the plugin:
          newPlug = value
        } else if (typeof value === 'string') {
          // We need to load the plugin:
          newPlug = require('./' + value)
        } // Endif.
      } // Endif.

      plugsmanObj.plugsList.push(new newPlug(plugsmanObj))
    })
  } // Endif.

  gen_private.type.msoffice.files_list.push(
    {
      name: '/xl/styles.xml',
      type:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
      clear: 'type'
    },
    {
      name: '/xl/workbook.xml',
      type:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      clear: 'type'
    }
  )

  gen_private.type.msoffice.rels_app.push(
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      target: 'styles.xml',
      clear: 'type'
    },
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: 'theme/theme1.xml',
      clear: 'type'
    }
  )

  gen_private.plugs.intAddAnyResourceToParse(
    'docProps\\app.xml',
    'buffer',
    null,
    cbMakeXlsApp,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'xl\\styles.xml',
    'buffer',
    null,
    cbMakeXlsStyles,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'xl\\workbook.xml',
    'buffer',
    null,
    cbMakeXlsWorkbook,
    true
  )

  gen_private.plugs.intAddAnyResourceToParse(
    'xl\\_rels\\workbook.xml.rels',
    'buffer',
    gen_private.type.msoffice.rels_app,
    gen_private.plugs.type.msoffice.cbMakeRels,
    true
  )

  // ----- API for Excel documents: -----

  /**
   * Create a new sheet.
   *
   * This method creating a new Excel sheet.
   */
  genobj.makeNewSheet = function () {
    var pageNumber = gen_private.pages.length

    // The sheet object that the user will use:
    var sheetObj = {
      data: [], // Place here all the data.
      width: []
    }

    gen_private.pages[pageNumber] = {}
    gen_private.pages[pageNumber].id = pageNumber
    gen_private.pages[pageNumber].relId =
      gen_private.type.msoffice.rels_app.length + 1
    gen_private.pages[pageNumber].sheet = sheetObj

    gen_private.type.msoffice.rels_app.push({
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      target: 'worksheets/sheet' + (pageNumber + 1) + '.xml',
      clear: 'data'
    })

    gen_private.type.msoffice.files_list.push({
      name: '/xl/worksheets/sheet' + (pageNumber + 1) + '.xml',
      type:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      clear: 'data'
    })

    sheetObj.setColumnWidth = function (colId, width) {
      var colRec = null

      colId = cbCellToNumber(colId + '1', false)

      sheetObj.width.every(function (value) {
        if (value.colId === colId) {
          colRec = value
          return false
        } // Endif.

        return true
      })

      if (!colRec) {
        sheetObj.width.push({
          colId: colId,
          width: width
        })
        return
      } // Endif.

      colRec.width = width
    }

    sheetObj.setColumnCenter = function (colId) {
      var colRec = null

      colId = cbCellToNumber(colId + '1', false)

      sheetObj.width.every(function (value) {
        if (value.colId === colId) {
          colRec = value
          return false
        } // Endif.

        return true
      })

      if (!colRec) {
        sheetObj.width.push({
          colId: colId,
          width: 9.140625,
          styleCode: 1
        })
        return
      } // Endif.

      colRec.styleCode = 1
    }

    sheetObj.setCell = function (position, data_val) {
      var rel_pos = cbCellToNumber(position, true)

      if (!sheetObj.data[rel_pos.row]) {
        sheetObj.data[rel_pos.row] = []
      } // Endif.

      sheetObj.data[rel_pos.row][rel_pos.column] = data_val
    }

    gen_private.plugs.intAddAnyResourceToParse(
      'xl\\worksheets\\sheet' + (pageNumber + 1) + '.xml',
      'buffer',
      gen_private.pages[pageNumber],
      cbMakeXlsSheet,
      false
    )

    // Signal to the plugins about a new sheet:
    gen_private.features.type.xlsx.emitEvent('newPage', {
      genobj: genobj,
      page: sheetObj,
      pageData: gen_private.pages[pageNumber],
      pageNumber: pageNumber
    })

    return sheetObj
  }

  // Tell all the features (plugins) to add extra API:
  gen_private.features.type.xlsx.emitEvent('makeDocApi', genobj)

  return this
}

baseobj.plugins.registerDocType(
  'xlsx',
  makeXlsx,
  {},
  baseobj.docType.SPREADSHEET,
  'Microsoft Excel Document'
)
