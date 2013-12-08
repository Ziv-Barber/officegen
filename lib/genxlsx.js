//
// officegen: All the code to generate XLSX files.
//
// Please refer to README.md for this module's documentations.
//
// NOTE:
// - Before changing this code please refer to the hacking the code section on README.md.
//
// Copyright (c) 2013 Ziv Barber;
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// 'Software'), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
// IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
// CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
// TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

var baseobj = require("./basicgen.js");
var msdoc = require("./msofficegen.js");

if ( !String.prototype.encodeHTML ) {
	String.prototype.encodeHTML = function () {
		return this.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;');
	};
}

///
/// @brief Extend officegen object with XLSX support.
///
/// This method extending the given officegen object to create XLSX document.
///
/// @param[in] genobj The object to extend.
/// @param[in] new_type The type of object to create.
/// @param[in] options The object's options.
/// @param[in] gen_private Access to the internals of this object.
/// @param[in] type_info Additional information about this type.
///
function makeXlsx ( genobj, new_type, options, gen_private, type_info ) {
	///
	/// @brief Create the shared string resource.
	///
	/// This resource holding all the text strings of any Excel document.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeXlsSharedStrings ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + genobj.generate_data.total_strings + '" uniqueCount="' + genobj.generate_data.shared_strings.length + '">';

		for ( var i = 0, total_size = genobj.generate_data.shared_strings.length; i < total_size; i++ ) {
			outString += '<si><t>' + genobj.generate_data.shared_strings[i].encodeHTML () + '</t></si>';
		} // Endif.

		return outString + '</sst>';
	}

	///
	/// @brief Prepare everything to generate XLSX files.
	///
	/// This method working on all the Excel cells to find out information needed by the generator engine.
	///
	function cbPrepareXlsxToGenerate () {
		genobj.generate_data = {};
		genobj.generate_data.shared_strings = [];
		genobj.lookup_strings = {};
		genobj.generate_data.total_strings = 0;
		genobj.generate_data.cell_strings = [];

		// Create the share strings data:
		for ( var i = 0, total_size = gen_private.pages.length; i < total_size; i++ ) {
			if ( gen_private.pages[i] ) {
				for ( var rowId = 0, total_size_y = gen_private.pages[i].sheet.data.length; rowId < total_size_y; rowId++ ) {
					if ( gen_private.pages[i].sheet.data[rowId] ) {
						for ( var columnId = 0, total_size_x = gen_private.pages[i].sheet.data[rowId].length; columnId < total_size_x; columnId++ ) {
							if ( typeof gen_private.pages[i].sheet.data[rowId][columnId] != 'undefined' ) {
								switch ( typeof gen_private.pages[i].sheet.data[rowId][columnId] ) {
									case 'string':
										genobj.generate_data.total_strings++;

										if ( !genobj.generate_data.cell_strings[i] ) {
											genobj.generate_data.cell_strings[i] = [];
										} // Endif.

										if ( !genobj.generate_data.cell_strings[i][rowId] ) {
											genobj.generate_data.cell_strings[i][rowId] = [];
										} // Endif.

										var shared_str = gen_private.pages[i].sheet.data[rowId][columnId];

										if (shared_str in genobj.lookup_strings) {
											genobj.generate_data.cell_strings[i][rowId][columnId] = genobj.lookup_strings[ shared_str ];

										} else {
											var shared_str_position = genobj.generate_data.shared_strings.length;
											genobj.generate_data.cell_strings[i][rowId][columnId] = shared_str_position;
											genobj.lookup_strings[ shared_str ] = shared_str_position;
											genobj.generate_data.shared_strings[shared_str_position] = shared_str;
										} // Endif.
										break;
								} // End of switch.
							} // Endif.
						} // End of for loop.
					} // Endif.
				} // End of for loop.
			} // Endif.
		} // End of for loop.

		if ( genobj.generate_data.total_strings ) {
			gen_private.plugs.intAddAnyResourceToParse ( 'xl\\sharedStrings.xml', 'buffer', null, cbMakeXlsSharedStrings, false );
			gen_private.type.msoffice.files_list.push (
				{
					name: '/xl/sharedStrings.xml',
					type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
					clear: 'generate'
				}
			);

			gen_private.type.msoffice.rels_app.push (
				{
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
					target: 'sharedStrings.xml',
					clear: 'generate'
				}
			);

			// console.log ( genobj.generate_data.total_strings );
			// console.log ( genobj.generate_data.shared_strings.length );
		} // Endif.
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeXlsStyles ( data ) {
		return gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf applyAlignment="1" borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeXlsApp ( data ) {
		var pagesCount = gen_private.pages.length;
		var userName = genobj.options.creator || 'officegen';
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>' + pagesCount + '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' + pagesCount + '" baseType="lpstr">';

		for ( var i = 0, total_size = gen_private.pages.length; i < total_size; i++ ) {
			outString += '<vt:lpstr>Sheet' + (i + 1) + '</vt:lpstr>';
		} // End of for loop.

		outString += '</vt:vector></TitlesOfParts><Company>' + userName + '</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>';
		return outString;
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeXlsWorkbook ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4507"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="19095" windowHeight="7485"/></bookViews><sheets>';

		for ( var i = 0, total_size = gen_private.pages.length; i < total_size; i++ ) {
			var sheetName = gen_private.pages[i].sheet.name || 'Sheet' + (i + 1);
                        var rId = gen_private.pages[i].relId;
			outString += '<sheet name="' + sheetName + '" sheetId="' + (i + 1) + '" r:id="rId' + rId + '"/>';
		} // End of for loop.

		outString += '</sheets><calcPr calcId="125725"/></workbook>';
		return outString;
	}

	///
	/// @brief Translate from the Excel displayed row name into index number.
	///
	/// ???.
	///
	/// @param[in] cell_string Either the cell displayed position or the row displayed position.
	/// @return The cell's row Id.
	///
	function cbCellToNumber ( cell_string, ret_also_column ) {
		var cellNumber = 0;
		var cellIndex = 0;
		var cellMax = cell_string.length;
		var rowId = 0;

		// Converted from C++ (from DuckWriteC++):
		while ( cellIndex < cellMax )
		{
			var curChar = cell_string.charCodeAt ( cellIndex );
			if ( (curChar >= 0x30) && (curChar <= 0x39) )
			{
				rowId = parseInt ( cell_string.slice ( cellIndex ), 10 );
				rowId = (rowId > 0) ? (rowId - 1) : 0;
				break;

			} else if ( (curChar >= 0x41) && (curChar <= 0x5A) )
			{
				if ( cellIndex > 0 )
				{
					cellNumber++;
					cellNumber *= (0x5B-0x41);
				} // Endif.

				cellNumber += (curChar - 0x41);

			} else if ( (curChar >= 0x61) && (curChar <= 0x7A) )
			{
				if ( cellIndex > 0 )
				{
					cellNumber++;
					cellNumber *= (0x5B-0x41);
				} // Endif.

				cellNumber += (curChar - 0x61);
			} // Endif.

			cellIndex++;
		} // End of while loop.

		if ( ret_also_column ) {
			return { row: rowId, column: cellNumber };
		} // Endif.

		return cellNumber;
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] cell_number ???.
	/// @return ???.
	///
	function cbNumberToCell ( cell_number ) {
		var outCell = '';
		var curCell = cell_number;

		while ( curCell >= 0 )
		{
			outCell = String.fromCharCode ( (curCell % (0x5B-0x41)) + 0x41 ) + outCell;
			if ( curCell >= (0x5B-0x41) )
				curCell = Math.floor ( curCell / (0x5B-0x41) ) - 1;
			else
				break;
		} // End of while loop.

		return outCell;
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data The main sheet object.
	/// @return Text string.
	///
	function cbMakeXlsSheet ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		var maxX = 0;
		var maxY = 0;
		var curColMax;
		var rowId;
		var columnId;

		// Find the maximum cells area:
		maxY = data.sheet.data.length ? (data.sheet.data.length - 1) : 0;
		for ( var rowId = 0, total_size_y = data.sheet.data.length; rowId < total_size_y; rowId++ ) {
			if ( data.sheet.data[rowId] ) {
				curColMax = data.sheet.data[rowId].length ? (data.sheet.data[rowId].length - 1) : 0;
				maxX = maxX < curColMax ? curColMax : maxX;
			} // Endif.
		} // End of for loop.

		outString += '<dimension ref="A1:' + cbNumberToCell ( maxX ) + '' + (maxY + 1) + '"/><sheetViews>';
		outString += '<sheetView tabSelected="1" workbookViewId="0"/>';
		// outString += '<selection activeCell="A1" sqref="A1"/>';
		outString += '</sheetViews><sheetFormatPr defaultRowHeight="15"/>';

		// BMK_TODO: <cols><col min="2" max="2" width="19" customWidth="1"/></cols>

		outString += '<sheetData>';

		for ( var rowId = 0, total_size_y = data.sheet.data.length; rowId < total_size_y; rowId++ ) {
			if ( data.sheet.data[rowId] ) {

                var rowLines = 1;
                data.sheet.data[rowId].forEach(function (cellData) {
                    if (typeof cellData === 'string') {
                        var candidate = cellData.split('\n').length;
                        rowLines = Math.max(rowLines, candidate);
                    }
                });
				outString += '<row r="' + (rowId + 1) + '" spans="1:' + (data.sheet.data[rowId].length) + '" ht="' + ( rowLines * 15 ) + '">';


				for ( var columnId = 0, total_size_x = data.sheet.data[rowId].length; columnId < total_size_x; columnId++ ) {
                    var cellData = data.sheet.data[rowId][columnId];
                    if ( typeof  cellData != 'undefined' ) {
						var isString = '';
						var cellOutData = '0';

						switch ( typeof cellData ) {
							case 'number':
								cellOutData = cellData;
								break;

							case 'string':
								cellOutData = genobj.generate_data.cell_strings[data.id][rowId][columnId];
                                if (cellData.indexOf('\n') >= 0) {
                                    isString = ' s="1" t="s"';
                                }
                                else {
                                    isString = ' t="s"';
                                }
								break;
						} // End of switch.

						outString += '<c r="' + cbNumberToCell ( columnId ) + '' + (rowId + 1) + '"' + isString + '><v>' + cellOutData + '</v></c>';
					} // Endif.
				} // End of for loop.

				outString += '</row>';
			} // Endif.
		} // End of for loop.

		outString += '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
		return outString;
	}

	// Prepare genobj for MS-Office:
	msdoc.makemsdoc ( genobj, new_type, options, gen_private, type_info );
	gen_private.plugs.type.msoffice.makeOfficeGenerator ( 'xl', 'workbook', {} );

	gen_private.features.page_name = 'sheets'; // This document type must have pages.

	// On each generate we'll prepare the shared strings list:
	genobj.on ( 'beforeGen', cbPrepareXlsxToGenerate );

	gen_private.type.msoffice.files_list.push (
		{
			name: '/xl/styles.xml',
			type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
			clear: 'type'
		},
		{
			name: '/xl/workbook.xml',
			type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
			clear: 'type'
		}
	);

	gen_private.type.msoffice.rels_app.push (
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
			target: 'styles.xml',
			clear: 'type'
		},
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
			target: 'theme/theme1.xml',
			clear: 'type'
		}
	);

	gen_private.plugs.intAddAnyResourceToParse ( 'docProps\\app.xml', 'buffer', null, cbMakeXlsApp, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'xl\\styles.xml', 'buffer', null, cbMakeXlsStyles, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'xl\\workbook.xml', 'buffer', null, cbMakeXlsWorkbook, true );

	gen_private.plugs.intAddAnyResourceToParse ( 'xl\\_rels\\workbook.xml.rels', 'buffer', gen_private.type.msoffice.rels_app, gen_private.plugs.type.msoffice.cbMakeRels, true );

	// ----- API for Excel documents: -----

	///
	/// @brief Create a new sheet.
	///
	/// This method creating a new Excel sheet.
	///
	genobj.makeNewSheet = function () {
		var pageNumber = gen_private.pages.length;
		var sheetObj = {}; // The sheet object that the user will use.

		sheetObj.data = []; // Place here all the data.

		gen_private.pages[pageNumber] = {};
		gen_private.pages[pageNumber].id = pageNumber;
		gen_private.pages[pageNumber].relId = gen_private.type.msoffice.rels_app.length + 1;
		gen_private.pages[pageNumber].sheet = sheetObj;

		gen_private.type.msoffice.rels_app.push (
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
				target: 'worksheets/sheet' + (pageNumber + 1) + '.xml',
				clear: 'data'
			}
		);

		gen_private.type.msoffice.files_list.push (
			{
				name: '/xl/worksheets/sheet' + (pageNumber + 1) + '.xml',
				type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
				clear: 'data'
			}
		);

		sheetObj.setCell = function ( position, data_val ) {
			var rel_pos = cbCellToNumber ( position, true );

			if ( !sheetObj.data[rel_pos.row] ) {
				sheetObj.data[rel_pos.row] = [];
			} // Endif.

			sheetObj.data[rel_pos.row][rel_pos.column] = data_val;
		};

		gen_private.plugs.intAddAnyResourceToParse ( 'xl\\worksheets\\sheet' + (pageNumber + 1) + '.xml', 'buffer', gen_private.pages[pageNumber], cbMakeXlsSheet, false );

		return sheetObj;
	};
}

baseobj.plugins.registerDocType ( 'xlsx', makeXlsx, {}, baseobj.docType.SPREADSHEET, "Microsoft Excel Document" );

