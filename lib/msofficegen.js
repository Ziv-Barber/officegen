//
// officegen: basic Microsoft Office common code
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

if ( !String.prototype.encodeHTML ) {
	String.prototype.encodeHTML = function () {
		return this.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;');
	};
}

///
/// @brief Extend officegen object with MS-Office support.
///
/// This method extending the given officegen object with the common MS-Office code.
///
/// @param[in] genobj The object to extend.
/// @param[in] new_type The type of object to create.
/// @param[in] options The object's options.
/// @param[in] gen_private Access to the internals of this object.
/// @param[in] type_info Additional information about this type.
///
function makemsdoc ( genobj, new_type, options, gen_private, type_info ) {
	///
	/// @brief Generate string of the current date and time.
	///
	/// This method generating a string with the current date and time in Office XML format.
	///
	/// @return String of the current date and time in Office XML format.
	///
	function getCurDateTimeForOffice () {
		var date = new Date ();

		var year = date.getFullYear ();
		var month = date.getMonth () + 1;
		var day = date.getDate ();
		var hour = date.getHours ();
		var min = date.getMinutes ();
		var sec = date.getSeconds ();

		month = (month < 10 ? "0" : "") + month;
		day = (day < 10 ? "0" : "") + day;
		hour = (hour < 10 ? "0" : "") + hour;
		min = (min < 10 ? "0" : "") + min;
		sec = (sec < 10 ? "0" : "") + sec;

		return year + "-" + month + "-" + day + "T" + hour + ":" + min + ":" + sec + 'Z';
	}

	///
	/// @brief Compact the given array.
	///
	/// This function compacting the given array.
	///
	/// @param[in] arr The array to compact.
	///
	function compactArray ( arr ) {
		var len = arr.length, i;

		for ( i = 0; i < len; i++ )
			arr[i] && arr.push ( arr[i] );  // Copy non-empty values to the end of the array.

		arr.splice ( 0 , len ); // Cut the array and leave only the non-empty values.
	}

	///
	/// @brief Create the main files list resource.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeMainFilesList ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data );
		outString += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';

		for ( var i = 0, total_size = gen_private.type.msoffice.files_list.length; i < total_size; i++ ) {
			if ( typeof gen_private.type.msoffice.files_list[i] != 'undefined' ) {
				if ( gen_private.type.msoffice.files_list[i].ext )
				{
					
					// Fixed by @author:vtloc @date:2014Jan09
					// Reason: if we write out duplicate extension, office will raise error
					//
					if( outString.indexOf('<Default Extension="'+gen_private.type.msoffice.files_list[i].ext) == -1 )
					{ // check to make sure we don't write duplicate extension tag
						outString += '<Default Extension="' + gen_private.type.msoffice.files_list[i].ext + '" ContentType="' + gen_private.type.msoffice.files_list[i].type + '"/>';
					}
				} else
				{
					outString += '<Override PartName="' + gen_private.type.msoffice.files_list[i].name + '" ContentType="' + gen_private.type.msoffice.files_list[i].type + '"/>';
				} // Endif.
			} // Endif.
		} // End of for loop.

		outString += '</Types>\n';
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
	function cbMakeTheme ( data ) {
		return gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeCore ( data ) {
		var curDateTime = getCurDateTimeForOffice ();
		var userName = genobj.options.creator ? genobj.options.creator : 'officegen';
		var extraFields = '';

		// Work on all the properties:
		for ( infoRec in genobj.info ) {
			if ( genobj.info[infoRec] && genobj.info[infoRec].element && genobj.info[infoRec].data ) {
				extraFields += '<' + genobj.info[infoRec].element + '>' + genobj.info[infoRec].data.encodeHTML () + '</' + genobj.info[infoRec].element + '>';
			} // Endif.
		} // End of for loop.

		return gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + extraFields + '<dc:creator>' + userName + '</dc:creator><cp:lastModifiedBy>' + userName + '</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">' + curDateTime + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + curDateTime + '</dcterms:modified></cp:coreProperties>';
	}

	///
	/// @brief Remove selected records from the given array.
	///
	/// This method destroys records inside the given array of the given type.
	///
	/// @param[in] arr The array to work on it.
	/// @param[in] type_to_clear The type of records to clear.
	///
	function clearSmartArrayFromType ( arr, type_to_clear ) {
		var is_need_compact = false;

		for ( var i = 0, total_size = arr.length; i < total_size; i++ ) {
			if ( typeof arr[i] != 'undefined' ) {
				if ( arr[i].clear && (arr[i].clear == type_to_clear) ) {
					delete arr[i];
					is_need_compact = true;
				} // Endif.
			} // Endif.
		} // End of for loop.

		if ( is_need_compact ) {
			compactArray ( arr );
		} // Endif.
	}

	///
	/// @brief Clean after finishing to generate the document.
	///
	/// ???.
	///
	/// @param[in] err Generation error message (if there were any).
	/// @param[in] written Number of bytes been created.
	///
	function cbOfficeClearAfterGenerate ( err, written ) {
		clearSmartArrayFromType ( gen_private.type.msoffice.rels_main, 'generate' );
		clearSmartArrayFromType ( gen_private.type.msoffice.rels_app, 'generate' );
		clearSmartArrayFromType ( gen_private.type.msoffice.files_list, 'generate' );
	};

	///
	/// @brief Clear all the information of the current document.
	///
	/// ???.
	///
	function cbOfficeClearDocData () {
		clearSmartArrayFromType ( gen_private.type.msoffice.rels_main, 'data' );
		clearSmartArrayFromType ( gen_private.type.msoffice.rels_app, 'data' );
		clearSmartArrayFromType ( gen_private.type.msoffice.files_list, 'data' );

		for ( infoItem in genobj.info ) {
			genobj.info[infoItem].data = genobj.info[infoItem].def_data;
		} // Endif.

		gen_private.type.msoffice.src_files_list = [];
	};

	// Basic API for plugins:

	gen_private.plugs.type.msoffice = {};

	///
	/// @brief Configure a new Office property type.
	///
	/// This method register a new type of property that the user can configure. This property must be 
	/// a valid MS-Office property as you can configure on the "files/properties" menu option on MS-Office.
	///
	/// @param[in] element_name The name of the XML element of this type.
	/// @param[in] def_data Default value of this type.
	/// @param[in] prop_name The name of the options property to configure this type.
	/// @param[in] user_access_func_name The name of the function to create to configure this type.
	///
	gen_private.plugs.type.msoffice.addInfoType = function ( element_name, def_data, prop_name, user_access_func_name ) {
		genobj.info[element_name] = {};
		genobj.info[element_name].element = element_name;
		genobj.info[element_name].data = def_data;
		genobj.info[element_name].def_data = def_data;

		// The user of officegen can configure this property using the options object:
		if ( genobj.options.prop_name )
		{
			genobj.info[element_name].data = genobj.options.prop_name;
		} // Endif.

		genobj[user_access_func_name] = function ( new_data ) {
			genobj.info[element_name].data = new_data;
		};
	};

	///
	/// @brief Get the string that opening every Office XML type.
	///
	/// Every Microsoft Office XML resource will have this header at the begining of the file.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml = function ( data ) {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
	}

	///
	/// @brief Generate rel based XML file.
	///
	/// ???.
	///
	/// @param[in] data Array filled with all the rels links.
	/// @return Text string.
	///
	gen_private.plugs.type.msoffice.cbMakeRels = function ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data );
		outString += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

		// Add all the rels records inside the data array:
		var realRel = 1;
		for ( var i = 0, total_size = data.length; i < total_size; i++ ) {
			if ( typeof data[i] != 'undefined' ) {
				outString += '<Relationship Id="rId' + realRel + '" Type="' + data[i].type + '" Target="' + data[i].target + '"/>';
				realRel++;
			} // Endif.
		} // End of for loop.

		outString += '</Relationships>\n';
		return outString;
	}

	///
	/// @brief Prepare the officegen object to MS-Office documents.
	///
	/// Every plugin that implementing gemenrating MS-Office document must call this method to initialize 
	/// the common stuff.
	///
	/// @param[in] main_path The name of the main folder holding the common resources of this type.
	/// @param[in] main_file The main resource file name of this type.
	/// @param[in] ext_opt Optional settings (unused right now).
	///
	gen_private.plugs.type.msoffice.makeOfficeGenerator = function ( main_path, main_file, ext_opt ) {
		gen_private.features.type.msoffice = {};
		gen_private.features.type.msoffice.main_path = main_path;
		gen_private.features.type.msoffice.main_path_file = main_file;
		gen_private.type.msoffice = {};
		gen_private.type.msoffice.rels_main = [];
		gen_private.type.msoffice.rels_app = [];
		gen_private.type.msoffice.files_list = [];
		gen_private.type.msoffice.src_files_list = [];

		// Holding all the Office properties:
		genobj.info = {};

		genobj.on ( 'afterGen', cbOfficeClearAfterGenerate );
		genobj.on ( 'clearDoc', cbOfficeClearDocData );

		gen_private.type.msoffice.rels_main.push (
			{
				target: 'docProps/app.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
				clear: 'type'
			},
			{
				target: 'docProps/core.xml',
				type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
				clear: 'type'
			},
			{
				target: gen_private.features.type.msoffice.main_path + '/' + gen_private.features.type.msoffice.main_path_file + '.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
				clear: 'type'
			}
		);

		gen_private.type.msoffice.files_list.push (
			{
				ext: 'rels',
				type: 'application/vnd.openxmlformats-package.relationships+xml',
				clear: 'type'
			},
			{
				ext: 'xml',
				type: 'application/xml',
				clear: 'type'
			},
			{
				ext: 'jpeg',
				type: 'image/jpeg',
				clear: 'type'
			},
			{
				ext: 'png',
				type: 'image/png',
				clear: 'type'
			},
			{
				name: '/docProps/app.xml',
				type: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
				clear: 'type'
			},
			{
				name: '/' + gen_private.features.type.msoffice.main_path + '/theme/theme1.xml',
				type: 'application/vnd.openxmlformats-officedocument.theme+xml',
				clear: 'type'
			},
			{
				name: '/docProps/core.xml',
				type: 'application/vnd.openxmlformats-package.core-properties+xml',
				clear: 'type'
			}
		);

		gen_private.plugs.intAddAnyResourceToParse ( '_rels\\.rels', 'buffer', gen_private.type.msoffice.rels_main, gen_private.plugs.type.msoffice.cbMakeRels, true );
		gen_private.plugs.intAddAnyResourceToParse ( '[Content_Types].xml', 'buffer', null, cbMakeMainFilesList, true );
		gen_private.plugs.intAddAnyResourceToParse ( 'docProps\\core.xml', 'buffer', null, cbMakeCore, true );
		gen_private.plugs.intAddAnyResourceToParse ( gen_private.features.type.msoffice.main_path + '\\theme\\theme1.xml', 'buffer', null, cbMakeTheme, true );
	};
};

baseobj.plugins.registerPrototype ( 'msoffice', makemsdoc, 'Microsoft Office Document Prototype' );

exports.makemsdoc = makemsdoc;

