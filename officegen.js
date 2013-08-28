//
// officegen - generating Office documents
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

require("setimmediate"); // To be compatible with all versions of node.js

var officegen_info = require('./package.json');
var archiver = require('archiver');
var fast_image_size = require('fast-image-size');
var fs = require('fs');
var path = require('path');
var Stream = require('stream'); // BMK_STREAM:

// Globals:

var int_officegen_globals = {}; // Our internal globals.

int_officegen_globals.settings = {};
int_officegen_globals.types = {};
int_officegen_globals.common_obj = {};

// *********************************************************************
// Private functions: (search for ***PUBLIC_CODE*** for the public API):
// *********************************************************************

// ---------------------------------------------
// Internal functions not depended on officegen:
// ---------------------------------------------

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

function compactArray ( arr ) {
	var len = arr.length, i;

	for ( i = 0; i < len; i++ )
		arr[i] && arr.push ( arr[i] );  // Copy non-empty values to the end of the array.

	arr.splice ( 0 , len ); // Cut the array and leave only the non-empty values.
}

if ( !String.prototype.encodeHTML ) {
	String.prototype.encodeHTML = function () {
		return this.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;');
	};
}

// ----------------------
// The Office gen object:
// ----------------------

///
/// @brief The constructor of the office generator object.
///
/// This constructor function is been called by makegen().
///
/// @b The @b Options:
///
/// The configuration options effecting the operation of the officegen object. Some of them can be only been 
/// declared on the 'options' object passed to the constructor object and the rest can be configured by either 
/// a property with the same name or by special function.
///
/// @b List @b Of @b Options:
///
/// - 'type' - the type of generator to create. Possible options: either 'pptx', 'docx' or 'xlsx'.
/// - 'creator' - the name of the document's author. The default is 'officegen'.
/// - 'onend' - callback that been fired after finishing to create the zip stream.
///
/// @param[in] options List of configuration options (see in the description of this function).
///
officegen = function ( options ) {
	var genobj = this;    // Can be accessed by all the functions been declared inside the officegen object.
	var gen_private = {}; // For all the private data that we don't want the user of officegen to access it.

	gen_private.perment = {}; // All stuff that is 100% unchangable after selecting the type to create.
	gen_private.thisDoc = {}; // All stuff that is 100% depended on the current document to create (all the stuff that 
	                          // been erased by calling to startNewDoc().
	gen_private.mixed = {}; // Mixed stuff (both perment and document depend).

	gen_private.perment.features = {}; // Features been configured by the type selector and you can't change them.
	// gen_private.perment.features.page_name
	// gen_private.perment.features.call_before_gen
	// gen_private.perment.features.call_after_gen
	// gen_private.perment.features.call_on_clear

	gen_private.thisDoc.pages = []; // Information about all the pages to create.
	gen_private.mixed.res_list = []; // List of all the resources to create inside the zip.
	gen_private.mixed.res_data = {}; // Information about all the resources to create.

	// From now until ***REST_OF_OFFICEGEN_CODE*** there are only function declarations so I'll not put 
	// code outside of the function until ***REST_OF_OFFICEGEN_CODE*** so don't worry!

	///
	/// @brief Prepare all the internal options.
	///
	/// This function configuring all the both public properties and internal options depending on the given options.
	///
	/// @param[in] options The options to configure.
	///
	function setOptions ( options )
	{
		// BMK_TODO: Temporary - a better way will be implemented later:
		genobj.options = options ? options : {};
	};

	///
	/// @brief Configure this object to generate the given type of document.
	///
	/// This function configuring the generator to create the given type of document.
	///
	/// @param[in] new_type The type of document to create.
	///
	function setGeneratorType ( new_type ) {
		gen_private.length = 0;

		if ( new_type ) {
			switch ( new_type ) {
				case 'pptx':
				case 'ppsx':
					makePptxGenerator ( new_type );
					break;

				case 'docx':
					makeDocxGenerator ();
					break;

				case 'xlsx':
					makeXlsxGenerator ();
					break;

				default:
					// BMK_TODO: One day all the code above will be moved to here:
					for ( var cur_type in int_officegen_globals.types ) {
						if ( (cur_type == new_type) && int_officegen_globals.types[cur_type] && int_officegen_globals.types[cur_type].create_obj ) {
							int_officegen_globals.types[cur_type].create_obj ( new_type, genobj.options, gen_private, int_officegen_globals.types[cur_type] );
							break;
						} // Endif.
					} // End of for loop.

					console.error ( '\nFATAL ERROR: Either unknown or unsupported file type - %s\n', options.type );
					throw 'FATAL ERROR: Invalid file type.';
			} // End of switch.
		} // Endif.
	};

	///
	/// @brief Add a resource to the list of resources to place inside the output zip file.
	///
	/// This method adding a resource to the list of resources to place inside the output document ZIP.
	///
	/// @param[in] resource_name The name of the resource (path).
	/// @param[in] type_of_res The type of this resource: either 'file' or 'buffer'.
	/// @param[in] res_data Optional data to use when creating this resource.
	/// @param[in] res_cb Callback to generate this resource (for 'buffer' mode only).
	/// @param[in] is_always Is true if this resource is perment for all the zip of this document type.
	///
	function intAddAnyResourceToParse ( resource_name, type_of_res, res_data, res_cb, is_always ) {
		var newRes = {};

		newRes.name = resource_name;
		newRes.type = type_of_res;
		newRes.data = res_data;
		newRes.callback = res_cb;
		newRes.is_perment = is_always;

		gen_private.mixed.res_list.push ( newRes );
	};

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] element_name ???.
	/// @param[in] def_data ???.
	/// @param[in] prop_name ???.
	/// @param[in] user_access_func_name ???.
	///
	function addInfoType ( element_name, def_data, prop_name, user_access_func_name ) {
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
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeMsOfficeBasicXml ( data ) {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Array filled with all the rels links.
	/// @return Text string.
	///
	function cbMakeRels ( data ) {
		var outString = cbMakeMsOfficeBasicXml ( data );
		outString += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

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
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeMainFilesList ( data ) {
		var outString = cbMakeMsOfficeBasicXml ( data );
		outString += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';

		for ( var i = 0, total_size = gen_private.mixed.files_list.length; i < total_size; i++ ) {
			if ( typeof gen_private.mixed.files_list[i] != 'undefined' ) {
				if ( gen_private.mixed.files_list[i].ext )
				{
					outString += '<Default Extension="' + gen_private.mixed.files_list[i].ext + '" ContentType="' + gen_private.mixed.files_list[i].type + '"/>';

				} else
				{
					outString += '<Override PartName="' + gen_private.mixed.files_list[i].name + '" ContentType="' + gen_private.mixed.files_list[i].type + '"/>';
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
		return cbMakeMsOfficeBasicXml ( data ) + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
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

		for ( infoRec in genobj.info ) {
			if ( genobj.info[infoRec] && genobj.info[infoRec].element && genobj.info[infoRec].data ) {
				extraFields += '<' + genobj.info[infoRec].element + '>' + genobj.info[infoRec].data + '</' + genobj.info[infoRec].element + '>';
			} // Endif.
		} // End of for loop.

		return cbMakeMsOfficeBasicXml ( data ) + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + extraFields + '<dc:creator>' + userName + '</dc:creator><cp:lastModifiedBy>' + userName + '</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">' + curDateTime + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + curDateTime + '</dcterms:modified></cp:coreProperties>';
	}

	// PowerPoint only:

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakePptxPresProps ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakePptxStyles ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakePptxViewProps ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="64" d="100"/><a:sy n="64" d="100"/></p:scale><p:origin x="-1392" y="-96"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="78028800" cy="78028800"/></p:viewPr>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakePptxLayout ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1"><p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="2130425"/><a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1371600" y="3886200"/><a:ext cx="6400800" cy="1752600"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl9pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master subtitle style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>6/13/2013</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakePptxPresentation ( data ) {
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>';

		for ( var i = 0, total_size = gen_private.thisDoc.pages.length; i < total_size; i++ ) {
			outString += '<p:sldId id="' + (i + 256) + '" r:id="rId' + (i + 2) + '"/>';
		} // End of for loop.

		outString += '</p:sldIdLst><p:sldSz cx="9144000" cy="6858000" type="screen4x3"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>';

		var curPos = 0;
		for ( var i = 1; i < 10; i++ )
		{
			outString += '<a:lvl' + i + 'pPr marL="' + curPos + '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl' + i + 'pPr>';
			curPos += 457200;
		} // End of for loop.

		outString += '</p:defaultTextStyle></p:presentation>';
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
	function cbMakePptxSlideMasters ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>6/13/2013</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3124200" y="6356350"/><a:ext cx="2895600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="6553200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:otherStyle></p:txStyles></p:sldMaster>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] color_info ???.
	/// @param[in] back_info ???.
	///
	function cMakePptxColorSelection ( color_info, back_info )
	{
		var outText = '';
		var colorVal;
		var fillType = 'solid';
		var internalElements = '';

		if ( back_info ) {
			outText += '<p:bg><p:bgPr>';

			outText += cMakePptxColorSelection ( back_info, false );

			outText += '<a:effectLst/>';
			// BMK_TODO: (add support for effects)
			
			outText += '</p:bgPr></p:bg>';
		} // Endif.

		if ( color_info ) {
			if ( typeof color_info == 'string' ) {
				colorVal = color_info;

			} else {
				if ( color_info.type ) {
					fillType = color_info.type;
				} // Endif.

				if ( color_info.color ) {
					colorVal = color_info.color;
				} // Endif.

				if ( color_info.alpha ) {
					internalElements += '<a:alpha val="' + (100 - color_info.alpha) + '000"/>';
				} // Endif.
			} // Endif.

			switch ( fillType )
			{
				case 'solid':
					outText += '<a:solidFill><a:srgbClr val="' + colorVal + '">' + internalElements + '</a:srgbClr></a:solidFill>';
					break;
			} // End of switch.
		} // Endif.

		return outText;
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] text_info Information how to display the text.
	/// @param[in] slide_obj The object of this slider.
	/// @return Text string.
	///
	function cMakePptxOutTextData ( text_info, slide_obj ) {
		var out_obj = {};

		out_obj.font_size = '';
		out_obj.bold = '';
		out_obj.underline = '';
		out_obj.rpr_info = '';

		if ( typeof text_info == 'object' )
		{
			if ( text_info.bold ) {
				out_obj.bold = ' b="1"';
			} // Endif.

			if ( text_info.underline ) {
				out_obj.underline = ' u="sng"';
			} // Endif.

			if ( text_info.font_size ) {
				out_obj.font_size = ' sz="' + text_info.font_size + '00"';
			} // Endif.

			if ( text_info.color ) {
				out_obj.rpr_info += cMakePptxColorSelection ( text_info.color );

			} else if ( slide_obj && slide_obj.color )
			{
				out_obj.rpr_info += cMakePptxColorSelection ( slide_obj.color );
			} // Endif.

			if ( text_info.font_face ) {
				out_obj.rpr_info += '<a:latin typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/><a:cs typeface="' + text_info.font_face + '" pitchFamily="34" charset="0"/>';
			} // Endif.

		} else {
			if ( slide_obj && slide_obj.color )
			{
				out_obj.rpr_info += cMakePptxColorSelection ( slide_obj.color );
			} // Endif.
		} // Endif.

		if ( out_obj.rpr_info != '' )
			out_obj.rpr_info += '</a:rPr>';

		return out_obj;
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] text_info Information how to display the text.
	/// @param[in] text_string The text string.
	/// @param[in] slide_obj The object of this slider.
	/// @return The PPTX code.
	///
	function cMakePptxOutTextCommand ( text_info, text_string, slide_obj ) {
		var area_opt_data = cMakePptxOutTextData ( text_info, slide_obj );
		return '<a:r><a:rPr lang="en-US"' + area_opt_data.font_size + area_opt_data.bold + area_opt_data.underline + ' dirty="0" smtClean="0"' + (area_opt_data.rpr_info != '' ? ('>' + area_opt_data.rpr_info) : '/>') + '<a:t>' + text_string.encodeHTML () + '</a:t></a:r>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] in_data_val ???.
	/// @param[in] max_value ???.
	/// @param[in] def_value ???.
	/// @param[in] auto_val ???.
	/// @return ???.
	///
	function parseSmartNumber ( in_data_val, max_value, def_value, auto_val, mul_val ) {
		var realNum = mul_val ? in_data_val * mul_val : in_data_val;
	
		if ( typeof in_data_val == 'undefined' ) {
			return (typeof def_value == 'number') ? def_value : 0;
		} // Endif.
	
		if ( typeof in_data_val == 'string' ) {
			if ( in_data_val.indexOf ( '%' ) != -1 ) {
				var realMax = (typeof max_value == 'number') ? max_value : 0;
				if ( realMax <= 0 ) return 0;

				var realVal = parseInt ( in_data_val, 10 );
				return (realMax / 100) * realVal;
			} // Endif.

			if ( in_data_val.indexOf ( '#' ) != -1 ) {
				var realVal = parseInt ( in_data_val, 10 );
				return realMax;
			} // Endif.

			var realAuto = (typeof auto_val == 'number') ? auto_val : 0;

			if ( in_data_val == '*' ) {
				return realAuto;
			} // Endif.

			if ( in_data_val == 'c' ) {
				return realAuto / 2;
			} // Endif.

			return (typeof def_value == 'number') ? def_value : 0;
		} // Endif.
	
		if ( typeof in_data_val == 'number' ) {
			return realNum;
		} // Endif.

		return (typeof def_value == 'number') ? def_value : 0;
	}

	///
	/// @brief Generate a slider resource.
	///
	/// This function generating a slider XML resource.
	///
	/// @param[in] data The main slide object.
	/// @return Text string.
	///
	function cbMakePptxSlide ( data ) {
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
		var objs_list = data.data;

		if ( !data.slide.show ) {
			outString += ' show="0"';
		} // Endif.

		outString += '><p:cSld>';

		if ( data.slide.back ) {
			outString += cMakePptxColorSelection ( false, data.slide.back );
		} // Endif.

		outString += '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

		// Loop on all the objects inside the slide to add it into the slide:
		for ( var i = 0, total_size = objs_list.length; i < total_size; i++ ) {
			var x = 0;
			var y = 0;
			var cx = 2819400;
			var cy = 369332;

			var moreStyles = '';
			var moreStylesAttr = '';
			var outStyles = '';
			var styleData = '';
			var shapeType = null;
			var locationAttr = '';

			if ( objs_list[i].options ) {
				if ( objs_list[i].options.cx ) {
					cx = parseSmartNumber ( objs_list[i].options.cx, 9144000, 2819400, 9144000, 10000 );
				} // Endif.

				if ( objs_list[i].options.cy ) {
					cy = parseSmartNumber ( objs_list[i].options.cy, 6858000, 369332, 6858000, 10000 );
				} // Endif.

				if ( objs_list[i].options.x ) {
					x = parseSmartNumber ( objs_list[i].options.x, 9144000, 0, 9144000 - cx, 10000 );
				} // Endif.

				if ( objs_list[i].options.y ) {
					y = parseSmartNumber ( objs_list[i].options.y, 6858000, 0, 6858000 - cy, 10000 );
				} // Endif.

				if ( objs_list[i].options.shape && (typeof objs_list[i].options.shape == 'string') ) {
					shapeType = objs_list[i].options.shape;
				} // Endif.

				if ( objs_list[i].options.flip_vertical ) {
					locationAttr += ' flipV="1"';
				} // Endif.

				if ( objs_list[i].options.rotate ) {
					var rotateVal = objs_list[i].options.rotate > 360 ? (objs_list[i].options.rotate - 360) : objs_list[i].options.rotate;
					rotateVal *= 60000;
					locationAttr += ' rot="' + rotateVal + '"';
				} // Endif.
			} // Endif.

			switch ( objs_list[i].type ) {
				case 'text':
				case 'cxn':
					if ( shapeType == null ) shapeType = 'rect';

					if ( objs_list[i].type == 'cxn' ) {
						outString += '<p:cxnSp><p:nvCxnSpPr>';
						outString += '<p:cNvPr id="' + (i + 2) + '" name="Object ' + (i + 1) + '"/><p:nvPr/></p:nvCxnSpPr>';

					} else {
						outString += '<p:sp><p:nvSpPr>';
						outString += '<p:cNvPr id="' + (i + 2) + '" name="Object ' + (i + 1) + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>';
					} // Endif.

					outString += '<p:spPr>';

					outString += '<a:xfrm' + locationAttr + '>';

					outString += '<a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="' + shapeType + '"><a:avLst/></a:prstGeom>';

					if ( objs_list[i].options ) {
						if ( objs_list[i].options.fill ) {
							outString += cMakePptxColorSelection ( objs_list[i].options.fill );

						} else {
							outString += '<a:noFill/>';
						} // Endif.

						if ( objs_list[i].options.line ) {
							var lineAttr = '';

							if ( objs_list[i].options.line_size ) {
								lineAttr += ' w="' + (objs_list[i].options.line_size * 12700) + '"';
							} // Endif.

							// cmpd="dbl"

							outString += '<a:ln' + lineAttr + '>';
							outString += cMakePptxColorSelection ( objs_list[i].options.line );

							if ( objs_list[i].options.line_head ) {
								outString += '<a:headEnd type="' + objs_list[i].options.line_head + '"/>';
							} // Endif.

							if ( objs_list[i].options.line_tail ) {
								outString += '<a:tailEnd type="' + objs_list[i].options.line_tail + '"/>';
							} // Endif.

							outString += '</a:ln>';
						} // Endif.

					} else {
						outString += '<a:noFill/>';
					} // Endif.

					outString += '</p:spPr>';

					if ( objs_list[i].options ) {
						if ( objs_list[i].options.align ) {
							switch ( objs_list[i].options.align )
							{
								case 'right':
									moreStylesAttr += ' algn="r"';
									break;

								case 'center':
									moreStylesAttr += ' algn="ctr"';
									break;

								case 'justify':
									moreStylesAttr += ' algn="just"';
									break;
							} // End of switch.
						} // Endif.

						if ( objs_list[i].options.indentLevel > 0 ) {
								moreStylesAttr += ' lvl="' + objs_list[i].options.indentLevel + '"';
						} // Endif.
					} // Endif.

					if ( moreStyles != '' ) {
						outStyles = '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>';

					} else if ( moreStylesAttr != '' ) {
						outStyles = '<a:pPr' + moreStylesAttr + '/>';
					} // Endif.

					if ( styleData != '' ) {
						outString += '<p:style>' + styleData + '</p:style>';
					} // Endif.

					if ( typeof objs_list[i].text == 'string' ) {
						outString += '<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>' + outStyles;
						outString += cMakePptxOutTextCommand ( objs_list[i].options, objs_list[i].text, data.slide );

					} else if ( objs_list[i].text ) {
						outString += '<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>' + outStyles;

						for ( var j = 0, total_size_j = objs_list[i].text.length; j < total_size_j; j++ ) {
							if ( objs_list[i].text[j] ) {
								outString += cMakePptxOutTextCommand ( objs_list[i].text[j].options, objs_list[i].text[j].text, data.slide );
							} // Endif.
						} // Endif.
					} // Endif.

					if ( typeof objs_list[i].text != 'undefined' ) {
						var font_size = '';
						if ( objs_list[i].options && objs_list[i].options.font_size ) {
							font_size = ' sz="' + objs_list[i].options.font_size + '00"';
						} // Endif.

						outString += '<a:endParaRPr lang="en-US"' + font_size + ' dirty="0"/></a:p></p:txBody>';
					} // Endif.

					outString += objs_list[i].type == 'cxn' ? '</p:cxnSp>' : '</p:sp>';
					break;

				// Image:
				case 'image':
					outString += '<p:pic><p:nvPicPr><p:cNvPr id="' + (i + 2) + '" name="Object ' + (i + 1) + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId' + objs_list[i].rel_id + '" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm' + locationAttr + '><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
					break;

				// Paragraph:
				case 'p':
					if ( shapeType == null ) shapeType = 'rect';

					outString += '<p:sp><p:nvSpPr>';
					outString += '<p:cNvPr id="' + (i + 2) + '" name="Object ' + (i + 1) + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>';
					outString += '<p:spPr>';

					outString += '<a:xfrm' + locationAttr + '>';

					outString += '<a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="' + shapeType + '"><a:avLst/></a:prstGeom>';

					if ( objs_list[i].options ) {
						if ( objs_list[i].options.fill ) {
							outString += cMakePptxColorSelection ( objs_list[i].options.fill );

						} else {
							outString += '<a:noFill/>';
						} // Endif.

						if ( objs_list[i].options.line ) {
							outString += '<a:ln>';
							outString += cMakePptxColorSelection ( objs_list[i].options.line );

							if ( objs_list[i].options.line_head ) {
								outString += '<a:headEnd type="' + objs_list[i].options.line_head + '"/>';
							} // Endif.

							if ( objs_list[i].options.line_tail ) {
								outString += '<a:tailEnd type="' + objs_list[i].options.line_tail + '"/>';
							} // Endif.

							outString += '</a:ln>';
						} // Endif.

					} else {
						outString += '<a:noFill/>';
					} // Endif.

					outString += '</p:spPr>';

					if ( styleData != '' ) {
						outString += '<p:style>' + styleData + '</p:style>';
					} // Endif.

					outString += '<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>';

					for ( var j = 0, total_size_j = objs_list[i].data.length; j < total_size_j; j++ ) {
						if ( objs_list[i].data[j] ) {
							moreStylesAttr = '';
							moreStyles = '';
							
							if ( objs_list[i].data[j].options ) {
								if ( objs_list[i].data[j].options.align ) {
									switch ( objs_list[i].data[j].options.align )
									{
										case 'right':
											moreStylesAttr += ' algn="r"';
											break;

										case 'center':
											moreStylesAttr += ' algn="ctr"';
											break;

										case 'justify':
											moreStylesAttr += ' algn="just"';
											break;
									} // End of switch.
								} // Endif.

								if ( objs_list[i].data[j].options.indentLevel > 0 ) {
									moreStylesAttr += ' lvl="' + objs_list[i].data[j].options.indentLevel + '"';
								} // Endif.

								if ( objs_list[i].data[j].options.listType == 'number' ) {
									moreStyles += '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
								} // Endif.
							} // Endif.

							if ( moreStyles != '' ) {
								outStyles = '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>';

							} else if ( moreStylesAttr != '' ) {
								outStyles = '<a:pPr' + moreStylesAttr + '/>';
							} // Endif.

							outString += '<a:p>' + outStyles;

							// if ( typeof objs_list[i].data[j].text == 'string' ) {
							outString += cMakePptxOutTextCommand ( objs_list[i].data[j].options, objs_list[i].data[j].text, data.slide );
							// BMK_TODO:
						} // Endif.
					} // Endif.

					var font_size = '';
					if ( objs_list[i].options && objs_list[i].options.font_size ) {
						font_size = ' sz="' + objs_list[i].options.font_size + '00"';
					} // Endif.

					outString += '<a:endParaRPr lang="en-US"' + font_size + ' dirty="0"/></a:p>';
					outString += '</p:txBody>';

					outString += '</p:sp>';
					break;
			} // End of switch.
		} // End of for loop.
		
		outString += '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
		return outString;
	}

	///
	/// @brief Generate the extended attributes file (app) for PPTX/PPSX documents.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakePptxApp ( data ) {
		var slidesCount = gen_private.thisDoc.pages.length;
		var userName = genobj.options.creator || 'officegen';
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Words>0</Words><Application>Microsoft Office PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>' + slidesCount + '</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>' + slidesCount + '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' + (slidesCount + 1) + '" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr>';

		for ( var i = 0, total_size = gen_private.thisDoc.pages.length; i < total_size; i++ ) {
			outString += '<vt:lpstr>' + gen_private.thisDoc.pages[i].slide.name.encodeHTML () + '</vt:lpstr>';
		} // End of for loop.

		outString += '</vt:vector></TitlesOfParts><Company>' + userName + '</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>';
		return outString;
	}

	// Word only:

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxFontsTable ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="A00002EF" w:usb1="4000207B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000009F" w:csb1="00000000"/></w:font><w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="20002A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="20002A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Cambria"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="A00002EF" w:usb1="4000004B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000009F" w:csb1="00000000"/></w:font></w:fonts>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxSettings ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="120"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat/><w:rsids><w:rsidRoot w:val="00A94AF2"/><w:rsid w:val="00A02F19"/><w:rsid w:val="00A94AF2"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="off"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-US" w:bidi="en-US"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="2050"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxWeb ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/></w:webSettings>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxStyles ( data ) {
		return cbMakeMsOfficeBasicXml ( data ) + '<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Revision" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Bibliography" w:uiPriority="37"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="00A02F19"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>';
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxApp ( data ) {
		var userName = genobj.options.creator || 'officegen';
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>1</TotalTime><Pages>1</Pages><Words>0</Words><Characters>0</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company>' + userName + '</Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>0</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>';
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
	function cbMakeDocxDocument ( data ) {
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<w:document xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:body>';
		var objs_list = data.data;

		for ( var i = 0, total_size = objs_list.length; i < total_size; i++ ) {
			outString += '<w:p w:rsidR="00A77427" w:rsidRDefault="00A77427">';
			var pPrData = '';

			if ( objs_list[i].options ) {
				if ( objs_list[i].options.align ) {
					switch ( objs_list[i].options.align ) {
						case 'center':
							pPrData += '<w:jc w:val="center"/>';
							break;

						case 'right':
							pPrData += '<w:jc w:val="right"/>';
							break;

						case 'justify':
							pPrData += '<w:jc w:val="both"/>';
							break;
					} // End of switch.
				} // Endif.

				if ( objs_list[i].options.list_type ) {
					pPrData += '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + objs_list[i].options.list_type + '"/></w:numPr>';
				} // Endif.
			} // Endif.

			if ( pPrData ) {
				outString += '<w:pPr>' + pPrData + '</w:pPr>';
			} // Endif.

			for ( var j = 0, total_size_j = objs_list[i].data.length; j < total_size_j; j++ ) {
				if ( objs_list[i].data[j] ) {
					var rExtra = '';
					var tExtra = '';
					var rPrData = '';

					if ( objs_list[i].data[j].options ) {
						if ( objs_list[i].data[j].options.color ) {
							rPrData += '<w:color w:val="' + objs_list[i].data[j].options.color + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.back ) {
							rPrData += '<w:shd w:val="clear" w:color="auto" w:fill="' + objs_list[i].data[j].options.back + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.bold ) {
							rPrData += '<w:b/><w:bCs/>';
						} // Endif.

						if ( objs_list[i].data[j].options.underline ) {
							rPrData += '<w:u w:val="single"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.font_face ) {
							rPrData += '<w:rFonts w:ascii="' + objs_list[i].data[j].options.font_face + '" w:hAnsi="' + objs_list[i].data[j].options.font_face + '" w:cs="' + objs_list[i].data[j].options.font_face + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.font_size ) {
							rPrData += '<w:sz w:val="' + objs_list[i].data[j].options.font_size + '"/><w:szCs w:val="' + objs_list[i].data[j].options.font_size + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.border ) {
							switch ( objs_list[i].data[j].options.border )
							{
								case 'single':
								case true:
									rPrData += '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
									break;
							} // End of switch.
						} // Endif.
					} // Endif.

					if ( objs_list[i].data[j].text ) {
						if ( objs_list[i].data[j].text[0] == ' ' ) {
							tExtra += ' xml:space="preserve"';
						} // Endif.

						outString += '<w:r' + rExtra + '>';

						if ( rPrData ) {
							outString += '<w:rPr>' + rPrData + '</w:rPr>';
						} // Endif.

						outString += '<w:t' + tExtra + '>' + objs_list[i].data[j].text.encodeHTML () + '</w:t></w:r>';

					} else if ( objs_list[i].data[j].page_break ) {
						outString += '<w:r><w:br w:type="page"/></w:r>';
					} // Endif.
				} // Endif.
			} // Endif.

			outString += '</w:p>';
		} // End of for loop.

		outString += '<w:p w:rsidR="00A02F19" w:rsidRDefault="00A02F19"/><w:sectPr w:rsidR="00A02F19" w:rsidSect="00A02F19"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
		return outString;
	}

	// Excel only:

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeXlsSharedStrings ( data ) {
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + genobj.generate_data.total_strings + '" uniqueCount="' + genobj.generate_data.shared_strings.length + '">';

		for ( var i = 0, total_size = genobj.generate_data.shared_strings.length; i < total_size; i++ ) {
			outString += '<si><t>' + genobj.generate_data.shared_strings[i].encodeHTML () + '</t></si>';
		} // Endif.

		return outString + '</sst>';
	}

	///
	/// @brief Prepare everything to generate XLSX files.
	///
	/// ???.
	///
	function cbPrepareXlsxToGenerate () {
		genobj.generate_data = {};
		genobj.generate_data.shared_strings = [];
		genobj.generate_data.total_strings = 0;
		genobj.generate_data.cell_strings = [];

		// Create the share strings data:
		for ( var i = 0, total_size = gen_private.thisDoc.pages.length; i < total_size; i++ ) {
			if ( gen_private.thisDoc.pages[i] ) {
				for ( var rowId = 0, total_size_y = gen_private.thisDoc.pages[i].sheet.data.length; rowId < total_size_y; rowId++ ) {
					if ( gen_private.thisDoc.pages[i].sheet.data[rowId] ) {
						for ( var columnId = 0, total_size_x = gen_private.thisDoc.pages[i].sheet.data[rowId].length; columnId < total_size_x; columnId++ ) {
							if ( typeof gen_private.thisDoc.pages[i].sheet.data[rowId][columnId] != 'undefined' ) {
								switch ( typeof gen_private.thisDoc.pages[i].sheet.data[rowId][columnId] ) {
									case 'string':
										genobj.generate_data.total_strings++;

										if ( !genobj.generate_data.cell_strings[i] ) {
											genobj.generate_data.cell_strings[i] = [];
										} // Endif.

										if ( !genobj.generate_data.cell_strings[i][rowId] ) {
											genobj.generate_data.cell_strings[i][rowId] = [];
										} // Endif.

										for ( var j = 0, total_size_j = genobj.generate_data.shared_strings.length; j < total_size_j; j++ ) {
											if ( gen_private.thisDoc.pages[i].sheet.data[rowId][columnId] == genobj.generate_data.shared_strings[j] ) {
												genobj.generate_data.cell_strings[i][rowId][columnId] = j;
											} // Endif.
										} // Endif.

										if ( typeof genobj.generate_data.cell_strings[i][rowId][columnId] == 'undefined' ) {
											genobj.generate_data.cell_strings[i][rowId][columnId] = genobj.generate_data.shared_strings.length;
											genobj.generate_data.shared_strings[genobj.generate_data.shared_strings.length] = gen_private.thisDoc.pages[i].sheet.data[rowId][columnId];
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
			intAddAnyResourceToParse ( 'xl\\sharedStrings.xml', 'buffer', null, cbMakeXlsSharedStrings, false );
			gen_private.mixed.files_list.push (
				{
					name: '/xl/sharedStrings.xml',
					type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
					clear: 'generate'
				}
			);

			gen_private.mixed.rels_app.push (
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
		return cbMakeMsOfficeBasicXml ( data ) + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
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
		var pagesCount = gen_private.thisDoc.pages.length;
		var userName = genobj.options.creator || 'officegen';
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>' + pagesCount + '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' + pagesCount + '" baseType="lpstr">';

		for ( var i = 0, total_size = gen_private.thisDoc.pages.length; i < total_size; i++ ) {
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
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4507"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="19095" windowHeight="7485"/></bookViews><sheets>';

		for ( var i = 0, total_size = gen_private.thisDoc.pages.length; i < total_size; i++ ) {
			var sheetName = gen_private.thisDoc.pages[i].sheet.name || 'Sheet' + (i + 1);
                        var rId = gen_private.thisDoc.pages[i].relId;
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
		var outString = cbMakeMsOfficeBasicXml ( data ) + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
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
				outString += '<row r="' + (rowId + 1) + '" spans="1:2">';

				for ( var columnId = 0, total_size_x = data.sheet.data[rowId].length; columnId < total_size_x; columnId++ ) {
					if ( typeof data.sheet.data[rowId][columnId] != 'undefined' ) {
						var isString = '';
						var cellOutData = '0';

						switch ( typeof data.sheet.data[rowId][columnId] ) {
							case 'number':
								cellOutData = data.sheet.data[rowId][columnId];
								break;

							case 'string':
								isString = ' t="s"';
								cellOutData = genobj.generate_data.cell_strings[data.id][rowId][columnId];
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

	//
	// Helper functions:
	//

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] arr ???.
	/// @param[in] type_to_clear ???.
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
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] err ???.
	/// @param[in] written ???.
	///
	function cbOfficeClearAfterGenerate ( err, written ) {
		clearSmartArrayFromType ( gen_private.mixed.rels_main, 'generate' );
		clearSmartArrayFromType ( gen_private.mixed.rels_app, 'generate' );
		clearSmartArrayFromType ( gen_private.mixed.files_list, 'generate' );

		if ( gen_private.perment.features.clear_gen_more ) {
			gen_private.perment.features.clear_gen_more ( err, written );
		} // Endif.
	};

	///
	/// @brief ???.
	///
	/// ???.
	///
	function cbOfficeClearDocData () {
		clearSmartArrayFromType ( gen_private.mixed.rels_main, 'data' );
		clearSmartArrayFromType ( gen_private.mixed.rels_app, 'data' );
		clearSmartArrayFromType ( gen_private.mixed.files_list, 'data' );

		if ( gen_private.perment.features.clear_data_more ) {
			gen_private.perment.features.clear_data_more ();
		} // Endif.

		for ( infoItem in genobj.info ) {
			genobj.info[infoItem].data = genobj.info[infoItem].def_data;
		} // Endif.
	};

	//
	// Create all types:
	//

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] ??? ???.
	///
	function makeOfficeGenerator ( main_path, main_file, ext_opt ) {
		gen_private.mixed.res_data.main_path = main_path;
		gen_private.mixed.res_data.main_path_file = main_file;
		gen_private.mixed.rels_main = [];
		gen_private.mixed.rels_app = [];
		gen_private.mixed.files_list = [];
		gen_private.thisDoc.embeddings = [];

		genobj.info = {};

		gen_private.perment.features.call_after_gen = cbOfficeClearAfterGenerate;
		gen_private.perment.features.call_on_clear = cbOfficeClearDocData;

		gen_private.mixed.rels_main.push (
			{
				target: 'docProps/app.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
			},
			{
				target: 'docProps/core.xml',
				type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
			},
			{
				target: gen_private.mixed.res_data.main_path + '/' + gen_private.mixed.res_data.main_path_file + '.xml',
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
			}
		);

		gen_private.mixed.files_list.push (
			{
				ext: 'rels',
				type: 'application/vnd.openxmlformats-package.relationships+xml'
			},
			{
				ext: 'xml',
				type: 'application/xml'
			},
			{
				name: '/docProps/app.xml',
				type: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
			},
			{
				name: '/' + gen_private.mixed.res_data.main_path + '/theme/theme1.xml',
				type: 'application/vnd.openxmlformats-officedocument.theme+xml'
			},
			{
				name: '/docProps/core.xml',
				type: 'application/vnd.openxmlformats-package.core-properties+xml'
			}
		);

		intAddAnyResourceToParse ( '_rels\\.rels', 'buffer', gen_private.mixed.rels_main, cbMakeRels, true );
		intAddAnyResourceToParse ( '[Content_Types].xml', 'buffer', null, cbMakeMainFilesList, true );
		intAddAnyResourceToParse ( 'docProps\\core.xml', 'buffer', null, cbMakeCore, true );
		intAddAnyResourceToParse ( gen_private.mixed.res_data.main_path + '\\theme\\theme1.xml', 'buffer', null, cbMakeTheme, true );
	};

	///
	/// @brief Configure a MS Office 2007 PowerPoint document.
	///
	/// ???.
	///
	function makePptxGenerator ( new_type ) {
		makeOfficeGenerator ( 'ppt', 'presentation', {} );

		gen_private.thisDoc.images_count = 0;

		gen_private.perment.features.page_name = 'slides'; // This document type must have pages.

		addInfoType ( 'dc:title', '', 'title', 'setDocTitle' );

		var type_of_main_doc = 'slideshow';
		if ( new_type != 'ppsx' )
		{
			type_of_main_doc = 'presentation';
		} // Endif.

		gen_private.mixed.files_list.push (
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
				name: '/ppt/slideMasters/slideMaster1.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
				clear: 'type'
			},
			{
				name: '/ppt/presProps.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
				clear: 'type'
			},
			{
				name: '/ppt/presentation.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.' + type_of_main_doc + '.main+xml',
				clear: 'type'
			},
			{
				name: '/ppt/slideLayouts/slideLayout1.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
				clear: 'type'
			},
			{
				name: '/ppt/tableStyles.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
				clear: 'type'
			},
			{
				name: '/ppt/viewProps.xml',
				type: 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
				clear: 'type'
			}
		);

		intAddAnyResourceToParse ( 'ppt\\presProps.xml', 'buffer', null, cbMakePptxPresProps, true );
		intAddAnyResourceToParse ( 'ppt\\tableStyles.xml', 'buffer', null, cbMakePptxStyles, true );
		intAddAnyResourceToParse ( 'ppt\\viewProps.xml', 'buffer', null, cbMakePptxViewProps, true );
		intAddAnyResourceToParse ( 'ppt\\presentation.xml', 'buffer', null, cbMakePptxPresentation, true );

		intAddAnyResourceToParse ( 'ppt\\slideLayouts\\slideLayout1.xml', 'buffer', null, cbMakePptxLayout, true );
		intAddAnyResourceToParse ( 'ppt\\slideLayouts\\_rels\\slideLayout1.xml.rels', 'buffer', [
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
				target: '../slideMasters/slideMaster1.xml'
			}
		], cbMakeRels, true );

		intAddAnyResourceToParse ( 'ppt\\slideMasters\\slideMaster1.xml', 'buffer', null, cbMakePptxSlideMasters, true );
		intAddAnyResourceToParse ( 'ppt\\slideMasters\\_rels\\slideMaster1.xml.rels', 'buffer', [
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
				target: '../slideLayouts/slideLayout1.xml'
			},
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				target: '../theme/theme1.xml'
			}
		], cbMakeRels, true );

		intAddAnyResourceToParse ( 'docProps\\app.xml', 'buffer', null, cbMakePptxApp, true );

		gen_private.mixed.rels_app.push (
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
				target: 'slideMasters/slideMaster1.xml',
				clear: 'type'
			}
		);

		intAddAnyResourceToParse ( 'ppt\\_rels\\presentation.xml.rels', 'buffer', gen_private.mixed.rels_app, cbMakeRels, true );

		///
		/// @brief ???.
		///
		/// ???.
		///
		/// @param[in] ??? ???.
		///
		genobj.makeNewSlide = function () {
			var pageNumber = gen_private.thisDoc.pages.length;
			var slideObj = { show: true }; // The slide object that the user will use.

			gen_private.thisDoc.pages[pageNumber] = {};
			gen_private.thisDoc.pages[pageNumber].slide = slideObj;
			gen_private.thisDoc.pages[pageNumber].data = [];
			gen_private.thisDoc.pages[pageNumber].rels = [
				{
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
					target: '../slideLayouts/slideLayout1.xml',
					clear: 'data'
				}
			];

			gen_private.mixed.files_list.push (
				{
					name: '/ppt/slides/slide' + (pageNumber + 1) + '.xml',
					type: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
					clear: 'data'
				}
			);

			gen_private.mixed.rels_app.push (
				{
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
					target: 'slides/slide' + (pageNumber + 1) + '.xml',
					clear: 'data'
				}
			);

			slideObj.getPageNumber = function () { return pageNumber; };

			slideObj.name = 'Slide ' + (pageNumber + 1);

			///
			/// @brief ???.
			///
			/// ???.
			///
			/// @param[in] ??? ???.
			///
			slideObj.addText = function ( text, opt, y_pos, x_size, y_size, opt_b ) {
				var objNumber = gen_private.thisDoc.pages[pageNumber].data.length;

				gen_private.thisDoc.pages[pageNumber].data[objNumber] = {};
				gen_private.thisDoc.pages[pageNumber].data[objNumber].type = 'text';
				gen_private.thisDoc.pages[pageNumber].data[objNumber].text = text;
				gen_private.thisDoc.pages[pageNumber].data[objNumber].options = typeof opt == 'object' ? opt : {};

				if ( typeof opt == 'string' ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.color = opt;

				} else if ( (typeof opt != 'object') && (typeof y_pos != 'undefined') ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.x = opt;
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.y = y_pos;

					if ( (typeof x_size != 'undefined') && (typeof y_size != 'undefined') ) {
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx = x_size;
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy = y_size;
					} // Endif.
				} // Endif.

				if ( typeof opt_b == 'object' ) {
					for ( var attrname in opt_b ) { gen_private.thisDoc.pages[pageNumber].data[objNumber].options[attrname] = opt_b[attrname]; }

				} else if ( (typeof x_size == 'object') && (typeof y_size == 'undefined') ) {
					for ( var attrname in x_size ) { gen_private.thisDoc.pages[pageNumber].data[objNumber].options[attrname] = x_size[attrname]; }
				} // Endif.
			};

			///
			/// @brief ???.
			///
			/// ???.
			///
			/// @param[in] ??? ???.
			///
			slideObj.addShape = function ( shape, opt, y_pos, x_size, y_size, opt_b ) {
				var objNumber = gen_private.thisDoc.pages[pageNumber].data.length;

				gen_private.thisDoc.pages[pageNumber].data[objNumber] = {};
				gen_private.thisDoc.pages[pageNumber].data[objNumber].type = 'text';
				gen_private.thisDoc.pages[pageNumber].data[objNumber].options = typeof opt == 'object' ? opt : {};
				gen_private.thisDoc.pages[pageNumber].data[objNumber].options.shape = shape;

				if ( typeof opt == 'string' ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.color = opt;

				} else if ( (typeof opt != 'object') && (typeof y_pos != 'undefined') ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.x = opt;
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.y = y_pos;

					if ( (typeof x_size != 'undefined') && (typeof y_size != 'undefined') ) {
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx = x_size;
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy = y_size;
					} // Endif.
				} // Endif.

				if ( typeof opt_b == 'object' ) {
					for ( var attrname in opt_b ) { gen_private.thisDoc.pages[pageNumber].data[objNumber].options[attrname] = opt_b[attrname]; }

				} else if ( (typeof x_size == 'object') && (typeof y_size == 'undefined') ) {
					for ( var attrname in x_size ) { gen_private.thisDoc.pages[pageNumber].data[objNumber].options[attrname] = x_size[attrname]; }
				} // Endif.
			};

			///
			/// @brief ???.
			///
			/// ???.
			///
			/// @param[in] ??? ???.
			///
			slideObj.addImage = function ( image_path, opt, y_pos, x_size, y_size, image_format_type ) {
				var objNumber = gen_private.thisDoc.pages[pageNumber].data.length;
				var image_type = (typeof image_format_type == 'string') ? image_format_type : 'png';
				var defWidth, defHeight = 0;

				if ( typeof image_path == 'string' ) {
					var ret_data = fast_image_size ( image_path );
					if ( ret_data.type == 'unknown' ) {
						var image_ext = path.extname ( image_path );

						switch ( image_ext ) {
							case '.bmp':
								image_type = 'bmp';
								break;

							case '.gif':
								image_type = 'gif';
								break;

							case '.jpg':
							case '.jpeg':
								image_type = 'jpeg';
								break;

							case '.emf':
								image_type = 'emf';
								break;

							case '.tiff':
								image_type = 'tiff';
								break;
						} // End of switch.

					} else {
						if ( ret_data.width ) {
							defWidth = ret_data.width;
						} // Endif.

						if ( ret_data.height ) {
							defHeight = ret_data.height;
						} // Endif.

						image_type = ret_data.type;
						if ( image_type == 'jpg' ) {
							image_type = 'jpeg';
						} // Endif.
					} // Endif.
				} // Endif.

				gen_private.thisDoc.pages[pageNumber].data[objNumber] = {};
				gen_private.thisDoc.pages[pageNumber].data[objNumber].type = 'image';
				gen_private.thisDoc.pages[pageNumber].data[objNumber].image = image_path;
				gen_private.thisDoc.pages[pageNumber].data[objNumber].options = typeof opt == 'object' ? opt : {};

				if ( !gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx && defWidth ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx = defWidth;
				} // Endif.

				if ( !gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy && defHeight ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy = defHeight;
				} // Endif.

				// console.log ( gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy );
				// console.log ( gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx );

				gen_private.thisDoc.pages[pageNumber].data[objNumber].image_id = gen_private.thisDoc.images_count++;
				gen_private.thisDoc.pages[pageNumber].data[objNumber].rel_id = gen_private.thisDoc.pages[pageNumber].rels.length + 1;

				intAddAnyResourceToParse ( 'ppt\\media\\image' + (gen_private.thisDoc.pages[pageNumber].data[objNumber].image_id + 1) + '.' + image_type, (typeof image_path == 'string') ? 'file' : 'stream', image_path, null, false );

				gen_private.thisDoc.pages[pageNumber].rels.push (
					{
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						target: '../media/image' + (gen_private.thisDoc.pages[pageNumber].data[objNumber].image_id + 1) + '.' + image_type,
						clear: 'data'
					}
				);

				if ( typeof opt == 'string' ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.color = opt;

				} else if ( (typeof opt != 'object') && (typeof y_pos != 'undefined') ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.x = opt;
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.y = y_pos;

					if ( (typeof x_size != 'undefined') && (typeof y_size != 'undefined') ) {
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx = x_size;
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy = y_size;
					} // Endif.
				} // Endif.
			};

			///
			/// @brief ???.
			///
			/// ???.
			///
			/// @param[in] ??? ???.
			///
			slideObj.addP = function ( text, opt, y_pos, x_size, y_size, opt_b ) {
				var objNumber = gen_private.thisDoc.pages[pageNumber].data.length;

				gen_private.thisDoc.pages[pageNumber].data[objNumber] = {};
				gen_private.thisDoc.pages[pageNumber].data[objNumber].type = 'p';
				gen_private.thisDoc.pages[pageNumber].data[objNumber].data = [];
				gen_private.thisDoc.pages[pageNumber].data[objNumber].options = typeof opt == 'object' ? opt : {};

				if ( typeof opt == 'string' ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.color = opt;

				} else if ( (typeof opt != 'object') && (typeof y_pos != 'undefined') ) {
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.x = opt;
					gen_private.thisDoc.pages[pageNumber].data[objNumber].options.y = y_pos;

					if ( (typeof x_size != 'undefined') && (typeof y_size != 'undefined') ) {
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cx = x_size;
						gen_private.thisDoc.pages[pageNumber].data[objNumber].options.cy = y_size;
					} // Endif.
				} // Endif.

				if ( typeof opt_b == 'object' ) {
					for ( var attrname in opt_b ) { gen_private.thisDoc.pages[pageNumber].data[objNumber].options[attrname] = opt_b[attrname]; }

				} else if ( (typeof x_size == 'object') && (typeof y_size == 'undefined') ) {
					for ( var attrname in x_size ) { gen_private.thisDoc.pages[pageNumber].data[objNumber].options[attrname] = x_size[attrname]; }
				} // Endif.

				// BMK_TODO:
				return gen_private.thisDoc.pages[pageNumber].data[objNumber].data;
			};

			intAddAnyResourceToParse ( 'ppt\\slides\\slide' + (pageNumber + 1) + '.xml', 'buffer', gen_private.thisDoc.pages[pageNumber], cbMakePptxSlide, false );
			intAddAnyResourceToParse ( 'ppt\\slides\\_rels\\slide' + (pageNumber + 1) + '.xml.rels', 'buffer', gen_private.thisDoc.pages[pageNumber].rels, cbMakeRels, false );		
			return slideObj;
		};
	};

	///
	/// @brief Configure a MS Office 2007 Word document.
	///
	/// ???.
	///
	function makeDocxGenerator () {
		makeOfficeGenerator ( 'word', 'document', {} );

		gen_private.perment.features.clear_data_more = function () {
			genobj.data.length = 0;
		};

		addInfoType ( 'dc:title', '', 'title', 'setDocTitle' );
		addInfoType ( 'dc:subject', '', 'subject', 'setDocSubject' );
		addInfoType ( 'cp:keywords', '', 'keywords', 'setDocKeywords' );
		addInfoType ( 'dc:description', '', 'description', 'setDescription' );

		gen_private.mixed.files_list.push (
			{
				name: '/word/settings.xml',
				type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
				clear: 'type'
			},
			{
				name: '/word/fontTable.xml',
				type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
				clear: 'type'
			},
			{
				name: '/word/webSettings.xml',
				type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
				clear: 'type'
			},
			{
				name: '/word/styles.xml',
				type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
				clear: 'type'
			},
			{
				name: '/word/document.xml',
				type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
				clear: 'type'
			}
		);

		gen_private.mixed.rels_app.push (
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
				target: 'styles.xml',
				clear: 'type'
			},
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
				target: 'settings.xml',
				clear: 'type'
			},
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
				target: 'webSettings.xml',
				clear: 'type'
			},
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
				target: 'fontTable.xml',
				clear: 'type'
			},
			{
				type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				target: 'theme/theme1.xml',
				clear: 'type'
			}
		);

		genobj.data = []; // All the data will be placed here.

		intAddAnyResourceToParse ( 'docProps\\app.xml', 'buffer', null, cbMakeDocxApp, true );
		intAddAnyResourceToParse ( 'word\\fontTable.xml', 'buffer', null, cbMakeDocxFontsTable, true );
		intAddAnyResourceToParse ( 'word\\settings.xml', 'buffer', null, cbMakeDocxSettings, true );
		intAddAnyResourceToParse ( 'word\\webSettings.xml', 'buffer', null, cbMakeDocxWeb, true );
		intAddAnyResourceToParse ( 'word\\styles.xml', 'buffer', null, cbMakeDocxStyles, true );
		intAddAnyResourceToParse ( 'word\\document.xml', 'buffer', genobj, cbMakeDocxDocument, true );

		intAddAnyResourceToParse ( 'word\\_rels\\document.xml.rels', 'buffer', gen_private.mixed.rels_app, cbMakeRels, true );

		///
		/// @brief ???.
		///
		/// ???.
		///
		/// @param[in] ??? ???.
		///
		genobj.createP = function ( options ) {
			var newP = {};

			newP.data = [];
			newP.options = options || {};

			newP.addText = function ( text_msg, opt, flag_data ) {
				newP.data[newP.data.length] = { text: text_msg, options: opt, ext_data: flag_data };
			};

			genobj.data[genobj.data.length] = newP;
			return newP;
		};

		///
		/// @brief ???.
		///
		/// ???.
		///
		/// @param[in] ??? ???.
		///
		genobj.createListOfDots = function ( options ) {
			var newP = genobj.createP ( options );

			newP.options.list_type = '1';

			return newP;
		};

		///
		/// @brief ???.
		///
		/// ???.
		///
		/// @param[in] ??? ???.
		///
		genobj.createListOfNumbers = function ( options ) {
			var newP = genobj.createP ( options );

			newP.options.list_type = '2';

			return newP;
		};

		///
		/// @brief ???.
		///
		/// ???.
		///
		/// @param[in] ??? ???.
		///
		genobj.putPageBreak = function () {
			var newP = {};

			newP.data = [ { 'page_break': true } ];

			genobj.data[genobj.data.length] = newP;
			return newP;
		};
	};

	///
	/// @brief Configure a MS Office 2007 Excel document.
	///
	/// ???.
	///
	function makeXlsxGenerator () {
		makeOfficeGenerator ( 'xl', 'workbook', {} );

		gen_private.perment.features.page_name = 'sheets'; // This document type must have pages.

		// On each generate we'll prepare the shared strings list:
		gen_private.perment.features.call_before_gen = cbPrepareXlsxToGenerate;

		gen_private.mixed.files_list.push (
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

		gen_private.mixed.rels_app.push (
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

		intAddAnyResourceToParse ( 'docProps\\app.xml', 'buffer', null, cbMakeXlsApp, true );
		intAddAnyResourceToParse ( 'xl\\styles.xml', 'buffer', null, cbMakeXlsStyles, true );
		intAddAnyResourceToParse ( 'xl\\workbook.xml', 'buffer', null, cbMakeXlsWorkbook, true );

		intAddAnyResourceToParse ( 'xl\\_rels\\workbook.xml.rels', 'buffer', gen_private.mixed.rels_app, cbMakeRels, true );

		///
		/// @brief ???.
		///
		/// ???.
		///
		/// @param[in] ??? ???.
		///
		genobj.makeNewSheet = function () {
			var pageNumber = gen_private.thisDoc.pages.length;
			var sheetObj = {}; // The sheet object that the user will use.

			sheetObj.data = []; // Place here all the data.

			gen_private.thisDoc.pages[pageNumber] = {};
			gen_private.thisDoc.pages[pageNumber].id = pageNumber;
			gen_private.thisDoc.pages[pageNumber].relId = gen_private.mixed.rels_app.length + 1;
			gen_private.thisDoc.pages[pageNumber].sheet = sheetObj;

			gen_private.mixed.rels_app.push (
				{
					type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
					target: 'worksheets/sheet' + (pageNumber + 1) + '.xml',
					clear: 'data'
				}
			);

			gen_private.mixed.files_list.push (
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

			intAddAnyResourceToParse ( 'xl\\worksheets\\sheet' + (pageNumber + 1) + '.xml', 'buffer', gen_private.thisDoc.pages[pageNumber], cbMakeXlsSheet, false );

			return sheetObj;
		};
	};

	// ***PUBLIC_CODE***

	// Public API - non plugin based:

	///
	/// @brief ???.
	///
	/// ???.
	///
	this.generate = function ( stream ) {
		if ( gen_private.perment.features.page_name ) {
			if ( gen_private.thisDoc.pages.length == 0 ) {
				throw 'ERROR: No ' + gen_private.perment.features.page_name + ' been found inside your document.';
			} // Endif.
		} // Endif.

		// Optional callback to prepare everything for generating:
		if ( gen_private.perment.features.call_before_gen )
		{
			gen_private.perment.features.call_before_gen ();
		} // Endif.

		var archive = archiver('zip');

		archive.on('error', function(err) {
			throw err;
		});

		archive.pipe ( stream );

		///
		/// @brief Add the next resource into the zip stream.
		///
		/// This function adding the next resource into the zip stream.
		///
		function generateNextResource ( cur_index )
		{
			var resStream;

			if ( cur_index < gen_private.mixed.res_list.length ) {
				if ( typeof gen_private.mixed.res_list[cur_index] != 'undefined' ) {
					switch ( gen_private.mixed.res_list[cur_index].type ) {
						case 'buffer':
							resStream = gen_private.mixed.res_list[cur_index].callback ( gen_private.mixed.res_list[cur_index].data );
							break;

						// BMK_STREAM: (***START***)
						// Using some kind of simple 'template' engine:
						case 'custom':
							resStream = new Stream ();
							resStream.readable = true;
							process.nextTick ( function() {
								// The callback should emit data events and then end event. The problem is that we can't 
								// call emit before the pipe starting to run. That's why we are not executing the callback 
								// immediately but using the process.nextTick trick to make it to run after the pipe is 
								// starting and someone is listening to our events.
								gen_private.mixed.res_list[cur_index].callback ( resStream, gen_private.mixed.res_list[cur_index].data );
							});
							break;
						// BMK_STREAM: (***END***)

						// Just copy the file as is:
						case 'file':
							resStream = fs.createReadStream ( gen_private.mixed.res_list[cur_index].data || gen_private.mixed.res_list[cur_index].name );
							break;

						// Just use this stream:
						case 'stream':
							resStream = gen_private.mixed.res_list[cur_index].data;
							break;
					} // End of switch.

					if ( typeof resStream != 'undefined' ) {
						if ( int_officegen_globals.settings.verbose ) {
							console.log ( 'Adding "' + gen_private.mixed.res_list[cur_index].name + '" (' + gen_private.mixed.res_list[cur_index].type + ')...' );
						} // Endif.

						archive.append ( resStream, { name: gen_private.mixed.res_list[cur_index].name }, function () {
							setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
						});
						

					} else {
						setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
					} // Endif.

				} else {
					setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
				} // Endif.

			} else {
				archive.finalize ( function ( err, written ) {
					// Optional callback to clean after us:
					if ( gen_private.perment.features.call_after_gen )
					{
						gen_private.perment.features.call_after_gen ( err, written );
					} // Endif.

					if (err) {
						throw err;
					} // Endif.

					if ( genobj.options && genobj.options.onend ) {
						genobj.options.onend ( written );
					} // Endif.
				});
			} // Endif.
		};

		// Start the process of generating the output zip stream:
		generateNextResource ( 0 );
	};

	///
	/// @brief Reuse this object for a new document of the same type.
	///
	/// Call this method if you want to start generating a new document of the same type using this object.
	///
	this.startNewDoc = function () {
		var kill = [];

		for ( var i = 0; i < gen_private.mixed.res_list.length; i++ ) {
			if ( !gen_private.mixed.res_list[i].is_perment ) kill.push ( i );
		} // End of for loop.

		for ( var i = 0; i < kill.length; i++ ) gen_private.mixed.res_list.splice ( kill[i] - i, 1 );

		gen_private.thisDoc.pages.length = 0;

		if ( gen_private.perment.features.call_on_clear ) {
			gen_private.perment.features.call_on_clear ();
		} // Endif.
	};

	// Public API - plugin API:

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] resource_name The name of the resource (path).
	/// @param[in] type_of_res The type of this resource: either 'file' or 'buffer'.
	/// @param[in] res_data Optional data to use when creating this resource.
	/// @param[in] res_cb Callback to generate this resource (for 'buffer' mode only).
	///
	this.addResourceToParse = function ( resource_name, type_of_res, res_data, res_cb ) {
		// We don't want the user to add perment resources to the list of resources:
		intAddAnyResourceToParse ( resource_name, type_of_res, res_data, res_cb, false );
	};

	// --- No more function declarations from here ---
	// ***REST_OF_OFFICEGEN_CODE***

	// See the officegen descriptions for the rules of the options:
	setOptions ( options );
	
	// Configure this object depending on the user's selected type:
	setGeneratorType ( genobj.options.type );
};

///
/// @brief ???.
///
/// ???.
///
/// @b Example:
///
/// @code
/// @endcode
///
function makegen ( options ) {
	try {
		return new officegen ( options );

	} catch ( err )
	{
		console.error ( err );
		throw err;
	}
};

///
/// @brief ???.
///
/// ???.
///
/// @b Example:
///
/// @code
/// @endcode
///
function setVerboseMode ( new_state ) {
	int_officegen_globals.settings.verbose = new_state;
};

exports.makegen = makegen;
exports.setVerboseMode = setVerboseMode;
// exports.registerDocType = ???;
exports.schema = int_officegen_globals.types;
exports.version = officegen_info.version;

