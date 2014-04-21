/// @author vtloc
/// @date 2014Jan09
/// This module's purpose is to transform 
/// 
var _ = require('underscore');
var xmlBuilder = require('xmlbuilder');

///
/// @brief Transform an array of string into an office's compliance structure
///
/// @param[in] region String
///		The reference cell of the string, for example: $A$1
/// @param[in] stringArr
///		An array of string, for example: ['foo', 'bar']
///
function strRefFromString(region, stringArr) {
	var obj = {
		'c:strRef': {
			'c:f' : region,
			'c:strCache': function() {
				var result = {};
				result['c:ptCount'] = { '@val' : stringArr.length };
				result['#list'] = [];
				for( var i=0; i<stringArr.length; i++ )
				{
					result['#list'].push({'c:pt' : { '@idx':i, 'c:v': stringArr[i] }});
				}
				
				return result;
			}
			
		}
	}
	
	return obj;
}

///
/// @brief Transform an array of string into an office's compliance structure
///
/// @param[in] region String
///		The reference cell of the string, for example: $A$1
/// @param[in] stringArr
///		An array of numArr, for example: [4, 7, 8]
/// @param[in] formatCode
///		A string describe the number's format. Example: General
///
function numRef(region, numArr, formatCode) {
	var obj = {
		'c:numRef' : {
			'c:f' : region,
			'c:numCache' : {
				'c:formatCode' : formatCode,
				'c:ptCount' : {'@val': ''+numArr.length},
				'#list' : function() {
					result= [];
					for( var i=0; i<numArr.length; i++ )
					{
						result.push({'c:pt' : { '@idx':i, 'c:v': numArr[i].toString() }});
					}
					
					return result;
				}
			}
		}
	};
	
	return obj;
}

///
/// @brief Transform an array of string into an office's compliance structure
///
/// @param[in] colorArr
///     An array of colorArr, for example: ['ff0000', '00ff00', '0000ff']
///
function colorRef(colorArr) {
    var arr = [];
    for( var i=0; i<colorArr.length; i++ )
    {
        arr.push({
            'c:dPt' : {
                'c:idx': {'@val': i},
                'c:bubble3D': {'@val': 0},
                'c:spPr': {
                    'a:solidFill': {
                        'a:srgbClr': {'@val': colorArr[i].toString()}
                    }
                }
            }
        });
    }

    return arr;
}

///
/// @brief Transform an array of string into an office's compliance structure
///
/// @param[in] row int
///		Row index.
/// @param[in] col int
///		Col index.
/// @param[in] isRowAbsolute boolean
///		Will add $ into cell's address if this parameter is true.
/// @param[in] isColAbsolute boolean
///		Will add $ into cell's address if this parameter is true.
///
function rowColToSheetAddress(row, col, isRowAbsolute, isColAbsolute) {
	var address = "";
	
	if( isColAbsolute )
		address += '$';
		
	// these lines of code will transform the number 1-26 into A->Z
	// used in excel's cell's coordination
	while(col > 0 )
	{
		var num = col % 26;
		col = (col - num ) / 26;
		address += String.fromCharCode(65+num-1);
	}
	
	if( isRowAbsolute )
		address += '$';
		
	address += row;
	
	return address;
}

///
/// @brief Transform a data object into column chart
///
/// @param[in] data object
///		{ 	
///			title: 'eSurvey chart',
///			data:  [ // array of series
///				{
///					name: 'Income',
///					labels: ['2005', '2006', '2007', '2008', '2009'],
///					values: [23.5, 26.2, 30.1, 29.5, 24.6],
///                 colors: ['ff0000', '00ff00', '0000ff', 'ffff00', '00ffff'] // optional
///				},
///				{
///					name: 'Expense',
///					labels: ['2005', '2006', '2007', '2008', '2009'],
///					values: [18.1, 22.8, 23.9, 25.1, 25],
///                 colors: ['ff0000', '00ff00', '0000ff', 'ffff00', '00ffff'] // optional
///				}
///			]
/// 	}
///
function makeColumnChartStringFromData( data ) {

	var series = [];	
	var rc2a = rowColToSheetAddress; // shortcut
	
	for( var i=0; i< data['data'].length; i++ )
	{
		var serie = data['data'][i];
		var serieData = {
			'c:ser' : {
				'c:idx' : {'@val': i},
				'c:order' : {'@val': i},
				'c:tx' : strRefFromString('Sheet1!' + rc2a(1,2+i,true, true), [serie.name]), // serie's value
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
        if (serie.color) {
            serieData['c:ser']['c:spPr'] = {
                    'a:solidFill': {
                        'a:srgbClr': {'@val': serie.color}
                    }
            };
        }
		var root = xmlBuilder.create(serieData, {headless: true});
	
		series.push(root.end({ pretty: true, indent: '  ', newline: '\n' }));
	}
	
	var seriesString = "";
	
	for( var i=0; i<data['data'].length; i++ )
	{
		seriesString += series[i] + '\n';
	}
	
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\
<c:date1904 val="1"/>\
<c:lang val="en-US"/>\
<c:chart>\
    <c:plotArea>\
        <c:layout/>\
        <c:barChart>\
            <c:barDir val="col"/>\
            <c:grouping val="clustered"/>'
            + seriesString + 
            '<c:axId val="45021824"/>\
            <c:axId val="45291008"/>\
        </c:barChart>\
        <c:catAx>\
            <c:axId val="45021824"/>\
            <c:scaling>\
                <c:orientation val="minMax"/>\
            </c:scaling>\
            <c:axPos val="b"/>\
            <c:numFmt formatCode="General" sourceLinked="1"/>\
            <c:tickLblPos val="nextTo"/>\
            <c:crossAx val="45291008"/>\
            <c:crosses val="autoZero"/>\
            <c:auto val="1"/>\
            <c:lblAlgn val="ctr"/>\
            <c:lblOffset val="100"/>\
        </c:catAx>\
        <c:valAx>\
            <c:axId val="45291008"/>\
            <c:scaling>\
                <c:orientation val="minMax"/>\
            </c:scaling>\
            <c:axPos val="l"/>\
            <c:majorGridlines/>\
            <c:numFmt formatCode="General" sourceLinked="1"/>\
            <c:tickLblPos val="nextTo"/>\
            <c:crossAx val="45021824"/>\
            <c:crosses val="autoZero"/>\
            <c:crossBetween val="between"/>\
        </c:valAx>\
    </c:plotArea>\
    <c:legend>\
        <c:legendPos val="r"/>\
        <c:layout/>\
    </c:legend>\
    <c:plotVisOnly val="1"/>\
</c:chart>\
<c:txPr>\
    <a:bodyPr/>\
    <a:lstStyle/>\
    <a:p>\
        <a:pPr>\
            <a:defRPr sz="1800"/>\
        </a:pPr>\
        <a:endParaRPr lang="en-US"/>\
    </a:p>\
</c:txPr>\
<c:externalData r:id="rId1"/>\
</c:chartSpace>'
}

///
/// @brief Transform a data object into pie chart
///
/// @param[in] data object
///		{ 	
///			title: 'eSurvey chart',
///			data:  [ // array of series
///				{
///					name: 'Income',
///					labels: ['2005', '2006', '2007', '2008', '2009'],
///					values: [23.5, 26.2, 30.1, 29.5, 24.6],
///                 colors: ['ff0000', '00ff00', '0000ff', 'ffff00', '00ffff'] // optional
///				}
///			]
/// 	}
///
function makePieChartStringFromData( data ) {
  var series = [];	
	var rc2a = rowColToSheetAddress; // shortcut
	
	for( var i=0; i< data['data'].length; i++ )
	{
		var serie = data['data'][i];
		var serieData = {
			'c:ser' : {
				'c:idx' : {'@val': i},
				'c:order' : {'@val': i},
				'c:tx' : strRefFromString('Sheet1!' + rc2a(1,2+i,true, true), [serie.name]), // serie's value
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
        if (serie.colors) {
            serieData['c:ser']['#list'] = colorRef(serie.colors);
        }
		var root = xmlBuilder.create(serieData, {headless: true});
	
		series.push(root.end({ pretty: true, indent: '  ', newline: '\n' }));
	}
	
	var seriesString = "";
	
	for( var i=0; i<data['data'].length; i++ )
	{
		seriesString += series[i] + '\n';
	}
	
	return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\
<c:lang val="en-US"/>\
<c:chart>\
    <c:title>\
        <c:layout/>\
    </c:title>\
    <c:plotArea>\
        <c:layout/>\
        <c:pieChart>\
            <c:varyColors val="1"/>'
            + seriesString +
            '<c:firstSliceAng val="0"/>\
        </c:pieChart>\
    </c:plotArea>\
    <c:legend>\
        <c:legendPos val="r"/>\
        <c:layout/>\
    </c:legend>\
    <c:plotVisOnly val="1"/>\
</c:chart>\
<c:txPr>\
    <a:bodyPr/>\
    <a:lstStyle/>\
    <a:p>\
        <a:pPr>\
            <a:defRPr sz="1800"/>\
        </a:pPr>\
        <a:endParaRPr lang="en-US"/>\
    </a:p>\
</c:txPr>\
<c:externalData r:id="rId1"/>\
</c:chartSpace>'
}

///
/// @brief Transform a data object into pie chart
///
/// @param[in] data object
///		{ 	
///			title: 'eSurvey chart',
///			data:  [ // array of series
///				{
///					name: 'Income',
///					labels: ['2005', '2006', '2007', '2008', '2009'],
///					values: [23.5, 26.2, 30.1, 29.5, 24.6],
///                 colors: ['ff0000', '00ff00', '0000ff', 'ffff00', '00ffff'] // optional
///				}
///			]
/// 	}
///
function makeGroupBarChartStringFromData( data ) {
  var series = [];	
	var rc2a = rowColToSheetAddress; // shortcut
	
	for( var i=0; i< data['data'].length; i++ )
	{
		var serie = data['data'][i];
		var serieData = {
			'c:ser' : {
				'c:idx' : {'@val': i},
				'c:order' : {'@val': i},
				'c:tx' : strRefFromString('Sheet1!' + rc2a(1,2+i,true, true), [serie.name]), // serie's value
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
        if (serie.color) {
            serieData['c:ser']['c:spPr'] = {
                    'a:solidFill': {
                        'a:srgbClr': {'@val': serie.color}
                    }
            };
        }
		var root = xmlBuilder.create(serieData, {headless: true});
	
		series.push(root.end({ pretty: true, indent: '  ', newline: '\n' }));
	}
	
	var seriesString = "";
	
	for( var i=0; i<data['data'].length; i++ )
	{
		seriesString += series[i] + '\n';
	}
  
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\
<c:lang val="en-US"/>\
<c:chart>\
    <c:plotArea>\
        <c:layout/>\
        <c:barChart>\
            <c:barDir val="bar"/>\
            <c:grouping val="percentStacked"/>'
            + seriesString +
            '<c:overlap val="100"/>\
            <c:axId val="65025536"/>\
            <c:axId val="65037824"/>\
        </c:barChart>\
        <c:catAx>\
            <c:axId val="65025536"/>\
            <c:scaling>\
                <c:orientation val="minMax"/>\
            </c:scaling>\
            <c:axPos val="l"/>\
            <c:tickLblPos val="nextTo"/>\
            <c:crossAx val="65037824"/>\
            <c:crosses val="autoZero"/>\
            <c:auto val="1"/>\
            <c:lblAlgn val="ctr"/>\
            <c:lblOffset val="100"/>\
        </c:catAx>\
        <c:valAx>\
            <c:axId val="65037824"/>\
            <c:scaling>\
                <c:orientation val="minMax"/>\
            </c:scaling>\
            <c:axPos val="b"/>\
            <c:majorGridlines/>\
            <c:numFmt formatCode="0%" sourceLinked="1"/>\
            <c:tickLblPos val="nextTo"/>\
            <c:crossAx val="65025536"/>\
            <c:crosses val="autoZero"/>\
            <c:crossBetween val="between"/>\
        </c:valAx>\
    </c:plotArea>\
    <c:legend>\
        <c:legendPos val="r"/>\
        <c:layout/>\
    </c:legend>\
    <c:plotVisOnly val="1"/>\
</c:chart>\
<c:txPr>\
    <a:bodyPr/>\
    <a:lstStyle/>\
    <a:p>\
        <a:pPr>\
            <a:defRPr sz="1800"/>\
        </a:pPr>\
        <a:endParaRPr lang="en-US"/>\
    </a:p>\
</c:txPr>\
<c:externalData r:id="rId1"/>\
</c:chartSpace>';
}


///
/// @brief Transform a data object into pie chart
///
/// @param[in] data object
///		{ 	
///			title: 'eSurvey chart',
///			data:  [ // array of series
///				{
///					name: 'Income',
///					labels: ['2005', '2006', '2007', '2008', '2009'],
///					values: [23.5, 26.2, 30.1, 29.5, 24.6],
///                 color: 'ff0000'
///				}
///			]
/// 	}
///
function makeBarChartStringFromData( data ) {
  var series = [];	
	var rc2a = rowColToSheetAddress; // shortcut
	
	for( var i=0; i< data['data'].length; i++ )
	{
		var serie = data['data'][i];
		var serieData = {
			'c:ser' : {
				'c:idx' : {'@val': i},
				'c:order' : {'@val': i},
				'c:tx' : strRefFromString('Sheet1!' + rc2a(1,2+i,true, true), [serie.name]), // serie's value
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
        if (serie.color) {
            serieData['c:ser']['c:spPr'] = {
                    'a:solidFill': {
                        'a:srgbClr': {'@val': serie.color}
                    }
            };
        }
		var root = xmlBuilder.create(serieData, {headless: true});
	
		series.push(root.end({ pretty: true, indent: '  ', newline: '\n' }));
	}
	
	var seriesString = "";
	
	for( var i=0; i<data['data'].length; i++ )
	{
		seriesString += series[i] + '\n';
	}
  
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\
<c:lang val="en-US"/>\
<c:chart>\
    <c:plotArea>\
        <c:layout/>\
        <c:barChart>\
            <c:barDir val="bar"/>\
            <c:grouping val="clustered"/>'
            + seriesString +
            '<c:axId val="64451712"/>\
            <c:axId val="64453248"/>\
        </c:barChart>\
        <c:catAx>\
            <c:axId val="64451712"/>\
            <c:scaling>\
                <c:orientation val="minMax"/>\
            </c:scaling>\
            <c:axPos val="l"/>\
            <c:tickLblPos val="nextTo"/>\
            <c:crossAx val="64453248"/>\
            <c:crosses val="autoZero"/>\
            <c:auto val="1"/>\
            <c:lblAlgn val="ctr"/>\
            <c:lblOffset val="100"/>\
        </c:catAx>\
        <c:valAx>\
            <c:axId val="64453248"/>\
            <c:scaling>\
                <c:orientation val="minMax"/>\
            </c:scaling>\
            <c:axPos val="b"/>\
            <c:majorGridlines/>\
            <c:numFmt formatCode="General" sourceLinked="1"/>\
            <c:tickLblPos val="nextTo"/>\
            <c:crossAx val="64451712"/>\
            <c:crosses val="autoZero"/>\
            <c:crossBetween val="between"/>\
        </c:valAx>\
    </c:plotArea>\
    <c:legend>\
        <c:legendPos val="r"/>\
        <c:layout/>\
    </c:legend>\
    <c:plotVisOnly val="1"/>\
</c:chart>\
<c:txPr>\
    <a:bodyPr/>\
    <a:lstStyle/>\
    <a:p>\
        <a:pPr>\
            <a:defRPr sz="1800"/>\
        </a:pPr>\
        <a:endParaRPr lang="en-US"/>\
    </a:p>\
</c:txPr>\
<c:externalData r:id="rId1"/>\
</c:chartSpace>';
}


exports.makePieChartStringFromData = makePieChartStringFromData;
exports.makeColumnChartStringFromData = makeColumnChartStringFromData;
exports.makeBarChartStringFromData = makeBarChartStringFromData;
exports.makeGroupBarChartStringFromData = makeGroupBarChartStringFromData;