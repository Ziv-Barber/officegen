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
				'c:ptCount' : {'@val': '4'},
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
///					values: [23.5, 26.2, 30.1, 29.5, 24.6]
///				},
///				{
///					name: 'Expense',
///					labels: ['2005', '2006', '2007', '2008', '2009'],
///					values: [18.1, 22.8, 23.9, 25.1, 25]
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
				'c:spPr' : {
					 'a:solidFill': {
						 'a:schemeClr' : {'@val': 'accent' + (i % 5 + 1)}, // adjust the colors
						  'a:ln' : {
							 'a:noFill': ''
						  },
						  'a:effectLst': ''
					 }
				 },
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
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
	<c:date1904 val="0"/>\
	<c:lang val="en-US"/>\
	<c:roundedCorners val="0"/>\
	<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\
		<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">\
			<c14:style val="102"/>\
		</mc:Choice>\
		<mc:Fallback>\
			<c:style val="2"/>\
		</mc:Fallback>\
	</mc:AlternateContent>\
	<c:chart>\
		<c:title>\
			<c:tx>\
				<c:rich>\
					<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
					<a:lstStyle/>\
					<a:p>\
						<a:pPr>\
							<a:defRPr sz="1862" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">\
								<a:solidFill>\
									<a:schemeClr val="tx1">\
										<a:lumMod val="65000"/>\
										<a:lumOff val="35000"/>\
									</a:schemeClr>\
								</a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							</a:defRPr>\
						</a:pPr>\
						<a:r>\
							<a:rPr lang="en-US" smtClean="0"/>\
							<a:t>' + data['title'] + '</a:t>\
						</a:r>\
						<a:endParaRPr lang="en-US" dirty="0"/>\
					</a:p>\
				</c:rich>\
			</c:tx>\
		</c:title>\
		<c:autoTitleDeleted val="0"/>\
		<c:plotArea>\
			<c:layout/>\
			<c:barChart>\
				<c:barDir val="col"/>\
				<c:grouping val="clustered"/>\
				<c:varyColors val="0"/>'
				+ seriesString +
				'<c:dLbls>\
					<c:showLegendKey val="0"/>\
					<c:showVal val="1"/>\
					<c:showCatName val="0"/>\
					<c:showSerName val="0"/>\
					<c:showPercent val="0"/>\
					<c:showBubbleSize val="0"/>\
				</c:dLbls>\
				<c:gapWidth val="219"/>\
				<c:overlap val="-27"/>\
				<c:axId val="192875056"/>\
				<c:axId val="192876232"/>\
			</c:barChart>\
			<c:catAx>\
				<c:axId val="192875056"/>\
				<c:scaling>\
					<c:orientation val="minMax"/>\
				</c:scaling>\
				<c:delete val="0"/>\
				<c:axPos val="b"/>\
				<c:numFmt formatCode="General" sourceLinked="1"/>\
				<c:majorTickMark val="none"/>\
				<c:minorTickMark val="none"/>\
				<c:tickLblPos val="nextTo"/>\
				<c:spPr>\
					<a:noFill/>\
					<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">\
						<a:solidFill>\
							<a:schemeClr val="tx1">\
								<a:lumMod val="15000"/>\
								<a:lumOff val="85000"/>\
							</a:schemeClr>\
						</a:solidFill>\
						<a:round/>\
					</a:ln>\
					<a:effectLst/>\
				</c:spPr>\
				<c:txPr>\
					<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
					<a:lstStyle/>\
					<a:p>\
						<a:pPr>\
							<a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
								<a:solidFill>\
									<a:schemeClr val="tx1">\
										<a:lumMod val="65000"/>\
										<a:lumOff val="35000"/>\
									</a:schemeClr>\
								</a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							</a:defRPr>\
						</a:pPr>\
						<a:endParaRPr lang="en-US"/>\
					</a:p>\
				</c:txPr>\
				<c:crossAx val="192876232"/>\
				<c:crosses val="autoZero"/>\
				<c:auto val="1"/>\
				<c:lblAlgn val="ctr"/>\
				<c:lblOffset val="100"/>\
				<c:noMultiLvlLbl val="0"/>\
			</c:catAx>\
			<c:valAx>\
				<c:axId val="192876232"/>\
				<c:scaling>\
					<c:orientation val="minMax"/>\
				</c:scaling>\
				<c:delete val="0"/>\
				<c:axPos val="l"/>\
				<c:majorGridlines>\
					<c:spPr>\
						<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">\
							<a:solidFill>\
								<a:schemeClr val="tx1">\
									<a:lumMod val="15000"/>\
									<a:lumOff val="85000"/>\
								</a:schemeClr>\
							</a:solidFill>\
							<a:round/>\
						</a:ln>\
						<a:effectLst/>\
					</c:spPr>\
				</c:majorGridlines>\
				<c:numFmt formatCode="General" sourceLinked="1"/>\
				<c:majorTickMark val="none"/>\
				<c:minorTickMark val="none"/>\
				<c:tickLblPos val="nextTo"/>\
				<c:spPr>\
					<a:noFill/>\
					<a:ln>\
						<a:noFill/>\
					</a:ln>\
					<a:effectLst/>\
				</c:spPr>\
				<c:txPr>\
					<a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
					<a:lstStyle/>\
					<a:p>\
						<a:pPr>\
							<a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
								<a:solidFill>\
									<a:schemeClr val="tx1">\
										<a:lumMod val="65000"/>\
										<a:lumOff val="35000"/>\
									</a:schemeClr>\
								</a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							</a:defRPr>\
						</a:pPr>\
						<a:endParaRPr lang="en-US"/>\
					</a:p>\
				</c:txPr>\
				<c:crossAx val="192875056"/>\
				<c:crosses val="autoZero"/>\
				<c:crossBetween val="between"/>\
			</c:valAx>\
			<c:spPr>\
				<a:noFill/>\
				<a:ln>\
					<a:noFill/>\
				</a:ln>\
				<a:effectLst/>\
			</c:spPr>\
		</c:plotArea>\
		<c:legend>\
			<c:legendPos val="b"/>\
			<c:layout/>\
			<c:overlay val="0"/>\
			<c:spPr>\
				<a:noFill/>\
				<a:ln>\
					<a:noFill/>\
				</a:ln>\
				<a:effectLst/>\
			</c:spPr>\
			<c:txPr>\
				<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
				<a:lstStyle/>\
				<a:p>\
					<a:pPr>\
						<a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
							<a:solidFill>\
								<a:schemeClr val="tx1">\
									<a:lumMod val="65000"/>\
									<a:lumOff val="35000"/>\
								</a:schemeClr>\
							</a:solidFill>\
							<a:latin typeface="+mn-lt"/>\
							<a:ea typeface="+mn-ea"/>\
							<a:cs typeface="+mn-cs"/>\
						</a:defRPr>\
					</a:pPr>\
					<a:endParaRPr lang="en-US"/>\
				</a:p>\
			</c:txPr>\
		</c:legend>\
		<c:plotVisOnly val="1"/>\
		<c:dispBlanksAs val="gap"/>\
		<c:showDLblsOverMax val="0"/>\
	</c:chart>\
	<c:spPr>\
		<a:noFill/>\
		<a:ln>\
			<a:noFill/>\
		</a:ln>\
		<a:effectLst/>\
	</c:spPr>\
	<c:txPr>\
		<a:bodyPr/>\
		<a:lstStyle/>\
		<a:p>\
			<a:pPr>\
				<a:defRPr/>\
			</a:pPr>\
			<a:endParaRPr lang="en-US"/>\
		</a:p>\
	</c:txPr>\
	<c:externalData r:id="rId3">\
		<c:autoUpdate val="0"/>\
	</c:externalData>\
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
///					values: [23.5, 26.2, 30.1, 29.5, 24.6]
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
				'#list' : function() {
					var result = [];
					var obj = {};
					for(var i=0; i<data['data'][0].values.length; i++){
						obj = {
							'c:dPt' : {
								'c:idx' : { '@val': i },
								'c:bubble3D' : { '@val': '0' },
								'c:spPr' : { 
									'a:solidFill' : {
										'a:schemeClr' : { '@val' : 'accent' + (i % 5 + 1) } // because we only have accent1 to accent6
									},
									'a:ln' : { 
										'@w' : '19050',
										'a:solidFill' : {
											'a:schemeClr' : { '@val' : 'lt1' }
										}
									},
									'a:effectLst' : ''
								},
							}
						};
						
						result.push(obj);
					}
					
					return result;
				},
				'c:spPr' : {
					 'a:solidFill': {
						 'a:schemeClr' : {'@val': 'accent' + (i % 5 + 1)}, // adjust the colors
						  'a:ln' : {
							 'a:noFill': ''
						  },
						  'a:effectLst': ''
					 }
				 },
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
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
	<c:date1904 val="0"/>\
	<c:lang val="en-US"/>\
	<c:roundedCorners val="0"/>\
	<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\
		<mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">\
			<c14:style val="102"/>\
		</mc:Choice>\
		<mc:Fallback>\
			<c:style val="2"/>\
		</mc:Fallback>\
	</mc:AlternateContent>\
	<c:chart>\
		<c:title>\
			<c:layout/>\
			<c:overlay val="0"/>\
			<c:spPr>\
				<a:noFill/>\
				<a:ln>\
					<a:noFill/>\
				</a:ln>\
				<a:effectLst/>\
			</c:spPr>\
			<c:txPr>\
				<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
				<a:lstStyle/>\
				<a:p>\
					<a:pPr>\
						<a:defRPr sz="1862" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">\
							<a:solidFill>\
								<a:schemeClr val="tx1">\
									<a:lumMod val="65000"/>\
									<a:lumOff val="35000"/>\
								</a:schemeClr>\
							</a:solidFill>\
							<a:latin typeface="+mn-lt"/>\
							<a:ea typeface="+mn-ea"/>\
							<a:cs typeface="+mn-cs"/>\
						</a:defRPr>\
					</a:pPr>\
					<a:endParaRPr lang="en-US"/>\
				</a:p>\
			</c:txPr>\
		</c:title>\
		<c:autoTitleDeleted val="0"/>\
		<c:plotArea>\
			<c:layout/>\
			<c:pieChart>\
				<c:varyColors val="1"/>'
				
				+ seriesString +
				
				'<c:dLbls>\
					<c:showLegendKey val="0"/>\
					<c:showVal val="0"/>\
					<c:showCatName val="0"/>\
					<c:showSerName val="0"/>\
					<c:showPercent val="0"/>\
					<c:showBubbleSize val="0"/>\
					<c:showLeaderLines val="1"/>\
				</c:dLbls>\
				<c:firstSliceAng val="0"/>\
			</c:pieChart>\
			<c:spPr>\
				<a:noFill/>\
				<a:ln>\
					<a:noFill/>\
				</a:ln>\
				<a:effectLst/>\
			</c:spPr>\
		</c:plotArea>\
		<c:legend>\
			<c:legendPos val="b"/>\
			<c:layout/>\
			<c:overlay val="0"/>\
			<c:spPr>\
				<a:noFill/>\
				<a:ln>\
					<a:noFill/>\
				</a:ln>\
				<a:effectLst/>\
			</c:spPr>\
			<c:txPr>\
				<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
				<a:lstStyle/>\
				<a:p>\
					<a:pPr>\
						<a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
							<a:solidFill>\
								<a:schemeClr val="tx1">\
									<a:lumMod val="65000"/>\
									<a:lumOff val="35000"/>\
								</a:schemeClr>\
							</a:solidFill>\
							<a:latin typeface="+mn-lt"/>\
							<a:ea typeface="+mn-ea"/>\
							<a:cs typeface="+mn-cs"/>\
						</a:defRPr>\
					</a:pPr>\
					<a:endParaRPr lang="en-US"/>\
				</a:p>\
			</c:txPr>\
		</c:legend>\
		<c:plotVisOnly val="1"/>\
		<c:dispBlanksAs val="gap"/>\
		<c:showDLblsOverMax val="0"/>\
	</c:chart>\
	<c:spPr>\
		<a:noFill/>\
		<a:ln>\
			<a:noFill/>\
		</a:ln>\
		<a:effectLst/>\
	</c:spPr>\
	<c:txPr>\
		<a:bodyPr/>\
		<a:lstStyle/>\
		<a:p>\
			<a:pPr>\
				<a:defRPr/>\
			</a:pPr>\
			<a:endParaRPr lang="en-US"/>\
		</a:p>\
	</c:txPr>\
	<c:externalData r:id="rId3">\
		<c:autoUpdate val="0"/>\
	</c:externalData>\
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
///					values: [23.5, 26.2, 30.1, 29.5, 24.6]
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
				'c:spPr' : {
					 'a:solidFill': {
						 'a:schemeClr' : {'@val': 'accent' + (i % 5 + 1)}, // adjust the colors
						  'a:ln' : {
							 'a:noFill': ''
						  },
						  'a:effectLst': ''
					 }
				 },
				 'c:invertIfNegative' : {'@val':'0'},
				 'c:cat' : strRefFromString('Sheet1!'+ rc2a(2,1,true, true) + ':' + rc2a(2+serie.labels.length-1,1,true, true), serie.labels),
				 'c:val' : numRef('Sheet1!'+rc2a(2,2+i,true, true) + ':' + rc2a(2+serie.labels.length-1,2+i,true, true), serie.values, "General")
			}
		};
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
  <c:date1904 val="0"/>\
  <c:lang val="en-US"/>\
  <c:roundedCorners val="0"/>\
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">\
    <mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">\
      <c14:style val="102"/>\
    </mc:Choice>\
    <mc:Fallback>\
      <c:style val="2"/>\
    </mc:Fallback>\
  </mc:AlternateContent>\
  <c:chart>\
    <c:title>\
			<c:tx>\
				<c:rich>\
					<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
					<a:lstStyle/>\
					<a:p>\
						<a:pPr>\
							<a:defRPr sz="1862" b="0" i="0" u="none" strike="noStrike" kern="1200" spc="0" baseline="0">\
								<a:solidFill>\
									<a:schemeClr val="tx1">\
										<a:lumMod val="65000"/>\
										<a:lumOff val="35000"/>\
									</a:schemeClr>\
								</a:solidFill>\
								<a:latin typeface="+mn-lt"/>\
								<a:ea typeface="+mn-ea"/>\
								<a:cs typeface="+mn-cs"/>\
							</a:defRPr>\
						</a:pPr>\
						<a:r>\
							<a:rPr lang="en-US" smtClean="0"/>\
							<a:t>' + data['title'] + '</a:t>\
						</a:r>\
						<a:endParaRPr lang="en-US" dirty="0"/>\
					</a:p>\
				</c:rich>\
			</c:tx>\
		</c:title>\
    <c:autoTitleDeleted val="0"/>\
    <c:plotArea>\
      <c:layout/>\
      <c:barChart>\
        <c:barDir val="bar"/>\
        <c:grouping val="stacked"/>\
        <c:varyColors val="0"/>'
        
        + seriesString +
        
        '<c:dLbls>\
          <c:showLegendKey val="0"/>\
          <c:showVal val="0"/>\
          <c:showCatName val="0"/>\
          <c:showSerName val="0"/>\
          <c:showPercent val="0"/>\
          <c:showBubbleSize val="0"/>\
        </c:dLbls>\
        <c:gapWidth val="150"/>\
        <c:overlap val="100"/>\
        <c:axId val="197796640"/>\
        <c:axId val="197797032"/>\
      </c:barChart>\
      <c:catAx>\
        <c:axId val="197796640"/>\
        <c:scaling>\
          <c:orientation val="minMax"/>\
        </c:scaling>\
        <c:delete val="0"/>\
        <c:axPos val="l"/>\
        <c:numFmt formatCode="General" sourceLinked="1"/>\
        <c:majorTickMark val="none"/>\
        <c:minorTickMark val="none"/>\
        <c:tickLblPos val="nextTo"/>\
        <c:spPr>\
          <a:noFill/>\
          <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">\
            <a:solidFill>\
              <a:schemeClr val="tx1">\
                <a:lumMod val="15000"/>\
                <a:lumOff val="85000"/>\
              </a:schemeClr>\
            </a:solidFill>\
            <a:round/>\
          </a:ln>\
          <a:effectLst/>\
        </c:spPr>\
        <c:txPr>\
          <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
          <a:lstStyle/>\
          <a:p>\
            <a:pPr>\
              <a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
                <a:solidFill>\
                  <a:schemeClr val="tx1">\
                    <a:lumMod val="65000"/>\
                    <a:lumOff val="35000"/>\
                  </a:schemeClr>\
                </a:solidFill>\
                <a:latin typeface="+mn-lt"/>\
                <a:ea typeface="+mn-ea"/>\
                <a:cs typeface="+mn-cs"/>\
              </a:defRPr>\
            </a:pPr>\
            <a:endParaRPr lang="en-US"/>\
          </a:p>\
        </c:txPr>\
        <c:crossAx val="197797032"/>\
        <c:crosses val="autoZero"/>\
        <c:auto val="1"/>\
        <c:lblAlgn val="ctr"/>\
        <c:lblOffset val="100"/>\
        <c:noMultiLvlLbl val="0"/>\
      </c:catAx>\
      <c:valAx>\
        <c:axId val="197797032"/>\
        <c:scaling>\
          <c:orientation val="minMax"/>\
        </c:scaling>\
        <c:delete val="0"/>\
        <c:axPos val="b"/>\
        <c:majorGridlines>\
          <c:spPr>\
            <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">\
              <a:solidFill>\
                <a:schemeClr val="tx1">\
                  <a:lumMod val="15000"/>\
                  <a:lumOff val="85000"/>\
                </a:schemeClr>\
              </a:solidFill>\
              <a:round/>\
            </a:ln>\
            <a:effectLst/>\
          </c:spPr>\
        </c:majorGridlines>\
        <c:numFmt formatCode="General" sourceLinked="1"/>\
        <c:majorTickMark val="none"/>\
        <c:minorTickMark val="none"/>\
        <c:tickLblPos val="nextTo"/>\
        <c:spPr>\
          <a:noFill/>\
          <a:ln>\
            <a:noFill/>\
          </a:ln>\
          <a:effectLst/>\
        </c:spPr>\
        <c:txPr>\
          <a:bodyPr rot="-60000000" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
          <a:lstStyle/>\
          <a:p>\
            <a:pPr>\
              <a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
                <a:solidFill>\
                  <a:schemeClr val="tx1">\
                    <a:lumMod val="65000"/>\
                    <a:lumOff val="35000"/>\
                  </a:schemeClr>\
                </a:solidFill>\
                <a:latin typeface="+mn-lt"/>\
                <a:ea typeface="+mn-ea"/>\
                <a:cs typeface="+mn-cs"/>\
              </a:defRPr>\
            </a:pPr>\
            <a:endParaRPr lang="en-US"/>\
          </a:p>\
        </c:txPr>\
        <c:crossAx val="197796640"/>\
        <c:crosses val="autoZero"/>\
        <c:crossBetween val="between"/>\
      </c:valAx>\
      <c:spPr>\
        <a:noFill/>\
        <a:ln>\
          <a:noFill/>\
        </a:ln>\
        <a:effectLst/>\
      </c:spPr>\
    </c:plotArea>\
    <c:legend>\
      <c:legendPos val="b"/>\
      <c:layout/>\
      <c:overlay val="0"/>\
      <c:spPr>\
        <a:noFill/>\
        <a:ln>\
          <a:noFill/>\
        </a:ln>\
        <a:effectLst/>\
      </c:spPr>\
      <c:txPr>\
        <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>\
        <a:lstStyle/>\
        <a:p>\
          <a:pPr>\
            <a:defRPr sz="1197" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">\
              <a:solidFill>\
                <a:schemeClr val="tx1">\
                  <a:lumMod val="65000"/>\
                  <a:lumOff val="35000"/>\
                </a:schemeClr>\
              </a:solidFill>\
              <a:latin typeface="+mn-lt"/>\
              <a:ea typeface="+mn-ea"/>\
              <a:cs typeface="+mn-cs"/>\
            </a:defRPr>\
          </a:pPr>\
          <a:endParaRPr lang="en-US"/>\
        </a:p>\
      </c:txPr>\
    </c:legend>\
    <c:plotVisOnly val="1"/>\
    <c:dispBlanksAs val="gap"/>\
    <c:showDLblsOverMax val="0"/>\
  </c:chart>\
  <c:spPr>\
    <a:noFill/>\
    <a:ln>\
      <a:noFill/>\
    </a:ln>\
    <a:effectLst/>\
  </c:spPr>\
  <c:txPr>\
    <a:bodyPr/>\
    <a:lstStyle/>\
    <a:p>\
      <a:pPr>\
        <a:defRPr/>\
      </a:pPr>\
      <a:endParaRPr lang="en-US"/>\
    </a:p>\
  </c:txPr>\
  <c:externalData r:id="rId3">\
    <c:autoUpdate val="0"/>\
  </c:externalData>\
</c:chartSpace>';
}

exports.makePieChartStringFromData = makePieChartStringFromData;
exports.makeColumnChartStringFromData = makeColumnChartStringFromData;
exports.makeBarChartStringFromData = makeBarChartStringFromData;