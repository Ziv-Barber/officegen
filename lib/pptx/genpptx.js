//
// officegen: All the code to generate PPTX/PPTS files.
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
 * Basicgen plugin to create pptx files (Microsoft PowerPoint).
 */

var baseobj = require('../core/index.js')
var OfficeChart = require('./officechart.js')
var msdoc = require('../msdoc/msofficegen.js')
var shapes = require('./shapes.js')
var pptxFields = require('./pptxfields.js')
var officeTable = require('./genofficetable')
var path = require('path')
var fast_image_size = require('fast-image-size')
var excelbuilder = require('./msexcel-builder.js')
var xmlBuilder = require('xmlbuilder')
var docplugman = require('../core/docplug')

// Officegen pptx plugins:
var plugWidescreen = require('./pptxplg-widescreen')
var plugSpeakernotes = require('./pptxplg-speakernotes')
var plugLayouts = require('./pptxplg-layouts')
// BMK_PPTX_PLUG:

var GLOBAL_CHART_COUNT = 0

/**
 * Extend officegen object with PPTX/PPSX support.
 *
 * This method extending the given officegen object to create PPTX/PPSX document.
 *
 * @param {object} genobj The object to extend.
 * @param {string} new_type The type of object to create.
 * @param {object} options The object's options.
 * @param {object} gen_private Access to the internals of this object.
 * @param {object} type_info Additional information about this type.
 * @constructor
 * @name makePptx
 */
function makePptx(genobj, new_type, options, gen_private, type_info) {
  /**
   * Prepare the default data.
   * @param {object} docpluginman Access to the document plugins manager.
   */
  function setDefaultDocValues(docpluginman) {
    var pptxData = docpluginman.getDataStorage()

    // Please put any setting that API can override here:
    pptxData.EMUS_PER_PT = 12700
    pptxData.pptWidth = 720 * pptxData.EMUS_PER_PT
    pptxData.pptHeight = 540 * pptxData.EMUS_PER_PT
    pptxData.pptType = 'screen4x3'

    // Rels im the main rels file that depended on the data and must be added after the slides:
    pptxData.extraMainRelList = []
  }

  // Allow you to control the view:
  genobj.view = {
    restoredLeft: 15620,
    restoredTop: 94660
  }

  // Use this method to create the options for any add* method:
  genobj.createShapeOptions = shapes.createShapeOptions

  genobj.shapes = shapes.shapes
  genobj.fields = pptxFields
  genobj.options = options && typeof options === 'object' ? options : {}

  // Temporary, I'll create a new code without the need to create a temp file (Ziv Barber, 2016-06-23):
  if (!genobj.options.tempDir) {
    genobj.options.tempDir = './'
  } // Endif.

  /**
   * Prepare everything to generate PPTX files.
   *
   * This method checking for extra resources needed to add by the generator engine.
   */
  function cbPreparePptxToGenerate() {
    genobj.generate_data = {}

    // Tell all the features (plugins) that we are about to generate a new document zip:
    gen_private.features.type.pptx.emitEvent('beforeGen', genobj)

    // Allow some plugins to do more stuff after all the plugins added their data:
    gen_private.features.type.pptx.emitEvent('beforeGenFinal', genobj)

    // Apple Keynote requires these added *after* all slides:
    gen_private.type.msoffice.rels_app.push(
      {
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
        target: 'presProps.xml',
        clear: 'type'
      },
      {
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
        target: 'viewProps.xml',
        clear: 'type'
      },
      {
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
        target: 'theme/theme1.xml',
        clear: 'type'
      }
    )

    // This event allow the plugins to add permanent rels that must be at the end:
    var extraMainRelList = []
    gen_private.features.type.pptx.emitEvent('addMainRels', {
      genobj: genobj,
      relsList: extraMainRelList
    })
    extraMainRelList.forEach(function (value) {
      gen_private.type.msoffice.rels_app.push(value)
    })

    // Add any extra rels needed by the plugins and depended on the data:
    if (
      gen_private.type.pptx.extraMainRelList &&
      typeof gen_private.type.pptx.extraMainRelList === 'object' &&
      gen_private.type.pptx.extraMainRelList.forEach
    ) {
      gen_private.type.pptx.extraMainRelList.forEach(function (value) {
        gen_private.type.msoffice.rels_app.push(value)
      })
    }

    gen_private.type.msoffice.rels_app.push({
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
      target: 'tableStyles.xml',
      clear: 'type'
    })
  }

  /**
   * Create the 'presProps.xml' resource.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxPresProps(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/></p:ext><p:ext uri="{D31A062A-798A-4329-ABDD-BBA856620510}"><p14:defaultImageDpi xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="220"/></p:ext><p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" val="1"/></p:ext></p:extLst></p:presentationPr>'
    )
  }

  /**
   * Create the 'tableStyles.xml' resource.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxStyles(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
    )
  }

  /**
   * Create the 'viewProps.xml' resource.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxViewProps(data) {
    var restoredLeft = data.view.restoredLeft || 15620
    var restoredTop = data.view.restoredTop || 94660
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="' +
      restoredLeft +
      '"/><p:restoredTop sz="' +
      restoredTop +
      '"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="64" d="100"/><a:sy n="64" d="100"/></p:scale><p:origin x="-1392" y="-96"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="78028800" cy="78028800"/></p:viewPr>'
    )
  }

  /**
   * Create the 'slideLayout1.xml' resource.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxLayout1(data) {
    if (!data || typeof data !== 'object') {
      data = {}
    } // Endif.

    // You can place here the title:
    var ph1 =
      '<a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>'
    if (data.ph1 && data.slide && data.ph1.length) {
      ph1 = createXmlSlideParagraph('', data.ph1, {}, data.slide)
    } // Endif.

    // The sub-title:
    var ph2 =
      '<a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master subtitle style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>'
    if (data.ph2 && data.slide && data.ph2.length) {
      ph2 = createXmlSlideParagraph('', data.ph2, {}, data.slide)
    } // Endif.

    var ph3 = createFieldText('DATE_TIME', 1, data.useDate)

    var footFull = ''
    var ft = ''
    var curElNum = 4

    if (data.isDate || !data.isRealSlide) {
      footFull +=
        '<p:sp><p:nvSpPr><p:cNvPr id="' +
        curElNum +
        '" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>' +
        ph3 +
        '</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
      curElNum++
    } // Endif.

    if (data.isFooter && data.ft && data.slide && data.ft.length) {
      ft = createXmlSlideParagraph('', data.ft, {}, data.slide)
      footFull +=
        '<p:sp><p:nvSpPr><p:cNvPr id="' +
        curElNum +
        '" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
        ft +
        '</p:txBody></p:sp>'
      curElNum++
    } else if (!data.isRealSlide) {
      footFull +=
        '<p:sp><p:nvSpPr><p:cNvPr id="' +
        curElNum +
        '" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
      curElNum++
    } // Endif.

    if (data.isSlideNum || !data.isRealSlide) {
      footFull +=
        '<p:sp><p:nvSpPr><p:cNvPr id="' +
        curElNum +
        '" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B1393E5F-521B-4CAD-9D3A-AE923D912DCE}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
      curElNum++
    } // Endif.

    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<p:sld' +
      (data.isRealSlide ? '' : 'Layout') +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
      (data.isRealSlide ? '' : ' type="title" preserve="1"') +
      '><p:cSld name="Title Slide"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="2130425"/><a:ext cx="7772400" cy="1470025"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>' +
      ph1 +
      '</p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1371600" y="3886200"/><a:ext cx="6400800" cy="1752600"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl9pPr></a:lstStyle>' +
      ph2 +
      '</p:txBody></p:sp>' +
      footFull +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld' +
      (data.isRealSlide ? '' : 'Layout') +
      '>'
    )
  }

  /**
   * Create the main presentation resource.
   *
   * This resource is the main resource of any PowerPoint document.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxPresentation(data) {
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'

    // Signal to the plugins to add any extra xml needed for this file:
    var dataToPlugs = {
      data: outString
    }
    gen_private.features.type.pptx.emitEvent('presentationGen', dataToPlugs)

    outString = dataToPlugs.data + '<p:sldIdLst>'

    for (
      var i = 0, total_size = gen_private.pages.length;
      i < total_size;
      i++
    ) {
      outString += '<p:sldId id="' + (i + 256) + '" r:id="rId' + (i + 2) + '"/>'
    } // End of for loop.

    outString +=
      '</p:sldIdLst><p:sldSz cx="' +
      (gen_private.type.pptx.pptWidthSLD || gen_private.type.pptx.pptWidth) +
      '" cy="' +
      (gen_private.type.pptx.pptHeightSLD || gen_private.type.pptx.pptHeight) +
      '" type="' +
      gen_private.type.pptx.pptType +
      '"/><p:notesSz cx="' +
      gen_private.type.pptx.pptHeight +
      '" cy="' +
      gen_private.type.pptx.pptWidth +
      '"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'

    var curPos = 0
    for (i = 1; i < 10; i++) {
      outString +=
        '<a:lvl' +
        i +
        'pPr marL="' +
        curPos +
        '" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl' +
        i +
        'pPr>'
      curPos += 457200
    } // End of for loop.

    outString += '</p:defaultTextStyle>'

    outString +=
      '<p:extLst><p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/></p:ext></p:extLst></p:presentation>'

    return outString
  }

  /**
   * Create the slides masters resource.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxSlideMasters(data) {
    var outData =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'

    // The master slide itself:
    if (gen_private.masterSlideXmlCode) {
      outData += '<p:cSld>' + gen_private.masterSlideXmlCode + '</p:cSld>'
    } else {
      outData +=
        '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master title style</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Click to edit Master text styles</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Second level</a:t></a:r></a:p><a:p><a:pPr lvl="2"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Third level</a:t></a:r></a:p><a:p><a:pPr lvl="3"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fourth level</a:t></a:r></a:p><a:p><a:pPr lvl="4"/><a:r><a:rPr lang="en-US" smtClean="0"/><a:t>Fifth level</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Date Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="dt" sz="half" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F8166F1F-CE9B-4651-A6AA-CD717754106B}" type="datetimeFigureOut"><a:rPr lang="en-US" smtClean="0"/><a:t>6/13/2013</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer Placeholder 4"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3124200" y="6356350"/><a:ext cx="2895600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Slide Number Placeholder 5"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="4"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="6553200" y="6356350"/><a:ext cx="2133600" cy="365125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum"><a:rPr lang="en-US" smtClean="0"/><a:t>�#�</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld>'
    } // Endif.

    outData +=
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst>'

    var curRelId = 1
    var curId = 2147483649
    var getDocData = plugsmanObj.getDataStorage()

    // Add all the slide layouts needed:
    outData += '<p:sldLayoutId id="' + curId + '" r:id="rId' + curRelId + '"/>'
    curRelId++
    curId++
    if (
      getDocData.slideLayouts &&
      typeof getDocData.slideLayouts === 'object'
    ) {
      for (var item in getDocData.slideLayouts) {
        if (getDocData.slideLayouts[item]) {
          outData +=
            '<p:sldLayoutId id="' +
            curId +
            '" r:id="rId' +
            getDocData.slideLayouts[item].relIdMaster +
            '"/>'
          curRelId++
          curId++
        } // Endif.
      } // End of for loop.
    } // Endif.

    outData +=
      '</p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="�"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr><a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr><a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr><a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr><a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr><a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr><a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr><a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr></p:otherStyle></p:txStyles></p:sldMaster>'
    return outData
  }

  /**
   * Translate field_name into the text real value.
   *
   * This method creating the text to display for the given field.
   *
   * @param {string} field_name the name of the field.
   * @param {number} slide_num current slide number.
   * @param {Date} useDate Optional date to use instead of the current date.
   * @return The text string data.
   */
  function createFieldText(field_name, slide_num, useDate) {
    var curDateTime = useDate ? new Date(useDate) : new Date()
    var dayInWeek = [
      'Sunday',
      'Monday',
      'Tuesday',
      'Wednesday',
      'Thursday',
      'Friday',
      'Saturday'
    ]
    var monthsList = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December'
    ]
    var monthsShortList = [
      'Jan',
      'Feb',
      'Mar',
      'Apr',
      'May',
      'Jun',
      'Jul',
      'Aug',
      'Sep',
      'Oct',
      'Nov',
      'Dec'
    ]
    var outValue = ''

    // curDateTime.getDate ()   Returns the day of the month (from 1-31)
    // curDateTime.getDay ()   Returns the day of the week (from 0-6)
    // curDateTime.getFullYear ()   Returns the year (four digits)
    // curDateTime.getHours ()   Returns the hour (from 0-23)
    // curDateTime.getMinutes ()   Returns the minutes (from 0-59)
    // curDateTime.getMonth ()   Returns the month (from 0-11)
    // curDateTime.getSeconds ()   Returns the seconds (from 0-59)

    switch (field_name) {
      // presentation slide number:
      case 'SLIDE_NUM':
      case 'slidenum':
        outValue += slide_num
        break

      // default date time format for the rendering application:
      case 'DATE_TIME':
      case 'datetime':
        outValue +=
          curDateTime.getMonth() +
          1 +
          '/' +
          curDateTime.getDate() +
          '/' +
          curDateTime.getFullYear()
        break

      // MM/DD/YYYY date time format (Example: 10/12/2007):
      case 'DATE_MM_DD_YYYY':
      case 'datetime1':
        outValue +=
          curDateTime.getMonth() +
          1 +
          '/' +
          curDateTime.getDate() +
          '/' +
          curDateTime.getFullYear()
        break

      // Day, Month DD, YYYY date time format (Example: Friday, October 12, 2007):
      case 'DATE_WD_MN_DD_YYYY':
      case 'datetime2':
        outValue +=
          dayInWeek[curDateTime.getDay()] +
          ', ' +
          monthsList[curDateTime.getMonth()] +
          ' ' +
          curDateTime.getDate() +
          ', ' +
          curDateTime.getFullYear()
        break

      // DD Month YYYY date time format (Example: 12 October 2007):
      case 'DATE_DD_MN_YYYY':
      case 'datetime3':
        outValue +=
          curDateTime.getDate() +
          ' ' +
          monthsList[curDateTime.getMonth()] +
          ' ' +
          curDateTime.getFullYear()
        break

      // Month DD, YYYY date time format (Example: October 12, 2007):
      case 'DATE_MN_DD_YYYY':
      case 'datetime4':
        outValue +=
          monthsList[curDateTime.getMonth()] +
          ' ' +
          curDateTime.getDate() +
          ', ' +
          curDateTime.getFullYear()
        break

      // DD-Mon-YY date time format (Example: 12-Oct-07):
      case 'DATE_DD_SMN_YY':
      case 'datetime5':
        outValue +=
          curDateTime.getDate() +
          '-' +
          monthsShortList[curDateTime.getMonth()] +
          '-' +
          (curDateTime.getFullYear() % 100)
        break

      // Month YY date time format (Example: October 07):
      case 'DATE_MM_YY':
      case 'datetime6':
        outValue +=
          monthsList[curDateTime.getMonth()] +
          ' ' +
          (curDateTime.getFullYear() % 100)
        break

      // Mon-YY date time format (Example: Oct-07):
      case 'DATE_SMN_YY':
      case 'datetime7':
        outValue +=
          monthsShortList[curDateTime.getMonth()] +
          '-' +
          (curDateTime.getFullYear() % 100)
        break

      // MM/DD/YYYY hh:mm AM/PM date time format (Example: 10/12/2007 4:28 PM):
      case 'DATE_TIME_DD_MM_YYYY_HH_MM_PM':
      case 'datetime8':
        outValue +=
          curDateTime.getMonth() +
          '/' +
          curDateTime.getDate() +
          '/' +
          curDateTime.getFullYear()
        outValue +=
          (curDateTime.getHours() % 12) + ':' + curDateTime.getMinutes()
        outValue += curDateTime.getHours() > 11 ? ' PM' : ' AM'
        break

      // MM/DD/YYYY hh:mm:ss AM/PM date time format (Example: 10/12/2007 4:28:34 PM):
      case 'DATE_TIME_DD_MM_YYYY_HH_MM_SC_PM':
      case 'datetime9':
        outValue +=
          curDateTime.getMonth() +
          '/' +
          curDateTime.getDate() +
          '/' +
          curDateTime.getFullYear()
        outValue +=
          (curDateTime.getHours() % 12) +
          ':' +
          curDateTime.getMinutes() +
          ':' +
          curDateTime.getSeconds()
        outValue += curDateTime.getHours() > 11 ? ' PM' : ' AM'
        break

      // hh:mm date time format (Example: 16:28):
      case 'TIME_HH_MM':
      case 'datetime10':
        outValue += curDateTime.getHours() + ':' + curDateTime.getMinutes()
        break

      // hh:mm:ss date time format (Example: 16:28:34):
      case 'TIME_HH_MM_SC':
      case 'datetime11':
        outValue +=
          curDateTime.getHours() +
          ':' +
          curDateTime.getMinutes() +
          ':' +
          curDateTime.getSeconds()
        break

      // hh:mm AM/PM date time format (Example: 4:28 PM):
      case 'TIME_HH_MM_PM':
      case 'datetime12':
        outValue +=
          (curDateTime.getHours() % 12) + ':' + curDateTime.getMinutes()
        outValue += curDateTime.getHours() > 11 ? ' PM' : ' AM'
        break

      // hh:mm:ss: AM/PM date time format (Example: 4:28:34 PM):
      case 'TIME_HH_MM_SC_PM':
      case 'datetime13':
        outValue +=
          (curDateTime.getHours() % 12) +
          ':' +
          curDateTime.getMinutes() +
          ':' +
          curDateTime.getSeconds()
        outValue += curDateTime.getHours() > 11 ? ' PM' : ' AM'
        break

      default:
        return null
    } // End of switch.

    return outValue
  }

  /**
   * ???.
   *
   * @param {object} text_info Information how to display the text.
   * @param {object} slide_obj The object of this slider.
   * @return Text string.
   */
  function createXmlSlideTextData(text_info, slide_obj) {
    var out_obj = {}

    out_obj.font_size = ''
    out_obj.bold = ''
    out_obj.italic = ''
    out_obj.strike = ''
    out_obj.underline = ''
    out_obj.rpr_info = ''
    out_obj.char_spacing = ''
    out_obj.baseline = ''

    if (typeof text_info === 'object') {
      if (text_info.bold) {
        out_obj.bold = ' b="1"'
      } // Endif.

      if (text_info.italic) {
        out_obj.italic = ' i="1"'
      } // Endif.

      if (text_info.strike) {
        out_obj.strike = ' strike="sngStrike"'
      } // Endif.

      if (text_info.underline) {
        out_obj.underline = ' u="sng"'
      } // Endif.

      if (text_info.font_size) {
        out_obj.font_size = ' sz="' + text_info.font_size + '00"'
      } // Endif.

      // psv 2015-01-21 Manually copied in from https://github.com/Ziv-Barber/officegen/pull/41/files
      if (text_info.char_spacing) {
        out_obj.char_spacing = ' spc="' + text_info.char_spacing * 100 + '"'

        // Must also disable kerning otherwise text won't actually expand:
        out_obj.char_spacing += ' kern="0"'
      } // Endif.

      if (text_info.baseline) {
        out_obj.baseline = ' baseline="' + text_info.baseline * 1000 + '"'
      } // Endif.

      if (text_info.color) {
        out_obj.rpr_info += shapes.createColorElements(text_info.color)
      } else if (slide_obj && slide_obj.color) {
        out_obj.rpr_info += shapes.createColorElements(slide_obj.color)
      } // Endif.

	  var pitchFamily = text_info.pitch_family || text_info.pitch_family === 0 ? text_info.pitch_family : 34
	  var charset = text_info.charset || text_info.charset === 0 ? text_info.charset : 0

      if (text_info.font_face) {
        out_obj.rpr_info +=
          '<a:latin typeface="' +
          text_info.font_face +
          '" pitchFamily="' + pitchFamily + '" charset="' + charset + '"/><a:cs typeface="' +
          text_info.font_face +
          '" pitchFamily="' + pitchFamily + '" charset="' + charset + '"/>'
      } // Endif.
    } else {
      if (slide_obj && slide_obj.color) {
        out_obj.rpr_info += shapes.createColorElements(slide_obj.color)
      } // Endif.
    } // Endif.

    if (out_obj.rpr_info !== '') {
      out_obj.rpr_info += '</a:rPr>'
    } // Endif.

    return out_obj
  }

  /**
   * Create a text object for adding into a slide.
   *
   * @param {object} text_info Information how to display the text.
   * @param {object} text_string The text string or requested field.
   * @param {object} slide_obj The object of this slider.
   * @param {object} slide_num Current slide number.
   * @param {string} out_styles Paragraph style used to style paragraphs generated by newlines
   * @return The PPTX code.
   */
  function createXmlSlideTextObject(
    text_info,
    text_string,
    slide_obj,
    slide_num,
    out_styles
  ) {
    text_info = text_info || {}

    var area_opt_data = createXmlSlideTextData(text_info, slide_obj)
    var textStyles = [
      'font_size',
      'strike',
      'italic',
      'bold',
      'underline',
      'char_spacing',
      'baseline'
    ].reduce(function (acc, attr) {
      return acc + area_opt_data[attr]
    }, '')
    var parsedText
    var startInfo =
      '<a:rPr lang="en-US"' +
      textStyles +
      ' dirty="0" smtClean="0"' +
      (area_opt_data.rpr_info !== '' ? '>' + area_opt_data.rpr_info : '/>') +
      '<a:t>'
    var endTag = '</a:r>'
    var outData = '<a:r>' + startInfo

    if (text_string.field) {
      endTag = '</a:fld>'
      var outTextField = pptxFields[text_string.field]
      if (outTextField === null) {
        for (var fieldIntName in pptxFields) {
          if (pptxFields[fieldIntName] === text_string.field) {
            outTextField = text_string.field
            break
          } // Endif.
        } // End of for loop.

        if (outTextField === null) {
          outTextField = 'datetime'
        } // Endif.
      } // Endif.

      outData =
        '<a:fld id="{' +
        gen_private.plugs.type.msoffice.makeUniqueID('5C7A2A3D') +
        '}" type="' +
        outTextField +
        '">' +
        startInfo
      outData += createFieldText(outTextField, slide_num)
    } else {
      // Automatic support for newline - split it into multi-p:
      parsedText = text_string.split('\n')
      if (parsedText.length > 1) {
        var outTextData = ''
        for (
          var i = 0, total_size_i = parsedText.length;
          i < total_size_i;
          i++
        ) {
          outTextData +=
            outData + gen_private.plugs.type.msoffice.escapeText(parsedText[i])

          if (i + 1 < total_size_i) {
            outTextData += '</a:t></a:r></a:p><a:p>'
            if (out_styles) outTextData += out_styles
          } // Endif.
        } // End of for loop.

        outData = outTextData
      } else {
        outData += gen_private.plugs.type.msoffice.escapeText(text_string)
      } // Endif.
    } // Endif.

    var outBreakP = ''
    if (text_info.breakLine) {
      outBreakP += '</a:p><a:p>'
    } // Endif.

    return outData + '</a:t>' + endTag + outBreakP
  }

  /**
   * Create all the objects inside a single paragraph.
   * @param {string} outString The string to add the output xml to it.
   * @param {Array} pData Array with all the parts of this paragraph.
   * @param {object} pOptions Paragraph options.
   * @param {object} slideObj The object of this slider.
   */
  function createXmlSlideParagraph(outString, pData, pOptions, slideObj) {
    var outStyles = ''
    var moreStylesAttr
    var moreStyles

    // Work on all the parts of this paragraph:
    for (var j = 0, total_size_j = pData.length; j < total_size_j; j++) {
      if (pData[j]) {
        moreStylesAttr = ''
        moreStyles = ''

        if (pData[j].options) {
          if (pData[j].options.align) {
            switch (pData[j].options.align) {
              case 'right':
                moreStylesAttr += ' algn="r"'
                break

              case 'center':
                moreStylesAttr += ' algn="ctr"'
                break

              case 'justify':
                moreStylesAttr += ' algn="just"'
                break
            } // End of switch.
          } // Endif.

          if (pData[j].options.indentLevel > 0) {
            moreStylesAttr += ' lvl="' + pData[j].options.indentLevel + '"'
          } // Endif.

          if (pData[j].options.listType === 'number') {
            moreStyles +=
              '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>'
          } else if (pData[j].options.listType === 'dot') {
            moreStyles += '<a:buChar char="•"/>'
          } else {
            if (pData[j].options.listType) {
              moreStyles +=
                '<a:buFont typeface="' + pData[j].options.listType + '"/>'
            } // Endif.

            if (pData[j].options.listTypeChar) {
              moreStyles +=
                '<a:buChar typeface="' + pData[j].options.listTypeChar + '"/>'
            } // Endif.
          } // Endif.
        } // Endif.

        if (moreStyles !== '') {
          outStyles = '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>'
        } else if (moreStylesAttr !== '') {
          outStyles = '<a:pPr' + moreStylesAttr + '/>'
        } // Endif.

        if (outStyles || !j) {
          if (j) {
            outString += '</a:p>'
          } // Endif.

          outString += '<a:p>' + outStyles
        } // Endif.

        outString += createXmlSlideTextObject(
          pData[j].options,
          pData[j].text,
          slideObj,
          slideObj.getPageNumber(),
          outStyles
        )
      } // Endif.
    } // End of for loop - adding all the objects inside the paragraph.

    var font_size = ''
    if (pOptions && pOptions.font_size) {
      font_size = ' sz="' + pOptions.font_size + '00"'
    } // Endif.

    outString += '<a:endParaRPr lang="en-US"' + font_size + ' dirty="0"/></a:p>'
    return outString
  }

  /**
   * ???.
   *
   * @param {object} in_data_val Input value as passed by the user.
   * @param {number} max_value Maximum value allowed.
   * @param {number} def_value Default value.
   * @param {number} auto_val ???.
   * @param {number} mul_val If you didn't provide a unit then we'll multiplay the given value with this value.
   * @return ???.
   */
  function parseSmartNumber(
    in_data_val,
    max_value,
    def_value,
    auto_val,
    mul_val
  ) {
    if (typeof in_data_val === 'undefined') {
      return typeof def_value === 'number' ? def_value : 0
    } // Endif.

    if (in_data_val === '') {
      in_data_val = 0
    } // Endif.

    var unitPos
    if (typeof in_data_val === 'string') {
      unitPos = in_data_val.search(/(inch|in)$/i)
      if (unitPos >= 0) {
        return parseInt(in_data_val.slice(0, unitPos), 10) * 914400
      } // Endif.

      unitPos = in_data_val.search(/cm$/i)
      if (unitPos >= 0) {
        return parseInt(in_data_val.slice(0, unitPos), 10) * 360000
      } // Endif.

      unitPos = in_data_val.search(/mm$/i)
      if (unitPos >= 0) {
        return parseInt(in_data_val.slice(0, unitPos), 10) * 36000
      } // Endif.

      unitPos = in_data_val.search(/emu{0,1}$/i)
      if (unitPos >= 0) {
        return parseInt(in_data_val.slice(0, unitPos), 10)
      } // Endif.

      unitPos = in_data_val.search(/(point|pt)$/i)
      if (unitPos >= 0) {
        return parseInt(in_data_val.slice(0, unitPos), 10) * 12700
      } // Endif.

      unitPos = in_data_val.search(/(pica|pc)$/i)
      if (unitPos >= 0) {
        return parseInt(in_data_val.slice(0, unitPos), 10) * 12700 * 12
      } // Endif.
    } // Endif.

    if (typeof in_data_val === 'string' && !isNaN(in_data_val)) {
      in_data_val = parseInt(in_data_val, 10)
    } // Endif.

    var realNum = Math.round(mul_val ? in_data_val * mul_val : in_data_val)

    var realVal
    var realMax

    if (typeof in_data_val === 'string') {
      if (in_data_val.indexOf('%') !== -1) {
        realMax = typeof max_value === 'number' ? max_value : 0
        if (realMax <= 0) {
          return 0
        } // Endif.

        realVal = parseInt(in_data_val, 10)
        return Math.round((realMax / 100) * realVal)
      } // Endif.

      if (in_data_val.indexOf('#') !== -1) {
        realVal = parseInt(in_data_val, 10)
        return realMax
      } // Endif.

      var realAuto = typeof auto_val === 'number' ? auto_val : 0

      if (in_data_val === '*') {
        return realAuto
      } // Endif.

      if (in_data_val === 'c') {
        return Math.round(realAuto / 2)
      } // Endif.
    } // Endif.

    if (typeof in_data_val === 'number') {
      return realNum
    } // Endif.

    return typeof def_value === 'number' ? def_value : 0
  }

  /**
   * Create the XML code of a single effect.
   *
   * This method creating the effect XML code for a single object.
   *
   * @param {object} effectData Effect data.
   * @param {string} effectName The name of the effect.
   */
  function generateEffects(effectData, effectName) {
    var outData = '<a:' + effectName + ' '
    var color = effectData.color || 'black'
    var alphaPer = 60
    var algnData = ''
    var blurRad = 50800
    var dist = 38100
    var dir = 13500000

    if (typeof effectData.transparency === 'number') {
      alphaPer = effectData.transparency
    } // Endif.

    if (alphaPer > 100 || alphaPer < 0) {
      alphaPer = 60
    }

    alphaPer = (100 - alphaPer) * 1000

    if (effectData.align) {
      if (effectData.align.top) {
        algnData += 't'
      }

      if (effectData.align.bottom) {
        algnData += 'b'
      }

      if (effectData.align.left) {
        algnData += 'l'
      }

      if (effectData.align.right) {
        algnData += 'r'
      }
    } // Endif.

    if (algnData === '') {
      algnData = 'br'
    } // Endif.

    // Size
    // Blur
    // Angle
    // Distance
    // BMK_TODO:

    outData +=
      ' blurRad="' +
      blurRad +
      '" dist="' +
      dist +
      '" dir="' +
      dir +
      '" algn="' +
      algnData +
      '" rotWithShape="0"'

    // sx="24000" sy="24000"
    // BMK_TODO:

    outData +=
      '><a:prstClr val="' +
      color +
      '"><a:alpha val="' +
      alphaPer +
      '"/></a:prstClr>'
    return outData + '</a:' + effectName + '>'
  }

  /**
   * Create the body properties code for text.
   *
   * This method creating the XML code of the body properties of a text.
   *
   * @return The body properties XML code.
   */
  function createBodyProperties(objOptions) {
    var bodyProperties = '<a:bodyPr'

    if (objOptions && objOptions.bodyProp) {
      // Set anchorPoints bottom, center or top:
      if (objOptions.bodyProp.anchor) {
        bodyProperties += ' anchor="' + objOptions.bodyProp.anchor + '"'
      } // Endif.

      if (objOptions.bodyProp.anchorCtr) {
        bodyProperties += ' anchorCtr="' + objOptions.bodyProp.anchorCtr + '"'
      } // Endif.

      // Enable or disable textwrapping none or square:
      if (objOptions.bodyProp.wrap) {
        bodyProperties += ' wrap="' + objOptions.bodyProp.wrap + '"'
      } else {
        bodyProperties += ' wrap="square"'
      } // Endif.

      // Box margins(padding):
      // BMK_TODO: I should pass a better value as the auto_val parameter of parseSmartNumber().
      if (objOptions.bodyProp.bIns) {
        bodyProperties +=
          ' bIns="' +
          parseSmartNumber(
            objOptions.bodyProp.bIns,
            gen_private.type.pptx.pptHeight,
            369332,
            gen_private.type.pptx.pptHeight,
            10000
          ) +
          '"'
      } // Endif.

      if (objOptions.bodyProp.lIns) {
        bodyProperties +=
          ' lIns="' +
          parseSmartNumber(
            objOptions.bodyProp.lIns,
            gen_private.type.pptx.pptWidth,
            2819400,
            gen_private.type.pptx.pptWidth,
            10000
          ) +
          '"'
      } // Endif.

      if (objOptions.bodyProp.rIns) {
        bodyProperties +=
          ' rIns="' +
          parseSmartNumber(
            objOptions.bodyProp.rIns,
            gen_private.type.pptx.pptWidth,
            2819400,
            gen_private.type.pptx.pptWidth,
            10000
          ) +
          '"'
      } // Endif.

      if (objOptions.bodyProp.tIns) {
        bodyProperties +=
          ' tIns="' +
          parseSmartNumber(
            objOptions.bodyProp.tIns,
            gen_private.type.pptx.pptHeight,
            369332,
            gen_private.type.pptx.pptHeight,
            10000
          ) +
          '"'
      } // Endif.

      bodyProperties += ' rtlCol="0">'

      if (objOptions.bodyProp.noAutofit) {
        bodyProperties += '<a:noAutofit/>'
      } else if (
        objOptions.bodyProp.normAutofit &&
        objOptions.bodyProp.normAutofitRed
      ) {
        bodyProperties +=
          '<a:normAutofit fontScale="' +
          objOptions.bodyProp.normAutofit +
          '" lnSpcReduction="' +
          objOptions.bodyProp.normAutofitRed +
          '"/>'
      } else if (objOptions.bodyProp.normAutofit) {
        bodyProperties +=
          '<a:normAutofit fontScale="' + objOptions.bodyProp.normAutofit + '"/>'
      } else if (objOptions.bodyProp.autoFit !== false) {
        bodyProperties += '<a:spAutoFit/>'
      } // Endif.

      bodyProperties += '</a:bodyPr>'

      // Default:
    } else {
      bodyProperties += ' wrap="square" rtlCol="0"></a:bodyPr>'
    } // Endif.

    return bodyProperties
  }

  /**
   * Generate a slider resource.
   *
   * This function generating a slider XML resource.
   *
   * @param {object} data The main slide object.
   * @param {boolean} makeOnlyObjects is true to turn this method to make only the objects in data.data.
   * @return Text string.
   */
  function cbMakePptxSlide(data, makeOnlyObjects) {
    var outString = ''
    var objs_list = data.data
    var timingData = ''

    // Slide with custom xml code generating (we are using it for slides that using the bulit-in layouts):
    if (
      data.slide.useLayout &&
      typeof data.slide.useLayout === 'object' &&
      data.slide.useLayout.mkResCb
    ) {
      // So in this case we are bypassing the normal slide generating and using a custom generating:
      data.slide.useLayout.slide = data.slide
      // To use a custom slide generator just declare data.slide.useLayout.mkResCb to be a cb that receive data.slide.useLayout itself.
      return data.slide.useLayout.mkResCb(data.slide.useLayout)
    } // Endif.

    // Allow you to turn this method into layout generator:
    var slideElement = data.slide.layoutName ? 'Layout' : ''
    var slideElementA = ''
    if (data.slide.layoutName) {
      if (data.slide.officeType) {
        slideElementA += ' type="' + data.slide.layoutName + '"'
      } // Endif.

      slideElementA += ' preserve="1"'

      if (!data.slide.officeType) {
        slideElementA += ' userDrawn="1"'
      } // Endif.
    } // Endif.

    // Create the header of the slide (only if we need that):
    if (!makeOnlyObjects) {
      outString =
        gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
        '<p:sld' +
        slideElement +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

      if (!data.slide.show && !data.slide.layoutName) {
        outString += ' show="0"'
      } // Endif.

      outString +=
        slideElementA +
        '><p:cSld' +
        (data.slide.name ? ' name="' + data.slide.name + '"' : '') +
        '>'

      if (data.slide.back) {
        outString += shapes.createColorElements(false, data.slide.back)
      } // Endif.

      outString +=
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    } // Endif.

    // Loop on all the objects inside the slide to add it into the slide:
    for (var i = 0, total_size = objs_list.length; i < total_size; i++) {
      var x = 0
      var y = 0
      var cx = 2819400
      var cy = 369332
      var xDef = true
      var yDef = true
      var cyDef = true
      var cxDef = true

      var moreStyles = ''
      var moreStylesAttr = ''
      var outStyles = ''
      var styleData = ''
      var shapeType = null
      var locationAttr = ''

      if (objs_list[i].options) {
        if (typeof objs_list[i].options.cx !== 'undefined') {
          cxDef = false
          if (objs_list[i].options.cx) {
            cx = parseSmartNumber(
              objs_list[i].options.cx,
              gen_private.type.pptx.pptWidth,
              cx,
              gen_private.type.pptx.pptWidth,
              10000
            )
          } else {
            cx = 1
          } // Endif.
        } // Endif.

        if (typeof objs_list[i].options.cy !== 'undefined') {
          cyDef = false
          if (objs_list[i].options.cy) {
            cy = parseSmartNumber(
              objs_list[i].options.cy,
              gen_private.type.pptx.pptHeight,
              cy,
              gen_private.type.pptx.pptHeight,
              10000
            )
          } else {
            cy = 1
          } // Endif.
        } // Endif.

        if (objs_list[i].options.x) {
          xDef = false
          x = parseSmartNumber(
            objs_list[i].options.x,
            gen_private.type.pptx.pptWidth,
            0,
            gen_private.type.pptx.pptWidth - cx,
            10000
          )
        } // Endif.

        if (objs_list[i].options.y) {
          yDef = false
          y = parseSmartNumber(
            objs_list[i].options.y,
            gen_private.type.pptx.pptHeight,
            0,
            gen_private.type.pptx.pptHeight - cy,
            10000
          )
        } // Endif.

        if (objs_list[i].options.shape) {
          shapeType = shapes.getShapeInfo(objs_list[i].options.shape)
        } // Endif.

        if (objs_list[i].options.flip_vertical) {
          locationAttr += ' flipV="1"'
        } // Endif.

        if (objs_list[i].options.flip_horizontal) {
          locationAttr += ' flipH="1"'
        } // Endif.

        if (objs_list[i].options.rotate) {
          var rotateVal =
            objs_list[i].options.rotate > 360
              ? objs_list[i].options.rotate - 360
              : objs_list[i].options.rotate
          rotateVal *= 60000
          locationAttr += ' rot="' + rotateVal + '"'
        } // Endif.
      } // Endif.

      // Check what type of object to add:
      switch (objs_list[i].type) {
        // TODO: remove hard code here:
        case 'table':
          var table_obj = officeTable.getTable(
            objs_list[i].data,
            objs_list[i].options
          )
          // console.log(JSON.stringify(table_obj, null, 2))
          var table_xml = xmlBuilder
            .create(table_obj, {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true
            })
            .toString({ pretty: true, indent: '  ', newline: '\n' })

          outString += table_xml
          break

        case 'chart':
          // loop through the charts
          //

          if (objs_list[i].renderType === 'pie') {
            outString +=
              '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="OfficeChart 19"/><p:cNvGraphicFramePr/><p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="4198609065"/></p:ext></p:extLst></p:nvPr></p:nvGraphicFramePr><p:xfrm><a:off x="' +
              (objs_list[i].options.x || 1524000) +
              '" y="' +
              (objs_list[i].options.y || 1397000) +
              '"/><a:ext cx="' +
              (objs_list[i].options.cx || 6096000) +
              '" cy="' +
              (objs_list[i].options.cy || 4064000) +
              '"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>'
          } else if (objs_list[i].renderType === 'column') {
            outString +=
              '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="OfficeChart 3"/><p:cNvGraphicFramePr/><p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1256887135"/></p:ext></p:extLst></p:nvPr></p:nvGraphicFramePr><p:xfrm><a:off x="' +
              (objs_list[i].options.x || 1524000) +
              '" y="' +
              (objs_list[i].options.y || 1397000) +
              '"/><a:ext cx="' +
              (objs_list[i].options.cx || 6096000) +
              '" cy="' +
              (objs_list[i].options.cy || 4064000) +
              '"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>'
          } // Endif.
          break
        case 'text':
        case 'cxn':
          var effectsList = ''

          if (shapeType == null) shapeType = shapes.getShapeInfo(null)

          // if ( objs_list[i].type == 'text' ) {
          //   if ( !objs_list[i].options || (!objs_list[i].options.cx && !objs_list[i].options.cx) ) {
          //     objs_list[i].options = objs_list[i].options ? objs_list[i].options : {}
          //     objs_list[i].options.bodyProp = objs_list[i].options.bodyProp ? objs_list[i].options.bodyProp : {}
          //     objs_list[i].options.bodyProp.autoFit = true
          //     cx = gen_private.type.pptx.pptWidth - x
          //     cy = gen_private.type.pptx.pptHeight - y
          //   } // Endif.
          // } // Endif.

          var pNvPr = '<p:nvPr/>'
          if (objs_list[i].options.ph) {
            var extraPhAttr = ''
            if (objs_list[i].options.phSz) {
              extraPhAttr += ' sz="' + objs_list[i].options.phSz + '"'
            } // Endif.

            if (objs_list[i].options.phIdx) {
              extraPhAttr += ' idx="' + objs_list[i].options.phIdx + '"'
            } // Endif.

            pNvPr =
              '<p:nvPr><p:ph type="' +
              objs_list[i].options.ph +
              '"' +
              extraPhAttr +
              (objs_list[i].options.nvPrAttrCode || '') +
              '/>' +
              (objs_list[i].options.nvPrCode || '') +
              '</p:nvPr>'
          } // Endif.

          var codeCNvPrAttr = ''
          if (objs_list[i].options.id) {
            codeCNvPrAttr += ' id="' + objs_list[i].options.id + '"'
          } else {
            codeCNvPrAttr += ' id="' + (i + 2) + '"'
          } // Endif.

          if (objs_list[i].options.name) {
            codeCNvPrAttr += ' name="' + objs_list[i].options.name + '"'
          } else {
            codeCNvPrAttr += ' name="Object ' + (i + 2) + '"'
          } // Endif.

          if (objs_list[i].options.title) {
            codeCNvPrAttr += ' title="' + objs_list[i].options.title + '"'
          } // Endif.

          if (objs_list[i].options.desc) {
            codeCNvPrAttr += ' descr="' + objs_list[i].options.desc + '"'
          } // Endif.

          if (objs_list[i].options.hidden) {
            codeCNvPrAttr += ' hidden="1"'
          } // Endif.

          if (objs_list[i].type === 'cxn') {
            outString += '<p:cxnSp><p:nvCxnSpPr>'
            outString +=
              '<p:cNvPr' + codeCNvPrAttr + '/>' + pNvPr + '</p:nvCxnSpPr>'
          } else {
            outString += '<p:sp><p:nvSpPr>'
            outString +=
              '<p:cNvPr' +
              codeCNvPrAttr +
              '/><p:cNvSpPr txBox="1"/>' +
              pNvPr +
              '</p:nvSpPr>'
          } // Endif.

          var bwMode = objs_list[i].options.bwMode
            ? ' bwMode="' + objs_list[i].options.bwMode + '"'
            : ''

          if (
            objs_list[i].options.ph &&
            xDef &&
            yDef &&
            cxDef &&
            cyDef &&
            !locationAttr
          ) {
            outString += '<p:spPr' + bwMode + '/>'
          } else {
            outString += '<p:spPr' + bwMode + '>'

            outString += '<a:xfrm' + locationAttr + '>'

            if (!objs_list[i].options.ph || !yDef || !xDef) {
              outString += '<a:off'

              if (!objs_list[i].options.ph || !xDef) {
                outString += ' x="' + x + '"'
              } // Endif.

              if (!objs_list[i].options.ph || !yDef) {
                outString += ' y="' + y + '"'
              } // Endif.

              outString += '/>'
            } // Endif.

            if (!objs_list[i].options.ph || !cyDef || !cxDef) {
              outString += '<a:ext'

              if (!objs_list[i].options.ph || !cxDef) {
                outString += ' cx="' + cx + '"'
              } // Endif.

              if (!objs_list[i].options.ph || !cyDef) {
                outString += ' cy="' + cy + '"'
              } // Endif.

              outString += '/>'
            } // Endif.

            outString += '</a:xfrm>'

            outString += '<a:prstGeom prst="' + shapeType.name + '">'

            // string changed to take into account change of shape that you do by moving the little yellow dot

            if (shapeType.avLst !== {}) {
              outString += '<a:avLst>'
              for (var adj in shapeType.avLst) {
                outString +=
                  '<a:gd name="' +
                  adj +
                  '" fmla="val ' +
                  shapeType.avLst[adj] +
                  '"/>'
              }
            }

            outString += '</a:avLst></a:prstGeom>'

            if (objs_list[i].options) {
              if (objs_list[i].options.fill) {
                outString += shapes.createColorElements(
                  objs_list[i].options.fill
                )
              } else if (!objs_list[i].options.disableFillSettings) {
                outString += '<a:noFill/>'
              } // Endif.

              if (objs_list[i].options.line) {
                var lineAttr = ''

                if (objs_list[i].options.line_size) {
                  lineAttr +=
                    ' w="' + objs_list[i].options.line_size * 12700 + '"'
                } // Endif.

                // cmpd="dbl"

                outString += '<a:ln' + lineAttr + '>'
                outString += shapes.createColorElements(
                  objs_list[i].options.line
                )

                if (objs_list[i].options.line_head) {
                  outString +=
                    '<a:headEnd type="' + objs_list[i].options.line_head + '"/>'
                } // Endif.

                if (objs_list[i].options.line_tail) {
                  outString +=
                    '<a:tailEnd type="' + objs_list[i].options.line_tail + '"/>'
                } // Endif.

                outString += '</a:ln>'
              } // Endif.
            } else if (!objs_list[i].options.disableFillSettings) {
              outString += '<a:noFill/>'
            } // Endif.

            if (objs_list[i].options.effects) {
              for (
                var ii = 0, total_size_ii = objs_list[i].options.effects.length;
                ii < total_size_ii;
                ii++
              ) {
                switch (objs_list[i].options.effects[ii].type) {
                  case 'outerShadow':
                    effectsList += generateEffects(
                      objs_list[i].options.effects[ii],
                      'outerShdw'
                    )
                    break

                  case 'innerShadow':
                    effectsList += generateEffects(
                      objs_list[i].options.effects[ii],
                      'innerShdw'
                    )
                    break
                } // End of switch.
              } // End of for loop.
            } // Endif.

            if (effectsList !== '') {
              outString += '<a:effectLst>' + effectsList + '</a:effectLst>'
            } // Endif.

            outString += '</p:spPr>'
          } // Endif.

          if (objs_list[i].options) {
            if (objs_list[i].options.align) {
              switch (objs_list[i].options.align) {
                case 'right':
                  moreStylesAttr += ' algn="r"'
                  break

                case 'center':
                  moreStylesAttr += ' algn="ctr"'
                  break

                case 'justify':
                  moreStylesAttr += ' algn="just"'
                  break
              } // End of switch.
            } // Endif.

            if (objs_list[i].options.indentLevel > 0) {
              moreStylesAttr +=
                ' lvl="' + objs_list[i].options.indentLevel + '"'
            } // Endif.
          } // Endif.

          if (moreStyles !== '') {
            outStyles =
              '<a:pPr' + moreStylesAttr + '>' + moreStyles + '</a:pPr>'
          } else if (moreStylesAttr !== '') {
            outStyles = '<a:pPr' + moreStylesAttr + '/>'
          } // Endif.

          if (styleData !== '') {
            outString += '<p:style>' + styleData + '</p:style>'
          } // Endif.

          if (typeof objs_list[i].text === 'string') {
            outString +=
              '<p:txBody>' +
              createBodyProperties(objs_list[i].options) +
              '<a:lstStyle/><a:p>' +
              outStyles
            outString += createXmlSlideTextObject(
              objs_list[i].options,
              objs_list[i].text,
              data.slide,
              data.slide.getPageNumber(),
              outStyles
            )
          } else if (typeof objs_list[i].text === 'number') {
            outString +=
              '<p:txBody>' +
              createBodyProperties(objs_list[i].options) +
              '<a:lstStyle/><a:p>' +
              outStyles
            outString += createXmlSlideTextObject(
              objs_list[i].options,
              objs_list[i].text + '',
              data.slide,
              data.slide.getPageNumber(),
              outStyles
            )
          } else if (objs_list[i].text && objs_list[i].text.length) {
            var outBodyOpt = createBodyProperties(objs_list[i].options)
            outString +=
              '<p:txBody>' + outBodyOpt + '<a:lstStyle/><a:p>' + outStyles

            for (
              var j = 0, total_size_j = objs_list[i].text.length;
              j < total_size_j;
              j++
            ) {
              if (
                typeof objs_list[i].text[j] === 'object' &&
                objs_list[i].text[j].text
              ) {
                outString += createXmlSlideTextObject(
                  objs_list[i].text[j].options || objs_list[i].options,
                  objs_list[i].text[j].text,
                  data.slide,
                  outBodyOpt,
                  outStyles,
                  data.slide.getPageNumber()
                )
              } else if (typeof objs_list[i].text[j] === 'string') {
                outString += createXmlSlideTextObject(
                  objs_list[i].options,
                  objs_list[i].text[j],
                  data.slide,
                  outBodyOpt,
                  outStyles,
                  data.slide.getPageNumber()
                )
              } else if (typeof objs_list[i].text[j] === 'number') {
                outString += createXmlSlideTextObject(
                  objs_list[i].options,
                  objs_list[i].text[j] + '',
                  data.slide,
                  outBodyOpt,
                  outStyles,
                  data.slide.getPageNumber()
                )
              } else if (
                typeof objs_list[i].text[j] === 'object' &&
                objs_list[i].text[j].field
              ) {
                outString += createXmlSlideTextObject(
                  objs_list[i].options,
                  objs_list[i].text[j],
                  data.slide,
                  outBodyOpt,
                  outStyles,
                  data.slide.getPageNumber()
                )
              } // Endif.
            } // Endif.
          } else if (typeof objs_list[i].text === 'object') {
            if (!objs_list[i].text) {
              objs_list[i].text = {}
            } // Endif.

            if (!objs_list[i].text.field) {
              objs_list[i].text.field = ''
            } // Endif.

            outString +=
              '<p:txBody>' +
              createBodyProperties(objs_list[i].options) +
              '<a:lstStyle/><a:p>' +
              outStyles
            outString += createXmlSlideTextObject(
              objs_list[i].options,
              objs_list[i].text,
              data.slide,
              data.slide.getPageNumber(),
              outStyles
            )
          } // Endif.

          // We must add that at the end of every paragraph with text:
          if (typeof objs_list[i].text !== 'undefined') {
            var font_size = ''
            if (objs_list[i].options && objs_list[i].options.font_size) {
              font_size = ' sz="' + objs_list[i].options.font_size + '00"'
            } // Endif.

            outString +=
              '<a:endParaRPr lang="en-US"' +
              font_size +
              ' dirty="0"/></a:p></p:txBody>'
          } // Endif.

          outString += objs_list[i].type === 'cxn' ? '</p:cxnSp>' : '</p:sp>'
          /* eslint-disable indent */
          break

        // Table:
        /*
        case 'table':
          outString += '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="' + (i + 2) + '" name="Object ' + (i + 1) + '"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
          outString += '<p:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></p:xfrm>'
          outString += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>'

          if ( objs_list[i].options && objs_list[i].options.tableStyleId ) {
            outString += objs_list[i].options.tableStyleId

          } else {
            outString += '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'
          } // Endif.

          outString += '</a:tableStyleId></a:tblPr><a:tblGrid>'
          // <a:gridCol w="3276600"/>
          outString += '</a:tblGrid>'
          // objs_list[i].options
          // objs_list[i].rows[][].text
          // BMK_TODO:
          break
        */
        /* eslint-enable indent */

        // Image:
        case 'image':
          // psv 2015-01-21 Manually copied this section from https://github.com/Ziv-Barber/officegen/pull/39/files?w=1
          // outString += '<p:pic><p:nvPicPr><p:cNvPr id="' + (i + 2) + '" name="Object ' + (i + 1) + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId' + objs_list[i].rel_id + '" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm' + locationAttr + '><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
          var parts = []

          parts.push(
            '<p:pic><p:nvPicPr><p:cNvPr id="' +
              (i + 2) +
              '" name="Object ' +
              (i + 1) +
              '"'
          )

          if (objs_list[i].link_rel_id) {
            parts.push(
              '><a:hlinkClick r:id="rId' +
                objs_list[i].link_rel_id +
                '"/></p:cNvPr>'
            )
          } else {
            parts.push('/>')
          } // Endif.

          parts.push(
            '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId' +
              objs_list[i].rel_id +
              '" cstate="print"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm' +
              locationAttr +
              '><a:off x="' +
              x +
              '" y="' +
              y +
              '"/><a:ext cx="' +
              cx +
              '" cy="' +
              cy +
              '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
          )
          outString += parts.join('')
          // -- psv 2015-01-21 end merge
          break

        // Paragraph:
        case 'p':
          if (shapeType == null) {
            shapeType = shapes.getShapeInfo(null)
          } // Endif.

          outString += '<p:sp><p:nvSpPr>'
          outString +=
            '<p:cNvPr id="' +
            (i + 2) +
            '" name="Object ' +
            (i + 1) +
            '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
          outString += '<p:spPr>'

          outString += '<a:xfrm' + locationAttr + '>'

          outString +=
            '<a:off x="' +
            x +
            '" y="' +
            y +
            '"/><a:ext cx="' +
            cx +
            '" cy="' +
            cy +
            '"/></a:xfrm><a:prstGeom prst="' +
            shapeType.name +
            '"><a:avLst/></a:prstGeom>'

          if (objs_list[i].options) {
            if (objs_list[i].options.fill) {
              outString += shapes.createColorElements(objs_list[i].options.fill)
            } else if (!objs_list[i].options.disableFillSettings) {
              outString += '<a:noFill/>'
            } // Endif.

            if (objs_list[i].options.line) {
              outString += '<a:ln>'
              outString += shapes.createColorElements(objs_list[i].options.line)

              if (objs_list[i].options.line_head) {
                outString +=
                  '<a:headEnd type="' + objs_list[i].options.line_head + '"/>'
              } // Endif.

              if (objs_list[i].options.line_tail) {
                outString +=
                  '<a:tailEnd type="' + objs_list[i].options.line_tail + '"/>'
              } // Endif.

              outString += '</a:ln>'
            } // Endif.
          } else if (!objs_list[i].options.disableFillSettings) {
            outString += '<a:noFill/>'
          } // Endif.

          outString += '</p:spPr>'

          if (styleData !== '') {
            outString += '<p:style>' + styleData + '</p:style>'
          } // Endif.

          outString +=
            '<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>'

          // Add all the paragraph objects:
          outString = createXmlSlideParagraph(
            outString,
            objs_list[i].data,
            objs_list[i].options,
            data.slide
          )

          outString += '</p:txBody>'

          outString += '</p:sp>'
          break

        // Custom xml code:
        case 'xml':
          // NOTE: While it allowing you to add any xml code directly, please try to use it only for custom layouts.
          outString += objs_list[i].data
          break
      } // End of switch.
    } // End of for loop.

    if (!makeOnlyObjects) {
      outString +=
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'

      if (timingData !== '') {
        outString += '<p:timing>' + timingData + '</p:timing>'
      } // Endif.

      outString += '</p:sld' + slideElement + '>'
    } // Endif.

    return outString
  }

  /**
   * Generate the extended attributes file (app) for PPTX/PPSX documents.
   *
   * @param {object} data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakePptxApp(data) {
    var slidesCount = gen_private.pages.length
    var userName =
      genobj.options.author || genobj.options.creator || 'officegen'
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Words>0</Words><Application>Microsoft Office PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>' +
      slidesCount +
      '</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>' +
      slidesCount +
      '</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="' +
      (slidesCount + 1) +
      '" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr>'

    for (
      var i = 0, total_size = gen_private.pages.length;
      i < total_size;
      i++
    ) {
      outString +=
        '<vt:lpstr>' +
        gen_private.plugs.type.msoffice.escapeText(
          gen_private.pages[i].slide.name
        ) +
        '</vt:lpstr>'
    } // End of for loop.

    outString +=
      '</vt:vector></TitlesOfParts><Company>' +
      userName +
      '</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>'
    return outString
  }

  /**
   * Create XML string for a chart description
   *
   * @param {object} chartInfo
   * @return Text Xml string.
   */
  function cbMakeCharts(chartInfo) {
    var chart = new OfficeChart(chartInfo)
    return chart.toXML()
  }

  function cbMakeChartDataExcel(data) {
    // PSV not sure why this method is here.  Presumably vestigial from earlier code.
  }

  // Prepare genobj for MS-Office:
  msdoc.makemsdoc(genobj, new_type, options, gen_private, type_info)
  gen_private.plugs.type.msoffice.makeOfficeGenerator('ppt', 'presentation', {})

  gen_private.features.page_name = 'slides' // This document type must have pages.

  gen_private.plugs.type.msoffice.addInfoType(
    'dc:title',
    '',
    'title',
    'setDocTitle'
  )

  genobj.on('beforeGen', cbPreparePptxToGenerate)

  var type_of_main_doc = 'slideshow'
  if (new_type !== 'ppsx') {
    type_of_main_doc = 'presentation'
  } // Endif.

  // Create the plugins manager:
  var plugsmanObj = new docplugman(
    genobj,
    gen_private,
    'pptx',
    setDefaultDocValues
  )

  // We'll register now any specific PowerPoint based plugin that we want to use:
  plugsmanObj.plugsList.push(new plugWidescreen(plugsmanObj))
  plugsmanObj.plugsList.push(new plugSpeakernotes(plugsmanObj))
  plugsmanObj.plugsList.push(new plugLayouts(plugsmanObj))
  // BMK_PPTX_PLUG:

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

  // Save some methods for the plugins:
  genobj.cbMakePptxLayout1 = cbMakePptxLayout1
  genobj.cbMakePptxSlide = cbMakePptxSlide
  genobj.createFieldText = createFieldText
  genobj.cMakePptxOutTextP = createXmlSlideParagraph

  gen_private.type.msoffice.files_list.push(
    {
      name: '/ppt/slideMasters/slideMaster1.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
      clear: 'type'
    },
    {
      name: '/ppt/presProps.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml',
      clear: 'type'
    },
    {
      name: '/ppt/presentation.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.' +
        type_of_main_doc +
        '.main+xml',
      clear: 'type'
    },
    {
      name: '/ppt/slideLayouts/slideLayout1.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
      clear: 'type'
    },
    {
      name: '/ppt/tableStyles.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
      clear: 'type'
    },
    {
      name: '/ppt/viewProps.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml',
      clear: 'type'
    }
  )

  genobj.slideMasterRels = [
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
      target: '../slideLayouts/slideLayout1.xml'
    },
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: '../theme/theme1.xml'
    }
  ]

  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\presProps.xml',
    'buffer',
    null,
    cbMakePptxPresProps,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\tableStyles.xml',
    'buffer',
    null,
    cbMakePptxStyles,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\viewProps.xml',
    'buffer',
    genobj,
    cbMakePptxViewProps,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\presentation.xml',
    'buffer',
    null,
    cbMakePptxPresentation,
    true
  )

  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\slideLayouts\\slideLayout1.xml',
    'buffer',
    null,
    cbMakePptxLayout1,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\slideLayouts\\_rels\\slideLayout1.xml.rels',
    'buffer',
    [
      {
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
        target: '../slideMasters/slideMaster1.xml'
      }
    ],
    gen_private.plugs.type.msoffice.cbMakeRels,
    true
  )

  // Slides master:
  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\slideMasters\\slideMaster1.xml',
    'buffer',
    null,
    cbMakePptxSlideMasters,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\slideMasters\\_rels\\slideMaster1.xml.rels',
    'buffer',
    genobj.slideMasterRels,
    gen_private.plugs.type.msoffice.cbMakeRels,
    true
  )

  gen_private.plugs.intAddAnyResourceToParse(
    'docProps\\app.xml',
    'buffer',
    null,
    cbMakePptxApp,
    true
  )

  gen_private.type.msoffice.rels_app.push({
    type:
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
    target: 'slideMasters/slideMaster1.xml',
    clear: 'type'
  })

  gen_private.plugs.intAddAnyResourceToParse(
    'ppt\\_rels\\presentation.xml.rels',
    'buffer',
    gen_private.type.msoffice.rels_app,
    gen_private.plugs.type.msoffice.cbMakeRels,
    true
  )

  // ----- API for PowerPoint documents: -----

  /**
   * Create a new slide.
   *
   * This method creating a new slide inside the presentation.
   *
   * @param {object} slideOptions Extra options how to create the new slide.
   * @return The new slide object.
   */
  genobj.makeNewSlide = function (slideOptions) {
    var pageNumber = gen_private.pages.length
    var slideObj = { show: true } // The slide object that the user will use.

    if (!slideOptions || typeof slideOptions !== 'object') {
      slideOptions = {}
    } // Endif.

    if (!slideOptions.basedLayout) {
      slideOptions.basedLayout = 1
    } // Endif.

    gen_private.pages[pageNumber] = {}
    gen_private.pages[pageNumber].slide = slideObj
    gen_private.pages[pageNumber].data = []
    gen_private.pages[pageNumber].rels = [
      {
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
        target:
          '../slideLayouts/slideLayout' + slideOptions.basedLayout + '.xml',
        clear: 'data'
      }
    ]

    gen_private.type.msoffice.files_list.push({
      name: '/ppt/slides/slide' + (pageNumber + 1) + '.xml',
      type:
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      clear: 'data'
    })

    gen_private.type.msoffice.rels_app.push({
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target: 'slides/slide' + (pageNumber + 1) + '.xml',
      clear: 'data'
    })

    slideObj.getPageNumber = function () {
      return pageNumber
    }
    slideObj.getRelFile = function () {
      return gen_private.pages[pageNumber].rels
    }

    slideObj.name = 'Slide ' + (pageNumber + 1)

    /**
     * ???.
     *
     * @param {object} prgObj Paragraph object.
     */
    function addParagraphApiForBasicOpt(prgObj) {
      if (!prgObj.api) {
        prgObj.api = {}
      } // Endif.

      prgObj.api.options = prgObj.options
    }

    /**
     * ???.
     *
     * @param {object} prgObj Paragraph object.
     */
    function addParagraphApiForEffects(prgObj) {
      if (!prgObj.api) {
        prgObj.api = {}
      } // Endif.

      /**
       * ???.
       *
       * @param {object} inType ???.
       * @param {object} inAlign ???.
       * @param {object} inColor ???.
       * @param {object} inTransparency ???.
       * @param {object} inSize ???.
       * @param {object} inBlur ???.
       * @param {object} inAngle ???.
       * @param {object} inDistance ???.
       */
      prgObj.api.setShadowEffect = function (
        inType,
        inAlign,
        inColor,
        inTransparency,
        inSize,
        inBlur,
        inAngle,
        inDistance
      ) {
        if (!prgObj.options.effects) {
          prgObj.options.effects = []
        }

        var newEffect = {
          type: inType,
          align: inAlign,
          color: inColor,
          transparency: inTransparency,
          size: inSize,
          blur: inBlur,
          angle: inAngle,
          distance: inDistance
        }

        prgObj.options.effects.push(newEffect)
      }
    }

    // Added 2015-01-02 PSV as a way for user to modify chart properties before generating the XML string:
    slideObj.createChart = function (chartInfo) {
      return new OfficeChart(chartInfo)
    }

    // Added 2015-01-08 PSV following model of addChart:
    slideObj.addTable = function (data, options) {
      var objNumber = gen_private.pages[pageNumber].data.length

      gen_private.pages[pageNumber].data[objNumber] = {
        type: 'table',
        data: data,
        options: options || {} // right now this isn't yet used
      }

      addParagraphApiForBasicOpt(gen_private.pages[pageNumber].data[objNumber])
      addParagraphApiForEffects(gen_private.pages[pageNumber].data[objNumber])
      return gen_private.pages[pageNumber].data[objNumber].api
    }

    /**
     * Generate the chart based on input data.
     *
     * @param {object} renderType should belong to: 'column', 'pie'
     * @param {object} data a JSON object with follow the following format
     *     {
     *      title: 'eSurvey chart',
     *      data:  [
     *        {
     *          name: 'Income',
     *          labels: ['2005', '2006', '2007', '2008', '2009'],
     *          values: [23.5, 26.2, 30.1, 29.5, 24.6]
     *        },
     *        {
     *          name: 'Expense',
     *          labels: ['2005', '2006', '2007', '2008', '2009'],
     *          values: [18.1, 22.8, 23.9, 25.1, 25]
     *        }
     *      ]
     *   }
     * @author vtloc
     * @date 2014Jan09
     */
    slideObj.addChart = function (data, callback, errorCallback) {
      var objNumber = gen_private.pages[pageNumber].data.length
      GLOBAL_CHART_COUNT += 1

      gen_private.pages[pageNumber].data[objNumber] = {}
      gen_private.pages[pageNumber].data[objNumber].type = 'chart'
      gen_private.pages[pageNumber].data[objNumber].renderType = 'column'
      gen_private.pages[pageNumber].data[objNumber].title = data.title
      gen_private.pages[pageNumber].data[objNumber].data = data.data
      gen_private.pages[pageNumber].data[objNumber].options = data // include generic options like x,y,cx,cy, etc.
      // data['renderType'] = renderType

      // First, generate a temporatory excel file for storing the chart's data
      var workbook = excelbuilder.createWorkbook(
        genobj.options.tempDir,
        'sample' + GLOBAL_CHART_COUNT + '.xlsx'
      )

      // Create a new worksheet with 10 columns and 12 rows
      // number of columns: data['data'].length+1 -> equaly number of series
      // number of rows: data['data'][0].values.length+1
      var sheet1 = workbook.createSheet(
        'Sheet1',
        data.data.length + 1,
        data.data[0].values.length + 1
      )
      var headerrow = 1

      // Write header using serie name:
      for (var j = 0; j < data.data.length; j++) {
        sheet1.set(j + 2, headerrow, data.data[j].name)
      } // End of for loop.

      // Write category column in the first column:
      for (j = 0; j < data.data[0].labels.length; j++) {
        sheet1.set(1, j + 2, data.data[0].labels[j])
      } // End of for loop.

      // For each serie, write out values in its row:
      for (var i = 0; i < data.data.length; i++) {
        for (j = 0; j < data.data[i].values.length; j++) {
          // col i+2
          // row j+1
          sheet1.set(i + 2, j + 2, data.data[i].values[j])
        } // End of for loop.
      } // End of for loop.

      // Fill some data
      // Save it
      var localEmbeddingExcelFile =
        'ppt\\embeddings\\Microsoft_Excel_Worksheet' +
        GLOBAL_CHART_COUNT +
        '.xlsx'
      var tmpExcelFile =
        genobj.options.tempDir + 'sample' + GLOBAL_CHART_COUNT + '.xlsx'
      // will copy the tmpExcelFile into localEmbeddingExcelFile
      gen_private.plugs.intAddAnyResourceToParse(
        localEmbeddingExcelFile,
        'file',
        tmpExcelFile,
        cbMakeChartDataExcel,
        false,
        true
      )
      gen_private.plugs.intAddAnyResourceToParse(
        'ppt\\charts\\chart' + GLOBAL_CHART_COUNT + '.xml',
        'buffer',
        data,
        cbMakeCharts,
        true
      )
      gen_private.plugs.intAddAnyResourceToParse(
        'ppt\\charts\\_rels\\chart' + GLOBAL_CHART_COUNT + '.xml.rels',
        'buffer',
        [
          {
            type:
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
            target:
              '../embeddings/Microsoft_Excel_Worksheet' +
              GLOBAL_CHART_COUNT +
              '.xlsx'
          }
        ],
        gen_private.plugs.type.msoffice.cbMakeRels,
        true
      )
      gen_private.type.msoffice.files_list.push(
        {
          name: '/ppt/charts/chart' + GLOBAL_CHART_COUNT + '.xml',
          type:
            'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          clear: 'type'
        },
        {
          name:
            '/ppt/embeddings/Microsoft_Excel_Worksheet' +
            GLOBAL_CHART_COUNT +
            '.xlsx',
          type:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          ext: 'xlsx',
          clear: 'type'
        }
      )

      gen_private.pages[pageNumber].rels.push({
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
        target: '../charts/chart' + GLOBAL_CHART_COUNT + '.xml',
        clear: 'data'
      })

      workbook.save(function (err) {
        if (err) {
          workbook.cancel()
        } else {
          callback()
        }
      })
    }

    /**
     * ???.
     *
     * @param {object} text The text to add.
     * @param {object} opt ???.
     * @param {object} y_pos Y position.
     * @param {object} x_size X size.
     * @param {object} y_size Y size.
     * @param {object} opt_b ???.
     */
    slideObj.addText = function (text, opt, y_pos, x_size, y_size, opt_b) {
      var objNumber = gen_private.pages[pageNumber].data.length

      gen_private.pages[pageNumber].data[objNumber] = {}
      gen_private.pages[pageNumber].data[objNumber].type = 'text'
      gen_private.pages[pageNumber].data[objNumber].text = text || ''
      gen_private.pages[pageNumber].data[objNumber].options =
        typeof opt === 'object' ? opt : {}

      if (typeof opt === 'string') {
        gen_private.pages[pageNumber].data[objNumber].options.color = opt
      } else if (typeof opt !== 'object' && typeof y_pos !== 'undefined') {
        gen_private.pages[pageNumber].data[objNumber].options.x = opt
        gen_private.pages[pageNumber].data[objNumber].options.y = y_pos

        if (typeof x_size !== 'undefined' && typeof y_size !== 'undefined') {
          gen_private.pages[pageNumber].data[objNumber].options.cx = x_size
          gen_private.pages[pageNumber].data[objNumber].options.cy = y_size
        } // Endif.
      } // Endif.

      var attrname

      if (typeof opt_b === 'object') {
        for (attrname in opt_b) {
          gen_private.pages[pageNumber].data[objNumber].options[attrname] =
            opt_b[attrname]
        }
      } else if (typeof x_size === 'object' && typeof y_size === 'undefined') {
        for (attrname in x_size) {
          gen_private.pages[pageNumber].data[objNumber].options[attrname] =
            x_size[attrname]
        }
      } // Endif.

      addParagraphApiForBasicOpt(gen_private.pages[pageNumber].data[objNumber])
      addParagraphApiForEffects(gen_private.pages[pageNumber].data[objNumber])
      return gen_private.pages[pageNumber].data[objNumber].api
    }

    /**
     * ???
     *
     * @param {string} shape ???.
     * @param {object} opt ???.
     * @param {number} y_pos ???.
     * @param {number} x_size ???.
     * @param {number} y_size ???.
     * @param {object} opt_b ???.
     */
    slideObj.addShape = function (shape, opt, y_pos, x_size, y_size, opt_b) {
      var objNumber = gen_private.pages[pageNumber].data.length

      gen_private.pages[pageNumber].data[objNumber] = {}
      gen_private.pages[pageNumber].data[objNumber].type = 'text'
      gen_private.pages[pageNumber].data[objNumber].options =
        typeof opt === 'object' ? opt : {}
      gen_private.pages[pageNumber].data[objNumber].options.shape = shape

      if (typeof opt === 'string') {
        gen_private.pages[pageNumber].data[objNumber].options.color = opt
      } else if (typeof opt !== 'object' && typeof y_pos !== 'undefined') {
        gen_private.pages[pageNumber].data[objNumber].options.x = opt
        gen_private.pages[pageNumber].data[objNumber].options.y = y_pos

        if (typeof x_size !== 'undefined' && typeof y_size !== 'undefined') {
          gen_private.pages[pageNumber].data[objNumber].options.cx = x_size
          gen_private.pages[pageNumber].data[objNumber].options.cy = y_size
        } // Endif.
      } // Endif.

      var attrname

      if (typeof opt_b === 'object') {
        for (attrname in opt_b) {
          gen_private.pages[pageNumber].data[objNumber].options[attrname] =
            opt_b[attrname]
        }
      } else if (typeof x_size === 'object' && typeof y_size === 'undefined') {
        for (attrname in x_size) {
          gen_private.pages[pageNumber].data[objNumber].options[attrname] =
            x_size[attrname]
        }
      } // Endif.

      addParagraphApiForBasicOpt(gen_private.pages[pageNumber].data[objNumber])
      addParagraphApiForEffects(gen_private.pages[pageNumber].data[objNumber])
      return gen_private.pages[pageNumber].data[objNumber].api
    }

    /**
     * ???.
     *
     * @param {object} image_path ???.
     * @param {object} opt ???.
     * @param {object} y_pos ???.
     * @param {object} x_size ???.
     * @param {object} y_size ???.
     * @param {object} image_format_type ???.
     */
    slideObj.addImage = function (
      image_path,
      opt,
      y_pos,
      x_size,
      y_size,
      image_format_type
    ) {
      var objNumber = gen_private.pages[pageNumber].data.length
      var image_type =
        typeof image_format_type === 'string' ? image_format_type : 'png'
      var defWidth
      var defHeight = 0

      if (typeof image_path === 'string') {
        var ret_data = fast_image_size(image_path)
        if (ret_data.type === 'unknown') {
          var image_ext = path.extname(image_path)

          switch (image_ext) {
            case '.bmp':
              image_type = 'bmp'
              break

            case '.gif':
              image_type = 'gif'
              break

            case '.jpg':
            case '.jpeg':
              image_type = 'jpeg'
              break

            case '.emf':
              image_type = 'emf'
              break

            case '.tiff':
              image_type = 'tiff'
              break
          } // End of switch.
        } else {
          if (ret_data.width) {
            defWidth = ret_data.width
          } // Endif.

          if (ret_data.height) {
            defHeight = ret_data.height
          } // Endif.

          image_type = ret_data.type
          if (image_type === 'jpg') {
            image_type = 'jpeg'
          } // Endif.
        } // Endif.
      } // Endif.

      gen_private.pages[pageNumber].data[objNumber] = {}
      gen_private.pages[pageNumber].data[objNumber].type = 'image'
      gen_private.pages[pageNumber].data[objNumber].image = image_path
      gen_private.pages[pageNumber].data[objNumber].options =
        typeof opt === 'object' ? opt : {}

      if (
        !gen_private.pages[pageNumber].data[objNumber].options.cx &&
        defWidth
      ) {
        gen_private.pages[pageNumber].data[objNumber].options.cx = defWidth
      } // Endif.

      if (
        !gen_private.pages[pageNumber].data[objNumber].options.cy &&
        defHeight
      ) {
        gen_private.pages[pageNumber].data[objNumber].options.cy = defHeight
      } // Endif.

      var image_id = gen_private.type.msoffice.src_files_list.indexOf(
        image_path
      )
      var image_rel_id = -1

      if (image_id >= 0) {
        for (
          var j = 0, total_size_j = gen_private.pages[pageNumber].rels.length;
          j < total_size_j;
          j++
        ) {
          if (
            gen_private.pages[pageNumber].rels[j].target ===
            '../media/image' + (image_id + 1) + '.' + image_type
          ) {
            image_rel_id = j + 1
          } // Endif.
        } // Endif.
      } else {
        image_id = gen_private.type.msoffice.src_files_list.length
        gen_private.type.msoffice.src_files_list[image_id] = image_path
        gen_private.plugs.intAddAnyResourceToParse(
          'ppt\\media\\image' + (image_id + 1) + '.' + image_type,
          typeof image_path === 'string' ? 'file' : 'stream',
          image_path,
          null,
          false
        )
      } // Endif.

      if (image_rel_id === -1) {
        image_rel_id = gen_private.pages[pageNumber].rels.length + 1

        gen_private.pages[pageNumber].rels.push({
          type:
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          target: '../media/image' + (image_id + 1) + '.' + image_type,
          clear: 'data'
        })
      } // Endif.

      // -- psv 2015-01-21 Manually copied change below from https://github.com/Ziv-Barber/officegen/pull/39/files?w=1
      if ((opt || {}).link) {
        var link_rel_id = gen_private.pages[pageNumber].rels.length + 1

        gen_private.pages[pageNumber].rels.push({
          type:
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
          target: opt.link,
          targetMode: 'External'
        })

        gen_private.pages[pageNumber].data[objNumber].link_rel_id = link_rel_id
      } // Endif.
      // -- psv 2015-01-21 end edits

      gen_private.pages[pageNumber].data[objNumber].image_id = image_id
      gen_private.pages[pageNumber].data[objNumber].rel_id = image_rel_id

      if (typeof opt === 'string') {
        gen_private.pages[pageNumber].data[objNumber].options.color = opt
      } else if (typeof opt !== 'object' && typeof y_pos !== 'undefined') {
        gen_private.pages[pageNumber].data[objNumber].options.x = opt
        gen_private.pages[pageNumber].data[objNumber].options.y = y_pos

        if (typeof x_size !== 'undefined' && typeof y_size !== 'undefined') {
          gen_private.pages[pageNumber].data[objNumber].options.cx = x_size
          gen_private.pages[pageNumber].data[objNumber].options.cy = y_size
        } // Endif.
      } // Endif.

      addParagraphApiForBasicOpt(gen_private.pages[pageNumber].data[objNumber])
      addParagraphApiForEffects(gen_private.pages[pageNumber].data[objNumber])
      return gen_private.pages[pageNumber].data[objNumber].api
    }

    /**
     * ???.
     *
     * @param {object} text ???.
     * @param {object} opt ???.
     * @param {object} y_pos ???.
     * @param {object} x_size ???.
     * @param {object} y_size ???.
     * @param {object} opt_b ???.
     */
    slideObj.addP = function (text, opt, y_pos, x_size, y_size, opt_b) {
      var objNumber = gen_private.pages[pageNumber].data.length

      gen_private.pages[pageNumber].data[objNumber] = {}
      gen_private.pages[pageNumber].data[objNumber].type = 'p'
      gen_private.pages[pageNumber].data[objNumber].data = []
      gen_private.pages[pageNumber].data[objNumber].options =
        typeof opt === 'object' ? opt : {}

      if (typeof opt === 'string') {
        gen_private.pages[pageNumber].data[objNumber].options.color = opt
      } else if (typeof opt !== 'object' && typeof y_pos !== 'undefined') {
        gen_private.pages[pageNumber].data[objNumber].options.x = opt
        gen_private.pages[pageNumber].data[objNumber].options.y = y_pos

        if (typeof x_size !== 'undefined' && typeof y_size !== 'undefined') {
          gen_private.pages[pageNumber].data[objNumber].options.cx = x_size
          gen_private.pages[pageNumber].data[objNumber].options.cy = y_size
        } // Endif.
      } // Endif.

      var attrname

      if (typeof opt_b === 'object') {
        for (attrname in opt_b) {
          gen_private.pages[pageNumber].data[objNumber].options[attrname] =
            opt_b[attrname]
        }
      } else if (typeof x_size === 'object' && typeof y_size === 'undefined') {
        for (attrname in x_size) {
          gen_private.pages[pageNumber].data[objNumber].options[attrname] =
            x_size[attrname]
        }
      } // Endif.

      // BMK_TODO:

      return gen_private.pages[pageNumber].data[objNumber].data
    }

    /**
     * API to add direct custom xml code into the slide.
     *
     * @param {object} code The xml code to add.
     */
    slideObj.addDirectXmlCode = function (code) {
      var objNumber = gen_private.pages[pageNumber].data.length

      gen_private.pages[pageNumber].data[objNumber] = {}
      gen_private.pages[pageNumber].data[objNumber].type = 'xml'
      gen_private.pages[pageNumber].data[objNumber].data = code
      gen_private.pages[pageNumber].data[objNumber].options = {}

      return gen_private.pages[pageNumber].data[objNumber].data
    }

    slideObj.addDateToHeader = function () {
      if (!gen_private.pages[pageNumber].header) {
        gen_private.pages[pageNumber].header = {}
      } // Endif.

      // gen_private.pages[pageNumber]
      // <a:fld id="{5C7A2A3D-B97F-4EB0-B937-FE8C3AFCAC1A}" type="datetime1">
      // BMK_TODO:
    }

    gen_private.plugs.intAddAnyResourceToParse(
      'ppt\\slides\\slide' + (pageNumber + 1) + '.xml',
      'buffer',
      gen_private.pages[pageNumber],
      cbMakePptxSlide,
      false
    )
    gen_private.plugs.intAddAnyResourceToParse(
      'ppt\\slides\\_rels\\slide' + (pageNumber + 1) + '.xml.rels',
      'buffer',
      gen_private.pages[pageNumber].rels,
      gen_private.plugs.type.msoffice.cbMakeRels,
      false
    )

    // Signal to the plugins about a new slide:
    gen_private.features.type.pptx.emitEvent('newPage', {
      genobj: genobj,
      page: slideObj,
      pageData: gen_private.pages[pageNumber],
      pageNumber: pageNumber,
      slideOptions: slideOptions
    })
    return slideObj
  }

  /**
   * Change the master slide.
   *
   * This method changing the master slide. Temporary code - don't use this method in real production code!
   *
   * @param {string} xmlCode - Master slide xml code.
   */
  genobj.setMasterSlideXml = function (xmlCode) {
    gen_private.masterSlideXmlCode = xmlCode
  }

  // Tell all the features (plugins) to add extra API:
  gen_private.features.type.pptx.emitEvent('makeDocApi', genobj)

  return this
}

baseobj.plugins.registerDocType(
  'pptx',
  makePptx,
  {},
  baseobj.docType.PRESENTATION,
  'Microsoft PowerPoint Document'
)
baseobj.plugins.registerDocType(
  'ppsx',
  makePptx,
  {},
  baseobj.docType.PRESENTATION,
  'Microsoft PowerPoint Slideshow Document'
)
