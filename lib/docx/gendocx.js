//
// officegen: All the code to generate DOCX files.
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
 * Basicgen plugin to create docx files (Microsoft World).
 */

var baseobj = require('../core/index.js')
var msdoc = require('../msdoc/msofficegen.js')
var docxP = require('./docx-p.js')
var docxTable = require('./docxtable.js')
var xmlBuilder = require('xmlbuilder')

var docplugman = require('../core/docplug')

// Officegen docx plugins:
var plugHeadfoot = require('./docxplg-headfoot')
// BMK_DOCX_PLUG:

var defaultStyleXML =
  '<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267"><w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/><w:lsdException w:name="toc 1" w:uiPriority="39"/><w:lsdException w:name="toc 2" w:uiPriority="39"/><w:lsdException w:name="toc 3" w:uiPriority="39"/><w:lsdException w:name="toc 4" w:uiPriority="39"/><w:lsdException w:name="toc 5" w:uiPriority="39"/><w:lsdException w:name="toc 6" w:uiPriority="39"/><w:lsdException w:name="toc 7" w:uiPriority="39"/><w:lsdException w:name="toc 8" w:uiPriority="39"/><w:lsdException w:name="toc 9" w:uiPriority="39"/><w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/><w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/><w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/><w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/><w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Revision" w:unhideWhenUsed="0"/><w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/><w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/><w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/><w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/><w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/><w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/><w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/><w:lsdException w:name="Bibliography" w:uiPriority="37"/><w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/></w:latentStyles><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rsid w:val="00A02F19"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>'

/**
 * Extend officegen object with DOCX support.
 * <br /><br />
 *
 * This method extending the given officegen object to create DOCX document.
 *
 * @param[in] genobj The object to extend.
 * @param[in] new_type The type of object to create.
 * @param[in] options The object's options.
 * @param[in] gen_private Access to the internals of this object.
 * @param[in] type_info Additional information about this type.
 * @constructor
 * @name makeDocx
 */
function makeDocx(genobj, new_type, options, gen_private, type_info) {
  /**
   * Prepare the default data.
   * @param {object} docpluginman Access to the document plugins manager.
   */
  function setDefaultDocValues(docpluginman) {
    // var pptxData = docpluginman.getDataStorage()
    // Please put any setting that API can override here:
  }

  /**
   * Prepare everything to generate a docx zip.
   */
  function cbPrepareDocxToGenerate() {
    // Tell all the features (plugins) that we are about to generate a new document zip:
    gen_private.features.type.docx.emitEvent('beforeGen', genobj)

    // Allow some plugins to do more stuff after all the plugins added their data:
    gen_private.features.type.docx.emitEvent('beforeGenFinal', genobj)
  }

  /**
   * ???.
   *
   * @param[in] data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeDocxFontsTable(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="A00002EF" w:usb1="4000207B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000009F" w:csb1="00000000"/></w:font><w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/><w:sig w:usb0="20002A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="20002A87" w:usb1="80000000" w:usb2="00000008" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/></w:font><w:font w:name="Cambria"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:sig w:usb0="A00002EF" w:usb1="4000004B" w:usb2="00000000" w:usb3="00000000" w:csb0="0000009F" w:csb1="00000000"/></w:font></w:fonts>'
    )
  }

  /**
   * ???.
   *
   * @param[in] data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeDocxSettings(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="120"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat/><w:rsids><w:rsidRoot w:val="00A94AF2"/><w:rsid w:val="00A02F19"/><w:rsid w:val="00A94AF2"/></w:rsids><m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="off"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr><w:themeFontLang w:val="en-US" w:bidi="en-US"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="2050"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>'
    )
  }

  // NJC - added to support bullets and multi-level ordered lists
  function cbMakeDocxNumbers(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0"><w:nsid w:val="6791653D" /><w:multiLevelType w:val="multilevel" /><w:tmpl w:val="32B23E04" /><w:lvl w:ilvl="0"><w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%1." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="720" /></w:tabs><w:ind w:left="720" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1" /><w:numFmt w:val="upperLetter" /><w:lvlText w:val="%2." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="1440" /></w:tabs><w:ind w:left="1440" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1" /><w:numFmt w:val="upperRoman" /><w:lvlText w:val="%3." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="2160" /></w:tabs><w:ind w:left="2160" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%4." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="2880" /></w:tabs><w:ind w:left="2880" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%5." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="3600" /></w:tabs><w:ind w:left="3600" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%6." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="4320" /></w:tabs><w:ind w:left="4320" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1" /><w:numFmt w:val="upperLetter" /><w:lvlText w:val="%7." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="5040" /></w:tabs><w:ind w:left="5040" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1" /><w:numFmt w:val="upperRoman" /><w:lvlText w:val="%8." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="5760" /></w:tabs><w:ind w:left="5760" w:hanging="720" /></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%9." /><w:lvlJc w:val="left" /><w:pPr><w:tabs><w:tab w:val="num" w:pos="6480" /></w:tabs><w:ind w:left="6480" w:hanging="720" /></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:nsid w:val="709F643A" /><w:multiLevelType w:val="hybridMultilevel" /><w:tmpl w:val="B8B464D0" /><w:lvl w:ilvl="0" w:tplc="04090001"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="720" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="o" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="1440" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="2160" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="2880" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="o" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="3600" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="4320" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5040" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="o" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5760" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New" w:hint="default" /></w:rPr></w:lvl><w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="1"><w:start w:val="1" /><w:numFmt w:val="bullet" /><w:lvlText w:val="" /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="6480" w:hanging="360" /></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default" /></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1" /></w:num><w:num w:numId="2"><w:abstractNumId w:val="0" /></w:num></w:numbering>'
    )
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
   * ???.
   *
   * @param[in] data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeDocxWeb(data) {
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:optimizeForBrowser/></w:webSettings>'
    )
  }

  /**
   * ???.
   *
   * @param[in] data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeDocxStyles(data) {
    var styleXML = (data && data.styleXML) || defaultStyleXML
    return (
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) + styleXML
    )
  }

  /**
   * ???.
   *
   * @param[in] data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeDocxApp(data) {
    var userName =
      genobj.options.author || genobj.options.creator || 'officegen'
    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>1</TotalTime><Pages>1</Pages><Words>0</Words><Characters>0</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company>' +
      userName +
      '</Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>0</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>'
    return outString
  }

  /**
   * Create the document's itself resource.
   *
   * @param[in] data Ignored by this callback function.
   * @return Text string.
   */
  function cbMakeDocxDocument(data) {
    data.docStartExtra = data.docStartExtra || ''
    data.docEndExtra = data.docEndExtra || ''

    var outString =
      gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml(data) +
      '<w:' +
      data.docType +
      ' xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">' +
      data.docStartExtra
    var objs_list = data.data
    var bookmarkId = 0

    // In case of an empty document - just place an empty paragraph:
    if (!objs_list.length) {
      outString += '<w:p w:rsidR="009F2180" w:rsidRDefault="009F2180">'
      if (data.pStyleDef) {
        outString += '<w:pPr><w:pStyle w:val="' + data.pStyleDef + '"/></w:pPr>'
      } // Endif.

      outString += '</w:p>'
    } // Endif.

    // BMK_DOCX_P: Work on all the stored paragraphs inside this document:
    for (var i = 0, total_size = objs_list.length; i < total_size; i++) {
      if (objs_list[i] && objs_list[i].type === 'table') {
        var table_obj = docxTable.getTable(
          objs_list[i].data,
          objs_list[i].options
        )
        var table_xml = xmlBuilder
          .create(table_obj, {
            version: '1.0',
            encoding: 'UTF-8',
            separateArrayItems: true,
            standalone: true
          })
          .toString({ pretty: true, indent: '  ', newline: '\n' })
        outString += table_xml
        continue
      } // Endif.

      outString += '<w:p w:rsidR="00A77427" w:rsidRDefault="007F1D13">'
      var pPrData = ''

      if (objs_list[i].options) {
        pPrData += '<w:ind'
        if (objs_list[i].options.indentLeft) {
          pPrData += ` w:left="${objs_list[i].options.indentLeft}"`
        }
        if (objs_list[i].options.indentFirstLine) {
          pPrData += ` w:firstLine="${objs_list[i].options.indentFirstLine}"`
        }
        pPrData += '/>'

        if (objs_list[i].options.align) {
          switch (objs_list[i].options.align) {
            case 'center':
              pPrData += '<w:jc w:val="center"/>'
              break

            case 'right':
              pPrData += '<w:jc w:val="right"/>'
              break

            case 'justify':
              pPrData += '<w:jc w:val="both"/>'
              break
          } // End of switch.
        } // Endif.

        if (objs_list[i].options.list_type) {
          pPrData +=
            '<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="' +
            objs_list[i].options.list_level +
            '"/><w:numId w:val="' +
            objs_list[i].options.list_type +
            '"/></w:numPr>'
        } // Endif.

        if (objs_list[i].options.backline) {
          pPrData +=
            '<w:pPr><w:shd w:val="solid" w:color="' +
            objs_list[i].options.backline +
            '" w:fill="auto"/></w:pPr>'
        } // Endif.

        if (objs_list[i].options.spacing) {
          var pSpacing = objs_list[i].options.spacing
          if (typeof pSpacing === 'object') {
            pPrData += '<w:spacing'
            for (var pSpacingKey of [
              'before',
              'after',
              'line',
              'lineRule',
              'beforeAutospacing',
              'afterAutospacing'
            ]) {
              if (pSpacingKey in pSpacing)
                pPrData +=
                  ' w:' + pSpacingKey + '="' + pSpacing[pSpacingKey] + '"'
            }
            pPrData += '/>'
          } // Endif.
        } // Endif.

        if (objs_list[i].options.rtl) {
          pPrData += '<w:bidi w:val="1"/>'
        } // Endif.

        if (objs_list[i].options.textAlignment) {
          pPrData +=
            '<w:textAlignment w:val="' +
            objs_list[i].options.textAlignment +
            '"/>'
        } // Endif.
      } // Endif.

      // Some resource types have default style in case that there's no style settings:
      var pStyleDef =
        (objs_list[i].options && objs_list[i].options.pStyleDef) ||
        data.pStyleDef
      if (!pPrData && pStyleDef) {
        pPrData = '<w:pStyle w:val="' + pStyleDef + '"/>'
      } else if (objs_list[i].options && objs_list[i].options.force_style) {
        pPrData = '<w:pStyle w:val="' + objs_list[i].options.force_style + '"/>'
      } // Endif.

      if (pPrData) {
        outString += '<w:pPr>' + pPrData + '</w:pPr>'
      } // Endif.

      // Work on all the objects in the document:
      for (
        var j = 0, total_size_j = objs_list[i].data.length;
        j < total_size_j;
        j++
      ) {
        if (objs_list[i].data[j]) {
          var rExtra = ''
          var tExtra = ''
          var rPrData = ''
          var colorCode
          var valType
          var sizeVal
          var hyperlinkOn = false

          if (objs_list[i].data[j].options) {
            if (objs_list[i].data[j].options.color) {
              rPrData +=
                '<w:color w:val="' + objs_list[i].data[j].options.color + '"/>'
            } // Endif.

            if (objs_list[i].data[j].options.back) {
              colorCode = objs_list[i].data[j].options.shdColor || 'auto'
              valType = objs_list[i].data[j].options.shdType || 'clear'

              rPrData +=
                '<w:shd w:val="' +
                valType +
                '" w:color="' +
                colorCode +
                '" w:fill="' +
                objs_list[i].data[j].options.back +
                '"/>'
            } // Endif.

            if (objs_list[i].data[j].options.highlight) {
              valType = 'yellow'

              if (typeof objs_list[i].data[j].options.highlight === 'string') {
                valType = objs_list[i].data[j].options.highlight
              } // Endif.

              rPrData += '<w:highlight w:val="' + valType + '"/>'
            } // Endif.

            if (objs_list[i].data[j].options.bold) {
              rPrData += '<w:b/><w:bCs/>'
            } // Endif.

            if (objs_list[i].data[j].options.italic) {
              rPrData += '<w:i/><w:iCs/>'
            } // Endif.

            if (objs_list[i].data[j].options.underline) {
              valType = 'single'

              if (typeof objs_list[i].data[j].options.underline === 'string') {
                valType = objs_list[i].data[j].options.underline
              } // Endif.

              rPrData += '<w:u w:val="' + valType + '"/>'
            } // Endif.

            // Since officegen 0.5.0 and later:
            if (objs_list[i].data[j].options.superscript) {
              rPrData += '<w:vertAlign w:val="superscript" />'
            } else if (objs_list[i].data[j].options.subscript) {
              rPrData += '<w:vertAlign w:val="subscript" />'
            } // Endif.

            if (objs_list[i].data[j].options.strikethrough) {
              rPrData += '<w:strike/>'
            } // Endif.

            var fontFaceInfo = ''
            if (objs_list[i].data[j].options.font_face) {
              fontFaceInfo +=
                ' w:ascii="' +
                objs_list[i].data[j].options.font_face +
                '" w:eastAsia="' +
                (objs_list[i].data[j].options.font_face_east ||
                  objs_list[i].data[j].options.font_face) +
                '" w:hAnsi="' +
                (objs_list[i].data[j].options.font_face_h ||
                  objs_list[i].data[j].options.font_face) +
                '" w:cs="' +
                (objs_list[i].data[j].options.font_face_cs ||
                  objs_list[i].data[j].options.font_face) +
                '"'
            } // Endif.

            if (
              objs_list[i].data[j].options.font_hint ||
              objs_list[i].data[j].options.font_rtl
            ) {
              fontFaceInfo +=
                ' w:hint="' +
                (objs_list[i].data[j].options.font_hint || 'cs') +
                '"'
            } // Endif.

            if (fontFaceInfo) {
              rPrData += '<w:rFonts' + fontFaceInfo + ' />'
            } // Endif.

            if (objs_list[i].data[j].options.font_size) {
              var fontSizeInHalfPoints =
                2 * objs_list[i].data[j].options.font_size
              rPrData +=
                '<w:sz w:val="' +
                fontSizeInHalfPoints +
                '"/><w:szCs w:val="' +
                fontSizeInHalfPoints +
                '"/>'
            } // Endif.

            if (objs_list[i].data[j].options.border) {
              colorCode = 'auto'
              valType = 'single'
              sizeVal = 4

              if (
                typeof objs_list[i].data[j].options.borderColor === 'string'
              ) {
                colorCode = objs_list[i].data[j].options.borderColor
              } // Endif.

              if (typeof objs_list[i].data[j].options.border === 'string') {
                valType = objs_list[i].data[j].options.border
              } // Endif.

              /* eslint-disable no-self-compare */
              if (
                typeof objs_list[i].data[j].options.borderSize === 'number' &&
                objs_list[i].data[j].options.borderSize &&
                objs_list[i].data[j].options.borderSize ===
                  objs_list[i].data[j].options.borderSize
              ) {
                sizeVal = objs_list[i].data[j].options.borderSize
              } // Endif.

              rPrData +=
                '<w:bdr w:val="' +
                valType +
                '" w:sz="' +
                sizeVal +
                '" w:space="0" w:color="' +
                colorCode +
                '"/>'
            } // Endif.

            // Hyperlink support:
            if (objs_list[i].data[j].options.hyperlink) {
              outString +=
                '<w:hyperlink w:anchor="' +
                objs_list[i].data[j].options.hyperlink +
                '">'
              hyperlinkOn = true

              if (!rPrData) {
                rPrData = '<w:rStyle w:val="Hyperlink"/>'
              } // Endif.
            } // Endif.

            if (objs_list[i].data[j].options.rtl) {
              rPrData += '<w:rtl w:val="1"/>'
            } // Endif.

            if (objs_list[i].data[j].options.lang) {
              outString += '<w:lang w:bidi="' + objs_list[i].data[j].lang + '">'
            } // Endif.
          } // Endif.

          // Field support:
          if (objs_list[i].data[j].fieldObj) {
            outString +=
              '<w:fldSimple w:instr="' + objs_list[i].data[j].fieldObj + '">'
          } // Endif.

          if (objs_list[i].data[j].text) {
            if (
              objs_list[i].data[j].text[0] === ' ' ||
              objs_list[i].data[j].text[
                objs_list[i].data[j].text.length - 1
              ] === ' '
            ) {
              tExtra += ' xml:space="preserve"'
            } // Endif.

            if (objs_list[i].data[j].link_rel_id) {
              outString +=
                '<w:hyperlink r:id="rId' +
                objs_list[i].data[j].link_rel_id +
                '">'
            }

            outString += '<w:r' + rExtra + '>'

            if (rPrData) {
              outString += '<w:rPr>' + rPrData + '</w:rPr>'
            } // Endif.

            outString +=
              '<w:t' +
              tExtra +
              '>' +
              gen_private.plugs.type.msoffice.escapeText(
                objs_list[i].data[j].text
              ) +
              '</w:t></w:r>'

            if (objs_list[i].data[j].link_rel_id) {
              outString += '</w:hyperlink>'
            }
          } else if (objs_list[i].data[j].page_break) {
            outString += '<w:r><w:br w:type="page"/></w:r>'
          } else if (objs_list[i].data[j].line_break) {
            outString += '<w:r><w:br/></w:r>'
          } else if (objs_list[i].data[j].horizontal_line) {
            outString +=
              '<w:r><w:pict><v:rect style="width:0height:.75pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#e0e0e0" stroked="f"/></w:pict></w:r>'

            // Bookmark start support:
          } else if (objs_list[i].data[j].bookmark_start) {
            outString +=
              '<w:bookmarkStart w:id="' +
              bookmarkId +
              '" w:name="' +
              objs_list[i].data[j].bookmark_start +
              '"/>'

            // Bookmark end support:
          } else if (objs_list[i].data[j].bookmark_end) {
            outString += '<w:bookmarkEnd w:id="' + bookmarkId + '"/>'
            bookmarkId++
          } else if (objs_list[i].data[j].image) {
            outString += '<w:r' + rExtra + '>'

            rPrData += '<w:noProof/>'

            if (rPrData) {
              outString += '<w:rPr>' + rPrData + '</w:rPr>'
            } // Endif.

            // 914400L / 96DPI
            var pixelToEmu = 9525

            outString += '<w:drawing>'
            outString += '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            outString +=
              '<wp:extent cx="' +
              parseSmartNumber(
                objs_list[i].data[j].options.cx,
                10000, // BMK_TODO: Change it to the maximum X size.
                320,
                10000, // BMK_TODO: Change it to the maximum X size.
                pixelToEmu
              ) +
              '" cy="' +
              parseSmartNumber(
                objs_list[i].data[j].options.cy,
                10000, // BMK_TODO: Change it to the maximum Y size.
                200,
                10000, // BMK_TODO: Change it to the maximum Y size.
                pixelToEmu
              ) +
              '"/>'
            outString += '<wp:effectExtent l="19050" t="0" r="9525" b="0"/>'

            outString +=
              '<wp:docPr id="' +
              (objs_list[i].data[j].image_id + 1) +
              '" name="Picture ' +
              objs_list[i].data[j].image_id +
              '" descr="Picture ' +
              objs_list[i].data[j].image_id +
              '">'
            if (objs_list[i].data[j].link_rel_id) {
              outString +=
                '<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:id="rId' +
                objs_list[i].data[j].link_rel_id +
                '"/>'
            }
            outString += '</wp:docPr>'

            outString += '<wp:cNvGraphicFramePr>'
            outString +=
              '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
            outString += '</wp:cNvGraphicFramePr>'
            outString +=
              '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            outString +=
              '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            outString +=
              '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            outString += '<pic:nvPicPr>'
            outString +=
              '<pic:cNvPr id="0" name="Picture ' +
              objs_list[i].data[j].image_id +
              '"/>'
            outString += '<pic:cNvPicPr/>'
            outString += '</pic:nvPicPr>'
            outString += '<pic:blipFill>'
            outString +=
              '<a:blip r:embed="rId' +
              objs_list[i].data[j].rel_id +
              '" cstate="print"/>'
            outString += '<a:stretch>'
            outString += '<a:fillRect/>'
            outString += '</a:stretch>'
            outString += '</pic:blipFill>'
            outString += '<pic:spPr>'
            outString += '<a:xfrm>'
            outString += '<a:off x="0" y="0"/>'
            outString +=
              '<a:ext cx="' +
              Math.round(objs_list[i].data[j].options.cx * pixelToEmu) +
              '" cy="' +
              Math.round(objs_list[i].data[j].options.cy * pixelToEmu) +
              '"/>'
            outString += '</a:xfrm>'
            outString += '<a:prstGeom prst="rect">'
            outString += '<a:avLst/>'
            outString += '</a:prstGeom>'
            outString += '</pic:spPr>'
            outString += '</pic:pic>'
            outString += '</a:graphicData>'
            outString += '</a:graphic>'
            outString += '</wp:inline>'
            outString += '</w:drawing>'

            outString += '</w:r>'
          } // Endif.

          // Field support:
          if (objs_list[i].data[j].fieldObj) {
            outString += '</w:fldSimple>'
          } // Endif.

          if (hyperlinkOn) {
            outString += '</w:hyperlink>'
          } // Endif.
        } // Endif.
      } // Endif.

      outString += '</w:p>'
    } // End of for loop.

    if (data.docType === 'document') {
      outString += '<w:p w:rsidR="00A02F19" w:rsidRDefault="00A02F19"/>'

      var margins
      let width = 11906
      let height = 16838
      if (options.pageSize) {
        if (typeof options.pageSize === 'string') {
          switch (options.pageSize) {
            case 'A3':
              width = 16838
              height = 23811
              break
            case 'A4':
              width = 11906
              height = 16838
              break
            case 'letter paper':
              width = 15840
              height = 12240
              break
            default:
              // default is A4
              width = 11906
              height = 16838
              break
          }
        }
      }

      // Landscape orientation support:
      if (options.orientation && options.orientation === 'landscape') {
        margins = options.pageMargins || {
          top: 1800,
          right: 1440,
          bottom: 1800,
          left: 1440
        }
        width = (options.pageSize && options.pageSize.height) || width
        height = (options.pageSize && options.pageSize.width) || height

        outString +=
          '<w:sectPr w:rsidR="00A02F19" w:rsidSect="00897086">' +
          (docxData.secPrExtra ? docxData.secPrExtra : '') +
          `<w:pgSz w:w="${height}" w:h="${width}" w:orient="landscape"/>` +
          '<w:pgMar w:top="' +
          margins.top +
          '" w:right="' +
          margins.right +
          '" w:bottom="' +
          margins.bottom +
          '" w:left="' +
          margins.left +
          '" w:header="720" w:footer="720" w:gutter="0"/>' +
          '<w:cols' +
          (options.columns ? ' w:num="' + options.columns + '"' : '') +
          ' w:space="720"/>' +
          '<w:docGrid w:linePitch="360"/>' +
          '</w:sectPr>'
      } else {
        margins = options.pageMargins || {
          top: 1440,
          right: 1800,
          bottom: 1440,
          left: 1800
        }
        width = (options.pageSize && options.pageSize.width) || width
        height = (options.pageSize && options.pageSize.height) || height

        outString +=
          '<w:sectPr w:rsidR="00A02F19" w:rsidSect="00A02F19">' +
          (docxData.secPrExtra ? docxData.secPrExtra : '') +
          `<w:pgSz w:w="${width}" w:h="${height}"/>` +
          '<w:pgMar w:top="' +
          margins.top +
          '" w:right="' +
          margins.right +
          '" w:bottom="' +
          margins.bottom +
          '" w:left="' +
          margins.left +
          '" w:header="720" w:footer="720" w:gutter="0"/>' +
          '<w:cols' +
          (options.columns ? ' w:num="' + options.columns + '"' : '') +
          ' w:space="720"/>' +
          '<w:docGrid w:linePitch="360"/>' +
          '</w:sectPr>'
      } // Endif.
    } // Endif.

    outString += data.docEndExtra + '</w:' + data.docType + '>'

    return outString
  }

  // Save it so if some plugin need to generate document style resource then it can use it:
  genobj.cbMakeDocxDocument = cbMakeDocxDocument

  // Prepare genobj for MS-Office:
  msdoc.makemsdoc(genobj, new_type, options, gen_private, type_info)
  gen_private.plugs.type.msoffice.makeOfficeGenerator('word', 'document', {})

  genobj.on('clearData', function () {
    genobj.data.length = 0
  })

  genobj.on('beforeGen', cbPrepareDocxToGenerate)

  // Add the document's properties:
  gen_private.plugs.type.msoffice.addInfoType(
    'dc:title',
    '',
    'title',
    'setDocTitle'
  )
  gen_private.plugs.type.msoffice.addInfoType(
    'dc:subject',
    '',
    'subject',
    'setDocSubject'
  )
  gen_private.plugs.type.msoffice.addInfoType(
    'cp:keywords',
    '',
    'keywords',
    'setDocKeywords'
  )
  gen_private.plugs.type.msoffice.addInfoType(
    'dc:description',
    '',
    'description',
    'setDescription'
  )
  gen_private.plugs.type.msoffice.addInfoType(
    'cp:category',
    '',
    'category',
    'setDocCategory'
  )
  gen_private.plugs.type.msoffice.addInfoType(
    'cp:contentStatus',
    '',
    'status',
    'setDocStatus'
  )

  // Create the plugins manager:
  var plugsmanObj = new docplugman(
    genobj,
    gen_private,
    'docx',
    setDefaultDocValues
  )

  // We'll register now any officegen internal plugin that we want to always use for Word based documents:
  plugsmanObj.plugsList.push(new plugHeadfoot(plugsmanObj))
  // BMK_DOCX_PLUG:

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

  // Access to our data:
  var docxData = plugsmanObj.getDataStorage()

  gen_private.type.msoffice.files_list.push(
    {
      name: '/word/settings.xml',
      type:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
      clear: 'type'
    },
    {
      name: '/word/fontTable.xml',
      type:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
      clear: 'type'
    },
    {
      name: '/word/webSettings.xml',
      type:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
      clear: 'type'
    },
    {
      name: '/word/styles.xml',
      type:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
      clear: 'type'
    },
    {
      name: '/word/document.xml',
      type:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      clear: 'type'
    },
    {
      // NJC - 20161231 - added to support bullets
      name: '/word/numbering.xml',
      type:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
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
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
      target: 'settings.xml',
      clear: 'type'
    },
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
      target: 'webSettings.xml',
      clear: 'type'
    },
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
      target: 'fontTable.xml',
      clear: 'type'
    },
    {
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: 'theme/theme1.xml',
      clear: 'type'
    },
    {
      // NJC - 20161231 - added to support bullets
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
      target: 'numbering.xml',
      clear: 'type'
    }
  )

  genobj.data = [] // All the data will be placed here.

  gen_private.plugs.intAddAnyResourceToParse(
    'docProps\\app.xml',
    'buffer',
    null,
    cbMakeDocxApp,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'word\\fontTable.xml',
    'buffer',
    null,
    cbMakeDocxFontsTable,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'word\\settings.xml',
    'buffer',
    null,
    cbMakeDocxSettings,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'word\\webSettings.xml',
    'buffer',
    null,
    cbMakeDocxWeb,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'word\\styles.xml',
    'buffer',
    { styleXML: options.styleXML },
    cbMakeDocxStyles,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'word\\document.xml',
    'buffer',
    {
      docType: 'document',
      docStartExtra: '<w:body>',
      docEndExtra: '</w:body>',
      data: genobj.data
    },
    cbMakeDocxDocument,
    true
  )
  gen_private.plugs.intAddAnyResourceToParse(
    'word\\numbering.xml',
    'buffer',
    null,
    cbMakeDocxNumbers,
    true
  ) // added to support bullets

  gen_private.plugs.intAddAnyResourceToParse(
    'word\\_rels\\document.xml.rels',
    'buffer',
    gen_private.type.msoffice.rels_app,
    gen_private.plugs.type.msoffice.cbMakeRels,
    true
  )

  for (var i = 1; i <= 3; i++) {
    gen_private.plugs.intAddAnyResourceToParse(
      'word\\_rels\\header' + i + '.xml.rels',
      'buffer',
      gen_private.type.msoffice.rels_app,
      gen_private.plugs.type.msoffice.cbMakeRels,
      true
    )
    gen_private.plugs.intAddAnyResourceToParse(
      'word\\_rels\\footer' + i + '.xml.rels',
      'buffer',
      gen_private.type.msoffice.rels_app,
      gen_private.plugs.type.msoffice.cbMakeRels,
      true
    )
  } // End of for loop.

  // ----- API for Word documents: -----

  /**
   * Create a new paragraph.
   *
   * @param {string} options Default options for all the objects inside this paragraph.
   */
  genobj.createP = function (options) {
    // Create a new instance of the paragraph object:
    return new docxP(genobj, gen_private, 'docx', genobj.data, {}, options)
  }

  /**
   * ???.
   *
   * @param {object} options ???.
   */
  genobj.createListOfDots = function (options) {
    var newP = genobj.createP(options)

    newP.options.list_type = '1'
    newP.options.list_level = 0

    return newP
  }

  /**
   * Create a nested unordered list based paragraph.
   *
   * @param {object} options ???.
   */
  genobj.createNestedUnOrderedList = function (options) {
    var newP = genobj.createP(options)

    newP.options.list_type = '1'
    if (!options || !options.level) {
      newP.options.list_level = 0
    } else {
      newP.options.list_level = options.level - 1
    }

    return newP
  }

  /**
   * Create a list of numbers based paragraph.
   *
   * @param {object} options ???.
   */
  genobj.createListOfNumbers = function (options) {
    var newP = genobj.createP(options)

    newP.options.list_type = '2'
    newP.options.list_level = 0

    return newP
  }

  /**
   * Create a nested list of numbers based paragraph.
   *
   * @param {object} options ???.
   */
  genobj.createNestedOrderedList = function (options) {
    var newP = genobj.createP(options)

    newP.options.list_type = '2'
    if (!options || !options.level) {
      newP.options.list_level = 0
    } else {
      newP.options.list_level = options.level - 1
    }

    return newP
  }

  /**
   * Add a page break.
   * <br /><br />
   *
   * This method add a page break to the current Word document.
   */
  genobj.putPageBreak = function () {
    var newP = {}

    newP.data = [{ page_break: true }]

    genobj.data[genobj.data.length] = newP
    return newP
  }

  /**
   * Add a page break.
   * <br /><br />
   *
   * This method add a page break to the current Word document.
   */
  genobj.addPageBreak = function () {
    var newP = {}

    newP.data = [{ page_break: true }]

    genobj.data[genobj.data.length] = newP
    return newP
  }

  /**
   * Create a table.
   * <br /><br />
   *
   * This method add a table to the current word document.
   *
   * @param {object} data ???.
   * @param {object} options ???.
   */
  genobj.createTable = function (data, options) {
    var newP = genobj.createP(options)
    newP.data = data
    newP.type = 'table'
    return newP
  }

  /**
   * Create Json.
   * <br /><br />
   *
   * @param {object} data ???.
   * @param {object} newP ???.
   */
  genobj.createJson = function (data, newP) {
    if (data.type !== 'table') {
      newP = newP || genobj.createP(data.lopt || {})
    }

    switch (data.type) {
      case 'text':
        newP.addText(data.val, data.opt)
        break
      case 'linebreak':
        newP.addLineBreak()
        break
      case 'horizontalline':
        newP.addHorizontalLine()
        break
      case 'image':
        // Improved by peizhuang in Aug 2016 (added data.opt):
        // data.imagetype been added by Ziv Barber in Aug 2016.
        newP.addImage(data.path, data.opt || {}, data.imagetype)
        break
      case 'pagebreak':
        newP = genobj.putPageBreak()
        break
      case 'table':
        newP = genobj.createTable(data.val, data.opt)
        break
      case 'numlist':
        newP = genobj.createListOfNumbers()
        break
      case 'dotlist':
        newP = genobj.createListOfDots()
        break
    }

    return newP
  }

  /**
   * Create a document by json data.
   * <br /><br />
   *
   * @param {array} dataArray ???.
   */
  genobj.createByJson = function (dataArray) {
    var newP = {}
    dataArray = [].concat(dataArray || [])
    dataArray.forEach(function (data) {
      if (Array.isArray(data)) {
        newP = genobj.createP(data[0] || {})
        data.forEach(function (d) {
          newP = genobj.createJson(d, newP)
        })
      } else {
        newP = genobj.createJson(data)
      }
    })
    return newP
  }

  // Tell all the features (plugins) to add extra API:
  gen_private.features.type.docx.emitEvent('makeDocApi', genobj)

  return this
}

baseobj.plugins.registerDocType(
  'docx',
  makeDocx,
  {},
  baseobj.docType.TEXT,
  'Microsoft Word Document'
)
