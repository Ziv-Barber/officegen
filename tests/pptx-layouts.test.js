//
// officegen: pptx layouts plugin tests
//
// Please put here all the pptx layouts plugin tests.
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

var assert = require('assert')
var officegen = require('../')
var fs = require('fs')
var path = require('path')

// var pluginLayouts = require('../lib/pptxplg-layouts')

var outDir = path.join(__dirname, '../tmp/')

// Common error method
var onError = function(err) {
  console.log(err)
  assert(false)
}

describe('PPTX Layouts plugin', function() {
  this.slow(1000)

  before(function(done) {
    fs.mkdir(outDir, 0o777, function(err) {
      if (err) {
      } // Endif.

      done()
    })
  })

  it('creates a presentation with the title layout', function(done) {
    var slide
    var pptx = officegen({
      type: 'pptx',
      extraPlugs: [
        // pluginLayouts // The 'pptxplg-layouts' plugin.
      ]
    })
    pptx.on('error', onError)

    pptx.setDocTitle('Testing Layouts')

    //
    // Slide #1:
    //

    slide = pptx.makeNewSlide({
      useLayout: 'title'
    })

    //
    // Slide #2:
    //

    slide = pptx.makeNewSlide({
      useLayout: 'title'
    })

    slide.setTitle('The title')
    slide.setSubTitle('Another text')

    // Add a speaker note:
    slide.setSpeakerNote(
      'This is a speaker note! Using the new setSpeakerNote feature of the slide API.'
    )

    //
    // Slide #3:
    //

    slide = pptx.makeNewSlide({
      useLayout: 'title'
    })

    slide.setTitle([
      { text: 'Hello ', options: { font_size: 56 } },
      {
        text: 'World!',
        options: { font_size: 56, font_face: 'Arial', color: 'ffff00' }
      }
    ])
    slide.setSubTitle('Another text')

    //
    // Slide #4:
    //

    slide = pptx.makeTitleSlide()

    //
    // Slide #5:
    //

    slide = pptx.makeTitleSlide('The title of this slide', 'Sub title')

    //
    // Slide #6:
    //

    slide = pptx.makeTitleSlide(
      [
        { text: 'Hello ', options: { font_size: 56 } },
        {
          text: 'World!',
          options: { font_size: 56, font_face: 'Arial', color: 'ffff00' }
        }
      ],
      'Sub title'
    )

    //
    // Slide #7:
    //

    slide = pptx.makeObjSlide('The title of slide 7', [
      { text: '', options: { listType: 'dot' } },
      { text: 'Some ', options: { font_size: 56 } },
      {
        text: 'data',
        options: { font_size: 56, font_face: 'Arial', color: 'ff8800' }
      }
    ])

    slide.useLayout.isDate = false
    slide.setFooter('Message in the footer')

    //
    // Slide #8:
    //

    slide = pptx.makeSecHeadSlide('The title of slide 8', 'Sub title')

    //
    // Create a custom layout:
    //

    pptx.makeNewLayout('1_Title Slide-2lines', {
      display: '1_Title Slide-2lines',
      back: {
        type: 'solid',
        color: 'tx2',
        scheme: true
      },
      xmlCode:
        '<p:sp><p:nvSpPr><p:cNvPr id="4" name="מלבן 3"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{FCF6DB08-1F1C-44AA-B6F7-C031B856AC29}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="5143500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent4"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-IL"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="9" name="מלבן 8"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{D005C2FB-DC1A-4D47-B65E-94719652F9D1}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="3425482"/><a:ext cx="9144000" cy="1718017"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent3"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-IL"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="888901" y="864893"/><a:ext cx="7375868" cy="1441518"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="t" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="4800" b="1"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="Open Sans Semibold" panose="020B0706030804020204" pitchFamily="34" charset="0"/><a:cs typeface="Open Sans Semibold" panose="020B0706030804020204" pitchFamily="34" charset="0"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:rPr lang="en-US" sz="4800" dirty="0"/><a:t>Add Title Here</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="13" name="Text Placeholder 2"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6A3C93B3-F38C-45C3-82A7-C60718709D9A}"/></a:ext></a:extLst></p:cNvPr><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="32" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="888901" y="3961326"/><a:ext cx="7375868" cy="646329"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="0" tIns="45719" rIns="0" bIns="45719" rtlCol="0" anchor="t" anchorCtr="0"><a:spAutoFit/></a:bodyPr><a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr lang="en-US" sz="3600" b="1" dirty="0"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="Open Sans Semibold" panose="020B0706030804020204" pitchFamily="34" charset="0"/><a:cs typeface="Open Sans Semibold" panose="020B0706030804020204" pitchFamily="34" charset="0"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:pPr marL="179384" lvl="0" indent="-179384"><a:spcBef><a:spcPct val="0"/></a:spcBef></a:pPr><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Full Name</a:t></a:r></a:p></p:txBody></p:sp>'
    })

    //
    // Generate the pptx file:
    //

    var outFilename = 'test-ppt-layouts-1.pptx'
    var out = fs.createWriteStream(path.join(outDir, outFilename))
    pptx.generate(out)
    out.on('close', function() {
      done()
    })
  })
})
