//
// officegen: Generate the shapes inside a slide.
//
// Please refer to README.md for this module's documentations.
//
// Copyright (c) 2019 Ziv Barber;
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

var pptxShapes = require('./pptxshapes.js')

/**
 * Convert shape name to shape information.
 *
 * This method convert the shape information received from the user to the real shape information object.
 *
 * @param {object} shapeName Either the name of the shape or the shape information.
 * @return Information about this shape.
 */
function getShapeInfo(shapeName) {
  if (!shapeName) {
    return pptxShapes.RECTANGLE
  } // Endif.

  if (
    typeof shapeName === 'object' &&
    shapeName.name &&
    shapeName.displayName &&
    shapeName.avLst
  ) {
    return shapeName
  } // Endif.

  if (pptxShapes[shapeName]) {
    return pptxShapes[shapeName]
  } // Endif.

  for (var shapeIntName in pptxShapes) {
    if (pptxShapes[shapeIntName].name === shapeName) {
      return pptxShapes[shapeIntName]
    } // Endif.

    if (pptxShapes[shapeIntName].displayName === shapeName) {
      return pptxShapes[shapeIntName]
    } // Endif.
  } // End of for loop.

  return pptxShapes.RECTANGLE
}

/**
 * Generate the XML code to describe colors.
 *
 * @param {object} color_info Foreground color information.
 * @param {object} back_info Background color information.
 */
function createColorElements(color_info, back_info) {
  var outText = ''
  var colorVal
  var fillType = 'solid'
  var internalElements = ''

  if (back_info) {
    outText += '<p:bg><p:bgPr>'

    outText += createColorElements(back_info, false)

    outText += '<a:effectLst/>'
    // BMK_TODO: (add support for effects)

    outText += '</p:bgPr></p:bg>'
  } // Endif.

  if (color_info) {
    if (typeof color_info === 'string') {
      colorVal = color_info
    } else {
      if (color_info.type) {
        fillType = color_info.type
      } // Endif.

      if (color_info.color) {
        colorVal = color_info.color
      } // Endif.

      if (color_info.alpha) {
        internalElements +=
          '<a:alpha val="' + (100 - color_info.alpha) + '000"/>'
      } // Endif.
    } // Endif.

    switch (fillType) {
      case 'solid':
        if (color_info.scheme) {
          outText +=
            '<a:solidFill><a:schemeClr val="' + colorVal + '"/></a:solidFill>'
          break
        } // Endif.

        outText +=
          '<a:solidFill><a:srgbClr val="' +
          colorVal +
          '">' +
          internalElements +
          '</a:srgbClr></a:solidFill>'
        break

      case 'gradient':
        outText += '<a:gradFill flip="none" rotWithShape="1"><a:gsLst>'

        for (var item in colorVal) {
          if (typeof colorVal[item] === 'string') {
            // Positions are inverted because they start with 100000:
            outText +=
              '<a:gs pos="' +
              (100000 - Math.round((100000 / (colorVal.length - 1)) * item)) +
              '"><a:srgbClr val="' +
              colorVal[item] +
              '">' +
              internalElements +
              '</a:srgbClr></a:gs>'
          } else {
            outText +=
              '<a:gs pos="' +
              colorVal[item].position * 1000 +
              '"><a:srgbClr val="' +
              colorVal[item].color +
              '">' +
              internalElements +
              '</a:srgbClr></a:gs>'
          } // Endif.
        } // End of for loop.

        if (typeof color_info.angle !== 'undefined') {
          outText +=
            '</a:gsLst><a:lin ang="' +
            color_info.angle * 100000 +
            '" scaled="1"/><a:tileRect/></a:gradFill>'
        } else {
          outText +=
            '</a:gsLst><a:path path="circle"><a:fillToRect l="100000" t="100000"/></a:path><a:tileRect r="-100000" b="-100000"/></a:gradFill>'
        } // Endif.
        break
    } // End of switch.
  } // Endif.

  return outText
}

function ShapeOptions(options) {
  var usedOptions = typeof options === 'object' ? options : {}

  function setCurValueInt(value, defValue) {
    if (typeof value === 'number' && !isNaN(value)) {
      return value
    } // Endif.

    return defValue
  }

  function setCurValueString(value, defValue) {
    if (typeof value === 'string') {
      return value
    } // Endif.

    return defValue
  }

  function setCurValueBool(value, defValue) {
    if (value === true || value === false) {
      return value
    } // Endif.

    return defValue
  }

  function setCurValueIntAndOff(value, defValue) {
    if (value === false) {
      return false
    } // Endif.

    return setCurValueInt(value, defValue)
  }

  function setCurValueStringOff(value, defValue) {
    if (value === false || value === null) {
      return null
    } // Endif.

    return setCurValueString(value, defValue)
  }

  this.x = setCurValueIntAndOff(usedOptions.x, false)
  this.y = setCurValueIntAndOff(usedOptions.y, false)
  this.cx = setCurValueIntAndOff(usedOptions.cx, false) // 2819400
  this.cy = setCurValueIntAndOff(usedOptions.cy, false) // 369332
  this.shape = setCurValueStringOff(usedOptions.shape, null)
  this.flip_vertical = setCurValueBool(usedOptions.flip_vertical, false)
  this.flip_horizontal = setCurValueBool(usedOptions.flip_horizontal, false)
  this.rotate = setCurValueIntAndOff(usedOptions.rotate, 0)
  this.ph = setCurValueStringOff(usedOptions.ph, null)
  this.phIdx = setCurValueInt(usedOptions.phIdx, 0)
  this.phSz = setCurValueString(usedOptions.phSz, '')
  this.fill = usedOptions.fill || null
  this.line = this.line || null
  this.line_size = setCurValueInt(usedOptions.line_size, 0)
  this.line_head = this.line_head || null
  this.line_tail = this.line_tail || null
  this.effects = this.effects || null // Array of effects
  this.align = setCurValueStringOff(usedOptions.align, null)
  this.indentLevel = setCurValueInt(usedOptions.indentLevel, 0)
  // this.bodyProp = {}
  return this
}

function createShapeOptions(options) {
  return new ShapeOptions(options)
}

module.exports = {
  createShapeOptions: createShapeOptions,
  getShapeInfo: getShapeInfo,
  createColorElements: createColorElements,
  shapes: pptxShapes
}
