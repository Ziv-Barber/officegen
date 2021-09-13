//
// officegen: paragraph API for docx
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

var fast_image_size = require('fast-image-size')

/**
 * This function implementing the paragraph API for docx based documents.
 *
 * @param {object} genobj The document object to work on it.
 * @param {object} genPrivate Access to the internals of this object.
 * @param {string} docType Must be the same as the new_type parameter to the document generator's constructor.
 * @param {object} dataContainer Optional container to store the new object.
 * @param {object} extraSettings Optional extra settings for the paragraph API itself.
 * @param {object} options Paragraph options.
 * @summary Implementation of the pptx document plugins system so it'll be easier to add new features.
 * @constructor
 * @name MakeDocxP
 */
function MakeDocxP(
  genobj,
  genPrivate,
  docType,
  dataContainer,
  extraSettings,
  options
) {
  // Save everything because we'll need it later:
  this.docType = docType
  this.genPrivate = genPrivate
  this.ogPluginsApi = genPrivate.plugs // Generic officegen API for plugins.
  this.msPluginsApi = genPrivate.plugs.type.msoffice // msoffice plugins API.
  this.genobj = genobj
  this.data = []
  // the paragraph data array syntax:
  // - text (string)
  // - options (object)
  //   - color
  //   - back
  //   - highlight
  //   - superscript
  //   - hyperlink
  //   - font_face
  //   - font_face_h
  //   - font_face_east
  //   - font_face_cs
  // - ext_data (object)
  // - line_break (boolean)
  // - horizontal_line (boolean)
  // - bookmark_start (string)
  // - bookmark_end (boolean)
  // - image (?)
  this.extraSettings = extraSettings || {}
  this.options = options || {}

  this.mainPath = genPrivate.features.type.msoffice.main_path // The "folder" name inside the document zip that all the specific resources of this document type are stored.
  this.mainPathFile = genPrivate.features.type.msoffice.main_path_file // The name of the main real xml resource of this document.
  this.relsMain = genPrivate.type.msoffice.rels_main // Main rels file.
  this.relsApp = genPrivate.type.msoffice.rels_app // Main rels file inside the specific document type "folder".
  this.filesList = genPrivate.type.msoffice.files_list // Resources list xml.
  this.srcFilesList = genPrivate.type.msoffice.src_files_list // For storing extra files inside the document zip.

  if (dataContainer) {
    dataContainer.push(this)
  } // Endif.

  return this
}

/**
 * Change the style of this paraphraph.
 *
 * @param {string} style_name The style name.
 */
MakeDocxP.prototype.setStyle = function (style_name) {
  this.options.force_style = style_name
}

/**
 * Insert text inside this paragraph.
 *
 * @param {string} text_msg The text message itself.
 * @param {object} opt ???.
 * @param {object} flag_data ???.
 */
MakeDocxP.prototype.addText = function (text_msg, opt, flag_data) {
  var newP = this
  var objNum = newP.data.length
  var textMsg = text_msg.toString().replace(/(?:\r\n|\r)/g, '\n').split('\n')

  textMsg.forEach(function (value, index) {
    newP.data[objNum] = {
      text: value,
      options: opt || {},
      ext_data: flag_data
    }

    if ((opt || {}).link) {
      var link_rel_id = newP.relsApp.length + 1

      newP.relsApp.push({
        type:
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: opt.link,
        targetMode: 'External'
      })

      newP.data[objNum].link_rel_id = link_rel_id
      // line_break: true
    } // Endif.

    objNum++
    if (index + 1 < textMsg.length) {
      newP.addLineBreak()
      objNum = newP.data.length
    } // Endif.
  })
}

/**
 * Insert a line break inside this paragraph.
 */
MakeDocxP.prototype.addLineBreak = function () {
  var newP = this
  newP.data[newP.data.length] = { line_break: true }
}

/**
 * Insert a horizontal line inside this paragraph.
 */
MakeDocxP.prototype.addHorizontalLine = function () {
  var newP = this
  newP.data[newP.data.length] = { horizontal_line: true }
}

/**
 * Insert a bookmark here.
 * @param {string} anchorName The anchor of this bookmark.
 */
MakeDocxP.prototype.startBookmark = function (anchorName) {
  var newP = this
  newP.data[newP.data.length] = { bookmark_start: anchorName }
}

/**
 * Close the previous placed bookmark.
 */
MakeDocxP.prototype.endBookmark = function () {
  var newP = this
  newP.data[newP.data.length] = { bookmark_end: true }
}

/**
 * Insert an image into the current paragraph.
 *
 * @param {object} image_path The image file to add.
 * @param {object} opt Additional options (cx, cy).
 * @param {object} image_format_type ???.
 */
MakeDocxP.prototype.addImage = function (image_path, opt, image_format_type) {
  var newP = this
  var image_ext =
    (/.*\.(.*?)$/.exec(image_path) && /.*\.(.*?)$/.exec(image_path)[1]) || null
  var image_type =
    typeof image_format_type === 'string'
      ? image_format_type
      : image_ext || 'png'
  var defWidth = 320
  var defHeight = 200

  if (typeof image_path === 'string') {
    var ret_data = fast_image_size(image_path)
    if (ret_data.type === 'unknown') {
      switch (image_type) {
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

  var objNum = newP.data.length
  newP.data[objNum] = { image: image_path, options: opt || {} }

  if (!newP.data[objNum].options.cx && defWidth) {
    newP.data[objNum].options.cx = defWidth
  } // Endif.

  if (!newP.data[objNum].options.cy && defHeight) {
    newP.data[objNum].options.cy = defHeight
  } // Endif.

  var image_id = newP.srcFilesList.indexOf(image_path)
  var image_rel_id = -1

  if (image_id >= 0) {
    for (var j = 0, total_size_j = newP.relsApp.length; j < total_size_j; j++) {
      if (
        newP.relsApp[j].target ===
        'media/image' + (image_id + 1) + '.' + image_type
      ) {
        image_rel_id = j + 1
      } // Endif.
    } // Endif.
  } else {
    image_id = newP.srcFilesList.length
    newP.srcFilesList[image_id] = image_path
    newP.ogPluginsApi.intAddAnyResourceToParse(
      newP.mainPath + '\\media\\image' + (image_id + 1) + '.' + image_type,
      typeof image_path === 'string' ? 'file' : 'stream',
      image_path,
      null,
      false
    )
  } // Endif.

  if (image_rel_id === -1) {
    image_rel_id = newP.relsApp.length + 1

    newP.relsApp.push({
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: 'media/image' + (image_id + 1) + '.' + image_type,
      clear: 'data'
    })
  } // Endif.

  if ((opt || {}).link) {
    var link_rel_id = newP.relsApp.length + 1

    newP.relsApp.push({
      type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: opt.link,
      targetMode: 'External'
    })

    newP.data[objNum].link_rel_id = link_rel_id
  } // Endif.

  newP.data[objNum].image_id = image_id
  newP.data[objNum].rel_id = image_rel_id
}

module.exports = MakeDocxP
