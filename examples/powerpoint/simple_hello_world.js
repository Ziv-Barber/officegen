var officegen = require('../../')

var async = require('async')
var fs = require('fs')
var path = require('path')

/**
 * Example how to create a simple PowerPoint file using officegen and node.js fs API.
 * We'll just write 'Hello World' into the generated file.
 */
function makePptxFile() {
  var pptx = officegen('pptx')

  // Example how to set the title (You can see it in the document properties):
  pptx.setDocTitle('Sample PPTX Document')

  // Let's create a new slide:
  var slide = pptx.makeNewSlide()

  slide.name = 'Hello World'

  // Change the background color:
  slide.back = '000000'

  // Declare the default color to use on this slide:
  slide.color = 'ffffff'

  // Basic way to add text string:
  slide.addText('Hello World!!!')

  // Create a file stream so we'll output the generated pptx data into this file:
  var out = fs.createWriteStream(
    path.join(__dirname, '../../tmp/simple_hello_world.pptx')
  )

  //
  // Generating part - let's do it into a file:
  //

  // This one catch only the officegen errors:
  pptx.on('error', function (err) {
    console.log(err)
  })

  // Catch fs errors:
  out.on('error', function (err) {
    console.log(err)
  })

  // End event after creating the PowerPoint file:
  out.on('close', function () {
    console.log('Finished to create the PowerPoint file')
  })

  // This method is working like a pipe - it'll generate the pptx data and put it into the output stream:
  pptx.generate(out)
}

async.series([makePptxFile])
