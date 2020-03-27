var async = require('async')
var officegen = require('../../') // In your code replace it with 'officegen'.

var fs = require('fs')
var path = require('path')

// We'll put the generated document in the tmp folder under the root folder of officegen:
var outDir = path.join(__dirname, '../../tmp/')

// Create the document object:
var docx = officegen({
  type: 'docx',
  orientation: 'portrait',
  pageMargins: { top: 1000, left: 1000, bottom: 1000, right: 1000 }
})

// Remove this comment if you want to debug Officegen:
// officegen.setVerboseMode ( true )

docx.on('error', function (err) {
  console.log(err)
})

var pObj = docx.createP()
pObj.setStyle('Heading1')
pObj.addText('Header!')

var out = fs.createWriteStream(path.join(outDir, 'example_word.docx'))

out.on('error', function (err) {
  console.log(err)
})

async.parallel(
  [
    function (done) {
      out.on('close', function () {
        console.log('Finish to create a DOCX file.')
        done(null)
      })
      docx.generate(out)
    }
  ],
  function (err) {
    if (err) {
      console.log('error: ' + err)
    } // Endif.
  }
)
