//======================================================================================================================
// TEST SUITE FOR OFFICEGEN
// This generates small individual files that test specific aspects of the API
// and compares them to reference files.
//
// The comparison is based on string comparisons of specified XML subdocuments.
// Comparing PPTX files for exact-bytewise equality fails because the doc properties include the creation date.
// This method tests a defined set of XML subdocuments for string equality.
//======================================================================================================================

var assert = require('assert');
var officegen = require('../');
var fs = require('fs');
var path = require('path');

var IMAGEDIR = __dirname + "/../examples/";
var OUTDIR = __dirname + "/../tmp/";

var AdmZip = require('adm-zip');


var docxEquivalent = function (path1, path2, subdocs) {
  var left = new AdmZip(path1);
  var right = new AdmZip(path2);
  for (var i = 0; i < subdocs.length; i++) {
    if (left.readAsText(subdocs[i]) != right.readAsText(subdocs[i])) {
      return false;
    }
  }
  return true;
}

// Common error method
var onError = function (err) {
  console.log(err);
  assert(false);
  done()
};


describe("DOCX generator", function () {
	this.timeout(1000);

  it("creates a document with text and styles", function (done) {

    var docx = officegen ( 'docx' );
	docx.on ( 'error', onError );

    var pObj = docx.createP ();

    pObj.addText ( 'Simple' );
    pObj.addText ( ' with color', { color: '000088' } );
    pObj.addText ( ' and back color.', { color: '00ffff', back: '000088' } );

    var pObj = docx.createP ();

    pObj.addText ( 'Bold + underline', { bold: true, underline: true } );

    var pObj = docx.createP ( { align: 'center' } );

    pObj.addText ( 'Center this text.' );

    var pObj = docx.createP ( { align: 'left', first_line: 600 } );

    pObj.addText ( 'First line support' );

    var pObj = docx.createP ();
    pObj.options.align = 'right';

    pObj.addText ( 'Align this text to the right.' );

    var pObj = docx.createP ();

    pObj.addText ( 'Those two lines are in the same paragraph,' );
    pObj.addLineBreak ();
    pObj.addText ( 'but they are separated by a line break.' );

    docx.putPageBreak ();

    var pObj = docx.createP ();

    pObj.addText ( 'Fonts face only.', { font_face: 'Arial' } );
    pObj.addText ( ' Fonts face and size.', { font_face: 'Arial', font_size: 40 } );

    docx.putPageBreak ();

    var pObj = docx.createListOfNumbers ();

    pObj.addText ( 'Option 1' );

    var pObj = docx.createListOfNumbers ();

    pObj.addText ( 'Option 2' );

    var FILENAME = "test-doc-1.docx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    docx.generate(out);
	out.on ( 'close', function () {
		done ();
	});
  });

  it("can handle text without spaces", function (done) {

    var docx = officegen ( 'docx' );
	docx.on ( 'error', onError );

    var pObj = docx.createP();
    pObj.addText('Hello,World');

    var FILENAME = "test-doc-3.docx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    docx.generate(out);
	out.on ( 'close', function () {
		done ();
	});
  });



  it("creates a document with images", function (done) {

    var docx = officegen ( 'docx' );
	docx.on ( 'error', onError );

    var pObj = docx.createP ();


    var pObj = docx.createP ();

    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/image3.png' ) );

    docx.putPageBreak ();

    var pObj = docx.createP ();

    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/image1.png' ) );

    var pObj = docx.createP ();

    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/sword_001.png' ) );
    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/sword_002.png' ) );
    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/sword_003.png' ) );
    pObj.addText ( '... some text here ...', { font_face: 'Arial' } );
    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/sword_004.png' ) );

    var pObj = docx.createP ();

    pObj.addImage ( path.resolve(IMAGEDIR, 'images_for_examples/image1.png' ) );

    var FILENAME = "test-doc-2.docx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
	out.on ( 'error', onError );

    docx.generate(out);
	out.on ( 'close', function () {
		done ();
	});
  });
});
