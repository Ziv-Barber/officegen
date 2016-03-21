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


var OUTDIR = '/tmp/';
var TGTDIR = __dirname + '/../test_files/';


var AdmZip = require('adm-zip');


var pptxEquivalent = function (path1, path2, subdocs) {
  var left = new AdmZip(path1);
  var right = new AdmZip(path2);
  for (var i = 0; i < subdocs.length; i++) {
    //console.log([subdocs[i], left.readAsText(subdocs[i]).length,right.readAsText(subdocs[i]).length])
    if (left.readAsText(subdocs[i]) != right.readAsText(subdocs[i])) return false;
  }
  return true;
}

// Common error method
var onError = function (err) {
  console.log(err);
  assert(false);
  done()
};


describe("PPTX generator", function () {

  it("creates a presentation with properties and text", function (done) {

    var pptx = officegen('pptx');
    pptx.setDocTitle('Sample PPTX Document');

    var slide = pptx.makeNewSlide();

    slide.name = 'The first slide!';

    // Change the background color:
    slide.back = '000000';

    // Declare the default color to use on this slide:
    slide.color = 'ffffff';

    // Basic way to add text string:
    slide.addText('Created using Officegen version ' + officegen.version);
    slide.addText('Fast position', 0, 20);
    slide.addText('Full line', 0, 40, '100%', 20);

    // Add text box with multi colors and fonts:
    slide.addText([
      { text: 'Hello ', options: { font_size: 56 } },
      { text: 'World!', options: { font_size: 56, font_face: 'Arial', color: 'ffff00' } }
    ], { cx: '75%', cy: 66, y: 150 });
    // Please note that you can pass object as the text parameter to addText.

    // For a single text just pass a text string to addText:
    slide.addText('Office generator', { y: 66, x: 'c', cx: '50%', cy: 60, font_size: 48, color: '0000ff' });

    pObj = slide.addText('Boom\nBoom!!!', { y: 100, x: 10, cx: '70%', font_face: 'Wide Latin', font_size: 54, color: 'cc0000', bold: true, underline: true });
    pObj.options.y += 150;

    var FILENAME = "test-ppt1.pptx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    pptx.generate(out, {
      'finalize': function (written) {
        setTimeout(function () {
          console.log("open \""+OUTDIR+FILENAME+"\"");
          assert(pptxEquivalent(OUTDIR + FILENAME, TGTDIR + FILENAME, ["ppt/slides/slide1.xml"]));
          done()
        }, 50); // give OS time to close the file
      }, 'error': onError
    });

  });

  it("creates slides with shapes", function (done) {
    var pptx = officegen('pptx');
    pptx.setDocTitle('Sample PPTX Document');
    pptx.setWidescreen(false);
    slide = pptx.makeNewSlide();

    slide.show = false;
    slide.addText('Red line', 'ff0000');
    slide.addShape(pptx.shapes.OVAL, { fill: { type: 'solid', color: 'ff0000', alpha: 50 }, line: 'ffff00', y: 50, x: 50 });
    slide.addText('Red box 1', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', line_size: 5, y: 100, rotate: 45 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 150, cy: 0, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 150, cy: 100, cx: 0 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 249, x: 150, cy: 0, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 449, cy: 100, cx: 0 });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 150, x: 150, cy: 100, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 150, x: 150, cy: 100, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 170, x: 150, cy: 100, cx: 300, line_head: 'triangle' });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 190, x: 150, cy: 100, cx: 300, line_tail: 'triangle' });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 210, x: 150, cy: 100, cx: 300, line_head: 'stealth', line_tail: 'stealth' });
    pObj = slide.addShape(pptx.shapes.LINE);
    pObj.options.line = '008888';
    pObj.options.y = 210;
    pObj.options.x = 150;
    pObj.options.cy = 100;
    pObj.options.cx = 300;
    pObj.options.line_head = 'stealth';
    pObj.options.line_tail = 'stealth';
    pObj.options.flip_vertical = true;
    slide.addText('Red box 2', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', y: 350, x: 200, shape: pptx.shapes.ROUNDED_RECTANGLE, indentLevel: 1 });

    var FILENAME = "test-ppt2.pptx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    pptx.generate(out, {
      'finalize': function (written) {
        setTimeout(function () {
          assert(pptxEquivalent(OUTDIR + FILENAME, TGTDIR + FILENAME,
              [
                "ppt/slides/slide1.xml",
                "ppt/presentation.xml"
              ]));
          done()
        }, 50); // give OS time to close the file
      }
    });

  });

  it("creates presentation to widescreen", function (done) {
    var pptx = officegen('pptx');
    pptx.setDocTitle('Sample PPTX Document');
    pptx.setWidescreen(true);
    slide = pptx.makeNewSlide();

    slide.show = false;
    slide.addText('Red line', 'ff0000');
    slide.addShape(pptx.shapes.OVAL, { fill: { type: 'solid', color: 'ff0000', alpha: 50 }, line: 'ffff00', y: 50, x: 50 });
    slide.addText('Red box 1', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', line_size: 5, y: 100, rotate: 45 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 150, cy: 0, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 150, cy: 100, cx: 0 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 249, x: 150, cy: 0, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '0000ff', y: 150, x: 449, cy: 100, cx: 0 });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 150, x: 150, cy: 100, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 150, x: 150, cy: 100, cx: 300 });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 170, x: 150, cy: 100, cx: 300, line_head: 'triangle' });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 190, x: 150, cy: 100, cx: 300, line_tail: 'triangle' });
    slide.addShape(pptx.shapes.LINE, { line: '000088', y: 210, x: 150, cy: 100, cx: 300, line_head: 'stealth', line_tail: 'stealth' });
    pObj = slide.addShape(pptx.shapes.LINE);
    pObj.options.line = '008888';
    pObj.options.y = 210;
    pObj.options.x = 150;
    pObj.options.cy = 100;
    pObj.options.cx = 300;
    pObj.options.line_head = 'stealth';
    pObj.options.line_tail = 'stealth';
    pObj.options.flip_vertical = true;
    slide.addText('Red box 2', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', y: 350, x: 200, shape: pptx.shapes.ROUNDED_RECTANGLE, indentLevel: 1 });

    var FILENAME = "test-ppt3.pptx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    pptx.generate(out, {
      'finalize': function (written) {
        setTimeout(function () {
          assert(pptxEquivalent(OUTDIR + FILENAME, TGTDIR + FILENAME,
              [
                "ppt/slides/slide1.xml",
                "ppt/presentation.xml"
              ]));
          done()
        }, 50); // give OS time to close the file
      }
    });
  });

  it("creates slides with images", function (done) {
    var pptx = officegen('pptx');
    pptx.setDocTitle('Sample PPTX Document');
    var IMAGEDIR = __dirname + "/../examples/";
    slide = pptx.makeNewSlide();


    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/image1.png'), { y: 'c', x: 'c' });

    slide = pptx.makeNewSlide();

    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/image2.jpg'), { y: 0, x: 0, cy: '100%', cx: '100%' });

    slide = pptx.makeNewSlide();
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/image3.png'), { y: 'c', x: 'c'});

    slide = pptx.makeNewSlide();

    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/image2.jpg'), { y: 0, x: 0, cy: '100%', cx: '100%' });

    slide = pptx.makeNewSlide();

    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/image2.jpg'), { y: 0, x: 0, cy: '100%', cx: '100%' });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_001.png'), { y: 10, x: 10 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_002.png'), { y: 10, x: 110 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_001.png'), { y: 110, x: 10 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_001.png'), { y: 110, x: 110 });

    slide = pptx.makeNewSlide();

    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/image2.jpg'), { y: 0, x: 0, cy: '100%', cx: '100%' });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_001.png'), { y: 10, x: 10 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_002.png'), 110, 10);
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_003.png'), { y: 10, x: 210 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_004.png'), { y: 110, x: 10 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_001.png'), { y: 110, x: 110 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_003.png'), { y: 110, x: 210 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_002.png'), { y: 210, x: 10 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_004.png'), { y: 210, x: 110 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_004.png'), { y: 210, x: 210 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_004.png'), { y: '310', x: 10 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_002.png'), { y: 310, x: 110 });
    slide.addImage(path.resolve(IMAGEDIR, 'images_for_examples/sword_003.png'), { y: 310, x: 210 });


    var FILENAME = "test-ppt-images.pptx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    pptx.generate(out, {
      'finalize': function (written) {
        setTimeout(function () {
          assert(pptxEquivalent(OUTDIR + FILENAME, TGTDIR + FILENAME,
              ["ppt/slides/slide1.xml",
                "ppt/slides/slide2.xml",
                "ppt/slides/slide3.xml",
                "ppt/slides/slide4.xml",
                "ppt/slides/slide5.xml",
                "ppt/slides/slide6.xml",
                "ppt/media/image1.png",
                "ppt/media/image2.png",
                "ppt/media/image3.png",
                "ppt/media/image4.png",
                "ppt/media/image5.png",
                "ppt/media/image6.png",
                "ppt/media/image7.png"
              ]));
          done()
        }, 50); // give OS time to close the file
      }
    });
  });

  it ("creates a native table", function(done) {

    var pptx = officegen('pptx');
    pptx.setDocTitle('Sample PPTX Document');
    var slide = pptx.makeNewSlide();

    var rows = [];
    for (var i = 0; i < 12; i++) {
      var row = [];
      for (var j = 0; j < 5; j++) {
        row.push("[" + i + "," + j + "]");
      }
      rows.push(row);
    }
    slide.addTable(rows, {});


    var FILENAME = "test-ppt-table-1.pptx";
    var out = fs.createWriteStream(OUTDIR + FILENAME);
    pptx.generate(out, {
      'finalize': function (written) {
        setTimeout(function () {

          assert(pptxEquivalent(OUTDIR + FILENAME, TGTDIR + FILENAME, ["ppt/slides/slide1.xml"]))
          done()
        }, 50); // give OS time to close the file
      },
      'error': onError
    });
  });
});
