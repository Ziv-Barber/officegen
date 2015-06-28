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
var chartsData = require('../test_files/charts-data.js');
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

  it ("creates a slides with charts", function(done) {

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

  chartsData.forEach(function (chartInfo, chartIdx) {
    it("creates a presentation with charts", function (done) {
      var officegen = require('../');
      var pptx = officegen('pptx');
      pptx.setDocTitle('Sample PPTX Document');
      var slide = pptx.makeNewSlide();
      slide.name = 'OfficeChart slide';
      slide.back = 'ffffff';

      slide.addChart(
          chartInfo,
          function () {

            var FILENAME = "test-ppt-chart" + chartIdx + ".pptx";
            var out = fs.createWriteStream(OUTDIR + FILENAME);
            pptx.generate(out, {
              'finalize': function (written) {
                setTimeout(function () {
                  assert(pptxEquivalent(OUTDIR + FILENAME, TGTDIR + FILENAME, ["ppt/slides/slide1.xml", "ppt/charts/chart" + (chartIdx+1) + ".xml"]))
                  done()
                }, 50); // give OS time to close the file
              },
              'error': onError
            });
          }, onError);
    })
  })
});
