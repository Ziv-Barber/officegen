var fs = require('fs');

var pptx = require('../officegen.js').makegen ( { 'type': 'pptx', 'onend': function ( written ) {
	console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
} } );

// You don't really have to call it:
pptx.startNewDoc ();

pptx.setDocTitle ( 'Sample PPTX Document' );

slide = pptx.makeNewSlide ();
slide.name = 'Ziv!';
slide.addText ( 'Ziv Barber', { x: 600000, y: 10000, font_size: 56, cx: 10000000 } );
slide.addText ( '222', { y: 850000, font_size: 48 } );
slide = pptx.makeNewSlide ();
slide = pptx.makeNewSlide ();
slide = pptx.makeNewSlide ();

var out = fs.createWriteStream ( 'out.pptx' );

pptx.generate ( out );

