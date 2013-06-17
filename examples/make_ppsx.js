var fs = require('fs');

var pptx = require('../officegen.js').makegen ( { 'type': 'ppsx', 'onend': function ( written ) {
	console.log ( 'Finish to create a PowerPoint slideshow file.\nTotal bytes created: ' + written + '\n' );
} } );

// You don't really have to call it:
pptx.startNewDoc ();

pptx.setDocTitle ( 'Sample PPTX Document' );

slide = pptx.makeNewSlide ();
slide.name = 'The first slide!';
slide.back = 'ff0000';
slide.addText ( 'Hello World!', { x: 600000, y: 10000, font_size: 56, cx: 10000000 } );
slide.addText ( 'Office generator', { y: 850000, font_size: 48 } );
slide = pptx.makeNewSlide ();
slide.back = { type: 'solid', color: '00ff00' };
slide = pptx.makeNewSlide ();
slide = pptx.makeNewSlide ();

var out = fs.createWriteStream ( 'out.ppsx' );

pptx.generate ( out );

