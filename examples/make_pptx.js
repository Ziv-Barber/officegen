var fs = require('fs');
var path = require('path');

var pptx = require('../officegen.js').makegen ( { 'type': 'pptx', 'onend': function ( written ) {
	console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
} } );

pptx.setDocTitle ( 'Sample PPTX Document' );

// Let's create a new slide:
slide = pptx.makeNewSlide ();

slide.name = 'The first slide!';

// Change the background color:
slide.back = '000000';

// Declare the default color to use on this slide:
slide.color = 'ffffff';

// Basic way to add text string:
slide.addText ( 'This is a test' );
slide.addText ( 'Fast position', 0, 20 );
slide.addText ( 'Full line', 0, 40, '100%', 20 );

// Add text box with multi colors and fonts:
slide.addText ( [
	{ text: 'Hello ', options: { font_size: 56 } },
	{ text: 'World!', options: { font_size: 56, font_face: 'Arial', color: 'ffff00' } }
	], { cx: '75%', cy: 66, y: 150 } );
// Please note that you can pass object as the text parameter to addText.

// For a single text just pass a text string to addText:
slide.addText ( 'Office generator', { y: 66, x: 'c', cx: '50%', cy: 60, font_size: 48, color: '0000ff' } );

slide.addText ( 'Boom!!!', { y: 250, x: 10, cx: '70%', font_face: 'Wide Latin', font_size: 54, color: 'cc0000', bold: true, underline: true } );

// 2nd slide:
slide = pptx.makeNewSlide ();

// For every color property (including the back color property) you can pass object instead of the color string:
slide.back = { type: 'solid', color: '004400' };
slide.addText ( 'Office generator', { y: 'c', x: 0, cx: '100%', cy: 66, font_size: 48, align: 'center', color: { type: 'solid', color: '008800' } } );

slide = pptx.makeNewSlide ();

slide.show = false;
slide.addText ( 'Red line', 'ff0000' );
slide.addShape ( 'ellipse', { fill: 'ff0000', line: 'ffff00', y: 50, x: 50 } );
slide.addText ( 'Red box', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', y: 100 } );
slide.addShape ( 'line', { line: '000088', y: 150, x: 150 } );
slide.addShape ( 'line', { line: '000088', y: 170, x: 150, line_head: 'triangle' } );
slide.addShape ( 'line', { line: '000088', y: 190, x: 150, line_tail: 'triangle' } );
slide.addShape ( 'line', { line: '000088', y: 210, x: 150, line_head: 'stealth', line_tail: 'stealth' } );
slide.addShape ( 'line', { line: '008888', y: 210, x: 150, line_head: 'stealth', line_tail: 'stealth', flip_vertical: true } );
slide.addText ( 'Red box', { color: 'ffffff', fill: 'ff0000', line: 'ffff00', y: 350, x: 200, shape: 'roundRect', indentLevel: 1 } );

slide = pptx.makeNewSlide ();

slide.addImage ( path.resolve(__dirname, 'image1.png' ), { y: 'c', x: 'c', cy: 294, cx: 420 } );

slide = pptx.makeNewSlide ();

var out = fs.createWriteStream ( 'out.pptx' );

pptx.generate ( out );

