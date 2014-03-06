# officegen [![Build Status](https://travis-ci.org/Ziv-Barber/officegen.png?branch=master)](https://travis-ci.org/Ziv-Barber/officegen) [![Dependencies Status](https://gemnasium.com/Ziv-Barber/officegen.png)](https://gemnasium.com/Ziv-Barber/officegen)

This module can generate Office Open XML files (the files been created by Microsoft Office 2007 and later). 
This module is not depend on any framework so you can use it for any kind of node.js application, even not 
web based. Also the output is a stream and not a file. This module should work on any environment that supporting 
node.js 0.10 including Linux, OS-X and Windows and it's not depending on any output tool.

## Announcement: ##

Donations:

I'm accepting tips through [Gittip](<https://www.gittip.com/Ziv-Barber>)

OpenOffice document generation support will be added in the future.
Please refer to the roadmap section for information on what will be added in the next versions.

## Contents: ##

- [Features](#a1)
- [Installation](#a2)
- [Public API](#a3)
- [Examples](#a4)
- [FAQ](#a5)
- [Hackers Wonderland](#a6)
- [Support](#a7)
- [Changelog](#a8)
- [Roadmap](#a9)
- [License](#a10)
- [Credit](#a11)
- [Donations](#a12)

<a name="a1"/>
## Features: ##

- Generating Microsoft PowerPoint document (.pptx file):
  - Create PowerPoint document with one or more slides.
  - Support both PPT and PPS.
  - Add text blocks.
  - Add images.
  - Can declare fonts, alignment, colors and background.
  - You can rotate objects.
  - Support shapes: Ellipse, Rectangle, Line, Arrows, etc.
  - Support hidden slides.
  - Support automatic fields like date, time and current slide number.
- Generating Microsoft Word document (.docx file):
  - Create Word document.
  - You can add one or more paragraphs to the document and you can set the fonts, colors, alignment, etc.
  - You can add images.
- Generating Microsoft Excel document (.xlsx file):
  - Create Excel document with one or more sheets. Supporting cells of type both number and string.

<a name="a2"/>
## Installation: ##

via Git:

```bash
$ git clone git://github.com/Ziv-Barber/officegen.git
```

via npm:

```bash
$ npm install officegen
```

This module is depending on:

- archiver
- setimmediate
- fast-image-size

<a name="a3"/>
## Public API: ##

### Creating the document object: ###

```js
var officegen = require('officegen');
```

There are two ways to use the officegen function:

```js
officegen ( '<type of document to create>' );

officegen ({
	'type': '<type of document to create>'
	// More options here (if needed)
});
```

Generating PowerPoint 2007 object:

```js
var pptx = officegen ( 'pptx' );
```

Generating Word 2007 object:

```js
var docx = officegen ( 'docx' );
```

Generating Excel 2007 object:

```js
var xlsx = officegen ( 'xlsx' );
```

General events of officegen:

- 'finalize' - been called after finishing to create the document.
- 'error' - been called on error.

Event examples:

```js
pptx.on ( 'finalize', function ( written ) {
			console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
		});

pptx.on ( 'error', function ( err ) {
			console.log ( err );
		});
```

Another way to register either 'finalize' or 'error' events:

```js
var pptx = officegen ({
    'type': 'pptx', // or 'xlsx', etc
    'onend': function ( written ) {
        console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
    },
    'onerr': function ( err ) {
        console.log ( err );
    }
});
```

If you are preferring to use callbacks instead of events you can pass your callbacks to the generate method
(see below).

Now you should fill the object with data (we'll see below) and then you should call generate with 
an output stream to create the output Office document.

Example with pptx:

```js
var out = fs.createWriteStream ( 'out.pptx' );

pptx.generate ( out );
```

Passing callbacks to generate:

```js
var out = fs.createWriteStream ( 'out.pptx' );

pptx.generate ( out, {
	'finalize': function ( written ) {
		console.log ( 'Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
	},
	'error': function ( err ) {
		console.log ( err );
	}
});
```

Generating HTTP stream (no file been created):

```js
var http = require("http");
var officegen = require('officegen');

http.createServer ( function ( request, response ) {
	response.writeHead ( 200, {
		"Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
		'Content-disposition': 'attachment; filename=surprise.pptx'
		});

	var pptx = officegen ( 'pptx' );

	pptx.on ( 'finalize', function ( written ) {
			// ...
			});

	pptx.on ( 'error', function ( err ) {
			// ...
			});

	// ... (fill pptx with data)

	pptx.generate ( response );
}).listen ( 3000 );
```

### Put data inside the document object: ###

#### MS-Office document properties (for all document types): ###

The default Author of all the documents been created by officegen is 'officegen'. If you want to put anything else please 
use the 'creator' option when calling the officegen function:

```js
var pptx = officegen ({
    'type': 'pptx', // or 'xlsx', etc
	'creator': '<your project name here>'
});
```

Change the document title (pptx,ppsx,docx):

```js
var pptx = officegen ({
    'type': 'pptx',
	'title': '<title>'
});

// or

pptx.setDocTitle ( '<title>' );
```

For Word only:

```js
var docx = officegen ({
    'type': 'docx',
	'subject': '...',
	'keywords': '...',
	'description': '...'
});

// or

docx.setDocSubject ( '...' );
docx.setDocKeywords ( '...' );
docx.setDescription ( '...' );
```

#### PowerPoint: ####

Creating a new slide:

```js
slide = pptx.makeNewSlide ();
```

The returned object from makeNewSlide representing a single slide. Use it to add objects into this slide. 
You must create at last one slide on your pptx/ppsx document.

Inside each slide you can place objects, for example: text box, shapes, images, etc.

Properties of the slide object itself:

- "name" - name for this slide.
- "back" - the background color.
- "color" - the default font color to use.
- "show" - change this property to false if you want to disable this slide.

The slide object supporting the following methods:

- addText ( text, options )
- addShape ( shape, options )
- addImage ( image, options )

Read only methods:

- getPageNumber - return the ID of this slide.

Common properties that can be added to the options object for all the add based methods:

- x - start horizontal position. Can be either number, percentage or 'c' to center this object (horizontal).
- y - start vertical position. Can be either number, percentage or 'c' to center this object (vertical).
- cx - the horizontal size of this object. Can be either number or percentage of the total horizontal size.
- cy - the vertical size of this object. Can be either number or percentage of the total vertical size.
- color - the font color for text.
- fill - the background color.
- line - border color / line color.
- flip_vertical: true - flip the object vertical.
- shape - see below.

Font properties:

- font_face
- font_size (in points)
- bold: true
- underline: true

Text alignment properties:

- align - can be either 'left' (default), 'right', 'center' or 'justify'.
- indentLevel - indent level (number: 0+, default = 0).

Line/border extra properties (only effecting if the 'line' property exist):

- 'line_size' - line width in pixels.
- 'line_head' - the shape name of the line's head side (either: 'triangle', 'stealth', etc).
- 'line_tail' - the shape name of the line's tail side (either: 'triangle', 'stealth', etc).

The 'shape' property:

Normally every object is a rectangle but you can change that for every object using the shape property, or in case that 
you don't need to write any text inside that object, you can use the addShape method instead of addText. Use the shape 
property only if you want to use a shape other then the default and you also want to add text inside it.

Shapes list:

- 'rect' (default) - rectangle.
- 'ellipse'
- 'roundRect' - round rectangle.
- 'triangle'
- 'line' - draw line.
- 'cloud'
- 'hexagon'
- 'flowChartInputOutput'
- 'wedgeEllipseCallout'
- (much more shapes already supported - I'll update this list later)

Please note that every color property can be either:

- String of the color code. For example: 'ffffff', '000000', '888800', etc.
- Color object:
  - 'type' - The type of the color fill to use. Right now only 'solid' supported.
  - 'color' - String with the color code to use.
  - 'alpha' - transparent level (0-100).

Adding images:

Just pass the image file name as the first parameter to addImage and the 2nd parameter, which is optional, is normal options objects 
and you can use all the common properties ('cx', 'cy', 'y', 'x', etc).

Examples:
  
Changing the background color of a slide:

```js
slide.back = '000088';
```

or:

```js
slide.back = { type: 'solid', color: '008800' };
```

Examples how to put text inside the new slide:

```js
// Change the background color:
slide.back = '000000';

// Declare the default color to use on this slide (default is black):
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

slide.addText ( 'Office generator', {
	y: 66, x: 'c', cx: '50%', cy: 60, font_size: 48, 
	color: '0000ff' } );

slide.addText ( 'Boom!!!', {
	y: 250, x: 10, cx: '70%', 
	font_face: 'Wide Latin', font_size: 54, 
	color: 'cc0000', bold: true, underline: true } );
```

Set the aspect ratio of the presentation using the `setWidescreen` method:

```js
pptx.setWidescreen(true);
```


#### Word: ####

All the text data in Word is saved in paragraphs. To add a new paragraph:

```js
var pObj = docx.createP ();
```

Paragraph options:

```js
pObj.options.align = 'center'; // Also 'right' or 'jestify'.
```

Every list item is also a paragraph so:

```js
var pObj = docx.createListOfDots ();

var pObj = docx.createListOfNumbers ();
```

Now you can fill the paragraph object with one or more text strings using the addText method:

```js
pObj.addText ( 'Simple' );

pObj.addText ( ' with color', { color: '000088' } );

pObj.addText ( ' and back color.', { color: '00ffff', back: '000088' } );

pObj.addText ( 'Bold + underline', { bold: true, underline: true } );

pObj.addText ( 'Fonts face only.', { font_face: 'Arial' } );

pObj.addText ( ' Fonts face and size.', { font_face: 'Arial', font_size: 40 } );
```

Add an image to a paragraph:

var path = require('path');

pObj.addImage ( path.resolve(__dirname, 'myFile.png' ) );
pObj.addImage ( path.resolve(__dirname, 'myFile.png', { cx: 300, cy: 200 } ) );

To add a line break;

```js
var pObj = docx.createP ();
pObj.addLineBreak ();
```

To add a page break:

```js
docx.putPageBreak ();
```

#### Excel: ####

```js
sheet = xlsx.makeNewSheet ();
sheet.name = 'My Excel Data';
```

Fill cells:

```js
// Using setCell:
sheet.setCell ( 'E7', 340 );
sheet.setCell ( 'G102', 'Hello World!' );

// Direct way:
sheet.data[0] = [];
sheet.data[0][0] = 1;
sheet.data[0][1] = 2;
sheet.data[1] = [];
sheet.data[1][3] = 'abc';
```

<a name="a4"/>
## Examples: ##

- examples/make_pptx.js - Example how to create PowerPoint 2007 presentation and save it into file.
- examples/make_xlsx.js - Example how to create Excel 2007 sheet and save it into file.
- examples/make_docx.js - Example how to create Word 2007 document and save it into file.
- examples/pptx_server.js - Example HTTP server that generating a PowerPoint file with your name without using files on the server side.

<a name="a5"/>
## Hackers Wonderland: ##

This section on the readme file will describe how to hack into the code. 
Right now please refer to the code itself. More information will be added later.

<a name="a6"/>
## FAQ: ##

- Q: Do you support also PPSX files?
- A: Yes! Just pass the type 'ppsx' to makegen instead of 'pptx'.

<a name="a7"/>
## Support: ##

Please visit the officegen Google Group:

https://groups.google.com/forum/?fromgroups#!forum/node-officegen

<a name="a8"/>
## Changelog: ##

[Changelog](https://github.com/Ziv-Barber/officegen/blob/master/CHANGELOG)

<a name="a9"/>
## Roadmap: ##

Features todo:

### Version 0.2.x: ###

- Excel basic styling.
- Word tables.
- PowerPoint lists and tables.

### Version 0.3.x: ###

- Better interface: (officegen will be a steam).
- Embedded document inside another document.

### Version 0.9.x: ###

- Unit tests and lots of testing.

### Version 1.0.x: ###

- Stable release with stable API.

<a name="a10"/>
## License: ##

(The MIT License)

Copyright (c) 2013 Ziv Barber;

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
'Software'), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

<a name="a11"/>
## Credit: ##

- For creating zip streams i'm using 'archiver' by cmilhench, dbrockman, paulj originally inspired by Antoine van Wel's zipstream.

<a name="a12"/>
## Donations: ##

I'm accepting tips through [Gittip](<https://www.gittip.com/Ziv-Barber>)

