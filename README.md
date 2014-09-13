# officegen-2 #

- by vtloc -

This module were built up-on the original module officegen which weren't published by me. 
In this module, I've support the feature of exporting chart ( pie, bar, column ). 
But the code is a bit hacky, that's why I've published this module separately. 
Used it for your own risks.

- end by vtloc -

This module can generate Office Open XML files (the files been created by Microsoft Office 2007 and later). 
This module is not depend on any framework so you can use it for any kind of node.js application, even not 
web based. Also the output is a stream and not a file. This module should work on any environment that supporting 
node.js including Linux, OS-X and Windows and it's not depending on any output tool.

This module is a Javascript porting of my 'DuckWriteC++' library which doing the same in C++.

## Announcement: ##

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
- [History](#a8)
- [Roadmap](#a9)
- [License](#a10)
- [Credit](#a11)

<a name="a1"/>
## Features: ##

- Generating Microsoft PowerPoint document (.pptx file):
  - Create PowerPoint document with one or more slides.
  - Support both PPT and PPS.
  - Add text blocks.
  - Add images.
  - Can declare fonts, alignment, colors and background.
  - Support shapes: Ellipse, Rectangle, Line, Arrows, etc.
  - Support hidden slides.
  - Support
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
$ git clone git://github.com/vtloc/officegen.git
```

via npm:

```bash
$ npm install officegen-2
```

This module is depending on:

- archiver
- setimmediate
- fast-image-size
- underscore
- xmlbuilder

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
- addPieChart ( data )
- addColumnChart ( data )
- addBarChart ( data )

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
- font_size
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
Examples how to add chart into the slide:
```js
// Column chart
slide = pptx.makeNewSlide();
slide.name = 'Chart slide';
slide.back = 'ffffff';
slide.addColumnChart(
	{ 	title: 'Column chart',
		data:  [ // each item is one serie
		{
			name: 'Income',
			labels: ['2005', '2006', '2007', '2008', '2009'],
			values: [23.5, 26.2, 30.1, 29.5, 24.6],
			color: 'ff0000' // optional
		},
		{
			name: 'Expense',
			labels: ['2005', '2006', '2007', '2008', '2009'],
			values: [18.1, 22.8, 23.9, 25.1, 25],
			color: '00ff00' // optional
		}]
	}
)

// Pie chart
slide = pptx.makeNewSlide();
slide.name = 'Pie Chart slide';
slide.back = 'ffff00';
slide.addPieChart(
	{ 	title: 'My production',
		data:  [ // each item is one serie
		{
			name: 'Oil',
			labels: ['Czech Republic', 'Ireland', 'Germany', 'Australia', 'Austria', 'UK', 'Belgium'],
			values: [301, 201, 165, 139, 128,  99, 60],
			colors: ['ff0000', '00ff00', '0000ff', 'ffff00', 'ff00ff', '00ffff', '000000'] // optional
		}]
	}
)

// Bar Chart
slide = pptx.makeNewSlide();
slide.name = 'Bar Chart slide';
slide.back = 'ff00ff';
slide.addBarChart(
	{ title: 'Sample bar chart',
		data:  [ // each item is one serie
      {
        name: 'europe',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.6, 2.8],
        color: 'ff0000' // optional
      },
      {
        name: 'namerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.5, 2.7, 2.9],
        color: '00ff00' // optional
      },
      {
        name: 'asia',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [2.1, 2.2, 2.4],
        color: '0000ff' // optional
      },
      {
        name: 'lamerica',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.3, 0.3, 0.3],
        color: 'ffff00' // optional
      },
      {
        name: 'meast',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.2, 0.3, 0.3],
        color: 'ff00ff' // optional
      },
      {
        name: 'africa',
        labels: ['Y2003', 'Y2004', 'Y2005'],
        values: [0.1, 0.1, 0.1],
        color: '00ffff' // optional
      }
    
    ]
	}
)
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
## History: ##
- Version 0.3.*:
  - PowerPoint:
    - Add pie chart
    - Add bar chart
    - Add column chart
- Version 0.2.6:
	- PowerPoint:
		- Automatically support line breaks.
		- Fixed a bug when using effects (shadows).
	- Excell:
		- Patch by arnesten: Automatically support line breaks if used in cell and also set appropriate row height depending on the number of line breaks.
- Version 0.2.5:
	- Internal design changes that should not effect current implementations. To support future features.
	- Bugs:
		- Small typo which makes it crash. oobjOptions should be objOptions on line 464 in genpptx.js (thanks Stefan Van Dyck!).
- Version 0.2.4:
	- PowerPoint:
		- Body properties like autoFit and margin now supported for text objects (thanks Stefan Van Dyck!).
		- You can pass now 0 to either cx or cy (good when drawing either horizontal or vertical lines).
	- Plugins developers:
		- You can now generate also tar and gzip based documents (or archive files).
		- You can generate your document resources using template engines (like jade, ejs, haml*, CoffeeKup, etc).
- Version 0.2.3:
	- PowerPoint:
		- You can now either read or change the options of a parahraph object after creating it.
		- You can add shadow effects (both outher and inner).
- Version 0.2.2:
	- Word:
		- You can now put images inside your document.
	- General features:
		- You can now pass callbacks to generate() instead of using node events.
	- Bugs / Optimization:
		- If you add the same image only one copy of it will be saved.
		- Missing requirement after the split of the code in version 0.2.x (thanks Seth Pollack!)
		- Fix the bug when you put number as a string for properties like y, x, cy and cx.
		- Generating invalid strings for MS-Office document properties.
		- Better shared string support in Excel (thanks vivocha!).
- Version 0.2.0:
	- Huge design change from 'quick patch' based code to real design with much better API while still supporting also 
	  the old API.
	- Bugs:
		- You can now listen on error events.
		- Missing files in the relationships list made the Excel files unreadable to the Numbers application on the Mac (lmalheiro).
		- Minor bug fixes on the examples and the documentation.
- Version 0.1.11:
	- PowerPoint:
		- Transparent level for solid color.
		- Rotate any object.
		- Flip vertical now working for any kind of object.
		- Line width.
	- Bugs:
		- Invalid PPTX file when adding more then one image of the same type.
- Version 0.1.10:
	- PowerPoint:
		- Supporting more image types.
		- Supporting hidden slides.
		- Allow changing the indent for text.
	- Bug: All the text messages for all type of documents can now have '<', '>', etc.
- Version 0.1.9:
	- Bug: Fix the invalid package.json main file.
	- PowerPoint: Allow adding shapes.
- Version 0.1.8:
	- PowerPoint: Allow adding images (png only).
- Version 0.1.7:
	- Excel 2007: addCell.
	- Many internal changes that are not effecting the user API.
- Version 0.1.6:
	- Excel 2007: finished supporting shared strings.
	- Excel 2007: The interface been changed.
	- A lot of changes in the design of this module.
- Version 0.1.5:
	- Word 2007 basic API now working.
- Version 0.1.4:
	- WARNING: The addText function for PowerPoint been changed since version 0.1.3.
	- Many new features for PowerPoint.
	- Minor bug fixes.
- Version 0.1.3:
	- Can generate also ppsx files.
	- Minor bug fixes.
- Version 0.1.2:
	- HTTP server demo.
	- Can generate very limited Excel file.
	- You can change the background color of slides.
	- Minor bug fixes.

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

