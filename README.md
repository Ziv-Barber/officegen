# officegen #

This module can generate Office Open XML files (the files been created by Microsoft Office 2007 and later). 
This module is not depend on any framework so you can use it for any kind of node.js application, even not 
web based. Also the output is a stream and not a file. This module should work on any environment that supporting 
node.js including Linux, OS-X and Windows and it's not depending on any output tool.

This module is a Javascript porting of my 'DuckWriteC++' library which doing the same in C++.

## Announcement: ##

This version only implementing basic features and there is no plugins API yet. You can fork this code if you 
want to but please beware that I'm in the middle of huge changing in the design of this module and it'll be 
better to wait for more stable releases if you want to improve it.

## Contents: ##

- [Features](#a1)
- [Installation](#a2)
- [Public API](#a3)
- [Examples](#a4)
- [Hackers Wonderland] (#a5)
- [Support] (#a6)
- [License](#a7)

<a name="a1"/>
## Features: ##

- Generating Microsoft PowerPoint document (.pptx file):
  - Create PowerPoint document with one or more slides.
  - Add text objects to each slide.
  - Can declare fonts, colors and background.
- Generating Microsoft Excel document (.xlsx file):
  - Create Excel document with one or more sheets (still needs some work).
- Generating Microsoft Word document (.docx file):
  - Not yet there.

<a name="a2"/>
## Installation: ##

via npm:

```bash
$ npm install officegen
```

This module is depending on:

- archiver
- setimmediate

<a name="a3"/>
## Public API: ##

Generating PowerPoint 2007 object:

```js
var pptx = require('../officegen.js').makegen ( { 'type': 'pptx' } );
```

Generating Word 2007 object:

```js
var docx = require('../officegen.js').makegen ( { 'type': 'docx' } );
```

Generating Excel 2007 object:

```js
var xlsx = require('../officegen.js').makegen ( { 'type': 'xlsx' } );
```

Now you should fill the object with data (we'll see below) and then you should call generate with 
an output stream to create the output Office document.

Example with pptx:

```js
var out = fs.createWriteStream ( 'out.pptx' );

pptx.generate ( out );
```

Generating HTTP stream (no file been created):

```js
var http = require("http");

http.createServer ( function ( request, response ) {
	var pptx = require('../officegen.js').makegen ( { 'type': 'pptx', 'onend': function ( written ) {
		// ... (called after finishing to serve the user)
	} } );

	// ... (fill pptx with data)

	pptx.generate ( response );
}).listen ( 3000 );
```

Creating new slides for pptx:

```js
slide = pptx.makeNewSlide ();
slide.back = '000088';
```

Changing the background color of a slide:

```js
slide.back = '000088';
```

or:

```js
slide.back = { type: 'solid', color: '008800' };
```

Example to put text inside the new slide:

```js
slide.addText ( 'Hello World!!!', { x: 600000, y: 10000, font_size: 56, cx: 10000000 } );
```

<a name="a4"/>
## Examples: ##

- examples/make_pptx.js - Example how to create a PowerPoint 2007 presentation and save it into file.
- examples/pptx_server.js - Example HTTP server that generating a PowerPoint file with your name without using files on the server side.
- examples/make_xlsx.js - Example how to create a Excel 2007 sheet and save it into file.

<a name="a5"/>
## Hackers Wonderland: ##

This section on the readme file will describe how to hack into the code. 
Right now please refer to the code itself. More information will be added later.

<a name="a6"/>
## Support: ##

Please visit the officegen Google Group:

https://groups.google.com/forum/?fromgroups#!forum/node-officegen

<a name="a7"/>
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

