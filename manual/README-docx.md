# Create Microsoft Office Word Document Reference

## Contents: ##

- [Creating the document object](#basic)
- [The document object's settings](#settings)
- [The paragraph API](#prgapi)

<a name="basic"></a>
## Creating the document object: ##

First, if you didn't have it yet, get access to the officegen module:

```js
const officegen = require('officegen')
```

Now you have few ways to use it to create a docx based document. The simple way is to use this code:

```js
let docx = officegen('docx')
```

But if you want to pass some settings then you should use the following format:

```js
let docx = officegen({
	type: 'docx', // We want to create a Microsoft Word document.
	... // Extra options goes here.
})
```

<a name="settings"></a>
### The document object's settings: ###

- author (string) - The document's author (part of the Document's Properties in Office).
- creator (string) - Alias. The document's author (part of the Document's Properties in Office).
- description (string) - The document's properties comments (part of the Document's Properties in Office).
- keywords (string) - The document's keywords (part of the Document's Properties in Office).
- orientation (string) - Either 'landscape' or 'portrait'. The default is 'portrait'.
- pageMargins (object) - Set document page margins. The default is { top: 1800, right: 1440, bottom: 1800, left: 1440 }
- subject (string) - The document's subject (part of the Document's Properties in Office).
- title (string) - The document's title (part of the Document's Properties in Office).

You can always change some of these settings after creating the docx object using there methods:

```js
docx.setDocTitle('...')
docx.setDocSubject('...')
docx.setDocKeywords('...')
docx.setDescription( '...')
docx.setDocCategory('...')
docx.setDocStatus('...')
```

<a name="prgapi"></a>
## The paragraph API: ##

To create a new paragraph in your document you need to create a parahpaph object from your main docx object:

```js
let pObj = docx.createP(options)
```

When the options are:

- align (string) - Can be either 'left' (the default), 'right', 'center' or 'justify'.

### Paragraph's methods: ###

```js
pObj.addText(textString, options)
```

When the options are:

- back (string) - background color code, for example: 'ffffff' (white) or '000000' (black).
	- shdType (string) - Optional pattern code to use: 'clear' (no pattern), 'pct10', 'pct12', 'pct15', 'diagCross', 'diagStripe', 'horzCross', 'horzStripe', 'nil', 'thinDiagCross', 'solid', etc.
	- shdColor (string) - The front color for the pattern (used with shdType).
- bold (boolean) - true to make the text bold.
- border (string) - the border type: 'single', 'dashDotStroked', 'dashed', 'dashSmallGap', 'dotDash', 'dotDotDash', 'dotted', 'double', 'thick', etc.
- color (string) - color code, for example: 'ffffff' (white) or '000000' (black).
- italic (boolean) - true to make the text italic.
- underline (boolean) - true to add underline.
- font_face (string) - the font to use.
- font_size (number) - the font size in points.
- highlight (string) - highlight color. Either 'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white' or 'yellow'.
- strikethrough (boolean) - true to add strikethrough.

All the text data in Word is saved in paragraphs. To add a new paragraph:

```javascript
var pObj = docx.createP ();
```

Paragraph options:

```javascript
pObj.options.align = 'center'; // Also 'right' or 'justify'.
pObj.options.indentLeft = 1440; // Indent left 1 inch
```

Every list item is also a paragraph so:

```javascript
var pObj = docx.createListOfDots ();

var pObj = docx.createListOfNumbers ();
```

Now you can fill the paragraph object with one or more text strings using the addText method:

```javascript
pObj.addText ( 'Simple' );

pObj.addText ( ' with color', { color: '000088' } );

pObj.addText ( ' and back color.', { color: '00ffff', back: '000088' } );

pObj.addText ( 'Bold + underline', { bold: true, underline: true } );

pObj.addText ( 'Fonts face only.', { font_face: 'Arial' } );

pObj.addText ( ' Fonts face and size. ', { font_face: 'Arial', font_size: 40 } );

pObj.addText ( 'External link', { link: 'https://github.com' } );

// Hyperlinks to bookmarks also supported:
pObj.addText ( 'Internal link', { hyperlink: 'myBookmark' } );
// ...
// Start somewhere a bookmark:
pObj.startBookmark ( 'myBookmark' );
// ...
// You MUST close your bookmark:
pObj.endBookmark ();
```

Add an image to a paragraph:

```
var path = require('path');

pObj.addImage ( path.resolve(__dirname, 'myFile.png' ) );
pObj.addImage ( path.resolve(__dirname, 'myFile.png', { cx: 300, cy: 200 } ) );
```

To add a line break;

```javascript
var pObj = docx.createP ();
pObj.addLineBreak ();
```

To add a page break:

```javascript
docx.putPageBreak ();
```

To add a horizontal line:

```javascript
var pObj = docx.createP ();
pObj.addHorizontalLine ();
```

To add a back line:

```javascript
var pObj = docx.createP ({ backline: 'E0E0E0' });
pObj.addText ( 'Backline text1' );
pObj.addText ( ' text2' );
```

To add a table:

```javascript
var table = [
  [{
    val: "No.",
    opts: {
      cellColWidth: 4261,
      b:true,
      sz: '48',
      shd: {
        fill: "7F7F7F",
        themeFill: "text1",
        "themeFillTint": "80"
      },
      fontFamily: "Avenir Book"
    }
  },{
    val: "Title1",
    opts: {
      b:true,
      color: "A00000",
      align: "right",
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  },{
    val: "Title2",
    opts: {
      align: "center",
      vAlign: "center",
      cellColWidth: 42,
      b:true,
      sz: '48',
      shd: {
        fill: "92CDDC",
        themeFill: "text1",
        "themeFillTint": "80"
      }
    }
  }],
  [1,'All grown-ups were once children',''],
  [2,'there is no harm in putting off a piece of work until another day.',''],
  [3,'But when it is a matter of baobabs, that always means a catastrophe.',''],
  [4,'You can include CR-LF inline\r\nfor multiple lines.',''],
  [5,['Or you can provide lines within', 'a cell in an array'],''],
  [6,'But when it is a matter of baobabs, that always means a catastrophe.',''],
  [7,'watch out for the baobabs!','END'],
]

var tableStyle = {
  tableColWidth: 4261,
  tableSize: 24,
  tableColor: "ada",
  tableAlign: "left",
  tableFontFamily: "Comic Sans MS",
  borders: true
}

docx.createTable (table, tableStyle);
```

Header and footer:

```javascript
// Add a header:
var header = docx.getHeader ().createP ();
header.addText ( 'This is the header' );
// Please note that the object header here is a paragraph object so you can use ANY of the paragraph API methods also for header and footer.
// The getHeader () method excepting a string parameter:
// getHeader ( 'even' ) - change the header for even pages.
// getHeader ( 'first' ) - change the header for the first page only.
// to do all of that for the footer, use the getFooter instead of getHeader.
// and sorry, right now only createP is supported (so only creating a paragraph) so no tables, etc.
```

To Create Word Document by json:

```javascript
var table = [
    [{
        val: "No.",
        opts: {
            cellColWidth: 4261,
            b:true,
            sz: '48',
            shd: {
                fill: "7F7F7F",
                themeFill: "text1",
                "themeFillTint": "80"
            },
            fontFamily: "Avenir Book"
        }
    },{
        val: "Title1",
        opts: {
            b:true,
            color: "A00000",
            align: "right",
            shd: {
                fill: "92CDDC",
                themeFill: "text1",
                "themeFillTint": "80"
            }
        }
    },{
        val: "Title2",
        opts: {
            align: "center",
            cellColWidth: 42,
            b:true,
            sz: '48',
            shd: {
                fill: "92CDDC",
                themeFill: "text1",
                "themeFillTint": "80"
            }
        }
    }],
    [1,'All grown-ups were once children',''],
    [2,'there is no harm in putting off a piece of work until another day.',''],
    [3,'But when it is a matter of baobabs, that always means a catastrophe.',''],
    [4,'watch out for the baobabs!','END'],
]

var tableStyle = {
    tableColWidth: 4261,
    tableSize: 24,
    tableColor: "ada",
    tableAlign: "left",
    tableFontFamily: "Comic Sans MS"
}

var data = [[{
        type: "text",
        val: "Simple"
    }, {
        type: "text",
        val: " with color",
        opt: { color: '000088' }
    }, {
        type: "text",
        val: "  and back color.",
        opt: { color: '00ffff', back: '000088' }
    }, {
        type: "linebreak"
    }, {
        type: "text",
        val: "Bold + underline",
        opt: { bold: true, underline: true }
    }], {
        type: "horizontalline"
    }, [{ backline: 'EDEDED' }, {
        type: "text",
        val: "  backline text1.",
        opt: { bold: true }
    }, {
        type: "text",
        val: "  backline text2.",
        opt: { color: '000088' }
    }], {
        type: "text",
        val: "Left this text.",
        lopt: { align: 'left' }
    }, {
        type: "text",
        val: "Center this text.",
        lopt: { align: 'center' }
    }, {
        type: "text",
        val: "Right this text.",
        lopt: { align: 'right' }
    }, {
        type: "text",
        val: "Fonts face only.",
        opt: { font_face: 'Arial' }
    }, {
        type: "text",
        val: "Fonts face and size.",
        opt: { font_face: 'Arial', font_size: 40 }
    }, {
        type: "table",
        val: table,
        opt: tableStyle
    }, [{ // arr[0] is common option.
        align: 'right'
    }, {
        type: "image",
        path: path.resolve(__dirname, 'images_for_examples/sword_001.png')
    },{
        type: "image",
        path: path.resolve(__dirname, 'images_for_examples/sword_002.png')
    }], {
        type: "pagebreak"
    }
]

docx.createByJson(data);
```
