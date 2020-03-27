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
- pageSize (string | object) - Set document page size. The default is A4 (support value: 'A4', 'A3', 'letter paper'). Or set customize size with { width: 11906, height: 16838 }
- subject (string) - The document's subject (part of the Document's Properties in Office).
- title (string) - The document's title (part of the Document's Properties in Office).
- columns (number) - The number of columns in each page. The default is 1 column.

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

To create a new paragraph in your document you need to create a paragraph object from your main docx object:

```js
let pObj = docx.createP(options)
```

When the options are:

- align (string) - Horizontal alignment, can be either 'left' (the default), 'right', 'center' or 'justify'.
- textAlignment (string) - Vertical alignment, can be 'center', 'top', 'bottom' or 'baseline'.

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
- font_face_east (string) - advanced setting: the font to use for east asian. You must set also font_face.
- font_face_cs (string) - advanced setting: the font to use (cs). You must set also font_face.
- font_face_h (string) - advanced setting: the font to use (hAnsi). You must set also font_face.
- font_hint (string) - optional. Either 'ascii' (the default), 'eastAsia', 'cs' or 'hAnsi'.
- font_size (number) - the font size in points.
- rtl (boolean) - add this to any text in rtl language.
- highlight (string) - highlight color. Either 'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white' or 'yellow'.
- strikethrough (boolean) - true to add strikethrough.
- superscript (boolean) - true to lower the text in this run below the baseline and change it to a smaller size, if a smallersize is available. Supported in officegen 0.5.0 and later.
- subscript (boolean) - true to raise the text in this run above the baseline and change it to a smaller size, if a smaller size is available. Supported in officegen 0.5.0 and later.

All the text data in Word is saved in paragraphs. To add a new paragraph:

```javascript
var pObj = docx.createP ();
```

Paragraph options:

```javascript
pObj.options.align = 'center'; // Also 'right' or 'justify'.
pObj.options.indentLeft = 1440; // Indent left 1 inch
pObj.options.indentFirstLine = 440; // Indent first line
```

Every list item is also a paragraph so:

```javascript
var pObj = docx.createListOfDots ();

var pObj = docx.createListOfNumbers ();
```

You can create ordered and ordered nested lists as below. You can pass the list level as input to the function.

```javascript
pObj = docx.createNestedUnOrderedList({
  "level":2
})

pObj = docx.createNestedOrderedList({
  "level":2
})
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

pObj.addImage(path.resolve(__dirname, 'myFile.png'))
pObj.addImage(path.resolve(__dirname, 'myFile.png'), {cx: 300, cy: 200})
```

To add a line break;

```javascript
var pObj = docx.createP()
pObj.addLineBreak()
```

To add a page break:

```javascript
docx.putPageBreak()
```

To add a horizontal line:

```javascript
var pObj = docx.createP()
pObj.addHorizontalLine()
```

To add a back line:

```javascript
var pObj = docx.createP({backline: 'E0E0E0'})
pObj.addText('Backline text1')
pObj.addText(' text2')
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
      spacingBefore: 120,
      spacingAfter: 120,
      spacingLine: 240,
      spacingLineRule: 'atLeast',
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
  spacingBefor: 120, // default is 100
  spacingAfter: 120, // default is 100
  spacingLine: 240, // default is 240
  spacingLineRule: 'atLeast', // default is atLeast
  indent: 100, // table indent, default is 0
  fixedLayout: true, // default is false
  borders: true, // default is false. if true, default border size is 4
  borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
  columns: [{ width: 4261 }, { width: 1 }, { width: 42 }], // Table logical columns
}
docx.createTable (table, tableStyle);
```

If you want to customize the border style, you can use the 'borderStyle' option:
```javascript
  const style = {
    '@w:val': 'single',
    '@w:sz': '3',
    '@w:space': '1',
    '@w:color': 'DF0000'
  }
  const borderStyle = {
    'w:top': style,
    'w:bottom': style,
    'w:left': style,
    'w:right': style,
    'w:insideH': style,
    'w:insideV': style,
  }
  const tableStyle = {
    tableColWidth: 4261,
    tableSize: 24,
    tableColor: 'ada',
    tableAlign: 'left',
    tableFontFamily: 'Comic Sans MS',
    borderStyle: borderStyle
  }
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
