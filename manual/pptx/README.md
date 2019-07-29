# Create Microsoft Office PowerPoint Document Reference

## Contents: ##

- [Creating the document object](#basic)
- [The document object's settings](#settings)

<a name="basic"></a>
## Creating the document object: ##

First, if you didn't have it yet, get access to the officegen module:

```js
const officegen = require('officegen')
```

Now you have few ways to use it to create a pptx based document. The simple way is to use this code:

```js
let pptx = officegen('pptx')
```

But if you want to pass some settings then you should use the following format:

```js
let pptx = officegen({
	type: 'pptx', // We want to create a Microsoft Powerpoint document.
	... // Extra options goes here.
})
```

To change the theme:

```js
let pptx = officegen({
	type: 'pptx', // We want to create a Microsoft Powerpoint document.
	themeXml: '... theme xml goes here ...' // Just copy the theme xml code from existing office document (stored in ppt\theme\theme1.xml).
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

You can always change some of these settings after creating the pptx object using there methods:

```js
pptx.setDocTitle('...')
pptx.setDocSubject('...')
pptx.setDocKeywords('...')
pptx.setDescription( '...')
pptx.setDocCategory('...')
pptx.setDocStatus('...')
```

Setting slide size:
```
pptx.setSlideSize(cx, cy, type, cxSLD, cySLD)
```

Arguments:
- cx - width of the slide and notes (in pixels)
- cy - height of the slide and notes (in pixels)
- cxSLD - width of the slide (in pixels) only if it's not the same as the notes size
- cySLD - height of the slide (in pixels) only if it's not the same as the notes size
- Supported types:
  - '35mm'
  - 'A3'
  - 'A4'
  - 'B4ISO'
  - 'B4JIS'
  - 'B5ISO'
  - 'B5JIS'
  - 'banner'
  - 'custom'
  - 'hagakiCard'
  - 'ledger'
  - 'letter'
  - 'overhead'
  - 'screen16x10'
  - 'screen16x9'
  - 'screen4x3'

Notes:
  - cx, cy, cxSLD and cySLD are optional and you can pass 0 if you just want to use one of the standard sizes (anything except for 'custom').
  - If you are taking the values for cx, cy, cxSLD and cySLD from existing document then do: sourceValue / 12700
  - cxSLD and cySLD only working in version 0.5.0 or later of officegen.
  - if you didn't set cxSLD and cySLD then officegen will use the values of cx and cy also for cxSLD and cySLD.

Changing the displayed view when opening the document in PowerPoint:

```javascript
pptx.view.restoredLeft = 15620 // This is the default value.
pptx.view.restoredTop = 94660 // This is the default value. Set it to lower value if you want to see more of the speaker notes.
```

Creating a new slide:

```javascript
let slide = pptx.makeNewSlide();
```

Using the Microsoft Office built-in layouts:

```js
let slide = pptx.makeTitleSlide(title, subTitle)
let slide = pptx.makeObjSlide(title, objData)
let slide = pptx.makeSecHeadSlide(title, subTitle)
```

For creating a new slide using a layout (alternative):

```javascript
let slide = pptx.makeNewSlide({
  userLayout: 'title'
});
slide.setTitle('The title');
slide.setSubTitle('Another text'); // For either 'title' and 'secHead' only.
// for 'obj' layout use slide.setObjData(...) to change the object element inside the slide.
```

userLayout can be:

- 'title': the first layout of Office (title).
- 'obj': the 2nd layout of Office (with one title and one object).
- 'secHead': the 3rd layout of Office.

Or more advance example:

```javascript
slide = pptx.makeNewSlide({
  userLayout: 'title'
});

// Both setTitle and setSubTitle excepting all the parameters that you can pass to slide.addText - see below:
slide.setTitle([
  // This array is like a paragraph and you can use any settings that you pass for creating a paragraph,
  // Each object here is like a call to addText:
  {
    text: 'Hello ',
    options: {font_size: 56}
  },
  {
    text: 'World!',
    options: {
      font_size: 56,
      font_face: 'Arial',
      color: 'ffff00'
    }
  }
]);
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
- addChart ( chartInfo )
- addTable ( rowsSpec, options )

Read only methods:

- getPageNumber - return the ID of this slide.

Common properties that can be added to the options object for all the add based methods:

- x - start horizontal position. Can be either number, percentage or 'c' to center this object (horizontal).
- y - start vertical position. Can be either number, percentage or 'c' to center this object (vertical).
- cx - the horizontal size of this object. Can be either number or percentage of the total horizontal size.
- cy - the vertical size of this object. Can be either number or percentage of the total vertical size.
- id - optional custom ID (must be unique number).
- name - optional custom name.
- title - optional custom title.
- desc - optional custom description.
- hidden - optional boolean flag. If true then this shape will be hidden.
- color - the font color for text.
- fill - the background color.
- line - border color / line color.
- flip_vertical: true - flip the object vertical.
- flip_horizontal: true - flip the object horizontal
- shape - see below.

Font properties:

- font_face
- font_size (in points)
- bold: true
- underline: true
- char_spacing: floating point number (kerning)
- baseline: percent (of font size). Used for superscript and subscript.

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
  - 'type' - The type of the color fill to use. Right now only 'solid' and 'gradient' supported.
  - 'color' - String with the color code to use.
  - 'alpha' - transparent level (0-100).
- For 'gradient' fill:
  - 'color' - Array of strings with the color code to use OR array of object, each object include "color" and "position" parameters. i.e. `[{"position": 0, "color": '000000'}, {}, ...]`
  - 'alpha' - transparent level (0-100).
  - 'angle' - (optional) the angle of gradient rotation

Adding images:

Just pass the image file name as the first parameter to addImage and the 2nd parameter, which is optional, is normal options objects
and you can use all the common properties ('cx', 'cy', 'y', 'x', etc).

Examples:

Changing the background color of a slide:

```javascript
slide.back = '000088';
```

or:

```javascript
slide.back = {type: 'solid', color: '008800'};
```

Examples how to put text inside the new slide:

```javascript
// Change the background color:
slide.back = '000000';

// Declare the default color to use on this slide (default is black):
slide.color = 'ffffff';

// Basic way to add text string:
slide.addText('This is a test');
slide.addText('Fast position', 0, 20);
slide.addText('Full line', 0, 40, '100%', 20);

// Add text box with multi colors and fonts:
slide.addText([
  {text: 'Hello ', options: {font_size: 56}},
  {text: 'World!', options: {font_size: 56, font_face: 'Arial', color: 'ffff00'}}
  ], {cx: '75%', cy: 66, y: 150});
// Please note that you can pass object as the text parameter to addText.

slide.addText('Office generator', {
  y: 66, x: 'c', cx: '50%', cy: 60, font_size: 48,
  color: '0000ff' } );

slide.addText('Big Red', {
  y: 250, x: 10, cx: '70%',
  font_face: 'Wide Latin', font_size: 54,
  color: 'cc0000', bold: true, underline: true } );
```

## Speaker notes:

PowerPoint slides can contain speaker notes, to do that use the setSpeakerNote method:

```javascript
slide.setSpeakerNote ( 'This is a speaker note!' );
```

## Charts:

PowerPoint slides can contain charts with embedded data.  To create a chart:

```javascript
slide.addChart(chartInfo)
```

Where `chartInfo` object is an object that takes the following attributes:

 - `data` -  an array of data, see examples below
 - `renderType` -  specifies base chart type, may be one of `"bar", "pie", "group-bar", "column", "stacked-column", "line"`
 - `title` -  chart title (default: none)
 - `valAxisTitle` -  value axis title (default: none)
 - `catAxisTitle` - category axis title (default: none)
 - `valAxisMinValue` - value axis min  (default: none)
 - `valAxisMaxValue` - vlaue axis max value (default: none)
 - `valAxisNumFmt` - value axis format, e.g `"$0"` or `"0%"` (default: none)
 - `valAxisMajorGridlines` - true|false (false)
 - `valAxisMinorGridlines` - true|false (false)
 - `valAxisCrossAtMaxCategory` - true|false (false)
 - `catAxisReverseOrder` - true|false (false)
 - `fontSize` - text size for chart, e.g. "1200" for 12pt type
 - `xml` - optional XML overrides to `<c:chart>` as a Javascript object that is mixed in

Also, the overall chart and  each data series take an an optional `xml` attribute, which specifies XML overrides to the `<c:series>` attribute.
* The `xml` argument for the `chartInfo` is mixed in to the `c:chartSpace` attribute.
* The `xml` argument for the `data` series is mixed into the `c:ser` attribute.

For instance, to specify the overall text size, you can specify the following on the `chartInfo` object.
The snippet below is what happens under the scenes when you specify `fontSize: 1200`

```javascript
chartInfo = {
 // ....
 "xml": {
      "c:txPr": {
        "a:bodyPr": {},
        "a:listStyle": {},
        "a:p": {
          "a:pPr": {
            "a:defRPr": {
              "@sz": "1200"
            }
          },
          "a:endParaRPr": {
            "@lang": "en-US"
          }
        }
      }
    }
```

Examples how to add chart into the slide:

```javascript
// Column chart
slide = pptx.makeNewSlide();
slide.name = 'Chart slide';
slide.back = 'ffffff';
slide.addChart(
  {   title: 'Column chart',
          renderType: 'column',
          valAxisTitle: 'Costs/Revenues ($)',
          catAxisTitle: 'Category',
          valAxisNumFmt: '$0',
                valAxisMaxValue: 24,
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
slide.addChart(
  {   title: 'My production',
      renderType: 'pie',
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
slide.addChart(
  { title: 'Sample bar chart',
    renderType: 'bar',
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

// Line Chart
slide = pptx.makeNewSlide();
slide.name = 'Line Chart slide';
slide.back = 'ff00ff';
slide.addChart(
  { title: 'Sample line chart',
    renderType: 'line',
      data:  [ // each item is one serie
      {
        name: 'europe',
        labels: ['Y2003', 'Y2004', 'Y2005', 'Y2006'],
        values: [2.5, 2.6, 2.8, 2.4],
        color: 'ff0000' // optional
      },
      {
        name: 'namerica',
        labels: ['Y2003', 'Y2004', 'Y2005', 'Y2006'],
        values: [2.5, 2.7, 2.9, 3.2],
        color: '00ff00' // optional
      },
      {
        name: 'asia',
        labels: ['Y2003', 'Y2004', 'Y2005', 'Y2006'],
        values: [2.1, 2.2, 2.4, 2.2],
        color: '0000ff' // optional
      }
    ]
  }
)
```

## Tables:

Add a table to a PowerPoint slide:

```javascript
 var rows = [];
  for (var i = 0; i < 12; i++) {
    var row = [];
    for (var j = 0; j < 5; j++) {
      row.push("[" + i + "," + j + "]");
    }
    rows.push(row);
  }
  slide.addTable(rows, {});
```

Specific options for tables (in addition to standard : x, y, cx, cy, etc.) :
- columnWidth : width of all columns (same size for all columns). Must be a number (~1 000 000)
- columnWidths : list of width for each columns (custom size per column). Must be array of number. This param will overwrite columnWidth if both are given

Formatting can also be applied directly to a cell:

```javascript
var rows = [];
rows.push([
	{
		val: "Category",
        	opts: {
          		font_face   : "Arial",
          		align       : "l",
          		bold        : 0
        	}
      },
      {
        	val  :"Average Score",
        	opts: {
          		font_face   : "Arial",
          		align       : "r",
          		bold        : 1,
          		font_color  : "000000",
          		fill_color  : "f5f5f5"
        	}
      }
]);
slide.addTable(rows, {});
```
