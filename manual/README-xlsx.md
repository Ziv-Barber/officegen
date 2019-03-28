# Create Microsoft Office Word Document Reference

## Contents: ##

- [Creating the document object](#basic)
- [The document object's settings](#settings)

<a name="basic"></a>
## Creating the document object: ##

First, if you didn't have it yet, get access to the officegen module:

```js
const officegen = require('officegen')
```

Now you have few ways to use it to create a xlsx based document. The simple way is to use this code:

```js
let xlsx = officegen('xlsx')
```

But if you want to pass some settings then you should use the following format:

```js
let xlsx = officegen({
	type: 'xlsx', // We want to create a Microsoft Excel document.
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

You can always change some of these settings after creating the xlsx object using there methods:

```js
xlsx.setDocTitle('...')
xlsx.setDocSubject('...')
xlsx.setDocKeywords('...')
xlsx.setDescription( '...')
xlsx.setDocCategory('...')
xlsx.setDocStatus('...')
```

Fill cells:

```javascript
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

