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

