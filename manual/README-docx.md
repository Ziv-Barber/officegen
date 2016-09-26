# Create Microsoft Office Word Document Reference

## Contents: ##

- [Creating the document object](#basic)
- [The document object's settings](#settings)

<a name="basic"></a>
## Creating the document object: ##

First, if you didn't have it yet, get access to the officegen module:

```js
var officegen = require ( 'officegen' );
```

Now you have few ways to use it to create a docx based document. The simple way is to use this code:

```js
var docx = officegen ( 'docx' );
```

But if you want to pass some settings then you should use the following format:

```js
var docx = officegen ({
	type: 'docx', // We want to create a Microsoft Word document.
	... // Extra options goes here.
});
```

<a name="settings"></a>
### The document object's settings: ###

- author (string) - The document's author (part of the Document's Properties in Office).
- creator (string) - Alias. The document's author (part of the Document's Properties in Office).
- description (string) - The document's properties comments (part of the Document's Properties in Office).
- keywords (string) - The document's keywords (part of the Document's Properties in Office).
- orientation (string) - Either 'landscape' or 'portrait'. The default is 'portrait'.
- subject (string) - The document's subject (part of the Document's Properties in Office).
- title (string) - The document's title (part of the Document's Properties in Office).

You can always change some of these settings after creating the docx object using there methods:

```js
docx.setDocTitle ( '...' );
docx.setDocSubject ( '...' );
docx.setDocKeywords ( '...' );
docx.setDescription ( '...' );
docx.setDocCategory ( '...' );
docx.setDocStatus ( '...' );
```

