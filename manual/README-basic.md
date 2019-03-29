# Creating an officegen stream object:

First, make sure to require the officegen module:

```javascript
var officegen = require('officegen')
```

There are two ways to use the officegen returned function to create an officegen stream:

```javascript
var myDoc = officegen('<type of document to create>')

// or:

var myDoc = officegen({
  'type': '<type of document to create>'
  // More options here (if needed)
})

// Supported types:
// 'pptx' or 'ppsx' - Microsoft PowerPoint based document.
// 'docx' - Microsoft Word based document.
// 'xlsx' - Microsoft Excel based document.
```

Generating an empty Microsoft PowerPoint officegen stream:

```javascript
var pptx = officegen('pptx')
```

Generating an empty Microsoft Word officegen stream:

```javascript
var docx = officegen('docx')
```

Generating an empty Microsoft Excel officegen stream:

```javascript
var xlsx = officegen('xlsx')
```

General events of the officegen stream:

- 'finalize' - been called after finishing to create the document.
- 'error' - been called on error.

Event examples:

```javascript
pptx.on('finalize', function (written) {
  console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n')
})

pptx.on('error', function (err) {
  console.log(err)
})
```

Another way to register either 'finalize' or 'error' events:

```javascript
var pptx = officegen({
  'type': 'pptx', // or 'xlsx', etc
  'onend': function (written) {
    console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n')
  },
  'onerr': function (err) {
    console.log(err)
  }
})
```

If you are preferring to use callbacks instead of events you can pass your callbacks to the generate method
(see below).

Now you should fill the object with data (we'll see below) and then you should call generate with
an output stream to create the output Office document.

Example with pptx:

```javascript
var out = fs.createWriteStream('out.pptx')

pptx.generate(out)
out.on('close', function () {
  console.log('Finished to create the PPTX file!')
})
```

Passing callbacks to generate:

```javascript
var out = fs.createWriteStream('out.pptx')

pptx.generate(out, {
  'finalize': function (written) {
    console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n')
  },
  'error': function (err) {
    console.log(err)
  }
})
```

Generating HTTP stream example (no file been created):

```javascript
var http = require('http')
var officegen = require('officegen')

http.createServer(function (request, response) {
  response.writeHead (200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'Content-disposition': 'attachment; filename=surprise.pptx'
  })

  var pptx = officegen('pptx')

  pptx.on('finalize', function (written) {
    // We don't really need it in this case.
  })

  pptx.on('error', function (err) {
    // Error handing...
  })

  // ... (fill pptx with data)

  // Generate the Powerpoint document and sent it to the client via http:
  pptx.generate(response)
}).listen(3000)
```

### Put data inside the document object: ###

#### MS-Office document properties (for all document types): ###

The default Author of all the documents been created by officegen is 'officegen'. If you want to put anything else please
use the 'creator' option when calling the officegen function:

```javascript
var pptx = officegen({
  'type': 'pptx', // or 'xlsx', etc.
  'creator': '<your project name here>'
})
```

Change the document title (pptx,ppsx,docx):

```javascript
var pptx = officegen({
  'type': 'pptx',
  'title': '<title>'
});

// or

pptx.setDocTitle('<title>')
```

For Word only:

```javascript
var docx = officegen({
  'type': 'docx',
  'subject': '...',
  'keywords': '...',
  'description': '...'
});

// or

docx.setDocSubject('...')
docx.setDocKeywords('...')
docx.setDescription('...')
```

## Debugging:

If needed, you can activate some verbose messages (warning: this does not cover all part of the lib yet) with :

```javascript
officegen.setVerboseMode(true)
```
