# Adding a new document type:

## Notes:

- If you want to add a new Microsoft Office document type then [click here](../mmdoc/README.md).
- If you want to add a new OpenOffice document type then [click here](../oodoc/README.md).

## Overview:

It's possible to add more documentation types into officegen! For example, some work been done to add even OpenOffice documents! 
Each document type is implemented by doing something like this:

```
var baseobj = require('./basicgen.js')

/**
 * Extend officegen object with some new document type support.
 *
 * @param {object} genobj The object to extend.
 * @param {string} new_type The type of object to create.
 * @param {object} options The object's options.
 * @param {object} gen_private Access to the internals of this object.
 * @param {object} type_info Additional information about this type.
 */
function makeSomeType(genobj, new_type, options, gen_private, type_info) {
  // ...
}

baseobj.plugins.registerDocType(
  'mytype', // The type code string.
  makeSomeType,
  {},
  baseobj.docType.PRESENTATION,
  'My Document'
)
```

TBD

- [Go back to the plugins documentation](../README.md)
- [Go back to the advanced topics documentation](../../README.md)
- [Go back to the main documentation](../../../README.md)
