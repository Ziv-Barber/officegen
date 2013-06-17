var fs = require('fs');

var docx = require('../officegen.js').makegen ( { 'type': 'docx', 'onend': function ( written ) {
	console.log ( 'Finish to create Word file.\nTotal bytes created: ' + written + '\n' );
} } );

// BMK_TODO:

var out = fs.createWriteStream ( 'out.docx' );

docx.generate ( out );

