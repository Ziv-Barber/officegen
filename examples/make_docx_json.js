
var officegen = require('../lib/index.js');

var fs = require('fs');
var path = require('path');

var docx = officegen ( 'docx' );

docx.on ( 'finalize', function ( written ) {
			console.log ( 'Finish to create Word file.\nTotal bytes created: ' + written + '\n' );
		});

docx.on ( 'error', function ( err ) {
			console.log ( err );
		});

var data = [[{ align: 'right' }, {
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
	}, [{}, {
		type: "image",
		path: path.resolve(__dirname, 'images_for_examples/sword_001.png')
	},{
		type: "image",
		path: path.resolve(__dirname, 'images_for_examples/sword_002.png')
	}], {
		type: "pagebreak"
	}, [{}, {
		type: "numlist"
	}, {
		type: "text",
		text: "numList1.",
	}, {
		type: "numlist"
	}, {
		type: "text",
		text: "numList2.",
	}], [{}, {
		type: "dotlist"
	}, {
		type: "text",
		text: "dotlist1.",
	}, {
		type: "dotlist"
	}, {
		type: "text",
		text: "dotlist2.",
	}], {
		type: "pagebreak"
	}
]

var pObj = docx.createByJson(data);

var out = fs.createWriteStream ( 'out.docx' );

out.on ( 'error', function ( err ) {
	console.log ( err );
});

docx.generate ( out );

