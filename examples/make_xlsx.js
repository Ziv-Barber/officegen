var fs = require('fs');

var xlsx = require('../officegen.js').makegen ( { 'type': 'xlsx', 'onend': function ( written ) {
	console.log ( 'Finish to create an Excel file.\nTotal bytes created: ' + written + '\n' );
} } );

sheet = xlsx.makeNewSheet ();
sheet.name = 'Excel Test';

// The direct option - two-dimensional array:
sheet.data[0] = [];
sheet.data[0][0] = 1;
sheet.data[1] = [];
sheet.data[1][3] = 'abc';
sheet.data[1][4] = 'bla bla';
sheet.data[1][8] = 'OK';
sheet.data[2] = [];
sheet.data[2][5] = 'abc';
sheet.data[2][6] = 900;
sheet.data[6] = [];
sheet.data[6][2] = 1972;

var out = fs.createWriteStream ( 'out.xlsx' );

xlsx.generate ( out );

