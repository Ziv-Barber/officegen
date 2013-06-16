var fs = require('fs');

var xlsx = require('../officegen.js').makegen ( { 'type': 'xlsx', 'onend': function ( written ) {
	console.log ( 'Finish to create an Excel file.\nTotal bytes created: ' + written + '\n' );
} } );

sheet = xlsx.makeNewSheet ();
// sheet.name = 'Excel Test';
sheet.data[1] = {};
sheet.data[1].A = 1;
sheet.data[2] = {};
sheet.data[2].D = 4;
sheet.data[3] = {};
sheet.data[3].G = 900;
sheet.data[7] = {};
sheet.data[7].C = 1972;

var out = fs.createWriteStream ( 'out.xlsx' );

xlsx.generate ( out );

