// Simple server that displaying form to ask the user name and then generate PowerPoint stream with the user's name 
// without using real files on the server side.

var fs = require('fs');
var http = require("http");
var querystring = require('querystring');

function postRequest ( request, response, callback ) {
    var queryData = "";
    if ( typeof callback !== 'function' ) return null;

    if ( request.method == 'POST' ) {
        request.on ( 'data', function ( data ) {
            queryData += data;
            if ( queryData.length > 100 ) {
                queryData = "";
                response.writeHead ( 413, {'Content-Type': 'text/plain'}).end ();
                request.connection.destroy ();
            }
        });

        request.on ( 'end', function () {
            response.post = querystring.parse ( queryData );
            callback ();
        });

    } else {
        response.writeHead ( 405, { 'Content-Type': 'text/plain' });
        response.end ();
    }
}

http.createServer ( function ( request, response ) {
	if ( request.method == 'GET' )
	{
		response.writeHead ( 200, "OK", { 'Content-Type': 'text/html' });
		response.write ( '<html>\n<head></head>\n<body>\n' );
		response.write ( '<h1>Please enter your name here:</H1>\n' );
		response.write ( '<form method="post" action="http://127.0.0.1:3000/"><input type="text" name="name"><input type="submit" value="Submit"></form>\n' );
		response.write ( '</body>\n</html>\n' );
		response.end ();

	} else
	{
		postRequest ( request, response, function () {
			// console.log ( response.post );

			response.writeHead ( 200, {
				"Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
				'Content-disposition': 'attachment; filename=surprise.pptx'
				});
			// .xlsx   application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
			// .xltx   application/vnd.openxmlformats-officedocument.spreadsheetml.template
			// .potx   application/vnd.openxmlformats-officedocument.presentationml.template
			// .ppsx   application/vnd.openxmlformats-officedocument.presentationml.slideshow
			// .pptx   application/vnd.openxmlformats-officedocument.presentationml.presentation
			// .sldx   application/vnd.openxmlformats-officedocument.presentationml.slide
			// .docx   application/vnd.openxmlformats-officedocument.wordprocessingml.document
			// .dotx   application/vnd.openxmlformats-officedocument.wordprocessingml.template
			// .xlam   application/vnd.ms-excel.addin.macroEnabled.12
			// .xlsb   application/vnd.ms-excel.sheet.binary.macroEnabled.12

			var pptx = require('../officegen.js').makegen ( { 'type': 'pptx', 'onend': function ( written ) {
				console.log ( 'Finish to create the surprise PowerPoint stream and send it to ' + response.post.name + '.\nTotal bytes created: ' + written + '\n' );
			} } );

			slide = pptx.makeNewSlide ();
			slide.back = 'cc88cc';
			slide.addText ( 'Hello ' + response.post.name + '!', { x: 600000, y: 10000, font_size: 56, cx: 10000000 } );
			pptx.generate ( response );
		});
	} // Endif.
}).listen ( 3000 );

console.log ( 'The PowerPoint server is listening on port 3000.\n' );

