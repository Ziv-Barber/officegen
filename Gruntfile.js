var path = require ( 'path' );

/**
 * List of grunt tasks.
 * @namespace gruntfile
 */
module.exports = function ( grunt ) {
	// We are checking how much time took each grunt task:
	require ( 'time-grunt' ) ( grunt );

	function lastModified ( minutes ) {
		return function ( filepath ) {
			var filemod = ( require ( 'fs' ).statSync ( filepath ) ).mtime;
			var timeago = ( new Date () ).setDate ( (new Date () ).getMinutes () - minutes );
			return ( filemod > timeago );
		};
	}

	grunt.initConfig ({
		pkg: grunt.file.readJSON ( 'package.json' ),

		jshint: {
			// List of all the source files to test:
			files: [ 'gruntfile.js', 'lib/**/*.js' ],

			// Configure JSHint (documented at http://www.jshint.com/docs/):
			options: {
				evil: false,
				multistr: true, // We need to take care about it.
				globals: {
					console: true,
					module: true
				}
			}
		},

		jsdoc : {
			dist : {
				src: ['gruntfile.js', 'lib/**/*.js'],
				options: {
					'destination': 'doc',
					'package': 'package.json',
					'readme': 'README.md',
					// template : "node_modules/grunt-jsdoc/node_modules/ink-docstrap/template",
					// configure : "node_modules/grunt-jsdoc/node_modules/ink-docstrap/template/jsdoc.conf.json"
				}
			}
		}
	});

	//
	// The default task:
	//

	/**
	 * The default grunt task.
	 * @name default
	 * @memberof gruntfile
	 * @kind function
	 */
	grunt.registerTask ( 'default', [
		'jshint'
	]);

	//
	// More Grunt tasks:
	//

	//
	// Load all the modules that we need:
	//

	grunt.loadNpmTasks ( 'grunt-contrib-jshint' );
	grunt.loadNpmTasks ( 'grunt-jsdoc' );
};
