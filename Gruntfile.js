module.exports = function(grunt) {
	var path = require("path");
	// Project configuration.
	grunt.initConfig({
		pkg: grunt.file.readJSON('package.json'),
		coffee: {
			dynamic_mappings: {
			    expand: true,
			    flatten: false,
			    cwd: 'src/',
			    src: ['**/*.coffee'],
			    dest: 'lib/',
			    ext: '.js'
			}
		},
		
		watch: {
			coffee: {
				files: ['src/*.coffee'],
				tasks: 'coffee'
			},
		}
	});
	
	grunt.event.on('watch', function(action, filepath, target) {
		console.log(filepath);
        grunt.config(['coffee', 'dynamic_mappings', 'files', 'src'], filepath);
    } );
	
	// Load the plugin that provides the "less" task.

	grunt.loadNpmTasks('grunt-contrib-watch');
	grunt.loadNpmTasks('grunt-contrib-coffee');
	
	grunt.registerTask('default', ['coffee:dynamic_mappings', 'watch']);
};