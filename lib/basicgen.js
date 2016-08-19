//
// officegen: basic common code
//
// Please refer to README.md for this module's documentations.
//
// NOTE:
// - Before changing this code please refer to the hacking the code section on README.md.
//
// Copyright (c) 2013 Ziv Barber;
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// 'Software'), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
// IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
// CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
// TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

require("setimmediate"); // To be compatible with all versions of node.js

var sys = require('util');
var events = require('events');

var Transform = require('stream').Transform || require('readable-stream/transform');

// Used by generate:
var archiver = require('archiver');
var fs = require('fs');
var PassThrough = require('stream').PassThrough || require('readable-stream/passthrough');

// Global data shared by all the officegen objects:

var int_officegen_globals = {}; // Our internal global objects.

int_officegen_globals.settings = {};
int_officegen_globals.types = {};
int_officegen_globals.docPrototypes = {};
int_officegen_globals.resParserTypes = {};

/**
 * The constructor of the office generator object.
 * <br /><br />
 * This constructor function is been called by makegen().
 *
 * <h3><b>The options:</b></h3>
 *
 * The configuration options effecting the operation of the officegen object. Some of them can be only been 
 * declared on the 'options' object passed to the constructor object and the rest can be configured by either 
 * a property with the same name or by special function.
 *
 * <h3><b>List of options:</b></h3>
 *
 * <ul>
 * <li>'type' - the type of generator to create. Possible options: either 'pptx', 'docx' or 'xlsx'.</li>
 * <li>'creator' - the name of the document's author. The default is 'officegen'.</li>
 * <li>'onend' - callback that been fired after finishing to create the zip stream.</li>
 * <li>'onerr' - callback that been fired on error.</li>
 * </ul>
 *
 * @param {object} options List of configuration options (see in the description of this function).
 * @constructor
 */
var officegen = function ( options ) {
    if ( false === ( this instanceof officegen )) {
		return new officegen ( options );
	} // Endif.

	events.EventEmitter.call ( this );
	// Transform.call ( this, { objectMode : true } );

	// Internal events for plugins - NOT for the user:
	// event 'beforeGen'
	// event 'afterGen'
	// event 'clearDoc'

	var genobj = this;    // Can be accessed by all the functions been declared inside the officegen object.

	/**
	 * For all the private data that we don't want the user of officegen to access it.
	 * Each officegen object has it's own copy of the private object so changes been done to the private object of one officegen document will not effect other objects.
	 * @namespace module:lib/basicgen.private
	 */
	var gen_private = {};

	/**
	 * API for plugins.
	 * <br /><br />
	 * Officegen plugins can extend officegen to support more document formats.
	 * To register a new format you must:
	 * <br /><br />
	 * var baseobj = require ( "officegen" );
	 * <br /><br />
	 * and then call baseobj.plugins.registerDocType to add your type.
	 * <br /><br />
	 * Examples how to do it can be found on lib/gendocx.js, lib/genpptx.js and lib/genxlsx.js.
	 * @namespace module:lib/basicgen.private.plugs
	 * @example <caption>Adding a new document type to officegen</caption>
	 * var baseobj = require ( "officegen" );
	 * baseobj.plugins.registerDocType ( 'mydoctype', makeMyDocConstructor, {}, baseobj.docType.TEXT, "My Special Document File Format" );
	 */
	gen_private.plugs = {};

	gen_private.features = {}; // Features been configured by the type selector and you can't change them.
	gen_private.features.type = {};
	gen_private.features.outputType = 'zip';
	// gen_private.features.page_name

	gen_private.pages = []; // Information about all the pages to create.
	gen_private.resources = []; // List of all the resources to create inside the zip.

	gen_private.type = {};

	/**
	 * Combine the given options and the default values.
	 * <br /><br />
	 * 
	 * This function creating the real options object.
	 * 
	 * @param {object} options The options to configure.
	 */
	function setOptions ( object, source ) {// what is source for
		object = object || {};

		var objectTypes = {
			'boolean': false,
			'function': true,
			'object': true,
			'number': false,
			'string': false,
			'undefined': false
		};

		function isObject (value) {
			return !!(value && objectTypes[typeof value]);
		}

		function keys (object) {
			if (!isObject(object)) {
				return [];
			}

			return Object.keys(object);
		}

		var index;
		var iterable = object;
		var result = iterable;

		var args = arguments;
		var argsIndex = 0;
		var argsLength = args.length;

	        //loop variables 
		var ownIndex = -1;
		var ownProps = objectTypes[typeof iterable] && keys(iterable);
		var length = ownProps ? ownProps.length : 0;
	    
		while (++argsIndex < argsLength) {
			iterable = args[argsIndex];

			if (iterable && objectTypes[typeof iterable]) {

				while (++ownIndex < length) {
					index = ownProps[ownIndex];

					if (typeof result[index] === 'undefined' || result[index] === null) {
						result[index] = iterable[index];

					} else if (isObject(result[index]) && isObject(iterable[index])) {
						result[index] = setOptions(result[index], iterable[index]);
					} // Endif.
				} // End of while loop.
			} // Endif.
		} // End of while loop.

		return result;
	}

	/**
	 * Configure this object to generate the given type of document.
	 * <br /><br />
	 * 
	 * Called by the document constructor to configure the new document object to the given type.
	 * 
	 * @param {string} new_type The type of document to create.
	 */
	function setGeneratorType ( new_type ) {
		gen_private.length = 0;
		var is_ok = false;

		if ( new_type ) {
			for ( var cur_type in int_officegen_globals.types ) {
				if ( (cur_type == new_type) && int_officegen_globals.types[cur_type] && int_officegen_globals.types[cur_type].createFunc ) {
					int_officegen_globals.types[cur_type].createFunc ( genobj, new_type, genobj.options, gen_private, int_officegen_globals.types[cur_type] );
					is_ok = true;
					break;
				} // Endif.
			} // End of for loop.

			if ( !is_ok ) {
				// console.error ( '\nFATAL ERROR: Either unknown or unsupported file type - %s\n', options.type );
				genobj.emit ( 'error', 'FATAL ERROR: Invalid file type.' );
			} // Endif.
		} // Endif.
	}

	/**
	 * Add a resource to the list of resources to place inside the output zip file.
	 * <br /><br />
	 * 
	 * This method adding a resource to the list of resources to place inside the output document ZIP.
	 * <br />
	 * Changed by vtloc in 2014Jan10.
	 * 
	 * @param {string} resource_name The name of the resource (path).
	 * @param {string} type_of_res The type of this resource: either 'file', 'buffer', 'stream' or 'officegen' (the last one allow you to put office document inside office document).
	 * @param {object} res_data Optional data to use when creating this resource.
	 * @param {function} res_cb Callback to generate this resource (for 'buffer' mode only).
	 * @param {boolean} is_always Is true if this resource is perment for all the zip of this document type.
	 * @param {boolean} removed_after_used Is true if we need to delete this file after used.
	 */
	gen_private.plugs.intAddAnyResourceToParse = function ( resource_name, type_of_res, res_data, res_cb, is_always, removed_after_used ) {
		var newRes = {};

		newRes.name = resource_name;
		newRes.type = type_of_res;
		newRes.data = res_data;
		newRes.callback = res_cb;
		newRes.is_perment = is_always;
    
    // delete the temporatory resources after used
    // @author vtloc
    // @date 2014Jan10
    if( removed_after_used )
      newRes.removed_after_used = removed_after_used;
    else
      newRes.removed_after_used = false;

		if ( int_officegen_globals.settings.verbose ) {
			console.log("[officegen] Push new res : ", newRes);
		}

		gen_private.resources.push ( newRes );
	};

	// Any additional plugin API must be placed here:
	gen_private.plugs.type = {};

	// Public API:

	/**
	 * Generating the output document stream.
	 * <br /><br />
	 * 
	 * The user of officegen must call this method after filling all the information about what to put inside 
	 * the generated document. This method is creating the output document directly into the given stream object.
	 * 
	 * The options parameters properties:
	 * 
	 * 'finalize' - callback to be called after finishing to generate the document.
	 * 'error' - callback to be called on error.
	 * 
	 * @param {object} output_stream The stream to receive the generated document.
	 * @param {object} options Way to pass callbacks.
	 */
	this.generate = function ( output_stream, options ) {
		if ( int_officegen_globals.settings.verbose ) {
			console.log("[officegen] Start generate() : ", {outputType: gen_private.features.outputType });
		}

		if ( typeof options == 'object' ) {
			if ( options.finalize ) {
				genobj.on ( 'finalize', options.finalize );
			} // Endif.

			if ( options.error ) {
				genobj.on ( 'error', options.error );
			} // Endif.
		} // Endif.

		if ( gen_private.features.page_name ) {
			if ( gen_private.pages.length == 0 ) {
				genobj.emit ( 'error', 'ERROR: No ' + gen_private.features.page_name + ' been found inside your document.' );
			} // Endif.
		} // Endif.

		// Allow the type generator to prepare everything:
		genobj.emit ( 'beforeGen', gen_private );

		var archive = archiver( gen_private.features.outputType == 'zip' ? 'zip' : 'tar' );

		/**
		 * Error handler.
		 * <br /><br />
		 * 
		 * This is our error handler method for creating archive.
		 * 
		 * @param {string} err The error string.
		 */
		function onArchiveError ( err ) {
			genobj.emit ( 'error', err );
		}

		archive.on ( 'error', onArchiveError );

		if ( gen_private.features.outputType == 'gzip' ) {
			var zlib = require('zlib');
			var gzipper = zlib.createGzip ();

			archive.pipe ( gzipper ).pipe ( output_stream );

		} else {
			archive.pipe ( output_stream );
		} // Endif.

		/**
		 * Add the next resource into the zip stream.
		 * <br /><br />
		 * 
		 * This function adding the next resource into the zip stream.
		 */
		function generateNextResource ( cur_index )
		{
			if ( int_officegen_globals.settings.verbose ) {
				console.log("[officegen] generateNextResource("+cur_index+") : ", gen_private.resources[cur_index]);
			}

			var resStream;

			if ( cur_index < gen_private.resources.length ) {
				if ( typeof gen_private.resources[cur_index] != 'undefined' ) {
					switch ( gen_private.resources[cur_index].type ) {
						// Generate the resource text data by calling to provided function:
						case 'buffer':
							resStream = gen_private.resources[cur_index].callback ( gen_private.resources[cur_index].data );
							break;

						// Just copy the file as is:
						case 'file':
							resStream = fs.createReadStream ( gen_private.resources[cur_index].data || gen_private.resources[cur_index].name );
							break;

						// Just use this stream:
						case 'stream':
							resStream = gen_private.resources[cur_index].data;
							break;

						// Officegen object:
						case 'officegen':
							resStream = new PassThrough ();
							gen_private.resources[cur_index].data.generate ( resStream );
							break;

						// Custom parser:
						default:
							for ( var cur_parserType in int_officegen_globals.resParserTypes ) {
								if ( (cur_parserType == gen_private.resources[cur_index].type) && int_officegen_globals.resParserTypes[cur_parserType] && int_officegen_globals.resParserTypes[cur_parserType].parserFunc ) {
									resStream = int_officegen_globals.resParserTypes[cur_parserType].parserFunc (
										genobj,
										gen_private.resources[cur_index].name,
										gen_private.resources[cur_index].callback, // Can be used as the template source for template engines.
										gen_private.resources[cur_index].data,     // The data for the template engine.
										int_officegen_globals.resParserTypes[cur_parserType].extra_data
									);
									break;
								} // Endif.
							} // End of for loop.
					} // End of switch.

					if ( typeof resStream != 'undefined' ) {
						if ( int_officegen_globals.settings.verbose ) {
							console.log ( '[officegen] Adding into archive : "' + gen_private.resources[cur_index].name + '" (' + gen_private.resources[cur_index].type + ')...' );
						} // Endif.

						archive.append ( resStream, { name: gen_private.resources[cur_index].name } );
						if( gen_private.resources[cur_index].removed_after_used )
						{
							// delete the temporatory resources after used
							// @author vtloc
							// @date 2014Jan10
							var fileName = gen_private.resources[cur_index].data || gen_private.resources[cur_index].name;
							fs.unlinkSync( fileName );
						} // Endif.

						generateNextResource ( cur_index + 1 );

					} else {
						if ( int_officegen_globals.settings.verbose ) {
							console.log("[officegen] resStream is undefined");   // is it normal ??
						}
						generateNextResource ( cur_index + 1 );
						// setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
					} // Endif.
          
				} else {
					// Removed resource - just ignore it:
					generateNextResource ( cur_index + 1 );
					// setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
				} // Endif.

			} else {
				// No more resources to add - close the archive:
				if ( int_officegen_globals.settings.verbose ) {
					console.log("[officegen] Finalizing archive ...");
				}
				archive.finalize ();

				// Event to the type generator:
				genobj.emit ( 'afterGen', gen_private, null, archive.pointer () );

				genobj.emit ( 'finalize', archive.pointer () );
			} // Endif.
		}

		// Start the process of generating the output zip stream:
		generateNextResource ( 0 );
	};

	/**
	 * Reuse this object for a new document of the same type.
	 * <br /><br />
	 * 
	 * Call this method if you want to start generating a new document of the same type using this object.
	 */
	this.startNewDoc = function () {
		var kill = [];

		for ( var i = 0; i < gen_private.resources.length; i++ ) {
			if ( !gen_private.resources[i].is_perment ) kill.push ( i );
		} // End of for loop.

		for ( var i = 0; i < kill.length; i++ ) gen_private.resources.splice ( kill[i] - i, 1 );

		gen_private.pages.length = 0;

		genobj.emit ( 'clearDoc', gen_private );
	};

	// Public API - plugin API:

	/**
	 * Register a new resource to add into the generated ZIP stream.
	 * <br /><br />
	 * 
	 * Using this method the user can add extra custom resources into the generated ZIP stream.
	 * 
	 * @param {string} resource_name The name of the resource (path).
	 * @param {string} type_of_res The type of this resource: either 'file' or 'buffer'.
	 * @param {object} res_data Optional data to use when creating this resource.
	 * @param {function} res_cb Callback to generate this resource (for 'buffer' mode only).
	 */
	this.addResourceToParse = function ( resource_name, type_of_res, res_data, res_cb ) {
		// We don't want the user to add permanent resources to the list of resources:
		gen_private.plugs.intAddAnyResourceToParse ( resource_name, type_of_res, res_data, res_cb, false );
	};

	if ( typeof options == 'string' ) {
		options = { 'type': options };
	} // Endif.

	// See the officegen descriptions for the rules of the options:
	genobj.options = setOptions ( options, { 'type': 'unknown' } );

	if ( genobj.options && genobj.options.onerr ) {
		genobj.on ( 'error', genobj.options.onerr );
	} // Endif.

	if ( genobj.options && genobj.options.onend ) {
		genobj.on ( 'finalize', genobj.options.onend );
	} // Endif.
	
	// Configure this object depending on the user's selected type:
	if ( genobj.options.type ) {
		setGeneratorType ( genobj.options.type );
	} // Endif.
};

sys.inherits ( officegen, events.EventEmitter );

/**
 * Create a new officegen object.
 * <br /><br />
 * 
 * This method creating a new officegen based object.
 * @module lib/basicgen
 */
module.exports = function ( options ) {
	return new officegen ( options );
};

/**
 * Change the verbose state of officegen.
 * <br /><br />
 * 
 * This is a global settings effecting all the officegen objects in your application. You should 
 * use it only for debugging.
 * 
 * @param {boolean} new_state Either true or false.
 */
module.exports.setVerboseMode = function setVerboseMode ( new_state ) {
	int_officegen_globals.settings.verbose = new_state;
}

/**
 * Plugin API effecting all the instances of the officegen object.
 *
 * @namespace module:lib/basicgen.plugins
 */
module.exports.plugins = {};

/**
 * Register a new type of document that we can generate.
 * <br /><br />
 * 
 * This method registering a new type of document that we can generate. You can extend officegen to support any 
 * type of document that based on resources files inside ZIP stream.
 * 
 * @param {string} typeName The type of the document file.
 * @param {function} createFunc The function to use to create this type of file.
 * @param {object} schema_data Information needed by Schema-API to generate this kind of document.
 * @param {string} docType Document type.
 * @param {string} displayName The display name of this type.
 */
module.exports.plugins.registerDocType = function ( typeName, createFunc, schema_data, docType, displayName ) {
	int_officegen_globals.types[typeName] = {};
	int_officegen_globals.types[typeName].createFunc = createFunc;
	int_officegen_globals.types[typeName].schema_data = schema_data;
	int_officegen_globals.types[typeName].type = docType;
	int_officegen_globals.types[typeName].display = displayName;
};

/**
 * Get a document type object by name.
 * <br /><br />
 * 
 * This method get a document type object.
 * 
 * @param {string} typeName The name of the document type.
 * @return The plugin object of the document type.
 */
module.exports.plugins.getDocTypeByName = function ( typeName ) {
	return int_officegen_globals.types[typeName];
};

/**
 * Register a document prototype object.
 * <br /><br />
 * 
 * This method registering a prototype document object. You can place all the common code needed by a group of document 
 * types in a single prototype object.
 * 
 * @param {string} typeName The name of the prototype object.
 * @param {object} baseObj The prototype object.
 * @param {string} displayName The display name of this type.
 */
module.exports.plugins.registerPrototype = function ( typeName, baseObj, displayName ) {
	int_officegen_globals.docPrototypes[typeName] = {};
	int_officegen_globals.docPrototypes[typeName].baseObj = baseObj;
	int_officegen_globals.docPrototypes[typeName].display = displayName;
};

/**
 * Get a document prototype object by name.
 * <br /><br />
 * 
 * This method get a prototype object.
 * 
 * @param {string} typeName The name of the prototype object.
 * @return The prototype plugin object.
 */
module.exports.plugins.getPrototypeByName = function getPrototypeByName ( typeName ) {
	return int_officegen_globals.docPrototypes[typeName];
};

/**
 * Register a new resource parser.
 * <br /><br />
 * 
 * This method registering a new resource parser. One use of this feature is in case that you are developing a new 
 * type of document and you want to extend officegen to use some kind of template engine as jade, ejs, haml* or CoffeeKup. 
 * In this case you can use a template engine to generate one or more of the resources inside the output archive. 
 * Another use of this method is to replace an existing plugin with different implementation.
 * 
 * @param {string} typeName The type of the parser plugin.
 * @param {function} parserFunc The resource generating function.
 * @param {object} extra_data Optional additional data that may be required by the parser function.
 * @param {string} displayName The display name of this type.
 */
module.exports.plugins.registerParserType = function ( typeName, parserFunc, extra_data, displayName ) {
	int_officegen_globals.resParserTypes[typeName] = {};
	int_officegen_globals.resParserTypes[typeName].parserFunc = parserFunc;
	int_officegen_globals.resParserTypes[typeName].extra_data = extra_data;
	int_officegen_globals.resParserTypes[typeName].display = displayName;
};

module.exports.schema = int_officegen_globals.types;

module.exports.docType = { "TEXT" : 1, "SPREADSHEET" : 2, "PRESENTATION" : 3 };

