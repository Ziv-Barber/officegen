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

var sys = require('sys');
var events = require('events');

var Transform = require('stream').Transform || require('readable-stream/transform');

// Used by generate:
var archiver = require('archiver');
var fs = require('fs');
var PassThrough = require('stream').PassThrough || require('readable-stream/passthrough');
var _ = require('underscore');
var async = require('async');
var startpoint = require('startpoint');
// var sanitizer = require('sanitizer');

// Globals:

var int_officegen_globals = {}; // Our internal globals.

int_officegen_globals.settings = {};
int_officegen_globals.types = {};
int_officegen_globals.docPrototypes = {};
int_officegen_globals.resParserTypes = {};

///
/// @brief The constructor of the office generator object.
///
/// This constructor function is been called by makegen().
///
/// @b The @b Options:
///
/// The configuration options effecting the operation of the officegen object. Some of them can be only been 
/// declared on the 'options' object passed to the constructor object and the rest can be configured by either 
/// a property with the same name or by special function.
///
/// @b List @b Of @b Options:
///
/// - 'type' - the type of generator to create. Possible options: either 'pptx', 'docx' or 'xlsx'.
/// - 'creator' - the name of the document's author. The default is 'officegen'.
/// - 'onend' - callback that been fired after finishing to create the zip stream.
/// - 'onerr' - callback that been fired on error.
///
/// @param[in] options List of configuration options (see in the description of this function).
///
officegen = function ( options ) {
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
	var gen_private = {}; // For all the private data that we don't want the user of officegen to access it.

	gen_private.plugs = {}; // API for plugins.
	gen_private.features = {}; // Features been configured by the type selector and you can't change them.
	gen_private.features.type = {};
	gen_private.features.outputType = 'zip';
	// gen_private.features.page_name

	gen_private.pages = []; // Information about all the pages to create.
	gen_private.resources = []; // List of all the resources to create inside the zip.

	gen_private.type = {};

	///
	/// @brief Combine the given options and the default values.
	///
	/// This function creating the real options object.
	///
	/// @param[in] options The options to configure.
	///
	function setOptions ( object, source ) {
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
		};

		function keys (object) {
			if (!isObject(object)) {
				return [];
			}

			return Object.keys(object);
		};

		var index;
		var iterable = object;
		var result = iterable;

		var args = arguments;
		var argsIndex = 0;
		var argsLength = args.length;

		while (++argsIndex < argsLength) {
			iterable = args[argsIndex];

			if (iterable && objectTypes[typeof iterable]) {
				var ownIndex = -1;
				var ownProps = objectTypes[typeof iterable] && keys(iterable);
				var length = ownProps ? ownProps.length : 0;

				while (++ownIndex < length) {
					index = ownProps[ownIndex];

					if (typeof result[index] === 'undefined' || result[index] == null) {
						result[index] = iterable[index];

					} else if (isObject(result[index]) && isObject(iterable[index])) {
						result[index] = setOptions(result[index], iterable[index]);
					} // Endif.
				} // End of while loop.
			} // Endif.
		} // End of while loop.

		return result;
	};

	///
	/// @brief Configure this object to generate the given type of document.
	///
	/// This function configuring the generator to create the given type of document.
	///
	/// @param[in] new_type The type of document to create.
	///
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
	};

	///
	/// @brief Add a resource to the list of resources to place inside the output zip file.
	///
	/// This method adding a resource to the list of resources to place inside the output document ZIP.
	///
	/// @param[in] resource_name The name of the resource (path).
	/// @param[in] type_of_res The type of this resource: either 'file' or 'buffer'.
	/// @param[in] res_data Optional data to use when creating this resource.
	/// @param[in] res_cb Callback to generate this resource (for 'buffer' mode only).
	/// @param[in] is_always Is true if this resource is perment for all the zip of this document type.
	///
	gen_private.plugs.intAddAnyResourceToParse = function ( resource_name, type_of_res, res_data, res_cb, is_always ) {
		var newRes = {};

		newRes.name = resource_name;
		newRes.type = type_of_res;
		newRes.data = res_data;
		newRes.callback = res_cb;
		newRes.is_perment = is_always;

		gen_private.resources.push ( newRes );
	};

	// Any additional plugin API must be placed here:
	gen_private.plugs.type = {};

	// Public API:

	///
	/// @brief Generating the output document stream.
	///
	/// The user of officegen must call this method after filling all the information about what to put inside 
	/// the generated document. This method is creating the output document directly into the given stream object.
	///
	/// The options parameters properties:
	///
	/// 'finalize' - callback to be called after finishing to generate the document.
	/// 'error' - callback to be called on error.
	///
	/// @param[in] output_stream The stream to receive the generated document.
	/// @param[in] options Way to pass callbacks.
	///
	this.generate = function ( output_stream, options ) {
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

		///
		/// @brief Error handler.
		///
		/// This is our error handler method for creating archive.
		///
		/// @param[in] err The error string.
		///
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

		var streamTransformers = {
			'buffer': function ( obj ) {
				return startpoint ( obj.callback ( obj.data ) );
			},
			'file': function ( obj ) {
				return fs.createReadStream ( obj.data || obj.name );
			},
			'stream': function ( obj ) {
				return obj.data;
			},
			'officegen': function ( obj ) {
				resStream = new PassThrough ();
				obj.data.generate ( resStream );
			},
			'custom': function ( obj ) {
				for ( var cur_parserType in int_officegen_globals.resParserTypes ) {
					if ( (obj.subType == obj.type) && int_officegen_globals.resParserTypes[cur_parserType] && int_officegen_globals.resParserTypes[cur_parserType].parserFunc ) {
						resStream = int_officegen_globals.resParserTypes[cur_parserType].parserFunc (
							genobj,
							obj.name,
							obj.callback, // Can be used as the template source for template engines.
							obj.data,     // The data for the template engine.
							int_officegen_globals.resParserTypes[cur_parserType].extra_data
						);
						break;
					} // Endif.
				} // End of for loop.
			}
		};

		var workQueue = async.queue ( function ( item, done ) {
			var result = streamTransformers[item.type]( item );
			if ( result != undefined ) {
				// The internal verbose mode to help debugging:
				if ( int_officegen_globals.settings.verbose ) {
					console.log ( 'Adding "' + item.name + '" (' + item.type + ')...' );
				} // Endif.

				archive.append ( result, { name: item.name }, function ( err ) {
					_.delay ( done, 1 ); // fix issue in node 10.x with corrupt files.
				});

			} else {
				_.delay ( done, 1 ); // fix issue in node 10.x with corrupt files.
			} // Endif.
		});

		workQueue.drain = function () {
			archive.finalize ( function ( err, written ) {
				// Event to the type generator:
				genobj.emit ( 'afterGen', gen_private, err, written );

				if ( err ) {
					onArchiveError ( err );
				} // Endif.

				genobj.emit ( 'finalize', written );
			});
		};

		_.each ( gen_private.resources, function ( d ) { workQueue.push ( d ); });
	};

	///
	/// @brief Reuse this object for a new document of the same type.
	///
	/// Call this method if you want to start generating a new document of the same type using this object.
	///
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

	///
	/// @brief Register a new resource to add into the generated ZIP stream.
	///
	/// Using this method the user can add extra custom resources into the generated ZIP stream.
	///
	/// @param[in] resource_name The name of the resource (path).
	/// @param[in] type_of_res The type of this resource: either 'file' or 'buffer'.
	/// @param[in] res_data Optional data to use when creating this resource.
	/// @param[in] res_cb Callback to generate this resource (for 'buffer' mode only).
	///
	this.addResourceToParse = function ( resource_name, type_of_res, res_data, res_cb ) {
		// We don't want the user to add perment resources to the list of resources:
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
// sys.inherits ( officegen, Transform );

// officegen.prototype._transform = function (chunk, encoding, callback) {
	// BMK_TODO:
// 	callback ();
// }

///
/// @brief Create a new officegen object.
///
/// This method creating a new officegen based object.
///
/// @b Example:
///
/// @code
/// @endcode
///
function makegen ( options ) {
	return new officegen ( options );
};

///
/// @brief Change the verbose state of officegen.
///
/// This is a global settings effecting all the officegen objects in your application. You should 
/// use it only for debugging.
///
/// @param[in] new_state Either true or false.
///
function setVerboseMode ( new_state ) {
	int_officegen_globals.settings.verbose = new_state;
};

///
/// @brief Register a new type of document that we can generate.
///
/// This method registering a new type of document that we can generate. You can extend officegen to support any 
/// type of document that based on resources files inside ZIP stream.
///
/// @param[in] typeName The type of the document file.
/// @param[in] createFunc The function to use to create this type of file.
/// @param[in] schema_data Information needed by Schema-API to generate this kind of document.
/// @param[in] docType Document type.
/// @param[in] displayName The display name of this type.
///
function registerDocType ( typeName, createFunc, schema_data, docType, displayName ) {
	int_officegen_globals.types[typeName] = {};
	int_officegen_globals.types[typeName].createFunc = createFunc;
	int_officegen_globals.types[typeName].schema_data = schema_data;
	int_officegen_globals.types[typeName].type = docType;
	int_officegen_globals.types[typeName].display = displayName;
};

///
/// @brief Register a document prototype object.
///
/// This method registering a prototype document object. You can place all the common code needed by a group of document 
/// types in a single prototype object.
///
/// @param[in] typeName The name of the prototype object.
/// @param[in] baseObj The prototype object.
/// @param[in] displayName The display name of this type.
///
function registerPrototype ( typeName, baseObj, displayName ) {
	int_officegen_globals.docPrototypes[typeName] = {};
	int_officegen_globals.docPrototypes[typeName].baseObj = baseObj;
	int_officegen_globals.docPrototypes[typeName].display = displayName;
};

///
/// @brief Get a document prototype object by name.
///
/// This method get a prototype object.
///
/// @param[in] typeName The name of the prototype object.
/// @return The name of the prototype object.
///
function getPrototypeByName ( typeName ) {
	return int_officegen_globals.docPrototypes[typeName];
};

///
/// @brief Register a new resource parser.
///
/// This method registering a new resource parser. One use of this feature is in case that you are developing a new 
/// type of document and you want to extand officegen to use some kind of template engine as jade, ejs, haml* or CoffeeKup. 
/// In this case you can use a template engine to generate one or more of the resources inside the output archive. 
/// Another use of this method is to replace an existing plugin with different implementation.
///
/// @param[in] typeName The type of the parser plugin.
/// @param[in] parserFunc The resource generating function.
/// @param[in] extra_data Optional additional data that may be required by the parser function.
/// @param[in] displayName The display name of this type.
///
function registerParserType ( typeName, parserFunc, extra_data, displayName ) {
	int_officegen_globals.resParserTypes[typeName] = {};
	int_officegen_globals.resParserTypes[typeName].parserFunc = parserFunc;
	int_officegen_globals.resParserTypes[typeName].extra_data = extra_data;
	int_officegen_globals.resParserTypes[typeName].display = displayName;
};

var ogen = module.exports = makegen;

ogen.makegen = makegen; // To support old versions.
ogen.setVerboseMode = setVerboseMode;
ogen.plugins = {};
ogen.plugins.registerDocType = registerDocType;
ogen.plugins.registerPrototype = registerPrototype;
ogen.plugins.getPrototypeByName = getPrototypeByName;
ogen.plugins.registerParserType = registerParserType;
ogen.schema = int_officegen_globals.types;

ogen.docType = { "TEXT" : 1, "SPREADSHEET" : 2, "PRESENTATION" : 3 };

