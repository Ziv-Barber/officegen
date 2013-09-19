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

var archiver = require('archiver');
var sys = require('sys');
var events = require('events');
var fs = require('fs');
var Stream = require('stream'); // BMK_STREAM:

// Globals:

var int_officegen_globals = {}; // Our internal globals.

int_officegen_globals.settings = {};
int_officegen_globals.types = {};
int_officegen_globals.common_obj = {};

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

	var genobj = this;    // Can be accessed by all the functions been declared inside the officegen object.
	var gen_private = {}; // For all the private data that we don't want the user of officegen to access it.

	gen_private.plugs = {}; // API for plugins.

	gen_private.perment = {}; // All stuff that is 100% unchangable after selecting the type to create.
	gen_private.thisDoc = {}; // All stuff that is 100% depended on the current document to create (all the stuff that 
	                          // been erased by calling to startNewDoc().
	gen_private.mixed = {}; // Mixed stuff (both perment and document depend).

	gen_private.perment.features = {}; // Features been configured by the type selector and you can't change them.
	// gen_private.perment.features.page_name
	// gen_private.perment.features.call_before_gen
	// gen_private.perment.features.call_after_gen
	// gen_private.perment.features.call_on_clear

	gen_private.thisDoc.pages = []; // Information about all the pages to create.
	gen_private.mixed.res_list = []; // List of all the resources to create inside the zip.
	gen_private.mixed.res_data = {}; // Information about all the resources to create.

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

		gen_private.mixed.res_list.push ( newRes );
	};

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] element_name ???.
	/// @param[in] def_data ???.
	/// @param[in] prop_name ???.
	/// @param[in] user_access_func_name ???.
	///
	function addInfoType ( element_name, def_data, prop_name, user_access_func_name ) {
		genobj.info[element_name] = {};
		genobj.info[element_name].element = element_name;
		genobj.info[element_name].data = def_data;
		genobj.info[element_name].def_data = def_data;

		// The user of officegen can configure this property using the options object:
		if ( genobj.options.prop_name )
		{
			genobj.info[element_name].data = genobj.options.prop_name;
		} // Endif.

		genobj[user_access_func_name] = function ( new_data ) {
			genobj.info[element_name].data = new_data;
		};
	};

	// Public API:

	///
	/// @brief Generating the output document stream.
	///
	/// The user of officegen must call this method after filling all the information about what to put inside 
	/// the generated document. This method is creating the output document directly into the given stream object.
	///
	/// @param[in] stream The stream to receive the generated document.
	///
	this.generate = function ( stream ) {
		if ( gen_private.perment.features.page_name ) {
			if ( gen_private.thisDoc.pages.length == 0 ) {
				genobj.emit ( 'error', 'ERROR: No ' + gen_private.perment.features.page_name + ' been found inside your document.' );
			} // Endif.
		} // Endif.

		// Optional callback to prepare everything for generating:
		if ( gen_private.perment.features.call_before_gen )
		{
			gen_private.perment.features.call_before_gen ();
		} // Endif.

		var archive = archiver('zip');

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

		archive.pipe ( stream );

		///
		/// @brief Add the next resource into the zip stream.
		///
		/// This function adding the next resource into the zip stream.
		///
		function generateNextResource ( cur_index )
		{
			var resStream;

			if ( cur_index < gen_private.mixed.res_list.length ) {
				if ( typeof gen_private.mixed.res_list[cur_index] != 'undefined' ) {
					switch ( gen_private.mixed.res_list[cur_index].type ) {
						case 'buffer':
							resStream = gen_private.mixed.res_list[cur_index].callback ( gen_private.mixed.res_list[cur_index].data );
							break;

						// BMK_STREAM: (***START***)
						// Using some kind of simple 'template' engine:
						case 'custom':
							resStream = new Stream ();
							resStream.readable = true;
							process.nextTick ( function() {
								// The callback should emit data events and then end event. The problem is that we can't 
								// call emit before the pipe starting to run. That's why we are not executing the callback 
								// immediately but using the process.nextTick trick to make it to run after the pipe is 
								// starting and someone is listening to our events.
								gen_private.mixed.res_list[cur_index].callback ( resStream, gen_private.mixed.res_list[cur_index].data );
							});
							break;
						// BMK_STREAM: (***END***)

						// Just copy the file as is:
						case 'file':
							resStream = fs.createReadStream ( gen_private.mixed.res_list[cur_index].data || gen_private.mixed.res_list[cur_index].name );
							break;

						// Just use this stream:
						case 'stream':
							resStream = gen_private.mixed.res_list[cur_index].data;
							break;
					} // End of switch.

					if ( typeof resStream != 'undefined' ) {
						if ( int_officegen_globals.settings.verbose ) {
							console.log ( 'Adding "' + gen_private.mixed.res_list[cur_index].name + '" (' + gen_private.mixed.res_list[cur_index].type + ')...' );
						} // Endif.

						archive.append ( resStream, { name: gen_private.mixed.res_list[cur_index].name }, function () {
							setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
						});
						

					} else {
						setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
					} // Endif.

				} else {
					setImmediate ( function() { generateNextResource ( cur_index + 1 ); });
				} // Endif.

			} else {
				archive.finalize ( function ( err, written ) {
					// Optional callback to clean after us:
					if ( gen_private.perment.features.call_after_gen )
					{
						gen_private.perment.features.call_after_gen ( err, written );
					} // Endif.

					if ( err ) {
						onArchiveError ( err );
					} // Endif.

					genobj.emit ( 'finalize', written );
				});
			} // Endif.
		};

		// Start the process of generating the output zip stream:
		generateNextResource ( 0 );
	};

	///
	/// @brief Reuse this object for a new document of the same type.
	///
	/// Call this method if you want to start generating a new document of the same type using this object.
	///
	this.startNewDoc = function () {
		var kill = [];

		for ( var i = 0; i < gen_private.mixed.res_list.length; i++ ) {
			if ( !gen_private.mixed.res_list[i].is_perment ) kill.push ( i );
		} // End of for loop.

		for ( var i = 0; i < kill.length; i++ ) gen_private.mixed.res_list.splice ( kill[i] - i, 1 );

		gen_private.thisDoc.pages.length = 0;

		if ( gen_private.perment.features.call_on_clear ) {
			gen_private.perment.features.call_on_clear ();
		} // Endif.
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
		intAddAnyResourceToParse ( resource_name, type_of_res, res_data, res_cb, false );
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
///
function registerDocType ( typeName, createFunc, schema_data ) {
	int_officegen_globals.types[typeName] = {};
	int_officegen_globals.types[typeName].createFunc = createFunc;
	int_officegen_globals.types[typeName].schema_data = schema_data;
};

var ogen = module.exports = makegen;

ogen.makegen = makegen; // To support old versions.
ogen.setVerboseMode = setVerboseMode;
ogen.registerDocType = registerDocType;
ogen.schema = int_officegen_globals.types;

