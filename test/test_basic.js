// Assert tools needed:
var chai = require ( 'chai' );
var assert = chai.assert;
var expect = chai.expect;

// The package under test:
var officegen = require('../');

describe ( 'Officegen internals test suits', function () {
	describe ( '<module> test suit', function () {
		// Executed before each test:
		beforeEach ( function ( done ) {
			done ();
		});

		it ( '<some test>', function ( done ) {
			// this.slow ( 500 );
			var myVar = 200;
			expect ( myVar ).to.equal ( 200 );
			assert ( myVar === 200, "How it can be?" );
			// include above below not true (.to.be.true;) also: .to.not.throw(Error);
			done ();
		});
	});
});
