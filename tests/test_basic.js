// Assert tools needed:
var chai = require('chai')
var assert = chai.assert
var expect = chai.expect

// The package under test:
var officegen = require('../')

describe('Officegen internals test suits', function () {
	describe('basicgen test suit', function () {
		it('plugins#getPrototypeByName ()', function (done) {
			expect(officegen.plugins.getPrototypeByName('msoffice')).to.be.an('object').to.be.ok
			done()
		})
	})
})
