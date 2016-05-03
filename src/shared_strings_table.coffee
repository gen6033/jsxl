Promise = require('bluebird')
xml2js = Promise.promisifyAll(require('xml2js'))
Utils = require('./utils')
parser = new xml2js.Parser()

class SharedStringsTable

	@parse:(xml)->
		return new Promise (resolve, reject)->
			parser.parseString xml,(err, result)->
				resolve(new SharedStringsTable(result))


	constructor:(xmlobj)->
		@_ = {}
		@_.xmlobj = xmlobj
		@_.count = xmlobj.sst.$.count
		@_.uniqueCount = xmlobj.sst.$.uniqueCount
		@_.strings = []
		for si in xmlobj.sst.si
			@_.strings.push si.t[0]


	Object.defineProperties @prototype,
		"count":
			get: -> @_.count
		"uniqueCount":
			get: -> @_.uniqueCount

	get: (id)->
		@_.strings[id]

	add: (str)->
		idx = @_.strings.indexOf str
		if idx == -1
			@_.strings.push str
			idx = @_.uniqueCount++
		++@_.count
		idx

	toXmlObj: ->
		obj = @_.xmlobj.sst
		obj.$.count = @count
		obj.$.uniqueCount = @uniqueCount
		obj.si = []
		for str in @_.strings
			obj.si.push {
				t:[str]
			}
		@_.xmlobj


module.exports = SharedStringsTable
