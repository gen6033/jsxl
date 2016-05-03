Utils = require('../utils')

class Font
	constructor:(id, sm, xmlobj)->
		@_ = {}
		@_.xmlobj = xmlobj
		@_.id = id
		@_.sm = sm
		@size = parseInt xmlobj.sz?[0].$.val
		@_.color = xmlobj.color?[0].$
		@name = xmlobj.name?[0].$.val
		@_.charset = parseInt xmlobj.charset?[0].$.val
		#@_.family = parseInt xmlobj.family?[0].$.val
		@_.scheme = xmlobj.scheme?[0].$.val

	Object.defineProperties @prototype,
		"id":
			get: -> @_.id
		"color":
			get: -> @_.color
		"charset":
			get: -> @_.charset
		"scheme":
			get: -> @_.scheme


	clone: ->
		@_.sm.cloneResource(this)


	toXmlObj: ->
		obj = {}

		if @size
			obj.sz = [{$:{val:@size}}]

		if @name
			obj.name = [{$:{val:@name}}]

		obj

module.exports = Font
