Utils = require('../utils')

class Style
	constructor:(id, sm, xmlobj)->
		attr = xmlobj.$
		@_ = {}
		@_.id = id
		@_.sm = sm
		@_.numFmtId = parseInt attr.numFmtId
		@_.font = @_.sm.getFont(parseInt attr.fontId)
		@_.fillId = parseInt attr.fillId
		@_.borderId = parseInt attr.borderId
		@_.applyFont = !!(parseInt attr.applyFont)
		@_.applyFill = !!(parseInt attr.applyFill)
		@_.applyBorder = !!(parseInt attr.applyBorder)
		@_.applyAlignment = !!(parseInt attr.applyAlignment)


	Object.defineProperties @prototype,
		"id":
			get: ->
				@_.id
		"font":
			get: ->
				f = @_.font.clone()
				@_.font = f
				f


	clone: ->
		@_.sm.cloneResource(this)

	toXmlObj:->
		obj = {$:{
			numFmtId:@_.numFmtId
			fontId:@_.font.id
			fillId:@_.fillId
			borderId:@_.borderId
		}}

		if @_.applyFont
			obj.$.applyFont = 1

		if @_.applyFill
			obj.$.applyFill = 1

		if @_.applyBorder
			obj.$.applyBorder = 1

		if @_.applyAlignment
			obj.$.applyAlignment = 1

		obj


module.exports = Style
