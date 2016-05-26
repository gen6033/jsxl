Utils = require('../utils')

class Style
	constructor:(id, sm, xmlobj)->
		attr = xmlobj.$
		@_ = {}
		@id = id
		@sm = sm
		@_.numFmtId = parseInt attr.numFmtId
		@font = @sm.getFont(parseInt attr.fontId)
		@fill = @sm.getFill(parseInt attr.fillId)
		@border = @sm.getBorder(parseInt attr.borderId)
		@_.applyFont = !!(parseInt attr.applyFont)
		@_.applyFill = !!(parseInt attr.applyFill)
		@_.applyBorder = !!(parseInt attr.applyBorder)
		@_.applyAlignment = !!(parseInt attr.applyAlignment)



	clone: ->
		@sm.cloneResource(this)

	toXmlObj:->
		obj = {$:{
			numFmtId:@_.numFmtId
			fontId:@font.id
			fillId:@fill.id
			borderId:@border.id
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
