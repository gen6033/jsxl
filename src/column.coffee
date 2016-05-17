StyleMixin = require('./style_mixin')

class Column
	StyleMixin.mixin(@prototype)

	constructor:(worksheet, colIndex, xmlobj)->
		@_ = {}
		attr = xmlobj?.$
		@_.worksheet = worksheet
		@_.colIndex = parseInt colIndex
		@width = @worksheet.defaultColWidth
		@hidden = false
		@bestFit = false
		@collapsed = false

		styleId = (parseInt attr?.style) || 0
		StyleMixin.bind(this, styleId)
		@_.styleMixin.onUpdate = =>
			@eachCell (cell)=>
				cell._.styleMixin.style = @_.styleMixin.style

		if attr
			@width = attr.width
			@hidden = !!(parseInt attr.hidden)
			@bestFit = !!(parseInt attr.bestFit)
			@collapsed = !!(parseInt attr.collapsed)

	getCell: (r)->
		@worksheet.getCell(r, @colIndex)

	eachCell: (f)->
		for idx in [@_.worksheet.top..@_.worksheet.bottom]
			f(@getCell(idx))

	Object.defineProperties @prototype,
		"workbook":
			get: -> @worksheet.workbook
		"worksheet":
			get: -> @_.worksheet
		"colIndex":
			get: -> @_.colIndex

	toXmlObj: ->
		obj = {$:{}}
		attr = obj.$

		attr.min = 0
		attr.max = 0
		attr.width = @width
		attr.style = @_.styleMixin.style.id
		attr.hidden = 1 if @hidden
		attr.bestFit = 1 if @bestFit
		attr.collapsed = 1 if @collapsed
		attr.customWidth = 1 if @width != @worksheet.defaultColWidth

		obj

module.exports = Column
