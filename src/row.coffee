Cell = require('./cell')
Utils = require('./utils')

class Row
	constructor:(worksheet, rowIndex, xmlobj)->
		@_ = {}
		@_.worksheet = worksheet
		@_.rowIndex = parseInt rowIndex
		@_.height = null
		@_.cells = []

		styleId = parseInt xmlobj?.$.s
		unless styleId
			styleId = 0
		@_.style = @workbook._.sm.getStyle(styleId)

		if xmlobj
			@_.heightCustomized = (parseInt xmlobj.$.customHeight) == 1
			@_.height = parseFloat xmlobj.$.ht

			unless xmlobj.c == undefined
				for cell_obj in xmlobj.c
					[tmp, col] = Utils.toRowCol cell_obj.$.r
					@_.cells[col] = new Cell(this, col, cell_obj)


	Object.defineProperties @prototype,
		"workbook":
			get: -> @worksheet.workbook
		"worksheet":
			get: -> @_.worksheet
		"rowIndex":
			get: -> @_.rowIndex
		"begin":
			get: -> @worksheet.left
		"end":
			get: -> @worksheet.right
		"height":
			get:->
				@_.height || @worksheet.defaultHeight
			set:(ht)->
				@_.heightCustomized = true
				@_.height = ht
		"style":
			get: ->
				s = @_.style.clone()
				@_.style = s
				s

	getCell: (c)->
		throw new Error "Index is out of range" unless c > 0
		@_.cells[c] || @_.cells[c] = new Cell(this, c)

	toXmlObj: ->
		obj = {
			$:{
				r:@_.rowIndex
				spans:@begin+":"+@end
				"x14ac:dyDescent":"0.3"
			}
			c:[]
		}
		obj.$.s = @_.style.id

		if @_.heightCustomized
			obj.$.ht = @_.height
			obj.$.customHeight = 1

		for c in @_.cells
			continue unless c
			obj.c.push c.toXmlObj()
		obj

module.exports = Row
