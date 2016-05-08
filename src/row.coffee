Cell = require('./cell')
Utils = require('./utils')

class Row
	constructor:(worksheet, rowIndex, xmlobj)->
		@_ = {}
		@_.worksheet = worksheet
		@_.rowIndex = parseInt rowIndex
		@_.begin = @_.end = 0
		@_.height = null
		@_.cells = []
		@_.styleId = 0

		if xmlobj
			[begin, end] = xmlobj.$.spans.split(":")
			@_.styleId = parseInt xmlobj.$.s
			@_.begin = parseInt begin
			@_.end = parseInt end
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

	getCell: (c)->
		throw new Error "Index is out of range" unless c > 0
		unless @_.cells[c]
			@_.cells[c] = new Cell(this, c)
			if c < @_.begin
				@_.begin = c
				@_.end = @_.begin if @_.begin > @_.end
			else if c > @_.end
				@_.end = c
		@_.cells[c]

	toXmlObj: ->
		obj = {
			$:{
				r:@_.rowIndex
				spans:@_.begin+":"+@_.end
				"x14ac:dyDescent":"0.3"
			}
			c:[]
		}
		obj.$.s = @_.styleId if @_.styleId

		if @_.heightCustomized
			obj.$.ht = @_.height
			obj.$.customHeight = 1

		for c in @_.cells
			continue unless c
			obj.c.push c.toXmlObj()
		obj

module.exports = Row
