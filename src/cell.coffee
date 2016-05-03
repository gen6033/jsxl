Utils = require('./utils')
{Range} = require('./range')

class Cell
	constructor:(@row, colIndex, xmlobj)->
		@_ = {}
		@_.rowIndex = @row.rowIndex
		@_.colIndex = colIndex
		@_.type = xmlobj?.$.t
		styleId = parseInt xmlobj?.$.s
		@_.value = xmlobj?.v?[0]
		@_.formula = xmlobj?.f?[0]

		unless styleId
			styleId = 0
		@_.style = @workbook._.sm.getStyle(styleId)

		if @_.value?._
			@_.value = @_.value._

		if @_.formula?.$
			attr = @_.formula.$
			@_.formula = @_.formula._
			if attr.t == "shared"
				if @_.formula
					range = new Range(attr.ref)
					@worksheet._.sft = {} unless @worksheet._.sft
					range.each (addr)=>
						@worksheet._.sft[addr] = @_.formula
				else
					@_.formula = @worksheet._.sft[@addr]


	Object.defineProperties @prototype,
		"workbook":
			get: -> @row.workbook

		"worksheet":
			get: -> @row.worksheet

		"value":
			get: ->
				switch @_.type
					when null
						@_.value
					else
						@workbook._.sst.get(@_.value)

			set: (val) ->
				switch typeof val
					when "number"
						@_.type = null
					else
						val = @workbook._.sst.add(val)
						@_.type = "s"

				@_.value = val

		"formula":
			get: -> @_.formula
			set: (val) ->
				@_.formula = val
				@_.type = null

		"rowIndex":
			get: -> @_.rowIndex

		"colIndex":
			get: -> @_.colIndex

		"addr":
			get: -> Utils.toAddr(@_.rowIndex, @_.colIndex)

		"style":
			get: ->
				s = @_.style.clone()
				@_.style = s
				s

	toXmlObj: ->

		obj = {
			$:{
				r:@addr
			}
		}

		if @_.formula
			obj.f = [@_.formula]

		value = @_.value
		if value
			if value == value.trim()
				obj.v = [value]
			else
				obj.v = [{$:{"xml:space":"preserve"}, _:value}]

		if @_.style
			obj.$.s = @_.style.id

		if @_.type
			obj.$.t = @_.type

		obj

module.exports = Cell
