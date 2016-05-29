Utils = require('./utils')
{Range} = require('./range')
StyleMixin = require('./style_mixin')

class Cell
	StyleMixin.mixin(@prototype)

	constructor:(row, colIndex, xmlobj)->
		@_ = {}
		@_.row = row
		@_.rowIndex = @row.rowIndex
		@_.colIndex = colIndex
		@_.type = xmlobj?.$.t
		@_.value = xmlobj?.v?[0]
		@_.formula = xmlobj?.f?[0]

		styleId = (parseInt xmlobj?.$.s) || 0
		StyleMixin.bind(this, styleId)

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
		"row":
			get: -> @_.row
		"column":
			get: ->
				@_.col = @_.col || @worksheet.getColumn(@_.colIndex)
		"value":
			get: ->
				switch @_.type
					when "s"
						@workbook._.sst.get(@_.value)
					else
						@_.value

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
			value = value + ""
			if value == value.trim()
				obj.v = [value]
			else
				obj.v = [{$:{"xml:space":"preserve"}, _:value}]

		obj.$.s = @_.styleMixin.style.id

		if @_.type
			obj.$.t = @_.type

		obj

module.exports = Cell
