Promise = require('bluebird')
{parseStringAsync} = Promise.promisifyAll(require('xml2js'))
processors = require('xml2js/lib/processors')
Row = require('./row')
Column = require('./column')
Validation = require('./validation')
Utils = require('./utils')
{BlockRange} = require('./range')
extend = require('extend')
require('js-object-clone')


class Worksheet

	@create:(workbook, id, rid, name, xml)->
		parseStringAsync(xml, {tagNameProcessors: [processors.stripPrefix]})
		.then (result)->
			new Worksheet(workbook, id, rid, name, result)

	constructor:(workbook, id, rid, name, xmlobj)->
		@_ = {}
		@_.workbook = workbook
		@_.xmlobj = xmlobj
		ws = xmlobj.worksheet
		@_.id = parseInt id
		@_.rid = parseInt rid
		@_.name = name
		@_.dimension = new BlockRange(ws.dimension[0].$.ref)
		@defaultColWidth = (parseFloat ws.sheetFormatPr[0].$.defaultColWidth) || 12
		@defaultRowHeight = (parseFloat ws.sheetFormatPr[0].$.defaultRowHeight) || 20
		@headers = []

		extLst = ws.extLst?[0]
		if extLst
			for ext in extLst.ext
				delete ext.$
				for k,v of ext
					ws[k] = [] unless ws[k]
					ws[k] = ws[k].concat(v)
		delete ws.extLst

		@validations = []
		if ws.dataValidations
			for dataValidations in ws.dataValidations
				for v in dataValidations.dataValidation
					@validations.push new Validation(v)


		@_.mergedCells = []
		mergedCells = ws.mergeCells?[0].mergeCell
		if mergedCells
			for cell in mergedCells
				@_.mergedCells.push new BlockRange(cell.$.ref)

		@_.rows = []
		if ws.sheetData[0]
			for row_obj in ws.sheetData[0].row
				rowIdx = row_obj.$.r
				@_.rows[rowIdx] = new Row(this, rowIdx, row_obj)


		@_.cols = []
		if ws.cols?[0]
			@_.cols[1] = new Column(this, 1)
			max = 0
			for col_obj in ws.cols[0].col
				begin = parseInt col_obj.$.min
				max = parseInt col_obj.$.max
				@_.cols[begin] = new Column(this, begin, col_obj)

			if max != 16384
				@_.cols[max+1] = new Column(this, max+1)


	Object.defineProperties @prototype,
		"workbook":
			get: -> @_.workbook
		"id":
			get: -> @_.id
		"rid":
			get: -> @_.rid
		"name":
			get: -> @_.name
		"top":
			get: -> @_.dimension.top
		"left":
			get: -> @_.dimension.left
		"bottom":
			get: -> @_.dimension.bottom
		"right":
			get: -> @_.dimension.right


	getRow: (r)->
		throw new Error "Index is out of range" unless r > 0
		@_.rows[r] = @_.rows[r] || new Row(this, r)

	getColumn: (c)->
		if isNaN(c)
			return @getColumn (@headers.indexOf(c) + 1)
		throw new Error "Index is out of range" unless c > 0
		cols = @_.cols
		if cols[c]
			unless cols[c+1]
				cols[c+1] = Object.clone(cols[c])
				cols[c+1]._ = Object.clone(cols[c]._)
				cols[c+1]._.colIndex = c+1

			return cols[c]
		if cols.length == 0
			return @_.cols[c] = new Column(this, c)

		i = c
		if cols.length <= c
			i = cols.length - 1

		col = null
		while i >= 0
			col = cols[i]
			break if col
			--i

		cols[c] = Object.clone(col)
		cols[c]._ = Object.clone(col._)
		cols[c]._.colIndex = c
		unless cols[c+1]
			cols[c+1] = Object.clone(col)
			cols[c+1]._ = Object.clone(col._)
			cols[c+1]._.colIndex = c+1
		cols[c]

	getCell:(r, c)->
		@getRow(r).getCell(c)

	eachRow:(f)->
		for r in [@top..@bottom]
			f(@getRow(r))

	eachColumn:(f)->
		for c in [@left..@right]
			f(@getColumn(c))

	merge:(top, left, bottom, right)->
		range = new BlockRange(top, left, bottom, right)
		for r in @_.mergedCells
			throw new Error range.toString() + " conflicts with " + r.toString() if r.conflicts range
		@_.mergedCells.push range

	unmerge:(top, left, bottom, right)->
		range = new BlockRange(top, left, bottom, right)
		for r,idx in @_.mergedCells
 			@_.mergedCells[idx].splice(idx, 1) if r.conflicts range

	toXmlObj: ->
		obj = {
			sheetPr:[]
			dimension:[]
			sheetViews:[]
			sheetFormatPr:[]
			cols:[]
			sheetData:[]
			mergeCells:[]
			phoneticPr:[]
			dataValidations:[]
			hyperlinks:[]
			pageMargins:[]
			pageSetup:[]
		}
		obj = extend obj,@_.xmlobj.worksheet
		obj.dimension[0].$.ref = @_.dimension.toString()
		formatPtr = obj.sheetFormatPr[0].$
		formatPtr.defaultColWidth = @defaultColWidth
		formatPtr.defaultRowHeight = @defaultRowHeight


		if @_.cols.length > 0
			cols = []
			col = null
			for c,idx in @_.cols
				if c
					if col
						col.$.max = idx - 1
					col = c.toXmlObj()
					col.$.min = idx
					cols.push col

			col.$.max = 16384
			obj.cols = [{col:cols}]

		obj.sheetData[0].row = []
		for r in @_.rows
			continue unless r
			obj.sheetData[0].row.push r.toXmlObj()


		obj.dataValidations = []
		if @validations.length > 0
			dataValidations = {$:{count:@validations.length}, dataValidation:[]}
			for v in @validations
				dataValidations.dataValidation.push v.toXmlObj()
			obj.dataValidations = [dataValidations]


		mCells = []
		for r in @_.mergedCells
			mCells.push {$:{ref:r.toString()}}
		if mCells.length > 0
			obj.mergeCells = [{$:{count:@_.mergedCells.length}, mergeCell:mCells}]

		@_.xmlobj.worksheet = obj
		@_.xmlobj

module.exports = Worksheet
