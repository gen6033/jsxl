Cell = require('./cell')
Utils = require('./utils')
StyleMixin = require('./style_mixin')

class Row
  StyleMixin.mixin(@prototype)

  constructor:(worksheet, rowIndex, xmlobj)->
    @_ = {}
    @_.worksheet = worksheet
    @_.rowIndex = parseInt rowIndex
    @_.height = null
    @_.cells = []



    if xmlobj
      @_.heightCustomized = (parseInt xmlobj.$.customHeight) == 1
      @_.height = parseFloat xmlobj.$.ht

      unless xmlobj.c == undefined
        for cell_obj in xmlobj.c
          [tmp, col] = Utils.toRowCol cell_obj.$.r
          @_.cells[col] = new Cell(this, col, cell_obj)

    styleId = (parseInt xmlobj?.$.s) || 0
    StyleMixin.bind(this, styleId)
    @_.styleMixin.onUpdate = =>
      for cell in @_.cells
        continue unless cell
        cell._.styleMixin.style = @_.styleMixin.style

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
    if isNaN(c)
      return @getCell (@worksheet.headers.indexOf(c) + 1)
    throw new Error "Index is out of range" unless c > 0
    @_.cells[c] || @_.cells[c] = new Cell(this, c)

  eachCell: (f)->
    for idx in [@_.worksheet.left..@_.worksheet.right]
      f(@getCell(idx))

  toXmlObj: ->
    obj = {
      $:{
        r:@_.rowIndex
        spans:@begin+":"+@end
        "x14ac:dyDescent":"0.3"
      }
      c:[]
    }
    styleId = @_.styleMixin.style.id
    obj.$.s = styleId
    obj.$.customFormat = 1 if styleId != 0

    if @_.heightCustomized
      obj.$.ht = @_.height
      obj.$.customHeight = 1

    for c in @_.cells
      continue unless c
      obj.c.push c.toXmlObj()
    obj

module.exports = Row
