Utils = require('./utils')
StyleMixin = require('./style_mixin')
Calculator = require('./calculator')

class Cell
  StyleMixin.mixin(@prototype)

  @BASE_TIME = new Date("1899/12/30") * 1

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
          @worksheet._.sft = {} unless @worksheet._.sft
          @worksheet._.sft[attr.si] = [@_.formula, @rowIndex, @colIndex]
        else
          [formula, rowIdx, colIdx] = @worksheet._.sft[attr.si]
          @_.formula = formula.replace /(".*?")|(\$?)([A-Z]+)(\$?)(\d+)/g, (m0, m1, m2, m3, m4, m5)=>
            return m1 if m1
            c = Utils.toDigit(m3)
            r = parseInt m5
            rr = @rowIndex
            cc = @colIndex
            unless m2
              c -= colIdx - cc
            unless m4
              r -= rowIdx - rr
            return Utils.toAddr(r, c, !!m4, !!m2)

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
        val = @_.value
        switch @_.type
          when "s"
            val = @workbook._.sst.get(@_.value)
        if @_.styleMixin.style.numberFormat.isDate()
          val = new Date(Cell.BASE_TIME + (parseFloat(val)*60*60*24*1000))
        val

      set: (val) ->
        if typeof val ==  "number"
          @_.type = null
        else if val instanceof Date
          @_.type = null
          val = (val - Cell.BASE_TIME)/(60*60*24*1000)
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

  calculate: ->
    Calculator(@worksheet, @formula)

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
