Utils = require('../utils')

class Style
  constructor:(id, sm, xmlobj)->
    attr = xmlobj.$
    @id = id
    @sm = sm
    @aaaa = attr.numFmtId
    @numberFormat = @sm.getNumberFormat(parseInt attr.numFmtId)
    @font = @sm.getFont(parseInt attr.fontId)
    @fill = @sm.getFill(parseInt attr.fillId)
    @border = @sm.getBorder(parseInt attr.borderId)
    @applyFont = !!(parseInt attr.applyFont)
    @applyFill = !!(parseInt attr.applyFill)
    @applyBorder = !!(parseInt attr.applyBorder)
    @applyAlignment = !!(parseInt attr.applyAlignment)



  clone: ->
    @sm.cloneResource(this)

  toXmlObj:->
    obj = {$:{
      numFmtId:@numberFormat?.id || 0
      fontId:@font.id
      fillId:@fill.id
      borderId:@border.id
    }}

    if @applyFont
      obj.$.applyFont = 1

    if @applyFill
      obj.$.applyFill = 1

    if @applyBorder
      obj.$.applyBorder = 1

    if @applyAlignment
      obj.$.applyAlignment = 1

    obj


module.exports = Style
