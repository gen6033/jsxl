Utils = require('../utils')
require('js-object-clone')

class Fill
  @DARK_DOWN = "darkDown"
  @DARK_GRAY = "darkGray"
  @DARK_GRID = "darkGrid"
  @DARK_HORIZONTAL = "darkHorizontal"
  @DARK_TRELLIS = "darkTrellis"
  @DARK_UP = "darkUp"
  @DARK_VERTICAL = "darkVertical"
  @GRAY_0625 = "gray0625"
  @GRAY_125 = "gray215"
  @LIGHT_DOWN = "lightDown"
  @LIGHT_GRAY = "lightGray"
  @LIGHT_GRID = "lightGrid"
  @LIGHT_HORIZONTAL = "lightHorizontal"
  @LIGHT_TRELLIS = "lightTrellis"
  @LIGHT_UP = "lightUp"
  @LIGHT_VERTICAL = "lightVertical"
  @NONE = "none"
  @SOLID = "solid"

  constructor:(id, @sm, @xmlobj)->
    @_ = {}
    @id = id
    fill = @xmlobj.patternFill[0]
    @type = fill.$.patternType
    fgColor = fill.fgColor?[0]
    if fgColor
      fgColor = @sm.getRGB(fgColor.$)
    @_.fgColor = fgColor

    bgColor = fill.bgColor?[0]
    if bgColor
      bgColor = @sm.getRGB(bgColor.$)
    @_.bgColor = bgColor

  Object.defineProperties @prototype,
    "fgColor":
      get: -> @_.fgColor
      set: (v)->
        @_.fgColor = v
        if @type == Fill.NONE
          @type = Fill.SOLID
    "bgColor":
      get: -> @_.bgColor
      set: (v)->
        @_.bgColor = v
        if @type == Fill.NONE
          @type = Fill.SOLID

  clone: ->
    @sm.cloneResource(this)


  toXmlObj: ->
    obj = {patternFill:[{$:{}}]}
    fill = obj.patternFill[0]
    fill.$.patternType = @type

    if @_.fgColor
      if @_.fgColor == "auto"
        fill.fgColor = [{$:{auto:1}}]
      else
        fill.fgColor = [{$:{rgb:@_.fgColor}}]

    if @_.bgColor
      if @_.bgColor == "auto"
        fill.bgColor = [{$:{auto:1}}]
      else
        fill.bgColor = [{$:{rgb:@_.bgColor}}]

    obj

module.exports = Fill
