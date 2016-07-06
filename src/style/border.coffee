Utils = require('../utils')
require('js-object-clone')

class Border
  @THIN = "thin"
  @MEDIUM = "medium"
  @THICK = "thick"
  @DOUBLE = "double"
  @DOTTED = "dotted"
  @HAIR = "hair"
  @DASHED = "dashed"
  @MEDIUM_DASHED = "mediumDashed"
  @DASH_DOT = "dashDot"
  @MEDIUM_DASH_DOT = "mediumDashDot"
  @DASH_DOT_DOT = "dashDotDot"
  @MEDIUM_DASH_DOT_DOT = "mediumDashDotDot"
  @SLANT_DASH_DOT = "slantDashDot"
  @NONE = "none"

  constructor:(@id, @sm, @xmlobj)->
    @diagonalUp = !!(@xmlobj.$?.diagonalUp)
    @diagonalDown = !!(@xmlobj.$?.diagonalDown)

    for direction in ["left", "right", "top", "bottom", "diagonal"]
      border = @xmlobj[direction]
      @[direction+"Style"] = border.$?.style
      @[direction+"Color"] = @sm.getRGB(border.color?.$)

  clone: ->
    @sm.cloneResource(this)

  toXmlObj: ->
    obj = {}

    obj.$ = {}
    if @diagonalUp
      obj.$.diagonalUp = 1
    if @diagonalDown
      obj.$.diagonalDown = 1

    for direction in ["left", "right", "top", "bottom", "diagonal"]
      obj[direction] = [{$:{}}]
      if @[direction+"Style"]
        obj[direction][0].$.style = @[direction+"Style"]

      color = @[direction+"Color"]
      if color
        if color == "auto"
          obj[direction][0].color = [{$:{auto:1}}]
        else
          obj[direction][0].color = [{$:{rgb:color}}]

    obj

module.exports = Border
