Utils = require('../utils')
require('js-object-clone')

class Font
  constructor:(id, sm, xmlobj)->
    @_ = {}
    @_.xmlobj = xmlobj
    @id = id
    @_.sm = sm
    @size = parseInt xmlobj.sz?[0].$.val
    colorAttrs = xmlobj.color?[0].$ || {}
    @color = sm.getRGB(colorAttrs)
    @name = xmlobj.name?[0].$.val
    #@_.charset = parseInt xmlobj.charset?[0].$.val
    #@_.family = parseInt xmlobj.family?[0].$.val
    #@_.scheme = xmlobj.scheme?[0].$.val


  clone: ->
    @_.sm.cloneResource(this)


  toXmlObj: ->
    obj = Object.clone(@_.xmlobj)

    if @size
      obj.sz = [{$:{val:@size}}]

    if @name
      obj.name = [{$:{val:@name}}]

    if @color
      if @color == "auto"
        obj.color = [{$:{auto:1}}]
      else
        obj.color = [{$:{rgb:@color}}]

    obj

module.exports = Font
