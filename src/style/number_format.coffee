Utils = require('../utils')

class NumberFormat

  @parse: (sm, xmlobj)->
    new NumberFormat(sm, xmlobj.$.numFmtId, xmlobj.$.formatCode)

  constructor:(@sm, @id, @format)->
    @_ = {}


  clone: ->
    @sm.cloneResource(this)

  isDate: ->
    @format.replace(/(\"[^\"]*\"|\[[^\]]*\]|[\\_*].)/gi, '').match(/[dmyhs]/i) != null


  toXmlObj: ->
    {$:{
      numFmtId:@id
      formatCode:@format
      }}

module.exports = NumberFormat
