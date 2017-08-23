Promise = require('bluebird')
xml2js = Promise.promisifyAll(require('xml2js'))
Utils = require('./utils')
parser = new xml2js.Parser()

class SharedStringsTable

  @parse:(xml)->
    return new Promise (resolve, reject)->
      unless xml
        resolve(new SharedStringsTable)
        return
      parser.parseString xml,(err, result)->
        resolve(new SharedStringsTable(result))


  constructor:(xmlobj)->
    @_ = {}
    @_.count = 0
    @_.uniqueCount = 0
    @_.strings = []
    if xmlobj && xmlobj.sst && xmlobj.sst.si
      @_.count = xmlobj.sst.$.count
      @_.uniqueCount = xmlobj.sst.$.uniqueCount
      for si in xmlobj.sst.si
        if si.t
          str = si.t[0]
          if str._
            str = str._
        else
          str = ""
          for r in si.r
            s = r.t[0]
            if s._
              s = s._
            str += s
        str += ""
        @_.strings.push str.replace(/\r/g, "")


  Object.defineProperties @prototype,
    "count":
      get: -> @_.count
    "uniqueCount":
      get: -> @_.uniqueCount

  get: (id)->
    @_.strings[id]

  add: (str)->
    idx = @_.strings.indexOf str
    if idx == -1
      @_.strings.push str
      idx = @_.uniqueCount++
    ++@_.count
    idx

  toXmlObj: ->
    obj = {$:{xmlns:"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}}
    obj.$.count = @count
    obj.$.uniqueCount = @uniqueCount
    obj.si = []
    for str in @_.strings
      obj.si.push {
        t:[str]
      }

    {sst:obj}


module.exports = SharedStringsTable
