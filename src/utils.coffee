moment = require("moment")
require("datejs") #Date.parseを上書きするので注意
BASE_TIME = moment("1899-12-30") * 1

module.exports = {
  toRowCol:(addr)->
    [tmp, col, row] = addr.match(/^\s*\$?([A-Z]+)\$?(\d+)\s*$/)
    [parseInt(row), @toDigit(col)]

  toAddr:(row, col, row_abs = false, col_abs = false)->
    (if col_abs then "$" else "")+@toAlphabet(col)+(if row_abs then "$" else "")+row

  toDigit: (alpha)->
    result = 0
    i = 0
    maxi = alpha.length
    while i < maxi
      result = result*26 + parseInt(alpha[i++], 36)-9
    result

  toAlphabet: (digit)->
    result = []
    i = 0
    while digit > 0
      result.unshift ((digit-1) % 26 + 10).toString(36).toUpperCase()
      digit = parseInt ((digit-1) / 26)
      ++i
    result.join("")

  offsetToDate: (offset)->
    new Date(BASE_TIME + (parseFloat(offset)*60*60*24*1000))

  dateToOffset: (date)->
    (date * 1 - BASE_TIME)/(60*60*24*1000)


  isString: (obj)->
    typeof (obj) == "string" || obj instanceof String

  isPureNumber: (x)->
    typeof(x) == 'number' || x instanceof Number

  isNumber: (x)->
    if typeof(x) != 'number' && typeof(x) != 'string'
      false
    else
      !isNaN(x - parseFloat(x))

  isBoolean: (obj)->
    typeof (obj) == "boolean" || obj instanceof Boolean || @isPureNumber(obj)

  isDate: (obj)->
    !@isNumber(obj) && (obj instanceof Date || @isDateString(obj))

  isDateString: (obj)->
    if @isString(obj)
      return Date.parse(obj) != null
    return false

  isISODateString: (obj)->
    if @isString(obj)
      return moment(obj, moment.ISO_8601).isValid()
    return false

  parseDate: (str)->
    #date.jsにより上書きされたDate.parse
    Date.parse(str)
}
