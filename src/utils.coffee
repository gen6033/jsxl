BASE_TIME = new Date("1899/12/30") * 1

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

  isNumber: (x)->
    if typeof(x) != 'number' && typeof(x) != 'string'
      false
    else
      !isNaN(x - parseFloat(x))

  isBoolean: (obj)->
    typeof (obj) == "boolean" || obj instanceof Boolean

  isDate: (obj)->
    if @isNumber(obj)
      return false
    if obj instanceof Date
      return true
    if @isString(obj)
      obj = obj.trim()
      d1 = Date.parse(obj)
      d2 = Date.parse(obj.substr(1))
      if isNaN(d1)
        return false
      else
        # Date.parseは前後に余計な文字が含まれていても受理されるがエクセルでは受理されないため，
        # 1文字切り取った値と元の値が等しければ余計が文字が含まれているとしてfalseを返す
        return d1 != d2

    return false
}
