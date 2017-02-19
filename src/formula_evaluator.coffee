Utils = require("./utils")

class FormulaEvaluator
  constructor: (@worksheet)->



  SUM: (args)->
    sum = 0
    args = [].concat(args)
    for expr in args
      if expr instanceof Range
        expr.each (r, c)=>
          cell = @worksheet.getCell(r, c);
          val = cell.value
          if cell.formula
            val = @expectNumber(cell.calculate())
          if Utils.isNumber(val)
            sum += Number(val)
      else
        sum += @expectNumber(expr)
    sum


  getValue:(range)->
    return range unless range instanceof Range

    size = range.size
    if size == 1
      cell = null
      range.each (r, c)=>
        cell = @worksheet.getCell(r, c)
      if cell.formula
        cell.calculate()
      else
        cell.value
    else if size == 0
      throw new Error("#NULL!")
    else
      throw new Error("#VALUE!")

  expectNumber:(x)->
    x = @getValue(x)
    unless Utils.isNumber(x)
      throw new Error("#VALUE!")
    x

  expectString:(x)->
    x = @getValue(x)
    unless Utils.isString(x)
      throw new Error("#VALUE!")
    x

  expectRange:(x)->
    unless x instanceof Range
      throw new Error("#VALUE!")
    x

  expectList:(x)->
    unless Array.isArray(x)
      throw new Error("#VALUE!")
    x

  checkArgumentSize:(expr, size)->
    if Array.isArray(expr)
      if expr.length == size
        return;
    else if size == 1
      return;

    throw new Error("Sizes of arguments do not match.")


module.exports = FormulaEvaluator
