Utils = require("./utils")

ERROR_DIV0 = "#DIV/0!"
ERROR_NA = "#N/A!"
ERROR_NAME = "#NAME!"
ERROR_NUM = "#NUM!"
ERROR_NULL = "#NULL!"
ERROR_VALUE = "#VALUE!"

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


  ROUND: (args)->
    @checkArgumentSize(args, 2)
    expr = @expectNumber(args[0])
    sign = Math.sign(expr)
    expr = Math.abs(expr)
    n = Math.floor(@expectNumber(args[1]))
    digit = Math.pow(10, n)
    sign * Math.round(expr*digit) / digit


  ROUNDUP: (args)->
    @checkArgumentSize(args, 2)
    expr = @expectNumber(args[0])
    sign = Math.sign(expr)
    expr = Math.abs(expr)
    n = Math.floor(@expectNumber(args[1]))
    digit = Math.pow(10, n)
    sign * @CEILING([expr, digit])

  ROUNDDOWN: (args)->
    @checkArgumentSize(args, 2)
    expr = @expectNumber(args[0])
    sign = Math.sign(expr)
    expr = Math.abs(expr)
    n = Math.floor(@expectNumber(args[1]))
    digit = Math.pow(10, n)
    sign * @FLOOR([expr, digit])

  CEILING: (args)->
    @checkArgumentSize(args, 2)
    expr = @expectNumber(args[0])
    digit = @expectNumber(args[1])
    if digit == 0
      return 0
    if expr >= 0 && digit < 0
      @error(ERROR_NUM)
    Math.ceil(expr/digit) * digit

  "CEILING.PRECISE": (args)->
    @checkArgumentSize(args, 2)
    args[1] = Math.abs(@expectNumber(args[1]))
    @CEILING(args)


  "CEILING.MATH": (args)->
    @checkArgumentSize(args, 2, 3)
    expr = @expectNumber(args[0])
    digit = Math.abs(@expectNumber(args[1]))
    sign = 1
    if digit == 0
      return 0
    if args.length == 3
      sign = Math.sign(expr)
      expr = Math.abs(expr)
    sign * @CEILING([expr, digit])



  FLOOR: (args)->
    @checkArgumentSize(args, 2)
    expr = @expectNumber(args[0])
    digit = @expectNumber(args[1])
    if digit == 0
      @error(ERROR_DIV0)
    if expr >= 0 && digit < 0
      @error(ERROR_NUM)
    Math.floor(expr/digit) * digit

  "FLOOR.PRECISE": (args)->
    @checkArgumentSize(args, 2)
    args[1] = Math.abs(@expectNumber(args[1]))
    if args[1] == 0
      return 0
    @FLOOR(args)


  "FLOOR.MATH": (args)->
    @checkArgumentSize(args, 2, 3)
    expr = @expectNumber(args[0])
    digit = Math.abs(@expectNumber(args[1]))
    sign = 1
    if digit == 0
      return 0
    if args.length == 3
      sign = Math.sign(expr)
      expr = Math.abs(expr)
    sign * @FLOOR([expr, digit])


  TRUNC: (args)->
    @checkArgumentSize(args, 1, 2)
    if args.length == 1
      args.push 0

    @ROUNDDOWN(args)




  error:(err)->
    throw new Error(err)


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
      @error(ERROR_NULL)
    else
      @error(ERROR_VALUE)

  expectNumber:(x)->
    x = @getValue(x)
    unless Utils.isNumber(x)
      @error(ERROR_VALUE)
    Number(x)

  expectString:(x)->
    x = @getValue(x)
    unless Utils.isString(x)
      @error(ERROR_VALUE)
    x

  expectRange:(x)->
    unless x instanceof Range
      @error(ERROR_VALUE)
    x

  expectList:(x)->
    unless Array.isArray(x)
      @error(ERROR_VALUE)
    x

  checkArgumentSize:(args, min_size, max_size=min_size)->
    if min_size <= args.length <= max_size
        return
    throw new Error("Sizes of arguments do not match.")


module.exports = FormulaEvaluator
