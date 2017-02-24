Utils = require("./utils")

ERROR_DIV0 = "#DIV/0!"
ERROR_NA = "#N/A!"
ERROR_NAME = "#NAME!"
ERROR_NUM = "#NUM!"
ERROR_NULL = "#NULL!"
ERROR_VALUE = "#VALUE!"

TYPE_NUMBER = 1
TYPE_INTEGER = 2
TYPE_STRING = 3



class FormulaEvaluator

  constructor: (@worksheet)->


  ABS: (args)->
    @checkArgumentSize(args, 1)
    Math.abs @expectNumber(args[0])

  ACOS: (args)->
    @checkArgumentSize(args, 1)
    Math.acos @expectNumber(args[0])

  ACOSH: (args)->
    @checkArgumentSize(args, 1)
    Math.acosh @expectNumber(args[0])

  ACOT: (args)->
    @checkArgumentSize(args, 1)
    Math.atan (1/@expectNumber(args[0]))

  ACOTH: (args)->
    @checkArgumentSize(args, 1)
    Math.atanh (1/@expectNumber(args[0]))

  ARABIC: (args)->
    @checkArgumentSize(args, 1)
    list = @expectString(args[0]).toUpperCase().trim().split("").reverse().map (c)->
      switch c
        when 'M'
          1000
        when 'D'
          500
        when 'C'
          100
        when 'L'
          50
        when 'X'
          10
        when 'V'
          5
        when 'I'
          1
    sum = 0
    max = 0
    for n in list
      if n >= max
        sum += n
        max = n
      else
        sum -= n
    sum

  ASIN: (args)->
    @checkArgumentSize(args, 1)
    Math.asin @expectNumber(args[0])

  ASINH: (args)->
    @checkArgumentSize(args, 1)
    Math.asinh @expectNumber(args[0])

  ATAN: (args)->
    @checkArgumentSize(args, 1)
    Math.atan @expectNumber(args[0])

  ATAN2: (args)->
    @checkArgumentSize(args, 1)
    Math.atan2 @expectNumber(args[0])

  ATANH: (args)->
    @checkArgumentSize(args, 1)
    Math.atanh @expectNumber(args[0])

  BASE: (args)->
    @checkArgumentSize(args, 2, 3)
    [n, base, len] = args
    res = @expectInteger(n).toString(@expectInteger(base))
    if len
      len = @expectInteger(len)
      res = "0".repeat(len - res.length) + res
    res

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
    @checkArgumentSize(args, 1, 2)
    if args.length == 1
      args.push 1
    args[1] = Math.abs(@expectNumber(args[1]))
    @CEILING(args)


  "CEILING.MATH": (args)->
    @checkArgumentSize(args, 1, 3)
    if args.length == 1
      args.push 1
    expr = @expectNumber(args[0])
    digit = Math.abs(@expectNumber(args[1]))
    sign = 1
    if digit == 0
      return 0
    if args.length == 3
      sign = Math.sign(expr)
      expr = Math.abs(expr)
    sign * @CEILING([expr, digit])

  COMBIN: (args)->
    @checkArgumentSize(args, 2)
    @PERMUT(args) / @FACT([args[1]])

  COMBINA: (args)->
    @checkArgumentSize(args, 2)
    n = @expectInteger(args[0])
    r = @expectInteger(args[1])
    @COMBIN([n+r-1, r]);

  COS: (args)->
    @checkArgumentSize(args, 1)
    Math.cos @expectNumber(args[0])

  COSH: (args)->
    @checkArgumentSize(args, 1)
    Math.cosh @expectNumber(args[0])

  COT: (args)->
    @checkArgumentSize(args, 1)
    1/ Math.tan(@expectNumber(args[0]))

  COTH: (args)->
    @checkArgumentSize(args, 1)
    1/ Math.tanh(@expectNumber(args[0]))

  CSC: (args)->
    @checkArgumentSize(args, 1)
    1/ Math.sin(@expectNumber(args[0]))

  CSCH: (args)->
    @checkArgumentSize(args, 1)
    1/ Math.sinh(@expectNumber(args[0]))

  DECIMAL: (args)->
    @checkArgumentSize(args, 2)
    num = @expectString(args[0])
    base = @expectInteger(args[1])
    parseInt(num, base)

  DEGREES: (args)->
    @checkArgumentSize(args, 1)
    @expectNumber(args[0]) * 180 / Math.PI

  EVEN: (args)->
    @checkArgumentSize(args, 1)
    args.push 2
    MODE = -1
    args.push MODE
    @["CEILING.MATH"](args)


  EXP: (args)->
    @checkArgumentSize(args, 1)
    num = @expectNumber(args[0])
    Math.exp(num)

  FACT: (args)->
    @checkArgumentSize(args, 1)
    n = @expectInteger(args[0])
    fact = (x)->
      if x <= 1
        1
      else
        x * fact(x-1)
    fact(n)

  FACTDOUBLE: (args)->
    @checkArgumentSize(args, 1)
    n = @expectInteger(args[0])
    fact2 = (x)->
      if x <= 1
        1
      else
        x * fact2(x-2)
    fact2(n)


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
    @checkArgumentSize(args, 1, 2)
    if args.length == 1
      args.push 1
    args[1] = Math.abs(@expectNumber(args[1]))
    if args[1] == 0
      return 0
    @FLOOR(args)


  "FLOOR.MATH": (args)->
    @checkArgumentSize(args, 1, 3)
    if args.length == 1
      args.push 1
    expr = @expectNumber(args[0])
    digit = Math.abs(@expectNumber(args[1]))
    sign = 1
    if digit == 0
      return 0
    if args.length == 3
      sign = Math.sign(expr)
      expr = Math.abs(expr)
    sign * @FLOOR([expr, digit])


  GCD: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_INTEGER, args)
    g = args[0]
    for i in [1...args.length]
      n = args[i]
      if n < 0
        @error(ERROR_NUM)
      g = @__gcd__(n, g)
    g

  IF: (args)->
    @checkArgumentSize(args, 3)
    [cond, true_expr, false_expr] = args
    cond = @getValue(cond)
    if Utils.isNumber(cond)
      cond = cond == 1
    else if !Utils.isBoolean(cond)
      @error(ERROR_VALUE)

    if cond
      true_expr
    else
      false_expr

  INT: (args)->
    @checkArgumentSize(args, 1)
    args.push 1
    @FLOOR(args)

  "ISO.CEILING":@::["CEILING.PRECISE"]

  LCM: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_INTEGER, args)
    l = args[0]
    for i in [1...args.length]
      n = args[i]
      if n < 0
        @error(ERROR_NUM)
      g = @__gcd__(n, l)
      l *= n
      l /= g
    l

  LN: (args)->
    @checkArgumentSize(args, 1)
    Math.log @expectNumber(args[0])

  LOG: (args)->
    @checkArgumentSize(args, 1, 2)
    if args.length == 1
      args.push 10
    Math.log(@expectNumber(args[0])) / Math.log(@expectNumber(args[1]))

  LOG10: (args)->
    @checkArgumentSize(args, 1)
    Math.log10 @expectNumber(args[0])

  ODD: (args)->
    @checkArgumentSize(args, 1)
    num = @expectNumber(args[0])
    if num >= 0
      @EVEN([num+1])-1
    else
      @EVEN([num-1])+1

  PERMUT: (args)->
    @checkArgumentSize(args, 2)
    n = @expectInteger(args[0])
    r = @expectInteger(args[1])
    if n <= 0 || r < 0 || n < r
      @error(ERROR_NUM)

    f = (nn, rr)->
      if rr == 0
        1
      else
        nn * f(nn-1, rr-1)
    f(n, r)

  PERMUTATIONA: (args)->
    @checkArgumentSize(args, 2)
    n = @expectInteger(args[0])
    r = @expectInteger(args[1])
    Math.pow(n, r)


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


  SUM: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_NUMBER, args)
    sum = 0
    for num in args
      sum += num
    sum


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
    if x == undefined
      return Number(0)
    unless Utils.isNumber(x)
      @error(ERROR_VALUE)
    Number(x)

  expectInteger:(x)->
    parseInt @expectNumber(x)

  expectString:(x)->
    x = @getValue(x)
    if x == undefined
      return ""
    unless Utils.isString(x)
      @error(ERROR_VALUE)
    x


  expect:(type, x)->
    switch type
      when TYPE_NUMBER
        @expectNumber(x)
      when TYPE_INTEGER
        @expectInteger(x)
      when TYPE_STRING
        @expectString(x)


  expandRange:(type, args)->
    expanded = []
    for arg in args
      if arg instanceof Range
        arg.each (r, c)=>
          cell = @worksheet.getCell(r, c);
          val = cell.value
          if cell.formula
            val = cell.calculate()
          expanded.push @expect(type, val)
      else
        expanded.push @expect(type, arg)
    expanded

  checkArgumentSize:(args, min_size, max_size=min_size)->
    if min_size <= args.length <= max_size
        return
    throw new Error("Sizes of arguments do not match.")

  __gcd__: (x, y)->
    return x if y == 0
    @__gcd__(y, x % y)

module.exports = FormulaEvaluator
