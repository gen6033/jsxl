
Utils = require("../utils")
math = require("mathjs")
moment = require("moment")
moji = require("moji")
multibyteLength = require('multibyte-length')
mbsubstr = require('mb-substr')
XRegExp = require("xregexp")
FormulaError = require("./error")
require('moment-weekday-calc')

TYPE_NUMBER = 1
TYPE_INTEGER = 2
TYPE_STRING = 3
TYPE_DATE = 4
TYPE_DATE_STRING = 5
TYPE_BOOLEAN = 6
TYPE_PURE_NUMBER = 7
TYPE_PURE_INTEGER = 8
TYPE_ANY = 100

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

  AND: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    res = true
    flag = false
    for arg in args
      if arg instanceof Range
        for x in @expandRange(TYPE_ANY, arg)
          if Utils.isBoolean(x)
            res &&= !!x
            flag = true
          else if x instanceof FormulaError
            throw x

      else
        x = @getValue(arg)
        if x instanceof FormulaError
          throw x
        unless Utils.isBoolean(x)
          @error(FormulaError.VALUE)
        res &&= !!x
        flag = true
    unless flag
      @error(FormulaError.VALUE)
    res

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

  ASC: (args)->
    @checkArgumentSize(args, 1)
    text = @expectString(args[0])
    moji(text).convert("ZE", "HE").convert("ZS", "HS").convert("ZK", "HK").toString()

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

  AVERAGE: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_ANY, args)
    sum = 0
    size = 0
    for arg in args
      if Utils.isPureNumber(arg) || Utils.isBoolean(arg)
        sum += Number(arg)
        size++
      else if arg instanceof FormulaError
        throw arg

    if size == 0
      @error(FormulaError.DIV0)
    sum / size

  AVEDEV: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    avg = @AVERAGE(args)
    args = @expandRange(TYPE_ANY, args)
    sum = 0
    size = 0
    for arg in args
      if Utils.isPureNumber(arg) || Utils.isBoolean(arg)
        sum += Math.abs(avg - Number(arg))
        size++
      else if arg instanceof FormulaError
        throw arg

    if size == 0
      @error(FormulaError.DIV0)
    sum / size

  AVERAGEA: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_ANY, args)
    sum = 0
    size = 0
    for arg in args
      if Utils.isPureNumber(arg) || Utils.isBoolean(arg)
        sum += Number(arg)
      else if arg instanceof FormulaError
        throw arg
      if arg != undefined
        size++

    if size == 0
      @error(FormulaError.DIV0)
    sum / size

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
      @error(FormulaError.NUM)
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


  CHAR: (args)->
    @checkArgumentSize(args, 1)
    code = @expectInteger(args[0])
    unless 1 <= code <= 255
      @error(FormulaError.VALUE)
    String.fromCharCode(code)

  CLEAN: (args)->
    @checkArgumentSize(args, 1)
    text = @expectString(args[0])
    text.replace(/[\x00-\x1F\x7F\x81\x8D\x8E\x9D]/g, "")

  CODE: (args)->
    @checkArgumentSize(args, 1)
    text = @expectString(args[0])
    text.charCodeAt(0)

  COMBIN: (args)->
    @checkArgumentSize(args, 2)
    math.combinations(@expectInteger(args[0]), @expectInteger(args[1]))

  COMBINA: (args)->
    @checkArgumentSize(args, 2)
    n = @expectInteger(args[0])
    r = @expectInteger(args[1])
    math.combinations(n+r-1, r)


  CONCAT: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_STRING, args)
    args.join("")

  CONCATENATE: (args)->
    @CONCAT(args)


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

  DATE: (args)->
    @checkArgumentSize(args, 3)
    year = @expectInteger(args[0])
    month = @expectInteger(args[1])
    day = @expectInteger(args[2])
    new Date(year, month-1, day)

  DATEDIF: (args)->
    @checkArgumentSize(args, 3)
    start = @expectMoment(args[0]).startOf("day")
    end = @expectMoment(args[1]).startOf("day")
    unit = @expectString(args[2]).toUpperCase()
    if start > end
      @error(FormulaError.NUM)
    switch unit
      when "Y"
        end.diff(start, "years")
      when "M"
        end.diff(start, "months")
      when "D"
        end.diff(start, "days")
      when "MD"
        start.year(1)
        end.year(1)
        if start.date() > end.date()
          start.month(end.month()-1)
        else
          start.month(end.month())
        end.diff(start, "days")
      when "YM"
        start.year(end.year()-1)
        end.diff(start, "months") % 12
      when "YD"
        end2 = end.clone()
        end2.year(start.year())
        if end2 >= start
          end2.diff(start, "days")
        else
          end3 = end.clone()
          end3.year(start.year() + 1)
          end3.diff(start, "days")

  DATEVALUE: (args)->
    @checkArgumentSize(args, 1)
    str = @expectDateString(args[0])
    Math.floor Utils.dateToOffset(Utils.parseDate(str))

  DAY: (args)->
    @checkArgumentSize(args, 1)
    @expectDate(args[0]).getDate()

  DAYS: (args)->
    @checkArgumentSize(args, 2)
    end = @expectMoment(args[0])
    start = @expectMoment(args[1])
    end.diff(start, "days")

  DAYS360: (args)->
    @checkArgumentSize(args, 2, 3)
    if args.length == 2
      args.push false
    start = @expectMoment(args[0])
    end = @expectMoment(args[1])
    flag = @expectBoolean(args[2])


    start_y = start.year()
    end_y = end.year()
    start_m = start.month()
    end_m = end.month()
    start_d = start.date()
    end_d = end.date()

    isLastDayOfMonth = (dt)->
      dt.date() == dt.endOf("month").date()

    if flag
      start_d = Math.min start_d, 30
      end_d = Math.min end_d, 30
    else
      start_is_last = isLastDayOfMonth(start)
      end_is_last = isLastDayOfMonth(end)
      #エクセルのバグで以下の処理は実行されない(NASD準拠ではなくPSA準拠になっている)
      #if start_is_last && end_is_last && end_m == 1#2月
      #  end_d = 30

      if start_is_last && start_m == 1 #2月
        start_d = 30
      if start_d >= 30 && end_d == 31
        end_d = 30
      if start_d == 31
        start_d = 30

    days = (end_y - start_y) * 360
    days += (end_m - start_m) * 30
    days += end_d - start_d
    days

  DBCS: (args)->
    @checkArgumentSize(args, 1)
    text = @expectString(args[0])
    moji(text).convert("HE", "ZE").convert("HS", "ZS").convert("HK", "ZK").toString()

  DECIMAL: (args)->
    @checkArgumentSize(args, 2)
    num = @expectString(args[0])
    base = @expectInteger(args[1])
    parseInt(num, base)

  DEGREES: (args)->
    @checkArgumentSize(args, 1)
    @expectNumber(args[0]) * 180 / Math.PI

  EDATE: (args)->
    @checkArgumentSize(args, 2)
    date = @expectMoment(args[0])
    month = @expectInteger(args[1])
    date.add(month, "months").startOf("day").toDate()

  EOMONTH: (args)->
    @checkArgumentSize(args, 2)
    date = @expectMoment(args[0])
    month = @expectInteger(args[1])
    date.add(month, "months").endOf("month").startOf("day").toDate()

  EVEN: (args)->
    @checkArgumentSize(args, 1)
    args.push 2
    MODE = -1
    args.push MODE
    @["CEILING.MATH"](args)

  EXACT: (args)->
    @checkArgumentSize(args, 2)
    a = @expectString(args[0])
    b = @expectString(args[1])
    a == b

  EXP: (args)->
    @checkArgumentSize(args, 1)
    num = @expectNumber(args[0])
    Math.exp(num)

  FACT: (args)->
    @checkArgumentSize(args, 1)
    math.factorial(@expectInteger(args[0]))

  FACTDOUBLE: (args)->
    @checkArgumentSize(args, 1)
    n = @expectInteger(args[0])
    fact2 = (x)->
      if x <= 1
        1
      else
        x * fact2(x-2)
    fact2(n)

  FALSE: (args)->
    @checkArgumentSize(args, 0)
    false


  FIND: (args)->
    @checkArgumentSize(args, 2, 3)
    search = @expectString(args[0])
    target = @expectString(args[1])
    start = 1
    if args.length == 3
      start = @expectInteger(args[2])
    start--

    pos = target.substr(start).search(search)
    if pos == -1
      @error(FormulaError.VALUE)
    start + pos + 1

  FINDB: (args)->
    @checkArgumentSize(args, 2, 3)
    search = @expectString(args[0])
    target = @expectString(args[1])
    start = 1
    if args.length == 3
      start = @expectInteger(args[2])
    #検索対象でない部分を切り出す
    prefix = mbsubstr(target, 0 , start)
    start = multibyteLength(prefix)
    #検索対象でない部分を取り除く
    target = target.replace(prefix, "")
    pos = target.search(search)
    if pos == -1
      @error(FormulaError.VALUE)
    s = target.substr(0, pos)
    start + multibyteLength(s) + 1

  FIXED: (args)->
    @checkArgumentSize(args, 1, 3)
    num = @expectNumber(args[0])
    digit = 2
    if args.length >= 2
      digit = @expectInteger(args[1])
    delim = false
    if args.length == 3
      delim = @expectBoolean(args[2])

    num = @ROUND([num, digit])
    if digit > 0
      num = num.toFixed(digit)
    else
      num = String(num)

    if delim
      return num
    else
      parts = num.split(".")
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",")
      return parts.join(".")

  FLOOR: (args)->
    @checkArgumentSize(args, 2)
    expr = @expectNumber(args[0])
    digit = @expectNumber(args[1])
    if digit == 0
      @error(FormulaError.DIV0)
    if expr >= 0 && digit < 0
      @error(FormulaError.NUM)
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
    math.gcd(args...)

  HOUR: (args)->
    @checkArgumentSize(args, 1)
    @expectDate(args[0]).getHours()

  ISOWEEKNUM: (args)->
    @checkArgumentSize(args, 1)
    @expectMoment(args[0]).isoWeek()

  IF: (args)->
    @checkArgumentSize(args, 3)
    [cond, true_expr, false_expr] = args
    cond = @getValue(cond)
    unless Utils.isBoolean(cond)
      @error(FormulaError.VALUE)
    #数値の場合があるのでboolにキャスト
    cond = !!cond
    if cond
      true_expr
    else
      false_expr

  IFERROR: (args)->
    @checkArgumentSize(args, 2)
    x = @getValue(args[0])
    if x instanceof FormulaError
      return @getValue(args[1])
    x

  IFNA: (args)->
    @checkArgumentSize(args, 2)
    x = @getValue(args[0])
    if x == FormulaError.NA
      return @getValue(args[1])
    x

  IFS: (args)->
    @checkArgumentSize(args, 2, Number.MAX_VALUE)
    args = @expandRange(TYPE_ANY, args)
    if args.length % 2 == 1
      @error(FormulaError.NA)

    for x,i in args when i % 2 == 0
      if @expectBoolean(x)
        return args[i+1]
    @error(FormulaError.NA)

  INT: (args)->
    @checkArgumentSize(args, 1)
    args.push 1
    @FLOOR(args)

  "ISO.CEILING":@::["CEILING.PRECISE"]

  JIS: (args)->
    @DBCS(args)

  LCM: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_INTEGER, args)
    math.lcm(args...)

  LEFT: (args)->
    @checkArgumentSize(args, 1, 2)
    str = @expectString(args[0])
    len = 1
    if args.length == 2
      len = @expectInteger(args[1])
      if len < 0
        @error(FormulaError.VALUE)
    str.substr(0, len)

  LEFTB: (args)->
    @checkArgumentSize(args, 1, 2)
    str = @expectString(args[0])
    len = 1
    if args.length == 2
      len = @expectInteger(args[1])
      if len < 0
        @error(FormulaError.VALUE)

    str = mbsubstr(str, 0, len)
    #2バイト文字の途中で終わっていた場合
    if multibyteLength(str) < len
      str += " "
    str

  LEN: (args)->
    @checkArgumentSize(args, 1)
    str = @expectString(args[0])
    str.length

  LENB: (args)->
    @checkArgumentSize(args, 1)
    str = @expectString(args[0])
    multibyteLength(str)

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

  LOWER: (args)->
    @checkArgumentSize(args, 1)
    str = @expectString(args[0])
    str.toLowerCase()

  MID: (args)->
    @checkArgumentSize(args, 3)
    str = @expectString(args[0])
    start = @expectInteger(args[1])
    len = @expectInteger(args[2])
    if start < 0 || len < 0
      @error(FormulaError.VALUE)
    str.substr(start-1, len)

  MIDB: (args)->
    @checkArgumentSize(args, 3)
    str = @expectString(args[0])
    start = @expectInteger(args[1])
    len = @expectInteger(args[2])
    --start
    if start < 0 || len < 0
      @error(FormulaError.VALUE)

    #start + len が文字列の長さを超える場合は補正する
    if start + len > multibyteLength(str)
      len = multibyteLength(str) - start

    prefix_space = ""
    #2バイト文字の途中から始まっていた場合
    if multibyteLength(mbsubstr(str, 0, start)) < start
      prefix_space = " "
    len2 = start + len
    suffix_space = ""
    #2バイト文字の途中で終わっていた場合
    if multibyteLength(mbsubstr(str, 0, len2)) < len2
      suffix_space = " "
    prefix_space + mbsubstr(str, start, len) + suffix_space

  MINUTE: (args)->
    @checkArgumentSize(args, 1)
    @expectDate(args[0]).getMinutes()

  MOD: (args)->
    @checkArgumentSize(args, 2)
    n = @expectNumber(args[1])
    if n == 0
      @error(FormulaError.DIV0)
    Math.sign(n) * Math.abs(@expectNumber(args[0]) % Math.abs(n))

  MONTH: (args)->
    @checkArgumentSize(args, 1)
    @expectDate(args[0]).getMonth()+1

  MROUND: (args)->
    @checkArgumentSize(args, 2)
    m = @expectNumber(args[0])
    n = @expectNumber(args[1])
    if m == 0 || n == 0
      return 0
    if Math.sign(m) != Math.sign(n)
      @error(FormulaError.NUM)
    Math.round(m/n) * n

  MULTINOMIAL: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_NUMBER, args)
    math.multinomial args


  NA: (args)->
    @checkArgumentSize(args, 0)
    FormulaError.NA

  NETWORKDAYS: (args)->
    @checkArgumentSize(args, 2, 3)
    args.splice(2, 0, 1)
    @["NETWORKDAYS.INTL"](args)

  "NETWORKDAYS.INTL": (args)->
    @checkArgumentSize(args, 2, 4)
    d1 = @expectDate(args[0])
    d2 = @expectDate(args[1])

    sign = Math.sign(d2 - d1)
    if d1 > d2
      [d1,d2] = [d2,d1]

    weekdays = [1,2,3,4,5]
    excludes = []
    if args.length >= 3
      weekends = @getValue(args[2])
      if Utils.isString(weekends)
        if weekends.length != 7
          @error(FormulaError.VALUE)
      else if Utils.isNumber(weekends)
        str = ""
        n = weekends
        if 1 <= weekends <= 7
          n--
          str = "0000011"
        else if 11 <= weekends <= 17
          n -= 11
          str = "0000001"
        else
          @error(FormulaError.NUM)
        weekends = str.substr(-n) + str.slice(0, -n)
      else
        @error(FormulaError.VALUE)
      weekdays = []
      for v,i in weekends.split("")
        if v == "0"
          weekdays.push i+1
        else if v != "1"
          @error(FormulaError.VALUE)
    if args.length == 4
      excludes = @expandRange(TYPE_DATE, args[3])
    sign * moment().isoWeekdayCalc({
      rangeStart:d1
      rangeEnd:d2
      weekdays:weekdays
      exclusions:excludes
    })

  NOT: (args)->
    @checkArgumentSize(args, 1)
    !@expectNumber(args[0])

  NOW: (args)->
    @checkArgumentSize(args, 0)
    new Date()

  NUMBERVALUE: (args)->
    @checkArgumentSize(args, 1, 3)
    str = @expectString(args[0]).replace(/\s/, "")
    decimal_delim = "."
    if args.length >= 2
      decimal_delim = @expectString(args[1])
    group_delim = ","
    if args.length == 3
      group_delim = @expectString(args[2])

    parts = str.split(decimal_delim)
    if parts.length >= 3
      @error(FormulaError.VALUE)
    parts[0] = parts[0].replace(new RegExp(group_delim, "g"), "")
    if parts.length == 2 && parts[1].search(group_delim) != -1
      @error(FormulaError.VALUE)

    @expectNumber(parts.join(decimal_delim))

  ODD: (args)->
    @checkArgumentSize(args, 1)
    num = @expectNumber(args[0])
    if num >= 0
      @EVEN([num+1])-1
    else
      @EVEN([num-1])+1

  OR: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    res = false
    flag = false
    for arg in args
      if arg instanceof Range
        for x in @expandRange(TYPE_ANY, arg)
          if Utils.isBoolean(x)
            res ||= !!x
            flag = true
          else if x instanceof FormulaError
            throw x

      else
        x = @getValue(arg)
        if x instanceof FormulaError
          throw x
        unless Utils.isBoolean(x)
          @error(FormulaError.VALUE)
        res ||= !!x
        flag = true
    unless flag
      @error(FormulaError.VALUE)
    res

  PI: (args)->
    Math.PI

  POWER: (args)->
    @checkArgumentSize(args, 2)
    Math.pow(@expectNumber(args[0]), @expectNumber(args[1]))


  PROPER_RE = XRegExp("(^|\\PL)(\\pL)", "g")#アルファベット(英語以外も含む)以外
  PROPER: (args)->
    @checkArgumentSize(args, 1)
    str = @expectString(args[0]).toLowerCase()
    str.replace PROPER_RE, (m0, m1, m2)->
      return String(m1)+m2.toUpperCase()

  PERMUT: (args)->
    @checkArgumentSize(args, 2)
    n = @expectInteger(args[0])
    r = @expectInteger(args[1])
    if n <= 0 || r < 0 || n < r
      @error(FormulaError.NUM)

    math.permutations(n, r)


  PERMUTATIONA: (args)->
    @checkArgumentSize(args, 2)
    n = @expectInteger(args[0])
    r = @expectInteger(args[1])
    Math.pow(n, r)

  QUOTIENT: (args)->
    @checkArgumentSize(args, 2)
    parseInt(@expectNumber(args[0]) / @expectNumber(args[1]))

  RADIANS: (args)->
    @checkArgumentSize(args, 1)
    @expectNumber(args[0]) * Math.PI / 180

  RAND: (args)->
    @checkArgumentSize(args, 0)
    Math.random()

  RANDBETWEEN: (args)->
    @checkArgumentSize(args, 2)
    min = @expectInteger(args[0])
    max = @expectInteger(args[1])
    Math.floor(Math.random() * (max-min+1)) + min

  REPLACE: (args)->
    @checkArgumentSize(args, 4)
    str = @expectString(args[0])
    start = @expectInteger(args[1])
    len = @expectInteger(args[2])
    replace_str = @expectString(args[3])
    if start < 1 || len < 0
      @error(FormulaError.VALUE)
    arr = str.split("")
    arr.splice(start-1, len, replace_str.split("")...)
    arr.join("")

  REPLACEB: (args)->
    @checkArgumentSize(args, 4)
    str = @expectString(args[0])
    start = @expectInteger(args[1])
    len = @expectInteger(args[2])
    replace_str = @expectString(args[3])
    left = @LEFTB([str, start - 1])
    right = @MIDB([str, start + len, multibyteLength(str)])
    left + replace_str + right


  REPT: (args)->
    @checkArgumentSize(args, 2)
    str = @expectString(args[0])
    n = @expectInteger(args[1])
    if n < 0
      @error(FormulaError.VALUE)
    str.repeat(n)

  RIGHT: (args)->
    @checkArgumentSize(args, 1, 2)
    str = @expectString(args[0])
    args[0] = str.split("").reverse().join("")
    @LEFT(args).split("").reverse().join("")

  RIGHTB: (args)->
    @checkArgumentSize(args, 1, 2)
    str = @expectString(args[0])
    args[0] = str.split("").reverse().join("")
    @LEFTB(args).split("").reverse().join("")


  ROMAN: (args)->
    @checkArgumentSize(args, 1, 2)
    n = @expectInteger(args[0])
    if n >= 4000
      @error(FormulaError.VALUE)
    opt = @getValue(args[1])
    if Utils.isBoolean(opt)
      if opt
        opt = 0
      else
        opt = 4
    else
      opt = @expectInteger(opt)

    codes = {1:'I', 4:'IV', 5:'V', 9:'IX', 10:'X', 40:'XL', 50:'L', 90:'XC', 100:'C', 400:'CD', 500:'D', 900:'CM', 1000:'M'}
    if opt > 0
      codes[45] = 'VL'
      codes[95] = 'VC'
      codes[450] = 'LD'
      codes[950] = 'LM'
    if opt > 1
      codes[49] = 'IL'
      codes[99] = 'IC'
      codes[490] = 'XD'
      codes[990] = 'XM'
    if opt > 2
      codes[495] = 'VD'
      codes[995] = 'VM'
    if opt > 3
      codes[499] = 'ID'
      codes[999] = 'IM'

    nums = Object.keys(codes).map (x)-> parseInt(x)
    nums.sort (a, b)-> b - a

    roman = ""
    for num in nums
      while n >= num
        roman += codes[num]
        n -= num

    roman

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

  SEARCH: (args)->
    @checkArgumentSize(args, 2, 3)
    search = @expectString(args[0])
    target = @expectString(args[1])
    start = 1
    if args.length == 3
      start = @expectInteger(args[2])
    --start
    if start < 0
      @error(FormulaError.VALUE)
    target = target.substr(start)
    pos = target.search(search)
    if pos == -1
      @error(FormulaError.VALUE)
    start + pos + 1

  SEARCHB: (args)->
    @checkArgumentSize(args, 2, 3)
    search = @expectString(args[0])
    target = @expectString(args[1])
    start = 1
    if args.length == 3
      start = @expectInteger(args[2])
    if start <= 0
      @error(FormulaError.VALUE)
    target = @MIDB([target, start, multibyteLength(target)])
    pos = target.search(search)
    if pos == -1
      @error(FormulaError.VALUE)
    start + multibyteLength(@LEFTB([target, pos]))

  SEC: (args)->
    @checkArgumentSize(args, 1)
    1 / Math.cos(@expectNumber(args[0]))

  SECOND: (args)->
    @checkArgumentSize(args, 1)
    @expectDate(args[0]).getSeconds()

  SECH: (args)->
    @checkArgumentSize(args, 1)
    1 / Math.cosh(@expectNumber(args[0]))

  SIGN: (args)->
    @checkArgumentSize(args, 1)
    Math.sign @expectNumber(args[0])

  SIN: (args)->
    @checkArgumentSize(args, 1)
    Math.sin @expectNumber(args[0])

  SINH: (args)->
    @checkArgumentSize(args, 1)
    Math.sinh @expectNumber(args[0])

  SQRT: (args)->
    @checkArgumentSize(args, 1)
    n = @expectNumber(args[0])
    if n < 0
      @error(FormulaError.NUM)
    Math.sqrt n

  SQRTPI: (args)->
    @checkArgumentSize(args, 1)
    n = @expectNumber(args[0])
    if n < 0
      @error(FormulaError.NUM)
    Math.sqrt n*Math.PI

  SUBSTITUTE: (args)->
    @checkArgumentSize(args, 3, 4)
    str = @expectString(args[0])
    search = @expectString(args[1])
    replace = @expectString(args[2])
    re = RegExp(search, "g")
    if args.length == 3
      str.replace(re, replace)
    else
      n = @expectInteger(args[3])
      c = 0
      str.replace re, (m0)->
        ++c
        if c == n
          replace
        else
          m0

  T: (args)->
    @checkArgumentSize(args, 1)
    x = @getValue(args[0])
    unless Utils.isString(x)
      x = ""
    String(x)

  TEXTJOIN: (args)->
    @checkArgumentSize(args, 3, Number.MAX_VALUE)
    delim = @expectString(args[0])
    ignore_empty = @expectBoolean(args[1])
    strings = @expandRange(TYPE_STRING, args.slice(2))
    if ignore_empty
      strings = strings.filter (str)->
        str.length > 0
    strings.join(delim)

  TRIM: (args)->
    @checkArgumentSize(args, 1)
    str = @expectString(args[0])
    str.replace(/\s+/g, " ").trim()

  TRUE: (args)->
    @checkArgumentSize(args, 0)
    true

  SUM: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_NUMBER, args)
    sum = 0
    for num in args
      sum += num
    sum

  SUMSQ: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    args = @expandRange(TYPE_NUMBER, args)
    sum = 0
    for num in args
      sum += num*num
    sum


  SWITCH: (args)->
    @checkArgumentSize(args, 3, Number.MAX_VALUE)
    args = @expandRange(TYPE_ANY, args)
    if args.length % 2 == 1
      @error(FormulaError.NA)

    x = args[0]
    for y,i in args when i % 2 == 1
      if x == y
        return args[i+1]
    args[args.length-1]


  TAN: (args)->
    @checkArgumentSize(args, 1)
    Math.tan @expectNumber(args[0])

  TANH: (args)->
    @checkArgumentSize(args, 1)
    Math.tanh @expectNumber(args[0])

  TIME: (args)->
    @checkArgumentSize(args, 3)
    hour = @expectInteger(args[0])
    minute = @expectInteger(args[1])
    second = @expectInteger(args[2])
    dt = moment(Utils.offsetToDate(0))
    dt.hour(hour)
    dt.minute(minute)
    dt.second(second)
    Utils.dateToOffset(dt)

  TIMEVALUE: (args)->
    @checkArgumentSize(args, 1)
    str = @expectDateString(args[0])
    dt = moment(Utils.parseDate(str))
    dt0 = moment(Utils.offsetToDate(0))
    dt.year(dt0.year())
    dt.month(dt0.month())
    dt.date(dt0.date())
    Utils.dateToOffset(dt)

  TODAY: (args)->
    @checkArgumentSize(args, 0)
    moment().startOf('day').toDate()

  TRUNC: (args)->
    @checkArgumentSize(args, 1, 2)
    if args.length == 1
      args.push 0

    @ROUNDDOWN(args)

  UNICHAR: (args)->
    @checkArgumentSize(args, 1)
    code = @expectInteger(args[0])
    if code < 1
      @error(FormulaError.VALUE)
    String.fromCharCode(code)

  UNICODE: (args)->
    @checkArgumentSize(args, 1)
    text = @expectString(args[0])
    text.charCodeAt(0)

  UPPER: (args)->
    @checkArgumentSize(args, 1)
    str = @expectString(args[0])
    str.toUpperCase()

  VALUE: (args)->
    @checkArgumentSize(args, 1)
    @expectString(args[0])


  WEEKDAY: (args)->
    @checkArgumentSize(args, 1, 2)
    dt = @expectMoment(args[0])
    weekday = dt.isoWeekday()
    type = 1
    if args.length == 2
      type = @expectInteger(args[1])
    switch type
      when 1,17
        return (weekday % 7) + 1
      when 2,11
        return weekday
      when 3
        return weekday - 1

    unless 11 <= type <= 17
      @error(FormulaError.NUM)

    weekday = (weekday + 18 - type) % 7
    if weekday == 0
      weekday = 7
    weekday


  WEEKNUM: (args)->
    @checkArgumentSize(args, 1, 2)
    dt = @expectMoment(args[0])
    doy = dt.dayOfYear()

    type = 1
    if args.length == 2
      type = @expectInteger(args[1])
      if type == 21
        return dt.isoWeek()

    switch type
      when 1
        type = 17
      when 2
        type = 11

    dt0 = dt.clone().startOf("year")
    offset = dt0.day() #1/1の曜日を取得する
    doy += (offset - (type - 10 - 7) % 7) % 7
    Math.ceil(doy / 7)

  "WORKDAY.INTL": (args)->
    @checkArgumentSize(args, 2, 4)
    d1 = @expectDate(args[0])
    workdays = @expectInteger(args[1])

    weekdays = [1,2,3,4,5]
    excludes = []
    if args.length >= 3
      weekends = @getValue(args[2])
      if Utils.isString(weekends)
        if weekends.length != 7
          @error(FormulaError.VALUE)
      else if Utils.isNumber(weekends)
        str = ""
        n = weekends
        if 1 <= weekends <= 7
          n--
          str = "0000011"
        else if 11 <= weekends <= 17
          n -= 11
          str = "0000001"
        else
          @error(FormulaError.NUM)
        weekends = str.substr(-n) + str.slice(0, -n)
      else
        @error(FormulaError.VALUE)
      weekdays = []
      for v,i in weekends.split("")
        if v == "0"
          weekdays.push i+1
        else if v != "1"
          @error(FormulaError.VALUE)
    if args.length == 4
      excludes = @expandRange(TYPE_DATE, args[3])

    moment(d1).isoAddWeekdaysFromSet({
      'workdays': workdays,
      'weekdays': weekdays,
      'exclusions': excludes
    }).toDate()


  WORKDAYS: (args)->
    @checkArgumentSize(args, 2, 3)
    args.splice(2, 0, 1)
    @["WORKDAYS.INTL"](args)

  XOR: (args)->
    @checkArgumentSize(args, 1, Number.MAX_VALUE)
    true_count = 0
    flag = false
    for arg in args
      if arg instanceof Range
        for x in @expandRange(TYPE_ANY, arg)
          if Utils.isBoolean(x)
            true_count++ if !!x
            flag = true
          else if x instanceof FormulaError
            throw x

      else
        x = @getValue(arg)
        if x instanceof FormulaError
          throw x
        unless Utils.isBoolean(x)
          @error(FormulaError.VALUE)
        true_count++ if !!x
        flag = true
    unless flag
      @error(FormulaError.VALUE)
    true_count % 2 == 1


  YEAR: (args)->
    @checkArgumentSize(args, 1)
    @expectDate(args[0]).getFullYear()

  YEARFRAC: (args)->
    @checkArgumentSize(args, 2, 3)
    start = @expectDate(args[0])
    end = @expectDate(args[1])
    mode = 0
    if args.length == 3
      mode = @expectInteger(args[2])

    if start > end
      [start, end] = [end, start]

    switch mode
      when 0
        @DAYS360([start, end, false]) / 360
      when 1
        start_m = moment(start)
        end_m = moment(end)
        days_in_year = (end_m.endOf("year").diff(start_m.startOf("year"), "days") + 1) / (end_m.year() - start_m.year() + 1)
        @DAYS([end, start]) / days_in_year
      when 2
        @DAYS([end, start]) / 360
      when 3
        @DAYS([end, start]) / 365
      when 4
        @DAYS360([start, end, true]) / 360
      else
        @error(FormulaError.NUM)

  error:(err)->
    throw err


  getValue:(range)->
    return range unless range instanceof Range

    size = range.size
    if size == 1
      cell = null
      range.each (r, c)=>
        cell = @worksheet.getCell(r, c)
      cell.calculate()
    else if size == 0
      @error(FormulaError.NULL)
    else
      @error(FormulaError.VALUE)

  expectNumber:(x)->
    x = @getValue(x)
    if x == undefined
      return Number(0)
    if Utils.isNumber(x) || Utils.isBoolean(x)
      return Number(x)
    if Utils.isDate(x)
      return Utils.dateToOffset(Utils.parseDate(x))
    if Utils.isString(x)
      x = x.replace(/\s*,\s*/, "")
      if Utils.isNumber(x)
        return Number(x)
    @error(FormulaError.VALUE)

  expectInteger:(x)->
    parseInt @expectNumber(x)

  expectString:(x)->
    x = @getValue(x)
    if x == undefined
      return ""
    if Utils.isString(x) || Utils.isNumber(x)
      return String(x)
    if Utils.isBoolean(x)
      return String(x*1)
    @error(FormulaError.VALUE)

  expectDate:(x)->
    x = @getValue(x)
    if Utils.isDate(x)
      return Utils.parseDate(x)
    if x == undefined
      x = 0
    if Utils.isNumber(x)
      return Utils.offsetToDate(x)
    @error(FormulaError.VALUE)

  expectMoment:(x)->
    moment(@expectDate(x))

  expectDateString:(x)->
    x = @getValue(x)
    if Utils.isDateString(x)
      return x

    @error(FormulaError.VALUE)

  expectBoolean:(x)->
    x = @getValue(x)
    if x == undefined
      return false
    unless Utils.isBoolean(x)
      @error(FormulaError.VALUE)
    !!x

  expect:(type, x)->
    switch type
      when TYPE_NUMBER
        @expectNumber(x)
      when TYPE_INTEGER
        @expectInteger(x)
      when TYPE_STRING
        @expectString(x)
      when TYPE_DATE
        @expectDate(x)
      when TYPE_DATE_STRING
        @expectDateString(x)
      when TYPE_BOOLEAN
        @expectBoolean(x)
      when TYPE_PURE_NUMBER
        @expectPureNumber(x)
      when TYPE_PURE_INTEGER
        @expectPureInteger(x)
      when TYPE_ANY
        @getValue(x)


  expandRange:(type, args)->
    args = [].concat(args)
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

module.exports = FormulaEvaluator
