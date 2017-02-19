Utils = require("./utils")
require('js-object-clone')

class BlockRange
  constructor:(top, left, bottom, right)->
    @_ = {}
    unless bottom
      if left
        [begin, end] = [top, left]
      else
        [begin, end] = top.split(":")
        end = begin unless end

      [top, left] = Utils.toRowCol(begin)
      [bottom, right] = Utils.toRowCol(end)

    [top, bottom] = [bottom, top] if bottom < top
    [left, right] = [right, left] if right < left
    @_.top = top
    @_.left = left
    @_.bottom = bottom
    @_.right = right

  Object.defineProperties @prototype,
    "top":
      get: -> @_.top
    "left":
      get: -> @_.left
    "bottom":
      get: -> @_.bottom
    "right":
      get: -> @_.right
    "size":
      get: -> (@_.right - @_.left + 1) * (@_.bottom - @_.top + 1)

  includes: (addr)->
    [row, col] = Utils.toRowCol(addr)
    @top <= row <= @bottom && @left <= col <= @right

  conflicts: (range)->
    !(@left > range.right || @top > range.bottom || @right < range.left || @bottom < range.top)

  intersection: (range)->
    @_.top = Math.max(@top, range.top)
    @_.bottom = Math.min(@bottom, range.bottom)
    @_.left = Math.max(@left, range.left)
    @_.right = Math.min(@right, range.right)
    this

  each: (f, row_major = true)->
    if row_major
      @rowMajorEach(f)
    else
      @columnMajorEach(f)

  rowMajorEach: (f)->
    for r in [@top..@bottom]
      for c in [@left..@right]
        f(r, c)

  columnMajorEach: (f)->
    for c in [@left..@right]
      for r in [@top..@bottom]
        f(r, c)

  toString:->
    return Utils.toAddr(@top, @left) if @top == @bottom && @left == @right
    Utils.toAddr(@top, @left) + ":" + Utils.toAddr(@bottom, @right)

class Range
  constructor:(range)->
    @_ = {}
    @_.blockRanges = []
    if Utils.isString(range)
      for r in range.split(",")
        @_.blockRanges.push new BlockRange(r)
    else
      ranges = [].concat(range)
      for r in ranges
        if Utils.isString(r)
          r = new BlockRange(r)
        @_.blockRanges.push r

  union: (range)->
    if range instanceof BlockRange
      @_.blockRanges.push range
    else if range instanceof Range
      @_.blockRanges = @_.blockRanges.concat(range._.blockRanges)
    this

  unify: ->
    top = []
    bottom = []
    left = []
    right = []
    @eachBlock (r)->
      top.push r.top
      bottom.push r.bottom
      left.push r.left
      right.push r.right
    @_.blockRanges = [new BlockRange(Math.min(top...), Math.min(left...), Math.max(bottom...), Math.max(right...))]
    this

  intersection: (range)->
    blocks = []
    @eachBlock (r1)->
      range.eachBlock (r2)->
        blocks.push Object.clone(r1, true).intersection(r2)
    @_.blockRanges = blocks
    this

  eachBlock: (f)->
    for r in @_.blockRanges
      break unless f(r)

  includes: (addr)->
    ret = false
    @eachBlock (r)->
      if r.includes addr
        ret = true
        return false
    ret

  conflicts: (range)->
    ret = false
    @eachBlock (r1)->
      range.eachBlock (r2)->
        if r1.conflicts r2
          ret = true
          return false
    ret


  Object.defineProperties @prototype,
    "size":
      get: ->
        size = 0
        @eachBlock (r)->
          size += r.size
        size


  each: (f, row_major)->
    @eachBlock (r)->
      r.each(f, row_major)

  rowMajorEach: (f, row_major)->
    @eachBlock (r)->
      r.rowMajorEach(f, row_major)

  columnMajorEach: (f, row_major)->
    @eachBlock (r)->
      r.columnMajorEach(f, row_major)

  toString:->
    res = []
    @eachBlock (r)->
      res.push r.toString()
    res.join(",")



module.exports = {Range:Range, BlockRange:BlockRange}
