
class FormulaError extends Error

  @DIV0 = new FormulaError("#DIV/0!")
  @NA = new FormulaError("#N/A!")
  @NAME = new FormulaError("#NAME?")
  @NUM = new FormulaError("#NUM!")
  @NULL = new FormulaError("#NULL!")
  @VALUE = new FormulaError("#VALUE!")

  constructor: (@name)->

  toString: ->
    @name

module.exports = FormulaError
