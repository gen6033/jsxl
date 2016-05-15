Utils = require('../utils')
require('js-object-clone')

COLOR_TABLE = [null, "000000", "FFFFFF", "0000FF", "00FF00", "FF0000", "00FFFF", "FF00FF", "FFFF00", "000080", "008000", "800000", "008080", "800080", "808000", "C0C0C0", "808080", "FF9999", "663399", "CCFFFF", "FFFFCC", "660066", "8080FF", "CC6600", "FFCCCC", "800000", "FF00FF", "00FFFF", "FFFF00", "800080", "000080", "808000", "FF0000", "FFCC00", "FFFFCC", "CCFFCC", "99FFFF", "FFCC99", "CC99FF", "FF99CC", "99CCFF", "FF6633", "CCCC33", "00CC99", "00CCFF", "0099FF", "0066FF", "996666", "969696", "663300", "669933", "003300", "003333", "003399", "663399", "993333", "333333"]

class Font
	constructor:(id, sm, xmlobj)->
		@_ = {}
		@_.xmlobj = xmlobj
		@id = id
		@_.sm = sm
		@size = parseInt xmlobj.sz?[0].$.val
		colorAttrs = xmlobj.color?[0].$ || {}
		color = null
		if colorAttrs.rgb
			color = colorAttrs.rgb
		else if colorAttrs.indexed
			color = COLOR_TABLE[colorAttrs.indexed]
		else if colorAttrs.theme
			@_.color_theme = colorAttrs.theme
			color = @_.sm._.theme.getColor(colorAttrs.theme)

		@color = color
		@name = xmlobj.name?[0].$.val
		#@_.charset = parseInt xmlobj.charset?[0].$.val
		#@_.family = parseInt xmlobj.family?[0].$.val
		#@_.scheme = xmlobj.scheme?[0].$.val


	clone: ->
		@_.sm.cloneResource(this)


	toXmlObj: ->
		obj = Object.clone(@_.xmlobj)

		if @size
			obj.sz = [{$:{val:@size}}]

		if @name
			obj.name = [{$:{val:@name}}]

		if @color
			obj.color = [{$:{rgb:@color}}]

		obj

module.exports = Font
