Promise = require('bluebird')
{parseStringAsync} = Promise.promisifyAll(require('xml2js'))
processors = require('xml2js/lib/processors')
Utils = require('./utils')
Style = require('./style/style')
Font = require('./style/font')
Fill = require('./style/fill')
Border = require('./style/border')
Theme = require('./theme')
require('js-object-clone')


COLOR_TABLE = [null, "000000", "FFFFFF", "0000FF", "00FF00", "FF0000", "00FFFF", "FF00FF", "FFFF00", "000080", "008000", "800000", "008080", "800080", "808000", "C0C0C0", "808080", "FF9999", "663399", "CCFFFF", "FFFFCC", "660066", "8080FF", "CC6600", "FFCCCC", "800000", "FF00FF", "00FFFF", "FFFF00", "800080", "000080", "808000", "FF0000", "FFCC00", "FFFFCC", "CCFFCC", "99FFFF", "FFCC99", "CC99FF", "FF99CC", "99CCFF", "FF6633", "CCCC33", "00CC99", "00CCFF", "0099FF", "0066FF", "996666", "969696", "663300", "669933", "003300", "003333", "003399", "663399", "993333", "333333"]

class StyleManager
	@parse:(styleXml, themeXml)->
		theme = null
		parseStringAsync(themeXml, {tagNameProcessors: [processors.stripPrefix]})
		.then (result)->
			theme = new Theme(result)
			parseStringAsync styleXml
		.then (result)->
			new StyleManager(result, theme)

	constructor:(@xmlobj, @theme)->
		@_ = {}
		ss = @xmlobj.styleSheet

		@_.fonts = []
		for font,idx in ss.fonts[0].font
			@_.fonts.push new Font(idx, this, font)

		@_.fills = []
		for fill,idx in ss.fills[0].fill
			@_.fills.push new Fill(idx, this, fill)

		@_.borders = []
		for border,idx in ss.borders[0].border
			@_.borders.push new Border(idx, this, border)

		@_.styles = []
		for xf,idx in ss.cellXfs[0].xf
			@_.styles.push new Style(idx, this, xf)

	getStyle:(id)->
		@_.styles[id]

	getFont:(id)->
		@_.fonts[id]

	getFill:(id)->
		@_.fills[id]

	getBorder:(id)->
		@_.borders[id]

	getRGB:(colorAttrs = {})->
		color = null
		if colorAttrs.rgb
			color = colorAttrs.rgb
		else if colorAttrs.indexed
			color = COLOR_TABLE[parseInt colorAttrs.indexed]
		else if colorAttrs.theme
			color = @theme.getColor(colorAttrs.theme)
		else if colorAttrs.auto
			color = "auto"
		color

	cloneResource:(resource)->
		newResource = Object.clone(resource)
		newResource._ = Object.clone(resource._)
		resources = null
		switch(resource.constructor.name)
			when "Style"
				resources = @_.styles
			when "Font"
				resources = @_.fonts
			when "Fill"
				resources = @_.fills
			when "Border"
				resources = @_.borders

		resources.push newResource
		newId = resources.length - 1
		newResource.id = newId
		newResource


	toXmlObj:->
		ss = @xmlobj.styleSheet

		ss.fonts = [{font:[]}]
		for font in @_.fonts
			ss.fonts[0].font.push font.toXmlObj()

		ss.fills = [{fill:[]}]
		for fill in @_.fills
			ss.fills[0].fill.push fill.toXmlObj()

		ss.borders = [{border:[]}]
		for border in @_.borders
			ss.borders[0].border.push border.toXmlObj()

		ss.cellXfs = [{xf:[]}]
		for style in @_.styles
			ss.cellXfs[0].xf.push style.toXmlObj()

		@xmlobj


module.exports = StyleManager
