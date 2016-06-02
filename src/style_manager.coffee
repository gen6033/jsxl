Promise = require('bluebird')
{parseStringAsync} = Promise.promisifyAll(require('xml2js'))
processors = require('xml2js/lib/processors')
Utils = require('./utils')
Style = require('./style/style')
Font = require('./style/font')
Fill = require('./style/fill')
Border = require('./style/border')
NumberFormat = require('./style/number_format')
Theme = require('./theme')
extend = require('extend')
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


		fmts = []
		fmts[0] = new NumberFormat(this, 0, '@')
		fmts[1] = new NumberFormat(this, 1, '0')
		fmts[2] = new NumberFormat(this, 2, '0.00')
		fmts[3] = new NumberFormat(this, 3, '#, ##0')
		fmts[4] = new NumberFormat(this, 4, '#, ##0.00')
		fmts[5] = new NumberFormat(this, 5, '$#, ##0_);($#, ##0)')
		fmts[6] = new NumberFormat(this, 6, '$#, ##0_);[Red]($#, ##0)')
		fmts[7] = new NumberFormat(this, 7, '$#, ##0.00_);($#, ##0.00)')
		fmts[8] = new NumberFormat(this, 8, '$#, ##0.00_);[Red]($#, ##0.00)')
		fmts[9] = new NumberFormat(this, 9, '0%')
		fmts[10] = new NumberFormat(this, 10, '0.00%')
		fmts[11] = new NumberFormat(this, 11, '0.00E+00')
		fmts[12] = new NumberFormat(this, 12, '# ?/?')
		fmts[13] = new NumberFormat(this, 13, '# ??/??')
		fmts[14] = new NumberFormat(this, 14, 'm/d/yyyy')
		fmts[15] = new NumberFormat(this, 15, 'd-mmm-yy')
		fmts[16] = new NumberFormat(this, 16, 'd-mmm')
		fmts[17] = new NumberFormat(this, 17, 'mmm-yy')
		fmts[18] = new NumberFormat(this, 18, 'h:mm AM/PM')
		fmts[19] = new NumberFormat(this, 19, 'h:mm:ss AM/PM')
		fmts[20] = new NumberFormat(this, 20, 'h:mm')
		fmts[21] = new NumberFormat(this, 21, 'h:mm:ss')
		fmts[22] = new NumberFormat(this, 22, 'm/d/yyyy h:mm')
		fmts[37] = new NumberFormat(this, 37, '#, ##0_);(#, ##0)')
		fmts[38] = new NumberFormat(this, 38, '#, ##0_);[Red](#, ##0)')
		fmts[39] = new NumberFormat(this, 39, '#, ##0.00_);(#, ##0.00)')
		fmts[40] = new NumberFormat(this, 40, '#, ##0.00_);[Red](#, ##0.00)')
		fmts[45] = new NumberFormat(this, 45, 'mm:ss')
		fmts[46] = new NumberFormat(this, 46, '[h]:mm:ss')
		fmts[47] = new NumberFormat(this, 47, 'mm:ss.0')
		fmts[48] = new NumberFormat(this, 48, '##0.0E+0')
		fmts[49] = new NumberFormat(this, 49, '@')
		fmts[169] = null
		if ss.numFmts?[0].numFmt
			for format in ss.numFmts[0].numFmt
				fmt = NumberFormat.parse(this, format)
				fmts[fmt.id] = fmt
		@_.numberFormats = fmts

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

	getNumberFormat:(id)->
		@_.numberFormats[id]

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
			when "NumberFormat"
				resources = @_.numberFormats

		resources.push newResource
		newId = resources.length - 1
		newResource.id = newId
		newResource


	toXmlObj:->
		ss = {
			$:{}
			numFmts:[]
			fonts:[]
			fills:[]
			borders:[]
		}
		ss = extend ss, @xmlobj.styleSheet
		@xmlobj.styleSheet = ss


		ss.numFmts = [{numFmt:[]}]
		for format in @_.numberFormats when format
			ss.numFmts[0].numFmt.push format.toXmlObj()


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
