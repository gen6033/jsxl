Promise = require('bluebird')
{parseStringAsync} = Promise.promisifyAll(require('xml2js'))
processors = require('xml2js/lib/processors')
Utils = require('./utils')
Style = require('./style/style')
Font = require('./style/font')
Theme = require('./theme')
require('js-object-clone')

class StyleManager
	@parse:(styleXml, themeXml)->
		theme = null
		parseStringAsync(themeXml, {tagNameProcessors: [processors.stripPrefix]})
		.then (result)->
			theme = new Theme(result)
			parseStringAsync styleXml
		.then (result)->
			new StyleManager(result, theme)

	constructor:(xmlobj, theme)->
		@_ = {}
		@_.xmlobj = xmlobj
		@_.theme = theme
		ss = xmlobj.styleSheet

		@_.fonts = []
		for font,idx in ss.fonts[0].font
			@_.fonts.push new Font(idx, this, font)

		@_.styles = []
		for xf,idx in ss.cellXfs[0].xf
			@_.styles.push new Style(idx, this, xf)

	getStyle:(id)->
		@_.styles[id]

	getFont:(id)->
		@_.fonts[id]

	cloneResource:(resource)->
		newResource = Object.clone(resource)
		newResource._ = Object.clone(resource._)
		resources = null
		switch(resource.constructor.name)
			when "Style"
				resources = @_.styles
			when "Font"
				resources = @_.fonts

		resources.push newResource
		newId = resources.length - 1
		newResource.id = newId
		newResource


	toXmlObj:->
		ss = @_.xmlobj.styleSheet
		#console.log ss.fonts
		ss.fonts = [{font:[]}]
		for font in @_.fonts
			ss.fonts[0].font.push font.toXmlObj()

		ss.cellXfs = [{xf:[]}]
		for style in @_.styles
			ss.cellXfs[0].xf.push style.toXmlObj()

		@_.xmlobj


module.exports = StyleManager
