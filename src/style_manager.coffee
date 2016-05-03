Promise = require('bluebird')
{parseStringAsync} = Promise.promisifyAll(require('xml2js'))
Utils = require('./utils')
Style = require('./style')
Font = require('./style/font')
require('js-object-clone')

class StyleManager
	@parse:(xml)->
		parseStringAsync xml
		.then (result)->
			new StyleManager(result)

	constructor:(xmlobj)->
		@_ = {}
		@_.xmlobj = xmlobj
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
		newResource._.id = newId
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
