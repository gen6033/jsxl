Promise = require('bluebird')
{parseStringAsync} = Promise.promisifyAll(require('xml2js'))
processors = require('xml2js/lib/processors')
Utils = require('./utils')

class Theme
	constructor:(xmlobj)->
		@_ = {}
		@_.xmlobj = xmlobj
		theme = xmlobj.theme
		colors = []
		elms = theme.themeElements[0]
		for k,elm of elms.clrScheme[0]
			continue if k == "$"
			if elm[0].sysClr
				colors.push elm[0].sysClr[0].$.lastClr
			else if elm[0].srgbClr
				colors.push elm[0].srgbClr[0].$.val
		[colors[0], colors[1], colors[2], colors[3]] = 	[colors[1], colors[0],colors[3], colors[2]]
		@_.colors = colors


	getColor: (id)->
		@_.colors[id]

	toXmlObj: ->
		@_.xmlobj

module.exports = Theme
