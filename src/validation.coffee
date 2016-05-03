class Validation
	constructor:(xmlobj)->
		attr = xmlobj.$
		@type = attr.type
		@allowBlank = attr.allowBlank
		@showInputMessage = attr.showInputMessage
		@showErrorMessage = attr.showErrorMessage
		addr = attr.sqref
		f1 = xmlobj.formula1?[0]
		if f1 != null && typeof f1 == "object"
			f1 = f1.f[0]
		f2 = xmlobj.formula2?[0]
		if f2 != null && typeof f2 == "object"
			f2 = f2.f[0]

		unless addr
			addr = xmlobj.sqref[0]

		@formula1 = f1
		@formula2 = f2

		@toXmlObj = ->
			obj = {
				$:{type:@type, allowBlank:@allowBlank, showInputMessage:@showInputMessage, showErrorMessage:@showErrorMessage, sqref:addr}
			}
			if @formula1
				obj.formula1 = [@formula1]
			if @formula2
				obj.formula2 = [@formula2]
			obj

module.exports = Validation
