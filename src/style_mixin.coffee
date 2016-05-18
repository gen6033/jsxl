require('js-object-clone')

class StyleMixin

	@mixin: (proto)->
		cloneStyle = (self)->
			return if self._.styleMixin.styleCloned
			mixin = self._.styleMixin = Object.clone(self._.styleMixin)
			mixin.style = mixin.sm.cloneResource(mixin.style)
			mixin.styleCloned = true

		cloneFont = (self)->
			return if self._.styleMixin.fontCloned
			cloneStyle(self)
			mixin = self._.styleMixin
			mixin.style.font = mixin.sm.cloneResource(mixin.style.font)
			mixin.fontCloned = true

		fontSetter = (self, key, val) ->
			cloneFont(self)
			mixin = self._.styleMixin
			mixin.style.font[key] = val
			mixin.onUpdate?()

		cloneFill = (self)->
			return if self._.styleMixin.fillCloned
			cloneStyle(self)
			mixin = self._.styleMixin
			mixin.style.fill = mixin.sm.cloneResource(mixin.style.fill)
			mixin.fillCloned = true

		fillSetter = (self, key, val) ->
			cloneFill(self)
			mixin = self._.styleMixin
			mixin.style.fill[key] = val
			mixin.onUpdate?()


		Object.defineProperties proto,
			"fontName":
				get: -> @_.styleMixin.style.font.name
				set: (val)->
					fontSetter(this, "name", val)

			"fontSize":
				get: -> @_.styleMixin.style.font.size
				set: (val)->
					fontSetter(this, "size", val)

			"fontColor":
				get: -> @_.styleMixin.style.font.color
				set: (val)->
					fontSetter(this, "color", val)

			"fillPattern":
				get: -> @_.styleMixin.style.fill.type
				set: (val)->
					fillSetter(this, "type", val)

			"fillPatternColor":
				get: -> @_.styleMixin.style.fill.fgColor
				set: (val)->
					fillSetter(this, "fgColor", val)

			"fillColor":
				get: -> @_.styleMixin.style.fill.bgColor
				set: (val)->
					fillSetter(this, "bgColor", val)

	@bind: (obj, styleId)->
		new StyleMixin(obj, styleId)

	constructor: (obj, styleId) ->
		obj._.styleMixin = this
		@sm = obj.workbook._.sm
		@style = @sm.getStyle(styleId)


module.exports = StyleMixin
