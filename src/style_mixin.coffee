

class StyleMixin

	@mixin: (proto)->
		cloneStyle = (self)->
			mixin = self._.styleMixin
			mixin.style = mixin.sm.cloneResource(mixin.style)

		cloneFont = (self)->
			cloneStyle(self)
			mixin = self._.styleMixin
			mixin.style.font = mixin.sm.cloneResource(mixin.style.font)

		fontSetter = (self, key, val) ->
			mixin = self._.styleMixin
			if mixin.cloned
				return mixin.style.font[key] = val
			cloneFont(self)
			mixin.style.font[key] = val
			mixin.cloned = true
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

	@bind: (obj, styleId)->
		new StyleMixin(obj, styleId)

	constructor: (obj, styleId) ->
		obj._.styleMixin = this
		@sm = obj.workbook._.sm
		@style = @sm.getStyle(styleId)


module.exports = StyleMixin
