

class StyleMixin

	@mixin: (proto)->
		cloneStyle = (self)->
			mixin = self._.styleMixin
			mixin.style = mixin.sm.cloneResource(mixin.style)

		cloneFont = (self)->
			cloneStyle(self)
			mixin = self._.styleMixin
			mixin.style.font = mixin.sm.cloneResource(mixin.style.font)

		Object.defineProperties proto,
			"fontName":
				get: -> @_.styleMixin.style.font.name
				set: (val)->
					cloneFont(this)
					mixin = @_.styleMixin
					mixin.style.font.name = val
					mixin.onUpdate?()
			"fontSize":
				get: -> @_.styleMixin.style.font.size
				set: (val)->
					cloneFont(this)
					mixin = @_.styleMixin
					mixin.style.font.size = val
					mixin.onUpdate?()
			"fontColor":
				get: -> @_.styleMixin.style.font.color
				set: (val)->
					cloneFont(this)
					mixin = @_.styleMixin
					mixin.style.font.color = val
					mixin.onUpdate?()

	@bind: (obj, styleId)->
		new StyleMixin(obj, styleId)
		
	constructor: (obj, styleId) ->
		obj._.styleMixin = this
		@sm = obj.workbook._.sm
		@style = @sm.getStyle(styleId)


module.exports = StyleMixin
