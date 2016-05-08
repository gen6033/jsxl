module.exports = {
	toRowCol:(addr)->
		[tmp, col, row] = addr.match(/^\s*([A-Z]+)(\d+)\s*$/)
		[parseInt(row), @toDigit(col)]

	toAddr:(row, col)->
		@toAlphabet(col)+row

	toDigit: (alpha)->
		result = 0
		i = 0
		maxi = alpha.length
		while i < maxi
			result = result*26 + parseInt(alpha[i++], 36)-9
		result

	toAlphabet: (digit)->
		result = []
		i = 0
		while digit > 0
			result.unshift ((digit-1) % 26 + 10).toString(36).toUpperCase()
			digit = parseInt ((digit-1) / 26)
			++i
		result.join("")

}
