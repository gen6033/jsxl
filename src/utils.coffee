module.exports = {
	toRowCol:(addr)->
		[tmp, col, row] = addr.match(/^\s*([A-Z]+)(\d+)\s*$/)
		result = 0
		i = 0
		maxi = col.length
		while i < maxi
			result = result*26 + parseInt(col[i++], 36)-9

		[parseInt(row), result]

	toAddr:(row, col)->
		result = []
		i = 0
		while col > 0
			result.unshift ((col-1) % 26 + 10).toString(36).toUpperCase()
			col = parseInt ((col-1) / 26)
			++i
		return result.join("")+row

}
