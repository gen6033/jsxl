Promise = require('bluebird')
fs = Promise.promisifyAll(require('fs'))
Zip = require('node-zip')
xml2js = Promise.promisifyAll(require('xml2js'))
SharedStringsTable = require('./shared_strings_table')
StyleManager = require('./style_manager')
Worksheet = require('./worksheet')
Theme = require('./theme')
Utils = require('./utils')
parser = new xml2js.Parser()
builder = new xml2js.Builder()

class Workbook

  @read: (excel_path)->
    zip = null
    workbook = new Workbook()
    fs.readFileAsync excel_path, 'binary'
    .then (data)->
      workbook._.zip = zip = new Zip(data, {base64: false, checkCRC32: true})
      jobs = []
      jobs.push StyleManager.parse(zip.files["xl/styles.xml"].asText(), zip.files["xl/theme/theme1.xml"].asText())
      jobs.push SharedStringsTable.parse(zip.files["xl/sharedStrings.xml"]?.asText())
      jobs.push parser.parseStringAsync zip.files["xl/workbook.xml"].asText()
      Promise.all(jobs)
    .then ([sm, sst, wb])->
      workbook._.sm = sm
      workbook._.sst = sst
      i = 1
      jobs = []
      workbook._.sheets = []
      workbook._.sheetIdxs = {}
      workbook.path = fs.realpathSync excel_path
      for s in wb.workbook.sheets[0].sheet
        sheet_name = s.$.name
        workbook._.sheetIdxs[sheet_name] = i-1
        jobs.push Worksheet.create(workbook, s.$.sheetId, s.$["r:id"].replace("rId", ""), sheet_name, zip.files["xl/worksheets/sheet#{i}.xml"].asText()).then (sheet)->
          workbook._.sheets.push sheet
        ++i

      Promise.all(jobs).then ->
        workbook

  constructor:->
    @_ = {}

  getSheet: (i)->
    @_.sheets[i]

  getSheetByName: (name)->
    @getSheet @_.sheetIdxs[name]

  save: (path = null)->
    path = @path if path == null
    @_.zip.file("xl/sharedStrings.xml", builder.buildObject @_.sst.toXmlObj())
    @_.zip.file("xl/styles.xml", builder.buildObject @_.sm.toXmlObj())

    i = 1
    for sheet in @_.sheets
      @_.zip.file("xl/worksheets/sheet#{i}.xml", builder.buildObject sheet.toXmlObj())
      ++i

    @_.zip.remove("xl/calcChain.xml")
    data = @_.zip.generate({base64:false,compression:'DEFLATE'});
    fs.writeFileSync(path, data, 'binary');

module.exports = Workbook
