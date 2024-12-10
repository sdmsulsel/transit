var sheetName = 'Sheet1'
var scriptProp = PropertiesService.getScriptProperties()

function intialSetup () {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  scriptProp.setProperty('key', activeSpreadsheet.getId())
// ID_FOLDER  ganti dengan id folder tujuan
  scriptProp.setProperty('folder', '1RyIoDeK-ZIjRngvjl9e_QOxnL1bNg7Ol')
}

function doPost (e) {
  var lock = LockService.getScriptLock()
  lock.tryLock(10000)

  try {
    // Simpan file upload pertama
     const files = Object.keys(e.parameter)?.filter(x => {
      return /^file/i.test(x)
    })
    if (files) {
      for (const x of files) {
        var data = e.parameter[x]
        var base64 = data.replace(/^data.*;base64,/gim, "")
        var mimetype = data.match(/(?<=data:).*?(?=;)/gim)?.[0]
        var decode = Utilities.base64Decode(base64, Utilities.Charset.UTF_8)
        //parameter nama dan tanggal tinggal disesuaikan
        var blob = Utilities.newBlob(decode, mimetype, `${x}-${e.parameter.nama}_${Date.now()}.${mimetype.split('/')[1]}`)
        var file = DriveApp.getFolderById(scriptProp.getProperty('folder')).createFile(blob)
        e.parameter[x] = file.getUrl()
      }
    }

    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
    var sheet = doc.getSheetByName(sheetName)

     const dataSheet = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues()
    const headers = dataSheet[0]
    let nextRow = sheet.getLastRow() + 1
    let updated = false
    dataSheet.forEach((v, i)=>{
      // ganti nama nik dgn nama di form yg menjadi dasar data double yg di edit
        if(v[headers.indexOf('nrp')] == e.parameter['nrp']){
            nextRow = i + 1
            updated = true
        }
    })
    var newRow = headers.map(function(header) {
        return header === 'timestamp' ? new Date() : e.parameter[header] 
    })
    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])

    return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'updated': updated })).setMimeType(ContentService.MimeType.JSON)
  }
  catch (e) {
    console.log(e)
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e.stack }))
      .setMimeType(ContentService.MimeType.JSON)
  }

  finally {
    lock.releaseLock()
  }
}