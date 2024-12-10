const sheetName = 'Sheet1'
const scriptProp = PropertiesService.getScriptProperties()

function intialSetup() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    scriptProp.setProperty('key', activeSpreadsheet.getId())
    // ID_FOLDER  ganti dengan id folder tujuan
    scriptProp.setProperty('folder', '1RyIoDeK-ZIjRngvjl9e_QOxnL1bNg7Ol')
}

function doPost(e) {
    // Kunci script untuk mencegah race condition
    const lock = LockService.getScriptLock()
    lock.tryLock(10000)

    try {
        const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'))
        const sheet = doc.getSheetByName(sheetName)
        const dataSheet = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues()
        const headers = dataSheet[0]
        const nextRow = sheet.getLastRow() + 1
        let dataExists = false
        const files = Object.keys(e.parameter)?.filter(x => {
            return /^file/i.test(x)
        })


        // Periksa apakah NRP dan NOTR sudah ada
        dataSheet.forEach((v, i) => {
            if (v[headers.indexOf('nrp')] == e.parameter['nrp'] && v[headers.indexOf('notr')] == e.parameter['notr']) dataExists = true
        })

        // Data NRP dan NOTR sudah ada
        if (dataExists) return ContentService.createTextOutput(JSON.stringify({ ok: false, message: `ANDA PERNAH INPUT DATA DENGAN NRP dan NOTR YANG SAMA` })).setMimeType(ContentService.MimeType.JSON)

        // Simpan file jika request memuat file
        if (files) {
            for (const x of files) {
                var data = e.parameter[x]
                var base64 = data.replace(/^data.*;base64,/gim, "")
                var mimetype = data.match(/(?<=data:).*?(?=;)/gim)?.[0]
                var decode = Utilities.base64Decode(base64, Utilities.Charset.UTF_8)
                //parameter nama dan tanggal tinggal disesuaikan
                var blob = Utilities.newBlob(decode, mimetype, `${x} - ${e.parameter.nama}_${Date.now()}.${mimetype.split('/')[1]}`)
                var file = DriveApp.getFolderById(scriptProp.getProperty('folder')).createFile(blob)
                e.parameter[x] = file.getUrl()
            }
        }

        const newRow = headers.map(function (header) {
            return header === 'timestamp' ? new Date() : e.parameter[header]
        })

        // Tulis data ke Spreadsheet
        sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow])

        return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON)
    } catch (e) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, 'message': 'Server error!' })).setMimeType(ContentService.MimeType.JSON)
    } finally {
        lock.releaseLock()
    }
}