function getColumnToLetter(column) {
  var temp, letter = ''
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26
  }
  return letter
}

function getLetterToColumn(letter) {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1)
  return column
}

function getLastRow(sheet) {
  const values = sheet.getRange('A1:A' + sheet.getLastRow()).getDisplayValues()
  let lastRow = 0
  for (let r = values.length - 1; r >= 0 ; r--) {
    if (!lastRow && !values[r].every(e => e == "")) {
      lastRow = r + 1
      break
    }
  }
  return lastRow
}

function getToday() {
  var today = new Date()
  var date = today.getFullYear() + '-' + (today.getMonth() + 1) + '-' + today.getDate()
  var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds()
  return date + ' ' + time
}

function getWorkerSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<google_sheet_name>')} 
function getWorkerDocTemplate() {return DriveApp.getFileById('<google_doc_file_id>') } // doc template for worker forms
function getWorkerRepository() {return DriveApp.getFolderById('<google_drive_folder_id>')} // directory where worker forms are stored

function getEmploymentSheet() {return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<google_sheet_name>')}
function getEmploymentDocTemplate() {return DriveApp.getFileById('<google_doc_file_id>') } // doc template for employment forms
function getEmploymentRepository() {return DriveApp.getFolderById('<google_drive_folder_id>')} // directory where employment forms are stored
