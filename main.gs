function arrangeFormData(e) { // start by event
  const sheet = e.source.getActiveSheet()
  const sheetName = sheet.getName()
  const lastRow = getLastRow(sheet)
  let assertTable = []
  let markers // markers map, where map and marker from gDoc file are mapped to replace
  let docTemplate // template doc file
  let repository // gDrive folder
  let workerEmail
  switch(sheetName) {
    case getWorkerSheet().getName():
      assertTable.push(insertValue(sheet, sheet.getRange(familyStateColWorkerSheet + lastRow), familyStateMarker))
      assertTable.push(insertValue(sheet, sheet.getRange(coursesStateColWorkerSheet + lastRow), coursesStateMarker))
      assertTable.push(insertValue(sheet, sheet.getRange(jobStateColWorkerSheet + lastRow), jobStateMarker))
      markers = workerMarkers
      docTemplate = getWorkerDocTemplate()
      repository = getWorkerRepository()
      workerEmail = sheet.getRange(emailColWorkerSheet + lastRow).getValue()
      break
    case getEmploymentSheet().getName():
      assertTable.push(insertValue(sheet, sheet.getRange(schoolStateColEmploymentSheet + lastRow), schoolStateMarker))
      assertTable.push(insertValue(sheet, sheet.getRange(coursesStateColEmploymentSheet + lastRow), coursesStateMarker))
      assertTable.push(insertValue(sheet, sheet.getRange(jobStateColEmploymentSheet + lastRow), jobStateMarker))
      markers = employmentMarkers
      docTemplate = getEmploymentDocTemplate()
      repository = getEmploymentRepository()
      workerEmail = sheet.getRange(emailColEmploymentSheet + lastRow).getValue()
      break
    default:
      console.warn('"' + sheetName + '" has not assigned any scripts')
      break
  }
  if (!assertTable.includes(false) && assertTable.length > 0) { // if any data has been changed
    sheet.getRange(timestampCol + lastRow).setValue(getToday()) // set timestamp
    let doc = createDocument(lastRow, sheet, markers, docTemplate, repository) // generate paper version of form
    sendMail(workerEmail, 'Kwestionariusz', DriveApp.getFileById(doc.getId()), sheet, lastRow)
  }
  else 
    console.warn('document not generated')
}

function insertValue(sheet, range, marker) {
  if (range.getValue().includes(newLineDelimeter) || range.getValue().includes(delimeter)) {
    let splittedCellValues = range.getValue().split(newLineDelimeter)
    range.setValue('')
    console.info('cell ' + range.getA1Notation() + ' to split')
    splittedCellValues.forEach(function(splittedCellValue) {
      console.log('value "' + splittedCellValue + '" splitting')
      splittedCellValue.split(delimeter).forEach(function(elem, i) {
        let currRange = sheet.getRange(getColumnToLetter(range.getColumn() + i) + range.getRow())
        let currValue = currRange.getValue()
        console.info('value "' + elem + "' to split")
        if (currValue === '')
          currRange.setValue(elem.toString())
        else 
          currRange.setValue(currValue + '\n' + elem.toString())
      })
    })
    range.setValue(range.getValue().slice(0, -1))
    return true
  }
  else if (range.getValue() === marker) { // means that this data has not been filled in form, empty
    range.setValue('')
    console.info('cell ' + range.getA1Notation() + ' cleared')
    return true
  }  
  else // means that this data has been already "inserted", splitted etc, just iterated by this function
    console.info('cell ' + range.getA1Notation() + ' skipped')
    return false
}

function createDocument(row, sheet, markers, docTemplate, repository) {
  let urlRange = sheet.getRange(getColumnToLetter(sheet.getLastColumn()) + row)
  urlRange.setValue('Tworzenie dokumentu...')
  let docTitle = sheet.getRange(workerNameCol + row).getValue() + ' (' + sheet.getName() + ')'
  const docNew = docTemplate.makeCopy(docTitle, repository)
  var doc = DocumentApp.openById(docNew.getId())
  var body = doc.getBody()
  for (const [column, marker] of Object.entries(markers)) {
    let text = sheet.getRange(column + row).getValue()
    if (text === '')
      text = '-'
    body.replaceText(marker, text)
    console.log(marker + " in document replaced to '" + text + "' from " + column + row) 
  }
  doc.saveAndClose()
  addUrlToSheet(doc.getUrl(), urlRange, docTitle)
  return doc
}

function addUrlToSheet(url, urlRange, title){
  var urlText = SpreadsheetApp.newRichTextValue().setText(title).setLinkUrl(url).build()
  urlRange.setRichTextValue(urlText)
  console.info('document saved in ' + url)
}

function sendMail(emailRecipent, subject, file, sheet, row) {
  try {
    MailApp.sendEmail(emailRecipent, sheet.getName(), '', {
      name: subject,
      noReply: true,
      attachments: [file.getAs(MimeType.PDF)]
    })
    console.info('document sent to ' + emailRecipent)
  }
  catch(error) {
    if (error.message.startsWith('Invalid email')) {
      console.error('incorrect email, note added to sheet')
      sheet.getRange(workerNameCol + row).setNote('Mail z kwestionariuszem nie został wysłany, wprowadzony e-mail był nieprawidłowy')
    }
    else
      console.error(error.name + ': ' + error.message)
  }
}
