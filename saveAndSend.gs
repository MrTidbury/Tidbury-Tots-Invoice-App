function saveAndSend() {

  var thisSS = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisSS.getActiveSheet();

  var email = thisSheet.getRange('C2').getValue();
  var month = thisSheet.getRange('D4').getValue();
  var name = thisSheet.getRange('D3').getValue();
  var year = new Date().getFullYear();
  var fileName = name+' - '+month+' '+year+'.pdf';
  var subject = 'Tidbury Tots MONTH invoice'.replace('MONTH', month);

  var body = ''
  body += '<p>Hi,</p>'
  body += '<p>Please find the attached invoice for NAME.</p>'.replace('NAME', name)
  body += '<p>As a reminder, please ensure that the funds have cleared my account by the due date or late fees may be incurred.</p>'
  body += '<p>Thanks,</p>'
  body += '<p>Judy</p>'

  var url = 'https://docs.google.com/spreadsheets/d/SS_ID/export?'.replace('SS_ID', thisSS.getId());

  const exportOptions =
    'exportFormat=pdf&format=pdf' + // export as pdf / csv / xls / xlsx
    '&size=A4' + // paper size legal / letter / A4
    '&portrait=true' + // orientation, false for landscape
    '&scale=3'+ // Fit to width
    '&source=labnol' + // fit to page width, false for actual size
    '&sheetnames=false&printtitle=false' + // hide optional headers and footers
    '&pagenumbers=false&gridlines=false' + // hide page numbers and gridlines
    '&fzr=false' + // do not repeat row headers (frozen rows) on each page
    '&top_margin=0.10&bottom_margin=0.10&left_margin=0.10&right_margin=0.10'+
    '&horizontal_alignment=CENTER' +
    '&gid='; // the sheet's Id

  var token = ScriptApp.getOAuthToken();
  var auth = 'Bearer '+token;

  // Convert individual worksheets to PDF
  const response = UrlFetchApp.fetch(url + exportOptions + thisSheet.getSheetId(), {
    headers: {
      Authorization: auth
    }
  });

  pdfBlob =  response.getBlob().setName(fileName);

  var invoiceFolder = DriveApp.getFolderById('')

  var nameFolders = invoiceFolder.getFoldersByName(name)
  var folder = null;
  while (nameFolders.hasNext()) {
    folder = nameFolders.next();
  }

  if (!folder) {
    folder = invoiceFolder.createFolder(name)
  }

  folder.createFile(pdfBlob)

  if (MailApp.getRemainingDailyQuota() > 0) {
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments: [pdfBlob]
    });
  }
}
