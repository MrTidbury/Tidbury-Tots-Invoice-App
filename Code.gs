var dataSheetId = ''

var ftSheetNum = 0
var ftRange = 'A4:N18'

var ptSheetNum = 1
var ptRange = 'A4:S18'

var months = ['January', 'Febuary', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

var fullTimeChildren = []
var partTimeChildren = []

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  //Initiate Menu.
  ui.createMenu('Tidbury Tots')
      .addItem('Create New Invoice', 'importChild')
      .addSeparator()
      .addItem('Save and Send', 'saveAndSend')
      .addToUi();
}

function getDataFromSheet(sheetId, range, sheetNum) {
 // Get the Spreadsheet that we store the Childrens Default Data in...
  var dataSS = SpreadsheetApp.openById(sheetId);
  // Open the first sheet as that is all we care about...
  var dataSheet = dataSS.getSheets()[sheetNum];
  var ftData = dataSheet.getRange(range).getValues();
  return ftData
}

function importChild() {
  deleteRedundantSheets()
  var ui = SpreadsheetApp.getUi()

  var ftData = getDataFromSheet(dataSheetId, ftRange, ftSheetNum)
  for (var row in ftData) { fullTimeChildren.push(ftData[row])}


  var ptData = getDataFromSheet(dataSheetId, ptRange, ptSheetNum)
  for (var row in ptData) { partTimeChildren.push(ptData[row])}

  var currentSheet = SpreadsheetApp.getActiveSpreadsheet()

  var htmlTemplate = HtmlService.createTemplateFromFile('importChildDialog')

  htmlTemplate.fullTimeChildren = ftData
  htmlTemplate.partTimeChildren = ptData
  htmlTemplate.months = months

  var html = htmlTemplate.evaluate().setWidth(700).setHeight(230);

  ui.showModalDialog(html, "Create New Invoice");

}

function instigateImport (childName, monthName) {

  var month = months.indexOf(monthName)
  var child = null;

  var fullTimeChildren = [];
  var partTimeChildren = [];

  var ftData = getDataFromSheet(dataSheetId, ftRange, ftSheetNum)
  for (var row in ftData) { fullTimeChildren.push(ftData[row])}


  var ptData = getDataFromSheet(dataSheetId, ptRange, ptSheetNum)
  for (var row in ptData) { partTimeChildren.push(ptData[row])}

  fullTimeChildren.forEach(function(row) { if (row[0] == childName) { child = row} });

  if (child) {
    importFulltime(child, month);
  } else {
    partTimeChildren.forEach(function(row) { if (row[0] == childName) { child = row} });
    if (child) {
      importPartTime(child, month);
    }
  }
}

function saveInvoice() {
  saveAndSend()
}

function deleteRedundantSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetName() != 'Home') {
      ss.deleteSheet(sheets[i]);
    }
  }
}
