var templateSheetId = ""
var months = ['January', 'Febuary', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

function importFulltime(child, month) {
  /*
  * Function to create a new invoice for a Full Time child using the full time template defined by the above id...
  */
  var currentYear = new Date().getFullYear();
  var startDate = new Date(currentYear, month);
  var endDate = new Date(currentYear, month+1);

  // Get the target sheets...
  var templateSS = SpreadsheetApp.openById(templateSheetId);
  var thisSS = SpreadsheetApp.getActiveSpreadsheet();

  // Open the first sheet from the target spreadsheet into memory...
  var templateSheet = templateSS.getSheets()[0];

  // Copy that into the current Spreadsheet, this function returns the sheet into memory...
  var newTemplateSheet = templateSheet.copyTo(thisSS);


  // Start on the First date collum and iterate the month so we can format it correctly.
  var startRow = 9
  var endRow = 40

  //Setup all of the needed vars for generating the spreadsheet...

  // Generic child params...
  var childName = child[0];
  var email = child[1];
  var cph = child[2];
  var hasMeals = child[3];

  // Get the Defualt times for this child...
  var monStart = child[4];
  var monEnd = child[5];
  var tueStart = child[6];
  var tueEnd = child[7];
  var wedStart = child[8];
  var wedEnd = child[9];
  var thurStart = child[10];
  var thurEnd = child[11];
  var friStart = child[12];
  var friEnd = child[13];

  // Set the name of this new sheet..
  newTemplateSheet.setName(childName+' - '+months[month])

  // Set some of the static values for the children...
  newTemplateSheet.getRange('D3').setValue(childName);
  newTemplateSheet.getRange('D4').setValue(months[month]);
  newTemplateSheet.getRange('C2').setValue(email);
  newTemplateSheet.getRange('B41').setValue(cph);

  // We will count how many meals that this child will have here...
  var mealCount = 0;

  // Iterate for every day in the month...
  while (startDate < endDate) {

    // Construct the Date String...
    var day = startDate.getDate();
    var month = startDate.getMonth() + 1;
    var year = startDate.getFullYear();
    var dateString = month + '/' + day + '/' + year;

    // Stamp the date...
    newTemplateSheet.getRange('A'+startRow).setValue(dateString)

    // Check monday...
    if (startDate.getDay() == 1) {
      newTemplateSheet.getRange('B'+startRow).setValue(monStart);
      newTemplateSheet.getRange('C'+startRow).setValue(monEnd);

      if (monStart || monEnd) {
        mealCount += 1;
      }
    }

    // Check tuesday...
    if (startDate.getDay() == 2) {
      newTemplateSheet.getRange('B'+startRow).setValue(tueStart);
      newTemplateSheet.getRange('C'+startRow).setValue(tueEnd);

      if (tueStart || tueEnd) {
        mealCount += 1;
      }
    }

    // Check wednesday...
    if (startDate.getDay() == 3) {
      newTemplateSheet.getRange('B'+startRow).setValue(wedStart);
      newTemplateSheet.getRange('C'+startRow).setValue(wedEnd);

      if (wedStart || wedEnd) {
        mealCount += 1;
      }
    }

    // Check thursday...
    if (startDate.getDay() == 4) {
      newTemplateSheet.getRange('B'+startRow).setValue(thurStart);
      newTemplateSheet.getRange('C'+startRow).setValue(thurEnd);

      if (thurStart || thurEnd) {
        mealCount += 1;
      }
    }

    // Check friday...
    if (startDate.getDay() == 5) {
      newTemplateSheet.getRange('B'+startRow).setValue(friStart);
      newTemplateSheet.getRange('C'+startRow).setValue(friEnd);

      if (friStart || friEnd) {
        mealCount += 1;
      }
    }

    // If it is a weekend, then lets set the background and wipe the values from the cell so it is clear that it is a weeekend...
    if (startDate.getDay() == 6 || startDate.getDay() == 0){
      var range = 'A'+startRow+':'+'E'+startRow
      newTemplateSheet.getRange(range).setBackground('#f1f1f1')
      newTemplateSheet.getRange(range).setValue('')

    }

    // Iterate the iterees...
    startDate.setDate(startDate.getDate() + 1);
    startRow += 1;

    // If it is the last day, calculate how many extra rows we have, the template has 31 rows, but we dont need all those for some months...
    if (sameDay(startDate, endDate)){
      var rowsToHide = endRow - startRow;
      if (rowsToHide > 0) {
        newTemplateSheet.hideRows(startRow, rowsToHide)
      }
    }

  }

  // If this child has meals then stamp the meal count to the sheet so we can coun them...
  if (hasMeals) {
    newTemplateSheet.getRange('B46').setValue(mealCount);
  }

  // Hide the gridlines and show the sheet as the active sheets...
  newTemplateSheet.setHiddenGridlines(true)
  thisSS.setActiveSheet(newTemplateSheet)

}

// Helper function to check if date is on the same day...
function sameDay(d1, d2) {
  return d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate();
}
