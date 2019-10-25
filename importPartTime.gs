var partTimeSheet = ""
var months = ['January', 'Febuary', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

function importPartTime(child, month) {
  /*
  * Function to create a new invoice for a Full Time child using the full time template defined by the above id...
  */
  var currentYear = new Date().getFullYear();
  var startDate = new Date(currentYear, month);
  var endDate = new Date(currentYear, month+1);

  // Get the target sheets...
  var templateSS = SpreadsheetApp.openById(partTimeSheet);
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

  // Get the Defualt times for this child...
  var monMorn = child[3];
  var monAfter = child[4];
  var monDay = child[5];

  var tueMorn = child[6];
  var tueAfter = child[7];
  var tueDay = child[8];

  var wedMorn = child[9];
  var wedAfter = child[10];
  var wedDay = child[11];

  var thurMorn = child[12];
  var thurAfter = child[13];
  var thurDay = child[14];

  var friMorn = child[15];
  var friAfter = child[16];
  var friDay = child[17];

  // Set the name of this new sheet..
  newTemplateSheet.setName(childName+' - '+months[month])

  // Set some of the static values for the children...
  newTemplateSheet.getRange('D3').setValue(childName);
  newTemplateSheet.getRange('D4').setValue(months[month]);
  newTemplateSheet.getRange('C2').setValue(email);
  newTemplateSheet.getRange('B46').setValue(cph);

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
      newTemplateSheet.getRange('B'+startRow).setValue(monMorn);
      newTemplateSheet.getRange('D'+startRow).setValue(monAfter);
      newTemplateSheet.getRange('F'+startRow).setValue(monDay);
    }

    // Check tuesday...
    if (startDate.getDay() == 2) {
      newTemplateSheet.getRange('B'+startRow).setValue(tueMorn);
      newTemplateSheet.getRange('D'+startRow).setValue(tueAfter);
      newTemplateSheet.getRange('F'+startRow).setValue(tueDay);
    }

    // Check wednesday...
    if (startDate.getDay() == 3) {
      newTemplateSheet.getRange('B'+startRow).setValue(wedMorn);
      newTemplateSheet.getRange('D'+startRow).setValue(wedAfter);
      newTemplateSheet.getRange('F'+startRow).setValue(wedDay);
    }

    // Check thursday...
    if (startDate.getDay() == 4) {
      newTemplateSheet.getRange('B'+startRow).setValue(thurMorn);
      newTemplateSheet.getRange('D'+startRow).setValue(thurAfter);
      newTemplateSheet.getRange('F'+startRow).setValue(thurDay);
    }

    // Check friday...
    if (startDate.getDay() == 5) {
      newTemplateSheet.getRange('B'+startRow).setValue(friMorn);
      newTemplateSheet.getRange('D'+startRow).setValue(friAfter);
      newTemplateSheet.getRange('F'+startRow).setValue(friDay);
    }

    // If it is a weekend, then lets set the background and wipe the values from the cell so it is clear that it is a weeekend...
    if (startDate.getDay() == 6 || startDate.getDay() == 0){
      var range = 'A'+startRow+':'+'J'+startRow
      newTemplateSheet.getRange(range).setBackground('#f1f1f1')
      newTemplateSheet.getRange(range).setDataValidation(null)
      newTemplateSheet.getRange(range).setValue('');

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
