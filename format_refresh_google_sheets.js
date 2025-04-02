/** @OnlyCurrentDoc */

function Refresh() {
  var spreadsheet = SpreadsheetApp.getActive();

  // Format column AB in the relevant sheets
  var sheetsToFormat = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4'];
  sheetsToFormat.forEach(function (sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.getRange('AB:AB').setNumberFormat('d"/"mm"/"yyyy" "hh":"mm" "am/pm');
    } else {
      Logger.log("Sheet '" + sheetName + "' does not exist.");
    }
  });

  // Insert and autofill columns for "Overview Dashboard"
  var dashboardSheet = spreadsheet.getSheetByName("Overview Dashboard"); // Reference the sheet by name

  if (!dashboardSheet) {
    Logger.log("Sheet 'Overview Dashboard' does not exist.");
    return; // Exit the function if the sheet is not found
  }

  // Get the last column and the date in Row 2 of the last column
  var lastColumn = dashboardSheet.getLastColumn();
  var lastDate = dashboardSheet.getRange(2, lastColumn).getValue(); // Get the date in Row 2 of the last column

  // Check if the lastDate is a valid date
  if (!(lastDate instanceof Date)) {
    Logger.log("Row 2 of the last column does not contain a valid date.");
    return; // Exit the function if the value is not a date
  }

  // Calculate the difference between the last date and the current date
  var currentDate = new Date();
  var timeDifference = lastDate.getTime() - currentDate.getTime(); // Difference in milliseconds
  var daysDifference = Math.ceil(timeDifference / (1000 * 60 * 60 * 24)); // Convert to days

  // Run the function only if the difference is less than 7 days
  if (daysDifference >= 7) {
    Logger.log("No need to add columns. The difference is more than or equal to 7 days.");
    return; // Exit if the condition is not met
  }

  Logger.log("Adding 20 columns because the difference is less than 7 days.");

  // Insert 20 new columns after the last column
  dashboardSheet.insertColumnsAfter(lastColumn, 20);

  // Get the range of the last column before the new columns
  var lastColumnRange = dashboardSheet.getRange(1, lastColumn, dashboardSheet.getLastRow());

  // Autofill formulas for the new columns
  for (var i = 1; i <= 20; i++) {
    // Copy formula from the last column to the new column
    var newColumnIndex = lastColumn + i;
    var newColumnRange = dashboardSheet.getRange(1, newColumnIndex, dashboardSheet.getLastRow());
    lastColumnRange.copyTo(newColumnRange);
  }

  // Autofill dates incrementally in Row 2
  for (var j = 1; j <= 20; j++) {
    var nextColumnIndex = lastColumn + j;
    var nextDate = new Date(lastDate); // Clone the last date
    nextDate.setDate(nextDate.getDate() + j); // Increment the date
    dashboardSheet.getRange(2, nextColumnIndex).setValue(nextDate); // Set the new date
  }
}
