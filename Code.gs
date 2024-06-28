var spreadsheetID = "id-here"; // Spreadsheet ID from URL goes here

// Clears availability ranges and increments the dates by one week
function clearAvailability() {
      var activeSpreadsheet = SpreadsheetApp.openById(spreadsheetID);
      var sheetActive = activeSpreadsheet.getSheetByName("Availability"); // Select Availability sheet

      // Clear availability of every employee
      sheetActive.getRange('B3:G9').clearContent();
      sheetActive.getRange('B12:G18').clearContent();

      // Update the dates listed for next week of availability
      var dateCell = sheetActive.getRange('A8');
      var dateValue = dateCell.getValue();
      dateCell.setValue('="'+ Utilities.formatDate(new Date(dateValue), "GMT-5", "MM/dd/yyyy") +'"+7');

      dateCell.setNumberFormat("ddd m/d");

    }

// Deletes the old schedule and creates a new one
function updateSchedules() {
  var activeSpreadsheet = SpreadsheetApp.openById(spreadsheetID);
  var sheetActive = activeSpreadsheet.getSheetByName("Template Sheet"); // Select Template Sheet
  var sheetList = activeSpreadsheet.getSheets();

  activeSpreadsheet.setActiveSheet(sheetActive);
  activeSpreadsheet.moveActiveSheet(sheetList.length); // Move Template Sheet to end of sheet list

  // Delete old schedule sheet
  sheetActive = sheetList[1];
  activeSpreadsheet.deleteSheet(sheetActive);

  // Get dates for new schedule sheet
  sheetActive = activeSpreadsheet.getSheetByName("Availability");
  
  var dateValue = sheetActive.getRange('A8').getValue();
  var startDate = sheetActive.getRange('A3').getDisplayValue();
  var endDate = sheetActive.getRange('A9').getDisplayValue();

  var newSheetName = "Game Schedule " + startDate.substring(4) + "-" + endDate.substring(4);

  // Duplicate Template Sheet and rename the duplicated sheet
  sheetActive = activeSpreadsheet.getSheetByName("Template Sheet");
  activeSpreadsheet.setActiveSheet(sheetActive);
  activeSpreadsheet.duplicateActiveSheet();
  activeSpreadsheet.renameActiveSheet(newSheetName);

  // Set dates in new schedule sheet
  sheetActive = activeSpreadsheet.getSheetByName(newSheetName);
  sheetActive.showSheet();
  var newDateCell = sheetActive.getRange('A31');
  newDateCell.setValue(dateValue);
  newDateCell.setNumberFormat("ddd m/d");

  sheetActive.protect();
}