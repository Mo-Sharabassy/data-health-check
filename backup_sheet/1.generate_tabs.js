function generateTabs() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Sheet1'); // Change 'Sheet1' to the name of your desired sheet
    
    for (var i = 14; i <= 30; i++) {
      var tabName = 'D' + i;
      spreadsheet.insertSheet(tabName);
    }
  }