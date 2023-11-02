function highlightDifferentTripsAndCopy() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheetByName("(metabase/4037)");
    var sheet2 = ss.getSheetByName("day_to_day_data");
    var editedTripsSheet = ss.getSheetByName("Edited Trips");
    var trips1 = sheet1.getDataRange().getValues();
    var trips2 = sheet2.getDataRange().getValues();
  
    // Clear existing formatting
    clearSheetFormatting(sheet1);
  
    // Clear existing contents and formatting in "Edited Trips" sheet starting from the second row
    clearSheetContentsAndFormatting(editedTripsSheet);
  
    for (var i = 1; i < trips1.length; i++) {
      var trip1 = trips1[i];
      var trip1Id = trip1[0];
  
      for (var j = 1; j < trips2.length; j++) {
        var trip2 = trips2[j];
        var trip2Id = trip2[0];
  
        if (trip1Id === trip2Id && areTripsDifferent(trip1, trip2)) {
          // Highlight cells and copy trips
          highlightAndCopyTrips(sheet1, sheet2, editedTripsSheet, i, j, trip1, trip2);
  
          break; // Stop looping through cells if a difference is found
        }
      }
    }
  }
  
  function clearSheetFormatting(sheet) {
    // Clear existing formatting
    sheet.getRange("A2:T").setBackground(null);
  }
  
  function clearSheetContentsAndFormatting(sheet) {
    // Clear existing contents and formatting in "Edited Trips" sheet starting from the second row
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear({ contentsOnly: true, formatOnly: true });
    }
  }
  
  function areTripsDifferent(trip1, trip2) {
    // Compare the trips based on their values and return true if they are different
    for (var i = 0; i < trip1.length; i++) {
      if (trip1[i].toString().trim() !== trip2[i].toString().trim()) {
        return true;
      }
    }
    return false;
  }
  
  function highlightAndCopyTrips(sheet1, sheet2, editedTripsSheet, i, j, trip1, trip2) {
    for (var k = 0; k < trip1.length; k++) {
      if (trip1[k].toString().trim() !== trip2[k].toString().trim()) {
        var cell1 = sheet1.getRange(i + 1, k + 1);
        cell1.setBackground("yellow");
        sheet1.getRange(i + 1, 1).setBackground("yellow"); // Highlight cell in column A
  
        copyTripWithFormatting(sheet2, editedTripsSheet, j);
        copyTripWithFormatting(sheet1, editedTripsSheet, i);
  
        break; // Stop looping through cells if a difference is found
      }
    }
  }
  
  function copyTripWithFormatting(sourceSheet, targetSheet, rowIndex) {
    var tripRange = sourceSheet.getRange(rowIndex + 1, 1, 1, sourceSheet.getDataRange().getNumColumns());
    var tripValues = tripRange.getValues();
    var tripFormats = tripRange.getBackgrounds();
  
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, 1, tripValues[0].length).setValues(tripValues);
    targetSheet.getRange(targetSheet.getLastRow(), 1, 1, tripFormats[0].length).setBackgrounds(tripFormats);
  }
  