function onInstall() {
    onOpen();
  }
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Metabase')
      .addItem('Import Question Trigger', 'automateImport')
      .addToUi();
  }
  
  function automateImportLOL() {
    var currentDate = new Date();
    var previousDate = new Date();
    previousDate.setDate(currentDate.getDate() - 1);
  
    var sheetName = 'D' + previousDate.getDate();
  
    var status = getQuestionAsCSV("QUESTION_NUMBER", sheetName);
    if (status.success === true) {
      Logger.log('QUESTION_NUMBER imported successfully to ' + sheetName);
    } else {
      Logger.log('Failed to import question QUESTION_NUMBER to ' + sheetName + ': ' + status.error);
    }
  }
  
  function getQuestionAsCSV(metabaseQuestionNum, sheetName) {
      var scriptProp = PropertiesService.getScriptProperties();
      var baseUrl = scriptProp.getProperty('BASE_URL');
      var username = scriptProp.getProperty('USERNAME');
      var password = scriptProp.getProperty('PASSWORD');
      var token = scriptProp.getProperty('TOKEN');
    
      if (!token) {
        token = getToken(baseUrl, username, password);
        scriptProp.setProperty('TOKEN', token);
      }
    
      status = getQuestionAndFillSheet(baseUrl, token, metabaseQuestionNum, sheetName);
      return status;
  }
  
  function getToken(baseUrl, username, password) {
      var sessionUrl = baseUrl + "api/session";
      var options = {
        "method": "post",
        "headers": {
          "Content-Type": "application/json"
        },
        "payload": JSON.stringify({
          username: username,
          password: password
        })
      };
      var response;
      try {
        response = UrlFetchApp.fetch(sessionUrl, options);
      } catch (e) {
        throw (e);
      }
      var token = JSON.parse(response).id;
      return token;
    }
  
    function getQuestionAndFillSheet(baseUrl, token, metabaseQuestionNum, sheetName) {
      var questionUrl = baseUrl + "api/card/" + metabaseQuestionNum + "/query/csv";
    
      var options = {
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
          "X-Metabase-Session": token
        },
        "muteHttpExceptions": true
      };
    
      var response;
      try {
        response = UrlFetchApp.fetch(questionUrl, options);
      } catch (e) {
        return {
          'success': false,
          'error': e
        };
      }
      var statusCode = response.getResponseCode();
    
      if (statusCode == 200 || statusCode == 202) {
        var values = Utilities.parseCsv(response.getContentText());
        try {
          fillSheet(values, sheetName);
          return {
            'success': true
          };
        } catch (e) {
          return {
            'success': false,
            'error': e
          };
        }
      } else if (statusCode == 401) {
        var scriptProp = PropertiesService.getScriptProperties();
        var username = scriptProp.getProperty('USERNAME');
        var password = scriptProp.getProperty('PASSWORD');
    
        var token = getToken(baseUrl, username, password);
        scriptProp.setProperty('TOKEN', token);
        var e = "Error: Could not retrieve question. Metabase says: '" + response.getContentText() + "'. Please try again in a few minutes.";
        return {
          'success': false,
          'error': e
        };
      } else {
        var e = "Error: Could not retrieve question. Metabase says: '" + response.getContentText() + "'. Please try again later.";
        return {
          'success': false,
          'error': e
        };
      }
    }
    
    function fillSheet(values, sheetName) {
      var colLetters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ", "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ", "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ", "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ"];
    
      var sheet;
      if (sheetName == false) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      } else {
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      }
    
      sheet.clear({
        contentsOnly: true
      });
    
      var rows = values;
      var header = rows[0];
      var minCol = colLetters[0];
      var maxCol = colLetters[header.length - 1];
      var minRow = 1;
      var maxRow = rows.length;
      var range = sheet.getRange(minCol + minRow + ":" + maxCol + maxRow);
      range.setValues(rows);
    }
  