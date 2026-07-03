// ========================================================================================
// WARNING: Test Version Daily Sheet Cleanup Script - SYED - Updated 12:06 02/07/2026
// ========================================================================================
/**
 * Clean-up Script Syed Testing
 * Created by: Syed
 * * ⚠️ DISCLAIMER / IMPORTANT WARNING:
 * Before running this script, please make sure to copy and save 
 * your current version. Having a backup ensures you can easily 
 * revert to your original work if the script produces unexpected results.
 */
 function cleanUpSheetsUltimate() {
  var targetSheets = ["SRLD2", "T-Series 1", "T-Series 2", "T-Series 3", "RedLight", "WRR", "ELK"];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. CALCULATE EXACTLY TOMORROW'S DATE ---
  var localDateString = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  var dateParts = localDateString.split('-');
  
  var localDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
  localDate.setDate(localDate.getDate() + 1); // Mathematically add exactly 1 day
  
  var nextYear = localDate.getFullYear();
  var nextMonth = ("0" + (localDate.getMonth() + 1)).slice(-2);
  var nextDay = ("0" + localDate.getDate()).slice(-2);
  var nextDayStr = nextYear + "-" + nextMonth + "-" + nextDay; // Format: YYYY-MM-DD

  // Helper: Mathematically detects ANY shade of Black or Dark Grey
  function isTextBlack(hex) {
    if (!hex) return true;
    var c = hex.toLowerCase().trim();
    var darks = ['#000000', '#000', 'black', '#434343', '#333333', '#666666', '#444444', '#222222', '#595959'];
    if (darks.indexOf(c) !== -1) return true;
    
    if (c.indexOf('#') === 0 && c.length === 7) {
      var r = parseInt(c.slice(1,3), 16);
      var g = parseInt(c.slice(3,5), 16);
      var b = parseInt(c.slice(5,7), 16);
      if (r < 100 && g < 100 && b < 100) return true; 
    }
    return false;
  }

  for (var s = 0; s < targetSheets.length; s++) {
    var sheet = ss.getSheetByName(targetSheets[s]);
    if (!sheet) continue;

    // --- THE DOUBLE-CHECK LOOP ---
    var keepCleaning = true;
    var safetyCounter = 0; 

    while (keepCleaning && safetyCounter < 5) {
      keepCleaning = false; 
      safetyCounter++;

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) break;

      var data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var fonts = sheet.getRange(1, 1, lastRow, lastCol).getFontColors();
      
      var headerRowIndex = 1;
      var frameColIndex = -1;
      var actionColIndex = -1;

      // Find Core Structure Columns (Frame & Action Required)
      for (var r = 0; r < Math.min(10, data.length); r++) {
        for (var c = 0; c < data[r].length; c++) {
          var cellText = String(data[r][c]).toLowerCase().replace(/[^a-z]/g, '');
          if (cellText.includes("framenumber")) frameColIndex = c;
          if (cellText.includes("actionrequired") || cellText.includes("actionreq") || cellText === "action") actionColIndex = c;
        }
        if (frameColIndex !== -1 && actionColIndex !== -1) {
          headerRowIndex = r + 1;
          break;
        }
      }

      // Column K is the 11th column. In zero-based programming code arrays, Column K is index 10.
      var kColumnIndex = 10; 

      // Find Bottom Block boundary line
      var bottomBlockRow = lastRow + 1;
      for (var r = lastRow - 1; r >= 0; r--) {
        var rowStr = data[r].join("").toLowerCase().replace(/[^a-z]/g, ''); 
        if (rowStr.includes("reactivemainten") || rowStr.includes("reactuvemainten")) {
          bottomBlockRow = r + 1;
          break;
        }
      }

      // Process Rows Backwards safely
      for (var r = bottomBlockRow - 1; r > headerRowIndex; r--) {
        var rowIndex = r - 1;
        var rawRowText = data[rowIndex].join(" ").toUpperCase();
        
        if (rawRowText.includes("PLEASE DO NOT DELETE")) continue;

        var deleteThisRow = false;

        // DIRECT CHECK COLUMN K: If it contains "REMOVE", flag for deletion immediately!
        if (kColumnIndex < lastCol) {
          var kCellValue = String(data[rowIndex][kColumnIndex]).toUpperCase().trim();
          if (kCellValue.indexOf("REMOVE") !== -1) {
            deleteThisRow = true;
          }
        }

        // 1. Check Action Required for N/A
        if (!deleteThisRow && actionColIndex !== -1 && actionColIndex < lastCol) {
          var rawAction = String(data[rowIndex][actionColIndex]).toUpperCase().trim();
          var actionFont = fonts[rowIndex][actionColIndex]; 
          
          if ((rawAction === "N/A" || rawAction === "NA" || rawAction === "NOT APPLICABLE" || rawAction.indexOf("N/A") === 0) && isTextBlack(actionFont)) {
            deleteThisRow = true;
          }
        }

        // 2. Check for "No further issue"
        if (!deleteThisRow && (rawRowText.includes("NO FURTHER ISSUE") || rawRowText.includes("NO FURTHER ACTION"))) {
          var rowIsBlack = true;
          for (var c = 0; c < lastCol; c++) {
            if (String(data[rowIndex][c]).trim() !== "") {
              if (!isTextBlack(fonts[rowIndex][c])) rowIsBlack = false; 
              break; 
            }
          }
          if (rowIsBlack) deleteThisRow = true;
        }

        if (deleteThisRow) {
          sheet.deleteRow(r);
          keepCleaning = true; 
        } else {
          // Clear Frame Numbers
          if (frameColIndex !== -1 && frameColIndex < lastCol && safetyCounter === 1) {
            sheet.getRange(r, frameColIndex + 1).clearContent();
          }
        }
      }
    } // End of While Loop

    // --- 5. INSERT 10 ROWS & UPDATE DATES ---
    var finalLastRow = sheet.getLastRow();
    if (finalLastRow < 1) finalLastRow = 1;
    var finalData = sheet.getRange(1, 1, finalLastRow, lastCol).getDisplayValues();
    
    var finalBottomBlockRow = -1;
    var realLastDataRow = headerRowIndex;
    var pdndRowIndex = -1;

    for (var r = finalLastRow - 1; r >= 0; r--) {
      var fullRowString = finalData[r].join("").trim();
      var strippedString = fullRowString.toLowerCase().replace(/[^a-z]/g, '');
      
      if (strippedString.includes("reactivemainten") || strippedString.includes("reactuvemainten")) {
        finalBottomBlockRow = r + 1; 
      }
      if (strippedString.includes("pleasedonotdelete")) {
        pdndRowIndex = r + 1; 
      }
      if (fullRowString !== "") {
        if (finalBottomBlockRow === -1) {
          if (realLastDataRow === headerRowIndex && r + 1 > headerRowIndex) {
            realLastDataRow = r + 1;
          }
        }
      }
    }

    // Apply Tomorrow's Date
    if (pdndRowIndex !== -1) {
       var endOfActiveData = (finalBottomBlockRow !== -1 && finalBottomBlockRow > headerRowIndex) ? (finalBottomBlockRow - 1) : realLastDataRow;
       
       for (var r = pdndRowIndex; r < endOfActiveData; r++) { 
          var rowHasText = finalData[r].join("").trim() !== "";
          if (rowHasText || r === pdndRowIndex) {
             sheet.getRange(r + 1, 1).setValue(nextDayStr);
          }
       }
    }

    // --- Insert the 10 Rows and Apply Logic/Formatting ---
    var insertStartRow;
    
    if (finalBottomBlockRow !== -1 && finalBottomBlockRow > headerRowIndex) {
       insertStartRow = finalBottomBlockRow;
       sheet.insertRowsBefore(insertStartRow, 10);
    } else {
       insertStartRow = realLastDataRow + 1;
       sheet.insertRowsAfter(realLastDataRow, 10);
    }

    // Identify the row right above the newly inserted rows to act as our "template"
    var templateRow = insertStartRow - 1;

    // Define ranges, ensuring the templateRow actually exists (is strictly greater than 0)
    if (templateRow > 0) {
      var sourceRange = sheet.getRange(templateRow, 1, 1, lastCol);
      var targetRange = sheet.getRange(insertStartRow, 1, 10, lastCol);

      // Copy everything (Data Validations, Conditional Formatting, Backgrounds, Borders)
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

      // Clear the actual text/values from the new rows so they are ready for fresh input
      targetRange.clearContent();
    }
  }
  
  SpreadsheetApp.getUi().alert("Complete! Auto-loop cleared EVERY target row, inserted 10 formatted rows, and dates accurately rolled over to tomorrow (" + nextDayStr + ").");
}