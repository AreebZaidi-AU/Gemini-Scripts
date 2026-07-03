/**
 * Weekend Script Syed Testing
 * Created by: Syed
 * * ⚠️ DISCLAIMER / IMPORTANT WARNING:
 * Before running this script, please make sure to copy and save 
 * your current version. Having a backup ensures you can easily 
 * revert to your original work if the script produces unexpected results.
 */
function prepareSheetsForNext3Days() {
  const targetSheets = ["SRLD2", "T-Series 1", "T-Series 2", "T-Series 3", "RedLight", "WRR", "ELK"];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  targetSheets.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log("Skipping: Sheet '" + sheetName + "' not found.");
      return; 
    }
    
    Logger.log("Processing: " + sheetName);

    // 1. Find the bottom section 
    let bottomSectionRow = findBottomSectionStart(sheet);
    if (bottomSectionRow === -1) {
      Logger.log("Skipping " + sheetName + ": Could not find bottom section.");
      return; 
    }

    // 2. Find the exact end of the top block
    let endOfFirstBlock = -1;
    let data = sheet.getDataRange().getValues();
    
    for (let i = bottomSectionRow - 2; i >= 0; i--) {
      if (data[i] && data[i].join("").trim() !== "") {
        endOfFirstBlock = i + 1; 
        break;
      }
    }

    if (endOfFirstBlock === -1) endOfFirstBlock = bottomSectionRow - 1; 

    // Define the exact size of the block (Starting from Row 1)
    let numRowsInBlock = endOfFirstBlock;
    let lastCol = sheet.getLastColumn();
    
    if (numRowsInBlock < 1 || lastCol < 1) return; 
    
    // The original block of data
    let sourceRange = sheet.getRange(1, 1, numRowsInBlock, lastCol);

    // 3. Smart Date Finder (Finds WHERE your Date Heading is FIRST)
    let dateOffset = -1;
    let originalDateObj = null;
    for (let r = 1; r <= 5; r++) {
      let val = sheet.getRange(r, 1).getValue();
      if (val instanceof Date) {
        dateOffset = r - 1; 
        originalDateObj = val;
        break;
      } else {
        let parsed = new Date(val);
        if (val && !isNaN(parsed.getTime()) && val.toString().length > 6) {
          dateOffset = r - 1;
          originalDateObj = parsed;
          break;
        }
      }
    }

    // 4. TIMEZONE-SAFE CALENDAR LOGIC
    let timeZone = ss.getSpreadsheetTimeZone();
    let todayStr = Utilities.formatDate(new Date(), timeZone, "MM/dd/yyyy");
    let localToday = new Date(todayStr);
    localToday.setHours(0, 0, 0, 0); 
    
    let tomorrow = new Date(localToday.getTime());
    tomorrow.setDate(localToday.getDate() + 1);

    // Update the FIRST (Original) Block strictly to Tomorrow
    if (dateOffset !== -1) {
      sheet.getRange(1 + dateOffset, 1).setValue(tomorrow);
    }

    // 5. Create a 10-row gap right above the bottom section FIRST
    let initialBottomRow = findBottomSectionStart(sheet);
    sheet.insertRowsBefore(initialBottomRow, 10);
    sheet.getRange(initialBottomRow, 1, 10, sheet.getMaxColumns()).clearFormat().clearDataValidations();

    // 6. Duplicate the block TWO more times
    for (let copyIndex = 1; copyIndex <= 2; copyIndex++) {
      
      let currentBottomRow = findBottomSectionStart(sheet);
      
      let rowsToInsert = numRowsInBlock + 10;
      sheet.insertRowsBefore(currentBottomRow, rowsToInsert);
      
      let pasteStartRow = currentBottomRow;
      let pasteRange = sheet.getRange(pasteStartRow, 1, numRowsInBlock, lastCol);
      sourceRange.copyTo(pasteRange);
      
      let gapRowStart = pasteStartRow + numRowsInBlock;
      sheet.getRange(gapRowStart, 1, 10, sheet.getMaxColumns()).clearFormat().clearDataValidations();
      
      // ==========================================
      // THE FIX: IMMEDIATE CLEANUP
      // Wipes the top rows of the freshly pasted block blank IMMEDIATELY.
      // This guarantees the script doesn't confuse them for the bottom section in the next loop!
      // ==========================================
      if (dateOffset > 0) {
        sheet.getRange(pasteStartRow, 1, dateOffset, sheet.getMaxColumns()).clear();
      }
      
      // Update the Date Banner for the copied block right away
      if (dateOffset !== -1) {
        let newDateCell = sheet.getRange(pasteStartRow + dateOffset, 1);
        let blockDate = new Date(tomorrow.getTime());
        blockDate.setDate(tomorrow.getDate() + copyIndex); 
        newDateCell.setValue(blockDate);
      }
    }

  });
  
  Logger.log("All sheets updated successfully!");
}

// ==========================================
// UPDATED FINDER
// ==========================================
function findBottomSectionStart(sheet) {
  let values = sheet.getRange("A:A").getValues();
  
  // Starts checking at Row 6 to safely skip past "Date+" at the top of the page
  for (let i = 5; i < values.length; i++) {
    let cellText = values[i][0].toString().toLowerCase();
    if (cellText.includes("reactive maintenance") || cellText.includes("date+") || cellText.includes("please do not delete")) {
      return i + 1; 
    }
  }
  return -1; 
}