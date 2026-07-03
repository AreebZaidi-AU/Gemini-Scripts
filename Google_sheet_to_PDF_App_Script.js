function exportSpecificSheetsToPDF() {
  // =========================================================================
  // 1. Output Folder ID
  // =========================================================================
  const FOLDER_ID = "1MQUrzQ7knjEt-HQC8qbH05xk7RWyrFQk";

  // =========================================================================
  // 2. Define our two groups of sheets
  // =========================================================================
  // Group A: These will be merged into ONE single PDF
  const sheetsToMerge = [
    "SRLD2",
    "T-series 1",
    "T-series 2",
    "T-series 3",
    "Redlight"
  ];
  
  // Group B: These will remain individual PDFs
  const sheetsToKeepIndividual = [
    "WRR",
    "ELK"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  
  // Fetch the name of the spreadsheet and replace the specific text
  const rawName = ss.getName(); 
  const ssName = rawName.replace("Health & Image Review", "H&I Review");
  
  // Connect to the destination folder
  let folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: Could not find the folder. Please double-check your folder permissions.");
    return;
  }

  const token = ScriptApp.getOAuthToken();
  const options = {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };

  // =========================================================================
  // PART 1: EXPORT THE MERGED PDF ("SRLD")
  // =========================================================================
  // Create a temporary spreadsheet in the background
  const tempSs = SpreadsheetApp.create("Temp_Export_File");
  const tempSsId = tempSs.getId();

  // Copy each targeted sheet into the temporary spreadsheet
  sheetsToMerge.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const isHidden = sheet.isSheetHidden();
      if (isHidden) sheet.showSheet();

      const copiedSheet = sheet.copyTo(tempSs);
      copiedSheet.setName(sheetName); // Remove the "Copy of " prefix

      if (isHidden) sheet.hideSheet();
    } else {
      Logger.log("Merge Skipped: Sheet not found - " + sheetName);
    }
  });

  // Delete the blank "Sheet1" that comes by default with new spreadsheets
  const defaultSheet = tempSs.getSheetByName("Sheet1");
  if (defaultSheet) tempSs.deleteSheet(defaultSheet);
  
  SpreadsheetApp.flush();

  // Export the ENTIRE temporary workbook. 
  // Note: sheetnames=true is turned ON so the tab name appears as a heading on the PDF.
  const mergedUrl = "https://docs.google.com/spreadsheets/d/" + tempSsId + "/export?" +
              "format=pdf&" +
              "size=A4&" +             
              "portrait=true&" +       
              "fitw=true&" +           
              "gridlines=false&" +     
              "printtitle=false&" +
              "sheetnames=true&" +     // <--- Shows tab headings
              "pagenum=false&" +
              "cb=" + new Date().getTime(); 

  const mergedResponse = UrlFetchApp.fetch(mergedUrl, options);
  
  if (mergedResponse.getResponseCode() === 200) {
    const mergedFileName = ssName + "_SRLD.pdf";
    const blob = mergedResponse.getBlob().setName(mergedFileName);
    folder.createFile(blob);
  }

  // Delete the temporary spreadsheet to keep your Drive clean
  DriveApp.getFileById(tempSsId).setTrashed(true);


  // =========================================================================
  // PART 2: EXPORT THE INDIVIDUAL PDFs ("WRR" & "ELK")
  // =========================================================================
  sheetsToKeepIndividual.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      const isHidden = sheet.isSheetHidden();
      if (isHidden) sheet.showSheet();

      ss.setActiveSheet(sheet);
      SpreadsheetApp.flush(); 

      const sheetId = sheet.getSheetId();
      
      // sheetnames=false is used here so individual pages don't get the extra heading
      const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?" +
                  "format=pdf&" +
                  "size=A4&" +             
                  "portrait=true&" +       
                  "fitw=true&" +           
                  "gridlines=false&" +     
                  "printtitle=false&" +
                  "sheetnames=false&" +    
                  "pagenum=false&" +
                  "gid=" + sheetId +
                  "&cb=" + new Date().getTime(); 

      const response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() === 200) {
        const pdfFileName = ssName + "_" + sheetName + ".pdf";
        const blob = response.getBlob().setName(pdfFileName);
        folder.createFile(blob);
      } else {
        Logger.log("Failed to export: " + sheetName);
      }

      if (isHidden) sheet.hideSheet();
      
    } else {
      Logger.log("Skipped: Sheet not found - " + sheetName);
    }
  });
  
  SpreadsheetApp.getUi().alert("Export complete! Check your Google Drive folder.");
}