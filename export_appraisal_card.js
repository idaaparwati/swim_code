function exportAppraisalPDF() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.includes("SWIM - Appraisal")) {
    SpreadsheetApp.getUi().alert("Please run this inside an Appraisal sheet.");
    return;
  }

  const studentName = sheet.getRange("D16").getDisplayValue().trim();
  const ageGroup = sheet.getRange("D17").getDisplayValue();
  const level = sheet.getRange("D18").getDisplayValue();

  if (!studentName) {
    SpreadsheetApp.getUi().alert("Please select student first.");
    return;
  }

  // ===== GET CENTER FROM FILE NAME =====
  const fileNameSheet = ss.getName();
  const parts = fileNameSheet.split(" - ");
  const center = parts.length > 1 ? parts.pop().trim() : "UNKNOWN";

  // ===== FILE NAME (UPDATED) =====
  const fileName =
    studentName +
    " - " +
    ageGroup +
    " Level " +
    level +
    ".pdf";

  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();

  const exportUrl =
    "https://docs.google.com/spreadsheets/d/" +
    ssId +
    "/export?" +
    "format=pdf" +
    "&gid=" + sheetId +
    "&range=B6:F65" +
    "&size=F4" +
    "&portrait=true" +
    "&fitw=true" +
    "&gridlines=false" +
    "&printtitle=false" +
    "&sheetnames=false" +
    "&pagenumbers=false" +
    "&top_margin=0.3" +
    "&bottom_margin=0.5" +
    "&left_margin=0.2" +
    "&right_margin=0.2";

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + token
    }
  });

  const blob = response.getBlob().setName(fileName);

  // ===== MASTER FOLDER =====
  const masterFolderName = "Master Appraisal";
  let masterFolder;

  const masterFolders = DriveApp.getFoldersByName(masterFolderName);

  if (masterFolders.hasNext()) {
    masterFolder = masterFolders.next();
  } else {
    masterFolder = DriveApp.createFolder(masterFolderName);
  }

  // ===== CENTER FOLDER =====
  let centerFolder;
  const centerFolders = masterFolder.getFoldersByName(center);

  if (centerFolders.hasNext()) {
    centerFolder = centerFolders.next();
  } else {
    centerFolder = masterFolder.createFolder(center);
  }

  // ===== AGE GROUP FOLDER =====
  let ageFolder;
  const ageFolders = centerFolder.getFoldersByName(ageGroup);

  if (ageFolders.hasNext()) {
    ageFolder = ageFolders.next();
  } else {
    ageFolder = centerFolder.createFolder(ageGroup);
  }

  // ===== LEVEL FOLDER =====
  let levelFolder;
  const levelFolders = ageFolder.getFoldersByName(level);

  if (levelFolders.hasNext()) {
    levelFolder = levelFolders.next();
  } else {
    levelFolder = ageFolder.createFolder(level);
  }

  // ===== REPLACE FILE =====
  const existingFiles = levelFolder.getFilesByName(fileName);

  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // ===== SAVE FILE =====
  levelFolder.createFile(blob);

  const folderUrl = levelFolder.getUrl();

  showDriveLink(studentName, folderUrl);
}



function showDriveLink(studentName, folderUrl) {

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial;padding:20px">
      <h3>✅ PDF created successfully!</h3>
      <p><b>Student:</b> ${studentName}</p>
      <p>
        <a href="${folderUrl}" target="_blank" style="
          font-size:16px;
          color:#1a73e8;
          text-decoration:none;
          font-weight:bold;
        ">
          📂 Open Appraisal Folder
        </a>
      </p>
    </div>
  `)
  .setWidth(420)
  .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, "Export Complete");
}
