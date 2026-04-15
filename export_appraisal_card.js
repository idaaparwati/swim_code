function exportAppraisalPDF() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.includes("SWIM - Appraisal")) {
    SpreadsheetApp.getUi().alert("Please run this inside an Appraisal sheet.");
    return;
  }

  const studentName = sheet.getRange("D15").getDisplayValue().trim();
  const ageGroup = sheet.getRange("D16").getDisplayValue();
  const level = sheet.getRange("D17").getDisplayValue();

  if (!studentName) {
    SpreadsheetApp.getUi().alert("Please select student first.");
    return;
  }

  const fileName =
    "Appraisal - " +
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
    "&range=B6:F55" +
    "&size=A4" +
    "&portrait=true" +
    "&fitw=true" +
    "&gridlines=false" +
    "&printtitle=false" +
    "&sheetnames=false" +
    "&pagenumbers=false" +
    "&top_margin=0" +
    "&bottom_margin=0" +
    "&left_margin=0" +
    "&right_margin=0";

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

  // ===== SUB FOLDER (AgeGroup + Level) =====

  const subFolderName = ageGroup + " " + level;

  let subFolder;

  const subFolders = masterFolder.getFoldersByName(subFolderName);

  if (subFolders.hasNext()) {
    subFolder = subFolders.next();
  } else {
    subFolder = masterFolder.createFolder(subFolderName);
  }

  // ===== Replace file jika sudah ada =====

  const existingFiles = subFolder.getFilesByName(fileName);

  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  const file = subFolder.createFile(blob);

  const folderUrl = subFolder.getUrl();

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