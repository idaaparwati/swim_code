function generateStudentProgressPerCenter() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = ss.getSheetByName("Control");
  const studentSheet = ss.getSheetByName("Master Student List SWIM");
  const templateSheet = ss.getSheetByName("Assesment Appraisal - Sparks Swim");

  const selectedCenter = controlSheet.getRange("B1").getValue().toString().trim();

  if (!selectedCenter) {
    SpreadsheetApp.getUi().alert("⚠️ Please select center in Control!B1 first.");
    return;
  }

  const outputName = "Student Progressing - " + selectedCenter;

  let outputSheet = ss.getSheetByName(outputName);

  if (outputSheet) {
    outputSheet.clear();
    outputSheet.clearFormats();
  } else {
    outputSheet = ss.insertSheet(outputName);
  }

  const students = studentSheet.getDataRange().getValues();
  const templates = templateSheet.getDataRange().getValues();

  const TOTAL_WEEKS = 48;

  // HEADER
  const header = [
    "Center",
    "Student ID",
    "Student Name",
    "Age Group",
    "Level",
    "Coach",
    "Metrics",
    "Skill"
  ];

  for (let i = 1; i <= TOTAL_WEEKS; i++) {
    header.push("Score - Week " + i);
  }

  header.push("Primary Key");

  outputSheet.getRange(1,1,1,header.length).setValues([header]);
  outputSheet.getRange(1,1,1,header.length).setFontWeight("bold");

  const result = [];

  for (let i = 1; i < students.length; i++) {

    const center = students[i][0]?.toString().trim();
    const studentID = students[i][1];
    const name = students[i][2];
    const ageGroup = students[i][3];
    const ageGroupLower = students[i][3]?.toString().trim().toLowerCase();
    const level = students[i][5];
    const levelLower = students[i][5]?.toString().trim().toLowerCase();
    const coach = students[i][6];
    const status = students[i][7]?.toString().trim().toLowerCase();

    if (center !== selectedCenter) continue;
    if (!name) continue;
    if (status !== "active") continue;
    if (!level || level === "0") continue;

    for (let j = 1; j < templates.length; j++) {
  


      const tAge = templates[j][0]?.toString().trim().toLowerCase();
      const tLevel = templates[j][1]?.toString().trim().toLowerCase();
      const metrics = templates[j][2];
      const skill = templates[j][3];


           Logger.log({
    student: name,
    studentAge: ageGroup,
    templateAge: tAge,
    studentLevel: level,
    templateLevel: tLevel
  });

      if (ageGroupLower === tAge && levelLower === tLevel) {

        const primaryKey = (
          name + "|" +
          ageGroup + "|" +
          level + "|" +
          metrics + "|" +
          skill
        ).toLowerCase().trim();

        const row = [
          center,
          studentID,
          name,
          ageGroup,
          level,
          coach,
          metrics,
          skill
        ];

        for (let w = 0; w < TOTAL_WEEKS; w++) {
          row.push("");
        }

        row.push(primaryKey);

        result.push(row);
      }
    }
  }

  if (result.length > 0) {

    outputSheet.getRange(2,1,result.length,header.length).setValues(result);

    const scoreStartColumn = 9;
    const scoreRange = outputSheet.getRange(2, scoreStartColumn, result.length, TOTAL_WEEKS);

    // DROPDOWN SCORE
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Mastered", "Progressing", "Exploring"], true)
      .setAllowInvalid(false)
      .build();

    scoreRange.setDataValidation(rule);

    // COLOR
    const rules = [

      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Mastered")
        .setBackground("#c6efce")
        .setRanges([scoreRange])
        .build(),

      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Progressing")
        .setBackground("#fff2cc")
        .setRanges([scoreRange])
        .build(),

      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Exploring")
        .setBackground("#f8cbad")
        .setRanges([scoreRange])
        .build()
    ];

    outputSheet.setConditionalFormatRules(rules);
  }

  outputSheet.autoResizeColumns(1, header.length);


 
  SpreadsheetApp.getUi().alert("✅ Student Progress generated for center: " + selectedCenter);

}

function syncActiveStudentsOnly() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = ss.getSheetByName("Control");
  const studentSheet = ss.getSheetByName("Master Student List SWIM");
  const templateSheet = ss.getSheetByName("Assesment Appraisal - Sparks Swim");

  const selectedCenter = controlSheet.getRange("B1").getValue().toString().trim();

  if (!selectedCenter) {
    SpreadsheetApp.getUi().alert("⚠️ Please select center in Control!B1 first.");
    return;
  }

  const outputName = "Student Progressing - " + selectedCenter;
  let outputSheet = ss.getSheetByName(outputName);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputName);
  }

  const students = studentSheet.getDataRange().getValues();
  const templates = templateSheet.getDataRange().getValues();

  const TOTAL_WEEKS = 48;

  // HEADER
  const header = [
    "Center","Student ID","Student Name","Age Group",
    "Level","Coach","Metrics","Skill"
  ];

  for (let i = 1; i <= TOTAL_WEEKS; i++) {
    header.push("Score - Week " + i);
  }

  header.push("Primary Key");

  // create header once
  if (outputSheet.getLastRow() === 0) {
    outputSheet.getRange(1,1,1,header.length).setValues([header]);
    outputSheet.getRange(1,1,1,header.length).setFontWeight("bold");
  }

  // existing PK
  const existingData = outputSheet.getDataRange().getValues();
  const existingPK = new Set();

  if (existingData.length > 1) {
    for (let i = 1; i < existingData.length; i++) {
      const pk = existingData[i][existingData[i].length - 1];
      if (pk) existingPK.add(pk);
    }
  }

  const result = [];

  for (let i = 1; i < students.length; i++) {

    const center = students[i][0]?.toString().trim();
    const studentID = students[i][1];
    const name = students[i][2];
    const ageGroupRaw = students[i][3];
    const levelRaw = students[i][5];
    const coach = students[i][6];
    const status = students[i][7]?.toString().trim().toLowerCase();

    if (center !== selectedCenter) continue;
    if (!name) continue;
    if (status !== "active") continue;
    if (!levelRaw || levelRaw === "0") continue;

    // normalize
    const ageGroup = ageGroupRaw?.toString().toLowerCase().trim();
    const level = levelRaw?.toString().toLowerCase().trim();

    let foundMatch = false;

    for (let j = 1; j < templates.length; j++) {

      const tAge = templates[j][0]?.toString().toLowerCase().trim();
      const tLevel = templates[j][1]?.toString().toLowerCase().trim();
      const metrics = templates[j][2];
      const skill = templates[j][3];

      // FLEXIBLE MATCH
      if (
        ageGroup.includes(tAge) &&
        level.includes(tLevel)
      ) {

        foundMatch = true;

        const primaryKey = (
          studentID + "|" + metrics + "|" + skill
        ).toLowerCase().trim();

        if (existingPK.has(primaryKey)) continue;

        const row = [
          center,
          studentID,
          name,
          ageGroupRaw,
          levelRaw,
          coach,
          metrics,
          skill
        ];

        for (let w = 0; w < TOTAL_WEEKS; w++) {
          row.push("");
        }

        row.push(primaryKey);

        result.push(row);
      }
    }

    // DEBUG kalau tidak ada template match
    if (!foundMatch) {
      Logger.log("❌ NO TEMPLATE → " + name + " | " + ageGroup + " | " + level);
    }
  }

  if (result.length > 0) {

    const lastRow = outputSheet.getLastRow();

    outputSheet.getRange(lastRow + 1, 1, result.length, header.length).setValues(result);

    const scoreStartColumn = 9;
    const scoreRange = outputSheet.getRange(lastRow + 1, scoreStartColumn, result.length, TOTAL_WEEKS);

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Mastered", "Progressing", "Exploring"], true)
      .setAllowInvalid(false)
      .build();

    scoreRange.setDataValidation(rule);

    SpreadsheetApp.getUi().alert("✅ New active students synced!");
  } else {
    SpreadsheetApp.getUi().alert("⚠️ No new students added. Check logs.");
  }

}