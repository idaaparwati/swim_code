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

function syncActiveStudentsByCenter(selectedCenter) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const studentSheet = ss.getSheetByName("Master Student List SWIM");
  const templateSheet = ss.getSheetByName("Assesment Appraisal - Sparks Swim");

  const outputName = "Student Progressing - " + selectedCenter;
  let outputSheet = ss.getSheetByName(outputName);

  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputName);
  }

  const students = studentSheet.getDataRange().getValues();
  const templates = templateSheet.getDataRange().getValues();

  const TOTAL_WEEKS = 48;

  // 🔥 SUPER CLEAN (FIX SEMUA MASALAH STRING)
  const clean = (val) =>
  val
    ?.toString()
    .toLowerCase()
    .replace(/\u00A0/g, " ")  // 🔥 ubah NBSP jadi spasi normal
    .replace(/\s+/g, " ")     // 🔥 rapihin spasi jadi 1
    .trim();                  // 🔥 hapus depan & belakang
  // ================= HEADER =================
  const header = [
    "Center","Student ID","Student Name","Age Group",
    "Level","Coach","Metrics","Skill"
  ];

  for (let i = 1; i <= TOTAL_WEEKS; i++) {
    header.push("Score - Week " + i);
  }

  header.push("Primary Key");

  if (outputSheet.getLastRow() <= 1) {
    outputSheet.getRange(1,1,1,header.length).setValues([header]);
    outputSheet.getRange(1,1,1,header.length).setFontWeight("bold");
  }

  // ================= AMBIL STUDENT ID YANG ADA =================
  const lastRow = outputSheet.getLastRow();
  const existingStudentIDs = new Set();

  if (lastRow > 1) {
    const idColumn = outputSheet.getRange(2, 2, lastRow - 1).getValues();

    idColumn.forEach(r => {
      if (r[0]) {
        existingStudentIDs.add(clean(r[0]));
      }
    });
  }

  const result = [];
  const selected = clean(selectedCenter);

  // ================= LOOP STUDENT =================
  for (let i = 1; i < students.length; i++) {

    const center = clean(students[i][0]);
    const studentID = students[i][1];
    const name = students[i][2];
    const ageGroupRaw = students[i][3];
    const levelRaw = students[i][5];
    const coach = students[i][6];
    const status = clean(students[i][7]);

    if (center !== selected) continue;
    if (!name) continue;
    if (!status.includes("active")) continue;
    if (!levelRaw || levelRaw === "0") continue;

    const currentID = clean(studentID);

    // 🔥 LOGIC UTAMA KAMU
    if (existingStudentIDs.has(currentID)) {
      Logger.log("⏭ SKIP → " + name);
      continue;
    }

    const ageGroup = clean(ageGroupRaw);
    const level = clean(levelRaw);

    let foundMatch = false;

    // ================= LOOP TEMPLATE =================
    for (let j = 1; j < templates.length; j++) {

      const tAge = clean(templates[j][0]);
      const tLevel = clean(templates[j][1]);
      const metrics = templates[j][2];
      const skill = templates[j][3];

      // 🔥 FLEXIBLE MATCH (ANTI ERROR DATA)
      const ageMatch =
        ageGroup.includes(tAge) || tAge.includes(ageGroup);

      const levelMatch =
        level.includes(tLevel) || tLevel.includes(level);

      if (ageMatch && levelMatch) {

        foundMatch = true;

       const primaryKey = [
        name,
        ageGroupRaw,
        levelRaw,
        coach,
        metrics,
        skill
      ]
        .map(v => clean(v))
        .join("|");

        const row = [
          students[i][0],
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

    if (!foundMatch) {
      Logger.log(`❌ NO TEMPLATE → ${name} | ${ageGroup} | ${level}`);
    }
  }

  // ================= WRITE =================
  if (result.length > 0) {

  const startRow = outputSheet.getLastRow() + 1;

  outputSheet.getRange(startRow, 1, result.length, header.length).setValues(result);

  const studentCount = new Set(result.map(r => r[1])).size;

  Logger.log("✅ Students added: " + studentCount);
  Logger.log("📊 Rows added: " + result.length);

} else {
  Logger.log("ℹ️ No new students found");
}
}

function syncAllCenters() {

  const centers = ["KLM", "KWC"]; // tambah kalau ada center lain

  for (let i = 0; i < centers.length; i++) {
    syncActiveStudentsByCenter(centers[i]);
  }

}

function dailyAutomation() {
  syncAllCenters();
  addAndSyncSessionData();
}


function runSync() {
  syncActiveStudentsByCenter("KLM");
}

function regeneratePrimaryKeyOnly() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  // 🔥 TARUH CLEAN DI ATAS
  const clean = (val) =>
    val
      ?.toString()
      .toLowerCase()
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
      .trim();


  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colName = headers.indexOf("Student Name");
  const colAge = headers.indexOf("Age Group");
  const colLevel = headers.indexOf("Level");
  const colCoach = headers.indexOf("Coach");
  const colMetrics = headers.indexOf("Metrics");
  const colSkill = headers.indexOf("Skill");
  const colPK = headers.indexOf("Primary Key");

  if (colPK === -1) {
    Logger.log("❌ Primary Key column not found");
    return;
  }

  const pkValues = [];

  for (let i = 1; i < data.length; i++) {

    const name = data[i][colName] || "";
    const ageGroup = data[i][colAge] || "";
    const level = data[i][colLevel] || "";
    const coach = data[i][colCoach] || "";
    const metrics = data[i][colMetrics] || "";
    const skill = data[i][colSkill] || "";

    const primaryKey = [
      name,
      ageGroup,
      level,
      coach,
      metrics,
      skill
    ]
      .map(v => clean(v))   // 🔥 pakai clean
      .join("|");           // 🔥 separator

    pkValues.push([primaryKey]);
  }

  sheet.getRange(2, colPK + 1, pkValues.length, 1).setValues(pkValues);

  Logger.log("✅ Primary Key regenerated with separator");
}