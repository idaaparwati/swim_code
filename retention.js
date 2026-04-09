/**
 * =========================================================
 * FINAL STABLE RETENTION SYSTEM (NO DUPLICATE, NO RESET BUG)
 * =========================================================
 */


/**
 * 🔧 Helper kolom
 */

function col(header, name) {
  const clean = h => h.toString().toLowerCase().replace(/\s+/g, '').trim();
  const idx = header.map(h => clean(h)).indexOf(clean(name));
  if (idx === -1) throw new Error("Kolom tidak ditemukan: " + name);
  return idx;
}

function getNextRow(sheet) {
  const data = sheet.getRange("A:A").getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "") return i + 2;
  }
  return 3;
}




// ON edit warning only (tidak push ke Ret Selanjutnya)

function onEdit(e) {

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.startsWith("Ret")) return;

  const rowIndex = e.range.getRow();
  if (rowIndex < 3) return;

  const header = sheet.getRange(2,1,1,sheet.getLastColumn()).getValues()[0];

  const statusCol = col(header,"Retention Status");
  const editedCol = e.range.getColumn();

  if (editedCol !== statusCol + 1) return;

  Utilities.sleep(300);

  const row = sheet.getRange(rowIndex,1,1,sheet.getLastColumn()).getValues()[0];

  const value = row[statusCol];
  const clean = value ? value.toString().trim().toLowerCase() : "";

  if (clean !== "paid") return;

  // VALIDASI
  const requiredFields = [
    "SA Retention",
    "Churn Reason",
    "Join Date Retention",
    "Retention Package"
  ];

  for (let field of requiredFields) {
    if (!row[col(header, field)]) {
      SpreadsheetApp.getUi().alert(`❌ Kolom "${field}" wajib diisi`);
      sheet.getRange(rowIndex, statusCol + 1).setValue("");
      return;
    }
  }

  if (!row[col(header,"Type Payment")]) {
    SpreadsheetApp.getUi().alert("❌ Type Payment wajib diisi");
    sheet.getRange(rowIndex, statusCol + 1).setValue("");
    return;
  }

  const fpFields = ["FP Date","FP Amount","FP Invoice Number"];
  const isFPFilled = fpFields.every(f => row[col(header,f)]);

  if (!isFPFilled) {
    SpreadsheetApp.getUi().alert("❌ FP harus lengkap");
    sheet.getRange(rowIndex, statusCol + 1).setValue("");
    return;
  }

  Logger.log("✅ VALIDASI OK");
}

function pushSingleRowSafe(source, target, sourceHeader, targetHeader, row, newKey, cycle) {

  Logger.log("🚀 PUSH START: " + newKey);

  const newRow = new Array(targetHeader.length).fill("");

  // COPY FIELD
  for (let i = 0; i < sourceHeader.length; i++) {
    let colName = sourceHeader[i];
    try {
      let targetIndex = col(targetHeader, colName);
      newRow[targetIndex] = row[i];
    } catch(err) {}
  }

  // SET KEY & CYCLE
  try {
    newRow[col(targetHeader,"Unique Key")] = newKey;
    newRow[col(targetHeader,"Cycle")] = cycle;
  } catch(err) {}

  // MAPPING
  const mapping = {
    "Previous Join Date": "Join Date Retention",
    "Previous Last Membership Date": "Actual Last Membership Date",
    "Previous Age": "Age Now",
    "Previous Age Group": "Age Group Now",
    "Previous Package": "Retention Package",
    "Previous Total Session": "Total Session Retention Package",
    "Previous FP Date": "FP Date",
    "SA Aquisition": "SA Retention"
  };

  for (let targetField in mapping) {
    try {
      let targetIndex = col(targetHeader, targetField);
      let sourceIndex = col(sourceHeader, mapping[targetField]);

      newRow[targetIndex] = row[sourceIndex];

      Logger.log("🔁 MAP: " + mapping[targetField] + " → " + targetField);

    } catch(err) {
      Logger.log("⚠️ MAP FAIL: " + targetField);
    }
  }

  // RESET FIELD
  const resetFields = [
    "SA Retention","Retention Status","Churn Reason","Response Notes",
    "Join Date Retention","Retention Package","Total Session Retention Package",
    "FP Date","FP Amount","FP Invoice Number","Type Payment",
    "Total Actual Payment","Impacted Holiday","Other Impact",
    "Last Membership Date","Actual Last Membership Date",
    "Age Now","Age Group Now"
  ];

  resetFields.forEach(field => {
    try {
      newRow[col(targetHeader, field)] = "";
    } catch(err) {}
  });

  const nextRow = getNextRow(target);

  target.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

  Logger.log("✅ PUSH SUCCESS: " + newKey + " → row " + nextRow);
}


// CORE System Time-Driven Push to Next Ret
function processRetentionSafe() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_RET = 5;

  Logger.log("========== START PROCESS ==========");

  for (let current = 1; current < MAX_RET; current++) {

    Logger.log("👉 PROCESS RET " + current + " → RET " + (current+1));

    let source = ss.getSheetByName(`Ret ${current}`);
    let target = ss.getSheetByName(`Ret ${current + 1}`);

    if (!source) {
      Logger.log("❌ SOURCE NOT FOUND: Ret " + current);
      continue;
    }

    if (!target) {
      Logger.log("🆕 CREATE SHEET: Ret " + (current+1));
      target = ss.insertSheet(`Ret ${current + 1}`);
      const headerRange = source.getRange(1,1,2,source.getLastColumn());
      headerRange.copyTo(target.getRange(1,1), {contentsOnly:false});
      target.setFrozenRows(2);
    }

    const data = source.getDataRange().getValues();
    const header = data[1];

    const targetData = target.getDataRange().getValues();
    const targetHeader = target.getRange(2,1,1,target.getLastColumn()).getValues()[0];

    const idx = {
      id: col(header,"Sparks ID"),
      status: col(header,"Retention Status"),
      joinRet: col(header,"Join Date Retention"),
      lastMember: col(header,"Actual Last Membership Date"),
      ageNow: col(header,"Age Now"),
      ageGroupNow: col(header,"Age Group Now"),
      packageRet: col(header,"Retention Package"),
      totalSessRet: col(header,"Total Session Retention Package"),
      fpRet: col(header,"FP Date"),
      sa: col(header,"SA Retention")
    };

    // BUILD MAP
    let targetMap = {};
    for (let i = 2; i < targetData.length; i++) {
      let key = targetData[i][col(targetHeader,"Unique Key")];
      if (key) {
        targetMap[key.toString().trim()] = i;
      }
    }

    Logger.log("📦 EXISTING DATA: " + Object.keys(targetMap).length);

    for (let i = 2; i < data.length; i++) {

      let row = data[i];

      let status = row[idx.status];
      let clean = status ? status.toString().trim().toLowerCase() : "";

      Logger.log("🔍 ROW " + i + " STATUS: [" + status + "]");

      if (!clean.includes("paid")) {
        Logger.log("⏭️ SKIP (NOT PAID)");
        continue;
      }

      let id = row[idx.id];
      if (!id) {
        Logger.log("⚠️ NO ID");
        continue;
      }

      id = id.toString().trim();
      let newKey = id + "-C" + (current + 1);

      Logger.log("🎯 PROCESS KEY: " + newKey);

      // PUSH
      if (!targetMap[newKey]) {

        Logger.log("🆕 NEW DATA → PUSH");

        pushSingleRowSafe(
          source,
          target,
          header,
          targetHeader,
          row,
          newKey,
          current + 1
        );

        continue;
      }

      // UPDATE
      Logger.log("🔄 UPDATE EXISTING");

      let rowIndex = targetMap[newKey] + 1;

      const mapping = {
        "Previous Join Date": row[idx.joinRet],
        "Previous Last Membership Date": row[idx.lastMember],
        "Previous Age": row[idx.ageNow],
        "Previous Age Group": row[idx.ageGroupNow],
        "Previous Package": row[idx.packageRet],
        "Previous Total Session": row[idx.totalSessRet],
        "Previous FP Date": row[idx.fpRet],
        "SA Aquisition": row[idx.sa]
      };

      for (let field in mapping) {
        try {
          let colIndex = col(targetHeader, field);
          target.getRange(rowIndex, colIndex + 1).setValue(mapping[field]);

          Logger.log("✔ UPDATE FIELD: " + field);

        } catch(err) {
          Logger.log("❌ UPDATE FAIL: " + field);
        }
      }
    }
  }

  Logger.log("========== END PROCESS ==========");
}


// Date validation khusus FP date dan Join date

function setupDateValidation() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Ret 1","Ret 2","Ret 3","Ret 4","Ret 5"];

  sheets.forEach(name => {

    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const header = sheet.getRange(2,1,1,sheet.getLastColumn()).getValues()[0];

    const dateCols = [
      "Join Date Retention",
      "FP Date"
    ];

    dateCols.forEach(field => {

      try {

        let colIndex = col(header, field) + 1;
        const lastRow = Math.max(sheet.getLastRow(), 1000);

        let range = sheet.getRange(3, colIndex, lastRow);

        range.clearDataValidations();

        let rule = SpreadsheetApp.newDataValidation()
          .requireDate()
          .setAllowInvalid(false)
          .build();

        range.setDataValidation(rule);

      } catch(err) {}
    });

  });

  SpreadsheetApp.getUi().alert("✅ Date validation aktif!");
}





function resetAllRetSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("Ret ")) {

      if (sheet.getLastRow() > 2) {
        sheet.getRange(3,1,sheet.getLastRow()-2,sheet.getLastColumn()).clearContent();
      }

      Logger.log("🧹 CLEAN: " + sheet.getName());
    }
  });
}


/**
 * 🔥 OPTIONAL BACKUP SYNC (TIDAK RESET SA)
 */
function runAllRetention() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_RET = 5;

  for (let current = 1; current < MAX_RET; current++) {

    let source = ss.getSheetByName(`Ret ${current}`);
    let target = ss.getSheetByName(`Ret ${current + 1}`);

    if (!source || !target) continue;

    const data = source.getDataRange().getValues();
    const header = data[1];

    const idx = {
      id: col(header,"Sparks ID"),
      unique: col(header,"Unique Key"),
      joinRet: col(header,"Join Date Retention")
    };

    const targetData = target.getDataRange().getValues();

    let targetMap = {};

    for (let i = 2; i < targetData.length; i++) {
      let key = targetData[i][idx.unique];
      if (key) targetMap[key.toString().trim()] = i;
    }

    data.slice(2).forEach(row => {

      let id = row[idx.id];
      if (!id) return;

      id = id.toString().trim();
      let key = id + "-C" + (current + 1);

      if (targetMap[key] !== undefined) {

        // 🔥 update field NON-SA saja
        // target.getRange(targetMap[key] + 1, idx.joinRet + 1)
        //       .setValue(row[idx.joinRet]);
      }

    });

  }

  Logger.log("✅ Backup sync jalan");
}

function rebuildRetention() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_RET = 5;

  // 🔥 LOOP RET
  for (let current = 1; current < MAX_RET; current++) {

    let sourceName = `Ret ${current}`;
    let targetName = `Ret ${current + 1}`;

    let source = ss.getSheetByName(sourceName);
    let target = ss.getSheetByName(targetName);

    if (!source) continue;

    const data = source.getDataRange().getValues();
    const header = data[1];

    // 🔥 CLEAR TARGET DULU (BIAR BERSIH)
    if (target) {
      if (target.getLastRow() > 2) {
        target.getRange(3,1,target.getLastRow()-2,target.getLastColumn()).clearContent();
      }
    }

    // 🔥 REBUILD SEMUA ROW
    for (let i = 2; i < data.length; i++) {

      let row = data[i];

      let status = row[col(header,"Retention Status")];
      if (!status) continue;

      if (status.toString().toLowerCase() !== "paid") continue;

      pushSingleRow(sourceName, targetName, current + 1, i + 1);
    }

  }

  SpreadsheetApp.getUi().alert("✅ Rebuild Retention selesai!");
}

function syncAllRetentionFixed() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_RET = 5;

  // 🎯 MAPPING FIELD (INI KUNCI)
  const fieldMapping = {
    "Join Date Retention": "Previous Join Date",
    "Actual Last Membership Date": "Previous Last Membership Date",
    "Age Now": "Previous Age",
    "Age Group Now": "Previous Age Group",
    "Retention Package": "Previous Package",
    "Total Session Retention Package": "Previous Total Session",
    "FP Date": "Previous FP Date",
    "SA Retention": "SA Aquisition"
  };

  for (let current = 1; current < MAX_RET; current++) {

    let source = ss.getSheetByName(`Ret ${current}`);
    let target = ss.getSheetByName(`Ret ${current + 1}`);

    if (!source || !target) continue;

    const data = source.getDataRange().getValues();
    const header = data[1];
    const targetData = target.getDataRange().getValues();

    const idx = {
      unique: col(header,"Unique Key")
    };

    let targetMap = {};

    for (let i = 2; i < targetData.length; i++) {
      let key = targetData[i][idx.unique];
      if (key) targetMap[key.toString().trim()] = i;
    }

    for (let i = 2; i < data.length; i++) {

      let row = data[i];
      let currentKey = row[idx.unique];
      if (!currentKey) continue;

      let baseId = currentKey.toString().split("-C")[0];
      let nextKey = baseId + "-C" + (current + 1);

      if (targetMap[nextKey] !== undefined) {

        for (let sourceField in fieldMapping) {

          const targetField = fieldMapping[sourceField];

          const sourceCol = col(header, sourceField);
          const targetCol = col(header, targetField);

          const sourceValue = row[sourceCol];
          const targetValue = targetData[targetMap[nextKey]][targetCol];

          if (sourceValue != targetValue) {
            target.getRange(targetMap[nextKey] + 1, targetCol + 1)
                  .setValue(sourceValue);
          }

        }

      }

    }

  }

  SpreadsheetApp.getUi().alert("✅ Sync mapping selesai!");
}

function resetAllSAColumns() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Ret 1", "Ret 2", "Ret 3", "Ret 4"];

  const startRow = 3;
  const startCol = 20; // kolom T

  sheets.forEach(name => {

    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < startRow) return;

    sheet.getRange(
      startRow,
      startCol,
      lastRow - 2,
      lastCol - startCol + 1
    ).clearContent();

  });

  SpreadsheetApp.getUi().alert("🔥 Semua SA column sudah di-reset!");
}

function resetRetentionExceptRet1() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_RET = 5;

  for (let i = 2; i <= MAX_RET; i++) {

    const sheet = ss.getSheetByName(`Ret ${i}`);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 3) continue;

    // 🔥 HAPUS SEMUA DATA (kecuali header)
    sheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();

    Logger.log(`🔥 Ret ${i} cleared`);
  }

  SpreadsheetApp.getUi().alert("🚀 Ret 2 ke atas sudah bersih! Siap go-live!");
}