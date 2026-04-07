/**
 * =========================================================
 * FINAL STABLE RETENTION SYSTEM (NO DUPLICATE, NO RESET BUG)
 * =========================================================
 */


/**
 * 🔧 Helper kolom
 */
function col(header, name) {
  const clean = h => h.toString().toLowerCase().replace(/\s+/g, ' ').trim();
  const idx = header.map(h => clean(h)).indexOf(clean(name));
  if (idx === -1) throw new Error("Kolom tidak ditemukan: " + name);
  return idx;
}


/**
 * 🔥 REALTIME (HANYA SAAT PAID)
 */
function onEdit(e) {
  Logger.log("EDIT DETECTED");

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.startsWith("Ret")) return;

  const rowIndex = e.range.getRow();
  if (rowIndex < 3) return;

  const header = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = col(header, "Retention Status");
  const editedCol = e.range.getColumn();
  const value = e.range.getValue();

  Logger.log("Value: " + value);
  Logger.log("Column: " + editedCol);

  // ✅ filter dulu (INI PENTING BANGET)
  if (editedCol !== statusCol + 1) return;
  if (!value || value.toString().trim().toLowerCase() !== "paid") return;

  // ✅ delay SEKALI aja (biar semua input settle)
  Utilities.sleep(800);

  // ✅ ambil row SETELAH delay
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 🔥 VALIDASI WAJIB
  const requiredFields = [
    "SA Retention",
    "Retention Status",
    "Churn Reason",
    "Join Date Retention",
    "Retention Package"
  ];

  for (let field of requiredFields) {
    let val = row[col(header, field)];
    if (!val) {
      SpreadsheetApp.getUi().alert(`❌ Kolom "${field}" wajib diisi sebelum lanjut.`);
      sheet.getRange(rowIndex, statusCol + 1).setValue("");
      return;
    }
  }
  const typePayment = row[col(header, "Type Payment")];

if (!typePayment) {
  SpreadsheetApp.getUi().alert("❌ Type Payment wajib diisi sebelum lanjut.");
  sheet.getRange(rowIndex, statusCol + 1).setValue("");
  return;
}


  const fpFields = [
    "FP Date",
    "FP Amount",
    "FP Invoice Number"
  ];

  const isFPFilled = fpFields.every(f => row[col(header, f)]);

  if (!isFPFilled) {
    SpreadsheetApp.getUi().alert(
      "❌ Harus isi salah satu:\n\nDP (Date, Amount, Invoice)\nATAU\nFP (Date, Amount, Invoice)"
    );
    sheet.getRange(rowIndex, statusCol + 1).setValue("");
    return;
  }

  // 🔥 PUSH
  const currentRet = parseInt(sheetName.replace("Ret ", ""));
  const nextRet = currentRet + 1;

  pushSingleRow(sheetName, `Ret ${nextRet}`, nextRet, rowIndex);
}


/**
 * 🔥 PUSH 1 ROW (ANTI DUPLICATE + NO OVERWRITE SOURCE)
 */
function pushSingleRow(sourceName, targetName, nextCycle, rowIndex) {

  const lock = LockService.getScriptLock();
  lock.waitLock(3000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const source = ss.getSheetByName(sourceName);

    const data = source.getDataRange().getValues();
    const header = data[1];
    const row = data[rowIndex - 1];

    const idx = {
      id: col(header,"Sparks ID"),
      cycle: col(header,"Cycle"),
      unique: col(header,"Unique Key"),

      joinRet: col(header,"Join Date Retention"),
      lastMember: col(header,"Actual Last Membership Date"),
      ageNow: col(header,"Age Now"),
      ageGroupNow: col(header,"Age Group Now"),
      packageRet: col(header,"Retention Package"),
      totalSessRet: col(header,"Total Session Retention Package"),
      fpRet: col(header,"FP Date"),
      fpAmount: col(header,"FP Amount"),
       saAqc: col(header,"SA Retention")
    };

    let id = row[idx.id];
    if (!id) return;

    id = id.toString().trim();
    let newKey = id + "-C" + nextCycle;

    let target = ss.getSheetByName(targetName);

    if (!target) {
      target = ss.insertSheet(targetName);
      target.getRange(2,1,1,header.length).setValues([header]);
      target.setFrozenRows(2);
    }

    const targetData = target.getDataRange().getValues();

    // 🔥 ANTI DUPLICATE
    for (let i = 2; i < targetData.length; i++) {
      let existingKey = targetData[i][idx.unique];
      if (existingKey && existingKey.toString().trim() === newKey) {
        return;
      }
    }

    let newRow = [...row];

    // 🔥 UPDATE KEY & CYCLE
    newRow[idx.cycle] = nextCycle;
    newRow[idx.unique] = newKey;

    // 🔥 MAPPING PREVIOUS DATA (INI YANG KAMU MAU)
    newRow[col(header,"Previous Join Date")] = row[idx.joinRet];
    newRow[col(header,"Previous Last Membership Date")] = row[idx.lastMember];
    newRow[col(header,"Previous Age")] = row[idx.ageNow];
    newRow[col(header,"Previous Age Group")] = row[idx.ageGroupNow];
    newRow[col(header,"Previous Package")] = row[idx.packageRet];
    newRow[col(header,"Previous Total Session")] = row[idx.totalSessRet];
    newRow[col(header,"Previous FP Date")] = row[idx.fpRet];
     newRow[col(header,"SA Aquisition")] = row[idx.saAqc];


    // 🔥 CLEAR KOLOM SA (TARGET SAJA)
    [
      "Retention Status",
      "Churn Reason",
      "Response Notes",
      "Join Date Retention",
      "Retention Package"
    ].forEach(c => newRow[col(header,c)] = "");

    // clear payment & SA related
    [
      "Total Session Retention Package",
      "Last Membership Date",
      "Actual Last Membership Date",
      "FP Date",
      "FP Amount",
      "FP Invoice Number",
      "Type Payment",
      "Total Actual Payment",
      "SA Retention",
      "Impacted Holiday",
      "Other Impact"
    ].forEach(c => newRow[col(header,c)] = "");

    // clear age now (biar dihitung ulang)
    newRow[col(header,"Age Now")] = "";
    newRow[col(header,"Age Group Now")] = "";

    target.appendRow(newRow);

  } finally {
    lock.releaseLock();
  }
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

  const columnsToClear = [
    "SA Retention",
    "Retention Status",
    "Churn Reason",
    "Response Notes",
    "Join Date Retention",
    "Retention Package",
    "FP Date",
    "FP Amount",
    "FP Invoice Number",
    "Type Payment",
    "Total Actual Payment",
    "Impacted Holiday",
    "Other Impact"
  ];

  for (let i = 2; i <= MAX_RET; i++) { // 🔥 mulai dari Ret 2

    let sheet = ss.getSheetByName(`Ret ${i}`);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const header = data[1];
    const lastRow = sheet.getLastRow();

    if (lastRow < 3) continue;

    for (let colName of columnsToClear) {
      let colIndex = col(header, colName);
      sheet.getRange(3, colIndex + 1, lastRow - 2).clearContent();
    }

  }

  SpreadsheetApp.getUi().alert("✅ Ret 2 ke atas sudah di-reset!");
}