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

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  if (!sheetName.startsWith("Ret")) return;

  const rowIndex = e.range.getRow();
  if (rowIndex < 3) return;

  const data = sheet.getDataRange().getValues();
  const header = data[1];
  const row = data[rowIndex - 1];

  const statusCol = col(header, "Retention Status");
  const editedCol = e.range.getColumn();
  const value = e.range.getValue();

  // 🔄 SYNC antar retention (edit biasa)
  syncToNextRetention(sheetName, header, row, rowIndex, e.range.getColumn());

  // hanya trigger kalau edit di kolom status
  if (editedCol !== statusCol + 1) return;

  
  if (!value || value.toString().trim().toLowerCase() !== "paid") return;

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
      sheet.getRange(rowIndex, statusCol + 1).setValue(""); // reset status
      return;
    }
  }

  // 🔥 VALIDASI PAYMENT (DP ATAU FP)
  const dpFields = [
    "DP Date",
    "DP Amount",
    "DP Invoice Number"
  ];

  const fpFields = [
    "FP Date",
    "FP Amount",
    "FP Invoice Number"
  ];

  const isDPFilled = dpFields.every(f => row[col(header, f)]);
  const isFPFilled = fpFields.every(f => row[col(header, f)]);

  if (!isDPFilled && !isFPFilled) {
    SpreadsheetApp.getUi().alert(
      "❌ Harus isi salah satu:\n\nDP (Date, Amount, Invoice)\nATAU\nFP (Date, Amount, Invoice)"
    );
    sheet.getRange(rowIndex, statusCol + 1).setValue(""); // reset status
    return;
  }

  // 🔥 LULUS VALIDASI → PUSH
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
      "DP Date",
      "DP Amount",
      "DP Invoice Number",
      "FP Date",
      "FP Amount",
      "FP Invoice Number",
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

function syncToNextRetention(sheetName, header, row, rowIndex, editedCol) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const currentRet = parseInt(sheetName.replace("Ret ", ""));
  const nextSheet = ss.getSheetByName(`Ret ${currentRet + 1}`);
  if (!nextSheet) return;

  const idx = {
    id: col(header,"Sparks ID"),
    unique: col(header,"Unique Key")
  };

const currentKey = row[idx.unique]; // ambil Unique Key (C1)
if (!currentKey) return;

const baseId = currentKey.toString().split("-C")[0]; // ambil ID tanpa cycle
const nextKey = baseId + "-C" + (currentRet + 1);

  const targetData = nextSheet.getDataRange().getValues();

  const syncFields = [
    "SA Retention",
    "Retention Status",
    "Churn Reason",
    "Response Notes",
    "Join Date Retention",
    "Retention Package",
    "FP Date",
    "FP Amount",
    "Impacted Holiday",
    "Other Impact"
  ];

  for (let i = 2; i < targetData.length; i++) {

    let key = targetData[i][idx.unique];

    if (key && key.toString().trim() === nextKey) {

      for (let field of syncFields) {

        const colIndex = col(header, field);

        if (editedCol === colIndex + 1) {

          let sourceValue = row[colIndex];

      // 🔥 override kalau kolom yang diedit
      if (editedCol === colIndex + 1) {
        sourceValue = SpreadsheetApp.getActiveSheet()
          .getRange(rowIndex, editedCol)
          .getValue();
      } 
          const targetValue = targetData[i][colIndex];

          if (sourceValue != targetValue) {
            nextSheet.getRange(i + 1, colIndex + 1).setValue(sourceValue);
          }

        }

      }

      break;
    }
  }
}