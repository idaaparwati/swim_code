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

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    Logger.log("⏭️ SKIP (SCRIPT LOCKED)");
    return;
  }

  const props = PropertiesService.getScriptProperties();

  // 🔥 ANTI DOUBLE RUN (TRIGGER + BUTTON)
  if (props.getProperty("PROCESS_RUNNING") === "true") {
    Logger.log("⏭️ SKIP (PROCESS SEDANG BERJALAN)");
    return;
  }

  props.setProperty("PROCESS_RUNNING", "true");

  try {

    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const MAX_RET = 5;

    Logger.log("========== START PROCESS ==========");

    // 🔥 ERROR TRACKER
    let errorLogs = {};

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

      let targetMap = {};
      for (let i = 2; i < targetData.length; i++) {
        let key = targetData[i][col(targetHeader,"Unique Key")];
        if (key) targetMap[key.toString().trim()] = i;
      }

      for (let i = 2; i < data.length; i++) {

        let row = data[i];

        let status = row[idx.status];
        let clean = status ? status.toString().trim().toLowerCase() : "";

        if (!clean.includes("paid")) continue;

        let id = row[idx.id];
        if (!id) continue;

        if (!row[idx.joinRet] || !row[idx.packageRet] || !row[idx.totalSessRet]) continue;

        id = id.toString().trim();
        let newKey = id + "-C" + (current + 1);

        try {

          if (!targetMap[newKey]) {

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
            } catch(err) {

              if (!errorLogs[id]) errorLogs[id] = [];
              errorLogs[id].push(`Ret ${current} Row ${i}: gagal update ${field}`);
            }
          }

        } catch (err) {

          if (!errorLogs[id]) errorLogs[id] = [];
          errorLogs[id].push(`Ret ${current} Row ${i}: ${err}`);
        }
      }
    }

    Logger.log("========== END PROCESS ==========");

    // 🔥 HANDLE ERROR
    if (Object.keys(errorLogs).length > 0) {
      logErrorToSheet(errorLogs);
      sendErrorEmail(errorLogs);
      sendWhatsAppAlert(errorLogs); // 🔥 WA juga
    }

  } catch (err) {

    Logger.log("💥 FATAL ERROR: " + err);

  } finally {

    props.deleteProperty("PROCESS_RUNNING"); // 🔥 RESET FLAG
    lock.releaseLock();
  }
}

function sendErrorEmail(errorLogs) {

  const emails = ["ida.parwati@seven-retail.com"];

  let message = "🚨 ERROR RETENTION SYSTEM\n\n";

  for (let id in errorLogs) {
    message += `🧑 ${id}\n`;
    errorLogs[id].forEach(e => message += "- " + e + "\n");
    message += "\n";
  }

  MailApp.sendEmail({
    to: emails.join(","),
    subject: "🚨 Retention Error",
    body: message
  });
}

function runRetentionManual() {

  const ui = SpreadsheetApp.getUi();

  try {

    ui.alert("🚀 Processing retention...");

    processRetentionSafe();

    ui.alert("✅ SUCCESS!\nData berhasil diproses");

  } catch (err) {

    ui.alert("❌ ERROR\n" + err);

  }
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



function stackRetentionFinal_daily() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let target = ss.getSheetByName("Master Stack");

  const sheets = ["Ret 1", "Ret 2", "Ret 3", "Ret 4"];
  const START_ROW = 3;

  const ID_COL = 3; // Sparks ID
  const STATUS_COL = 21;

  Logger.log("🚀 START MASTER STACK PROCESS");

  // =============================
  // AUTO CREATE MASTER
  // =============================
  if (!target) {
    target = ss.insertSheet("Master Stack");

    const sample = ss.getSheetByName("Ret 1");
    const headerRange = sample.getRange(1,1,2,sample.getLastColumn());

    headerRange.copyTo(target.getRange(1,1), {contentsOnly:false});

    // 🔥 REMOVE UNUSED COLUMN DARI HEADER
    const header = target.getRange(2,1,1,target.getLastColumn()).getValues()[0];

    const removeFields = ["Sisa Sesi","Last Class Date"];
    let indexesToRemove = [];

    removeFields.forEach(field => {
      try {
        const idx = col(header, field);
        indexesToRemove.push(idx + 1);
      } catch(e){}
    });

    indexesToRemove.sort((a,b)=>b-a);
    indexesToRemove.forEach(colIdx => target.deleteColumn(colIdx));

    const lastCol = target.getLastColumn();

    const extraHeader = [
      "Center",
      "From Ret",
      "To Ret"
    ];

    target
      .getRange(2, lastCol + 1, 1, extraHeader.length)
      .setValues([extraHeader]);

    target.setFrozenRows(2);

    Logger.log("🆕 Master Stack created");
  }

  const lastRow = target.getLastRow();
  const lastCol = target.getLastColumn();

  let existingData = [];
  let idMap = {};

  // =============================
  // LOAD EXISTING DATA
  // =============================
  if (lastRow >= START_ROW) {
    existingData = target
      .getRange(START_ROW, 1, lastRow - 2, lastCol)
      .getValues();

    existingData.forEach((row, i) => {
      const id = row[ID_COL - 1];
      const toRet = row[lastCol - 1];

      if (id) {
        idMap[`${id}-${toRet}`] = {
          rowIndex: i + START_ROW,
          rowData: row
        };
      }
    });
  }

  let appendData = [];

  // =============================
  // LOOP ALL RET
  // =============================
  sheets.forEach((sheetName, idx) => {

    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const lastRowSheet = sh.getLastRow();
    if (lastRowSheet < START_ROW) return;

    const header = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];

    const data = sh
      .getRange(START_ROW, 1, lastRowSheet - 2, sh.getLastColumn())
      .getValues();

    data.forEach((row, i) => {

      const id = row[ID_COL - 1];
      if (!id) return;

      const status = row[STATUS_COL - 1];
      const clean = status ? status.toString().toLowerCase().trim() : "";

      const center = String(id).substring(0,3);

      const toRet = sheetName;
      const fromRet = idx === 0 ? "Ret 1" : `Ret ${idx}`;

      const mapKey = `${id}-${toRet}`;

      // =============================
      // 🔥 REMOVE UNUSED COLUMN
      // =============================
      let finalRow = removeUnusedColumns(row, header);

      // =============================
      // 🔥 RESET SA FIELD IF NOT PAID
      // =============================
      if (sheetName !== "Ret 1" && clean !== "paid") {

        const resetFields = [
          "SA Retention",
          "Retention Status",
          "Churn Reason",
          "Response Notes",
          "Join Date Retention",
          "Retention Package",
          "Total Session Retention Package",
          "FP Date",
          "FP Amount",
          "FP Invoice Number",
          "Type Payment",
          "Total Actual Payment"
        ];

        resetFields.forEach(field => {
          try {
            const colIndex = col(header, field);
            finalRow[colIndex] = "";
          } catch(err) {}
        });

        Logger.log(`🧹 RESET SA FIELD → ${id} (${sheetName})`);
      }

      // =============================
      // UPDATE EXISTING
      // =============================
      if (idMap[mapKey]) {

        const existing = idMap[mapKey];
        const existingCore = existing.rowData.slice(0, finalRow.length);

        let isChanged = false;

        for (let j = 0; j < finalRow.length; j++) {
          if (existingCore[j] !== finalRow[j]) {
            isChanged = true;
            break;
          }
        }

        if (isChanged) {

          const newRow = [
            ...finalRow,
            center,
            fromRet,
            toRet
          ];

          Logger.log(`🔄 UPDATE ${id} | ${toRet}`);

          target
            .getRange(existing.rowIndex, 1, 1, newRow.length)
            .setValues([newRow]);
        }

      } 
      // =============================
      // APPEND NEW
      // =============================
      else {

        const newRow = [
          ...finalRow,
          center,
          fromRet,
          toRet
        ];

        Logger.log(`🆕 APPEND ${id} | ${fromRet} → ${toRet}`);

        appendData.push(newRow);
      }

    });
  });

  // =============================
  // BATCH INSERT
  // =============================
  if (appendData.length > 0) {

    const start = target.getLastRow() + 1;

    target
      .getRange(start, 1, appendData.length, appendData[0].length)
      .setValues(appendData);
  }

  Logger.log("✅ MASTER STACK DONE");
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

function removeUnusedColumns(row, header) {

  const removeFields = [
    "Sisa Sesi",
    "Last Class Date"
  ];

  let indexesToRemove = [];

  removeFields.forEach(field => {
    try {
      const idx = col(header, field);
      indexesToRemove.push(idx);
    } catch (e) {}
  });

  indexesToRemove.sort((a,b) => b - a);

  let newRow = [...row];

  indexesToRemove.forEach(idx => {
    newRow.splice(idx, 1);
  });

  return newRow;
}

function resetProcessFlag() {
  PropertiesService.getScriptProperties().deleteProperty("PROCESS_RUNNING");
}