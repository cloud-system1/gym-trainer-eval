
const SPREADSHEET_ID = "1XI89AYss_HIU9yG30c_mvY2sodcmJe-F1WJW0QM4hsQ";
const SHEET_NAME = "Sheet1";
const ADMIN_PASSWORD = "Ranga1";
const TIME_ZONE = "Asia/Colombo";

const WEIGHTS = {
  "Excellent": 3,
  "Good": 2,
  "Average": 1,
  "Needs Improvement": 0,
  "Poor": -1,
  "Not Familiar": null
};

// ==========================================
// JSON RESPONSE
// ==========================================

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// GOOGLE SHEET CONNECTION
// ==========================================

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(
    SPREADSHEET_ID
  );

  const sheet = spreadsheet.getSheetByName(
    SHEET_NAME
  );

  if (!sheet) {
    throw new Error(
      "Worksheet not found: " + SHEET_NAME
    );
  }

  return sheet;
}

// ==========================================
// CUSTOMER EVALUATION AND BACKUP
// ==========================================

function doPost(e) {

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    const sheet = getSheet();
    const data = JSON.parse(e.postData.contents);

    // ======================================
    // AUTOMATIC BACKUP AND CLEAR
    // ======================================

    if (
      data.action === "clear" ||
      data.action === "backupAndClear"
    ) {

      // VERIFY ADMIN PASSWORD

      if (data.password !== ADMIN_PASSWORD) {
        return jsonResponse({
          ok: false,
          error: "Invalid password"
        });
      }

      const lastRow = sheet.getLastRow();

      if (lastRow <= 1) {
        return jsonResponse({
          ok: false,
          error: "No evaluation data to back up"
        });
      }

      // READ EXISTING DATA

      const originalData = sheet
        .getDataRange()
        .getValues();

      const rowCount = originalData.length;
      const columnCount = originalData[0].length;

      // GENERATE BACKUP FILE NAME

      const backupName = "Trainer Evaluation - " +
        Utilities.formatDate(
          new Date(),
          TIME_ZONE,
          "MMMM dd, yyyy - HH-mm-ss"
        );

      // CREATE NEW GOOGLE SPREADSHEET
      // WITH SUFFICIENT ROWS AND COLUMNS

      const backup = SpreadsheetApp.create(
        backupName,
        rowCount,
        columnCount
      );

      backup.setSpreadsheetTimeZone(TIME_ZONE);

      backup.setSpreadsheetLocale(
        sheet.getParent().getSpreadsheetLocale()
      );

      // USE DEFAULT WORKSHEET
      // NO COPYING, RENAMING OR DELETING SHEETS

      const backupSheet = backup.getSheets()[0];

      // SAVE ALL EXISTING EVALUATION DATA

      backupSheet.getRange(
        1,
        1,
        rowCount,
        columnCount
      ).setValues(originalData);

      // FORMAT BACKUP HEADERS

      backupSheet.getRange(
        1,
        1,
        1,
        columnCount
      )
      .setFontWeight("bold")
      .setBackground("#17365D")
      .setFontColor("#FFFFFF");

      // FORMAT TIMESTAMP COLUMN

      if (rowCount > 1) {

        backupSheet.getRange(
          2,
          1,
          rowCount - 1,
          1
        ).setNumberFormat(
          "yyyy-MM-dd HH:mm:ss"
        );

      }

      // IMPROVE BACKUP APPEARANCE

      backupSheet.setFrozenRows(1);

      backupSheet.autoResizeColumns(
        1,
        columnCount
      );

      SpreadsheetApp.flush();

      // ====================================
      // VERIFY BACKUP BEFORE CLEARING
      // ====================================

      const backupData = backupSheet.getRange(
        1,
        1,
        rowCount,
        columnCount
      ).getValues();

      if (
        JSON.stringify(originalData) !==
        JSON.stringify(backupData)
      ) {

        throw new Error(
          "Backup verification failed. " +
          "Original data has not been cleared."
        );

      }

      // CONFIRM SOURCE DATA HAS NOT CHANGED

      const currentData = sheet
        .getDataRange()
        .getValues();

      if (
        JSON.stringify(originalData) !==
        JSON.stringify(currentData)
      ) {

        throw new Error(
          "Source data changed during backup. " +
          "Original data has not been cleared."
        );

      }

      // ====================================
      // CLEAR ORIGINAL DATA
      // ONLY AFTER SUCCESSFUL VERIFICATION
      // ====================================

      sheet.getRange(
        2,
        1,
        lastRow - 1,
        5
      ).clearContent();

      SpreadsheetApp.flush();

      return jsonResponse({

        ok: true,

        message:
          "Backup created and data cleared successfully",

        backupName: backupName,

        backupUrl: backup.getUrl(),

        backedUpRows: rowCount - 1

      });

    }

    // ======================================
    // VALIDATE REQUEST
    // ======================================

    if (data.action) {
      throw new Error("Invalid action");
    }

    if (
      !data.ratings ||
      typeof data.ratings !== "object"
    ) {

      throw new Error("No ratings received");

    }

    // ======================================
    // CUSTOMER RATING SUBMISSION
    // ======================================

    const evaluatorId =
      "CUST-" + Utilities.getUuid();

    const timestamp = new Date();

    const entries = [];

    Object.keys(data.ratings).forEach(
      function(key) {

        const rating = data.ratings[key];

        if (!rating || rating === "null") {
          return;
        }

        if (
          !Object.prototype.hasOwnProperty.call(
            WEIGHTS,
            rating
          )
        ) {

          throw new Error(
            "Invalid rating: " + rating
          );

        }

        // NOT FAMILIAR HAS NO MARKS

        const marks =
          rating === "Not Familiar"
          ? ""
          : WEIGHTS[rating];

        entries.push([

          timestamp,
          evaluatorId,
          key,
          rating,
          marks

        ]);

      }
    );

    if (entries.length === 0) {

      throw new Error(
        "Please select at least one rating"
      );

    }

    // SAVE ALL RATINGS IN ONE OPERATION

    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      entries.length,
      5
    ).setValues(entries);

    SpreadsheetApp.flush();

    return jsonResponse({

      ok: true,

      message:
        "Evaluation submitted successfully"

    });

  } catch (error) {

    return jsonResponse({

      ok: false,
      error: error.message

    });

  } finally {

    if (lock.hasLock()) {
      lock.releaseLock();
    }

  }

}

// ==========================================
// RESULTS DASHBOARD
// ==========================================

function doGet(e) {

  // ========================================
  // BACKEND HEALTH CHECK
  // ========================================

  if (
    e &&
    e.parameter &&
    e.parameter.action === "health"
  ) {

    return jsonResponse({

      ok: true,
      version: "backup-v3",
      backupMethod: "direct-values-v4",
      sheet: SHEET_NAME

    });

  }

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    const sheet = getSheet();

    const rows = sheet
      .getDataRange()
      .getValues();

    if (rows.length <= 1) {

      return jsonResponse({

        ok: true,
        data: []

      });

    }

    // REMOVE HEADER ROW

    rows.shift();

    const trainerStats = new Map();

    // ======================================
    // CALCULATE TRAINER PERFORMANCE
    // ======================================

    rows.forEach(function(row) {

      const trainerName =
        String(row[2]).trim();

      const rating = row[3];

      const marks = Number(row[4]) || 0;

      if (!trainerName) {
        return;
      }

      if (
        !Object.prototype.hasOwnProperty.call(
          WEIGHTS,
          rating
        )
      ) {

        return;

      }

      if (!trainerStats.has(trainerName)) {

        trainerStats.set(trainerName, {

          name: trainerName,

          clients: 0,

          marks: 0,

          "Excellent": 0,

          "Good": 0,

          "Average": 0,

          "Needs Improvement": 0,

          "Poor": 0,

          "Not Familiar": 0

        });

      }

      const stats = trainerStats.get(
        trainerName
      );

      stats[rating] += 1;

      // NOT FAMILIAR DOES NOT AFFECT MARKS

      if (rating !== "Not Familiar") {

        stats.clients += 1;

        stats.marks += marks;

      }

    });

    // ======================================
    // SORT TRAINERS BY PERFORMANCE
    // ======================================

    const results = Array.from(
      trainerStats.values()
    );

    results.sort(function(a, b) {

      return (

        Number(b.clients > 0) -
        Number(a.clients > 0)

      ) || (

        b.marks - a.marks

      );

    });

    // ASSIGN TRAINER RANKINGS

    let rank = 0;

    results.forEach(function(trainer) {

      trainer.rank = trainer.clients > 0
        ? ++rank
        : "—";

    });

    return jsonResponse({

      ok: true,
      data: results

    });

  } catch (error) {

    return jsonResponse({

      ok: false,
      error: error.message

    });

  } finally {

    if (lock.hasLock()) {
      lock.releaseLock();
    }

  }

}
