
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

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Worksheet not found: " + SHEET_NAME);
  }

  return sheet;
}

// CUSTOMER EVALUATION AND BACKUP
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const sheet = getSheet();
    const data = JSON.parse(e.postData.contents);

    // BACKUP AND CLEAR DATA
    if (
      data.action === "clear" ||
      data.action === "backupAndClear"
    ) {
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

      const originalData = sheet.getDataRange().getValues();

      const backupName = "Trainer Evaluation - " +
        Utilities.formatDate(
          new Date(),
          TIME_ZONE,
          "MMMM dd, yyyy - HH-mm-ss"
        );

      // CREATE NEW GOOGLE SPREADSHEET
      const backup = SpreadsheetApp.create(backupName);

      backup.setSpreadsheetTimeZone(TIME_ZONE);
      backup.setSpreadsheetLocale(
        sheet.getParent().getSpreadsheetLocale()
      );

     
// Copy the original worksheet
const backupSheet = sheet.copyTo(backup);

// Remove the default worksheet first
backup.getSheets().forEach(function(s) {
  if (s.getSheetId() !== backupSheet.getSheetId()) {
    backup.deleteSheet(s);
  }
});

// Rename the copied worksheet after deletion
backupSheet.setName(SHEET_NAME);

SpreadsheetApp.flush();


      // VERIFY BACKUP BEFORE CLEARING
      const backupData = backupSheet
        .getDataRange()
        .getValues();

      if (
        JSON.stringify(originalData) !==
        JSON.stringify(backupData)
      ) {
        throw new Error(
          "Backup verification failed. Original data retained."
        );
      }

      // CLEAR DATA BUT KEEP HEADERS
      sheet.getRange(2, 1, lastRow - 1, 5)
        .clearContent();

      SpreadsheetApp.flush();

      return jsonResponse({
        ok: true,
        message: "Backup created and data cleared successfully",
        backupName: backupName,
        backupUrl: backup.getUrl()
      });
    }

    if (data.action) {
      throw new Error("Invalid action");
    }

    // CUSTOMER RATING SUBMISSION
    if (!data.ratings || typeof data.ratings !== "object") {
      throw new Error("No ratings received");
    }

    const evaluatorId = "CUST-" + Utilities.getUuid();
    const timestamp = new Date();
    const entries = [];

    Object.keys(data.ratings).forEach(function(key) {
      const rating = data.ratings[key];

      if (!rating || rating === "null") {
        return;
      }

      if (!Object.prototype.hasOwnProperty.call(WEIGHTS, rating)) {
        throw new Error("Invalid rating: " + rating);
      }

      const marks = rating === "Not Familiar"
        ? ""
        : WEIGHTS[rating];

      entries.push([
        timestamp,
        evaluatorId,
        key,
        rating,
        marks
      ]);
    });

    if (entries.length === 0) {
      throw new Error("Please select at least one rating");
    }

    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      entries.length,
      5
    ).setValues(entries);

    SpreadsheetApp.flush();

    return jsonResponse({
      ok: true,
      message: "Evaluation submitted successfully"
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

// RESULTS DASHBOARD
function doGet(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return jsonResponse({
        ok: true,
        data: []
      });
    }

    rows.shift();

    const trainerStats = new Map();

    rows.forEach(function(row) {
      const trainerName = String(row[2]).trim();
      const rating = row[3];
      const marks = Number(row[4]) || 0;

      if (!trainerName) return;

      if (!Object.prototype.hasOwnProperty.call(WEIGHTS, rating)) {
        return;
      }

      if (!trainerStats.has(trainerName)) {
        trainerStats.set(trainerName, {
          name: trainerName,
          clients: 0,
          marks: 0,
          Excellent: 0,
          Good: 0,
          Average: 0,
          "Needs Improvement": 0,
          Poor: 0,
          "Not Familiar": 0
        });
      }

      const stats = trainerStats.get(trainerName);

      stats[rating] += 1;

      // NOT FAMILIAR DOES NOT AFFECT PERFORMANCE
      if (rating !== "Not Familiar") {
        stats.clients += 1;
        stats.marks += marks;
      }
    });

    const results = Array.from(trainerStats.values());

    results.sort(function(a, b) {
      return (
        Number(b.clients > 0) -
        Number(a.clients > 0)
      ) || (b.marks - a.marks);
    });

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
