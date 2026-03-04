// =============================================
// Google Apps Script для Workout Tracker
// =============================================
// Инструкция:
// 1. Создай новую Google Таблицу
// 2. Extensions → Apps Script
// 3. Вставь этот код (замени содержимое)
// 4. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Скопируй URL и вставь в настройки приложения
// =============================================

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('workouts');
  if (!sheet) {
    return jsonResponse({ error: 'Sheet "workouts" not found' });
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return jsonResponse([]);
  }

  // Row 1 is header: id, number, date, exercises_json
  const workouts = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && row[0] !== 0) continue;
    try {
      workouts.push({
        id: row[0],
        number: row[1],
        date: row[2],
        exercises: JSON.parse(row[3] || '[]')
      });
    } catch (err) {
      // skip bad rows
    }
  }

  workouts.sort((a, b) => a.number - b.number);
  return jsonResponse(workouts);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const action = payload.action;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('workouts');

  if (!sheet) {
    sheet = ss.insertSheet('workouts');
    sheet.appendRow(['id', 'number', 'date', 'exercises_json']);
  }

  if (action === 'save') {
    // Save a single workout
    const w = payload.workout;
    const rowIdx = findRowById(sheet, w.id);
    const rowData = [w.id, w.number, w.date, JSON.stringify(w.exercises)];

    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, 4).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    sortSheet(sheet);
    return jsonResponse({ ok: true });

  } else if (action === 'delete') {
    const rowIdx = findRowById(sheet, payload.id);
    if (rowIdx > 0) {
      sheet.deleteRow(rowIdx);
    }
    return jsonResponse({ ok: true });

  } else if (action === 'sync') {
    // Full sync — replace all data
    const workouts = payload.workouts;
    // Clear everything except header
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
    }
    for (const w of workouts) {
      sheet.appendRow([w.id, w.number, w.date, JSON.stringify(w.exercises)]);
    }
    sortSheet(sheet);
    return jsonResponse({ ok: true, count: workouts.length });
  }

  return jsonResponse({ error: 'Unknown action' });
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) return i + 1; // 1-indexed
  }
  return -1;
}

function sortSheet(sheet) {
  if (sheet.getLastRow() <= 1) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).sort({ column: 2, ascending: true });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
