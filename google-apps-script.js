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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workouts = readWorkouts(ss);
  const exercises = readExercises(ss);
  return jsonResponse({ workouts, exercises });
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  const action = payload.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Workouts ──
  if (action === 'save') {
    const sheet = getOrCreateSheet(ss, 'workouts', ['id', 'number', 'date', 'exercises_json']);
    const w = payload.workout;
    const rowIdx = findRowById(sheet, w.id);
    const rowData = [w.id, w.number, w.date, JSON.stringify(w.exercises)];
    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 1, 1, 4).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    sortSheet(sheet, 2);
    return jsonResponse({ ok: true });

  } else if (action === 'delete') {
    const sheet = getOrCreateSheet(ss, 'workouts', ['id', 'number', 'date', 'exercises_json']);
    const rowIdx = findRowById(sheet, payload.id);
    if (rowIdx > 0) sheet.deleteRow(rowIdx);
    return jsonResponse({ ok: true });

  } else if (action === 'sync') {
    const sheet = getOrCreateSheet(ss, 'workouts', ['id', 'number', 'date', 'exercises_json']);
    // Clear old data
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    // Batch write all rows at once (fast!)
    const rows = (payload.workouts || []).map(w =>
      [w.id, w.number, w.date, JSON.stringify(w.exercises)]
    );
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
    sortSheet(sheet, 2);
    // Clean up extra empty rows
    cleanEmptyRows(sheet, rows.length);
    return jsonResponse({ ok: true, count: rows.length });

  // ── Exercises ──
  } else if (action === 'sync_exercises') {
    const sheet = getOrCreateSheet(ss, 'exercises', ['name', 'group', 'group2']);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    const rows = (payload.exercises || []).map(ex =>
      [ex.name, ex.group || '', ex.group2 || '']
    );
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
    sortSheet(sheet, 1);
    cleanEmptyRows(sheet, rows.length);
    return jsonResponse({ ok: true, count: rows.length });

  } else if (action === 'get_exercises') {
    return jsonResponse(readExercises(ss));
  }

  return jsonResponse({ error: 'Unknown action' });
}

// ── Helpers ──

function readWorkouts(ss) {
  const sheet = ss.getSheetByName('workouts');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const workouts = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && row[0] !== 0) continue;
    try {
      workouts.push({ id: row[0], number: row[1], date: row[2], exercises: JSON.parse(row[3] || '[]') });
    } catch (err) {}
  }
  workouts.sort((a, b) => a.number - b.number);
  return workouts;
}

function readExercises(ss) {
  const sheet = ss.getSheetByName('exercises');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const exercises = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    exercises.push({ name: data[i][0], group: data[i][1] || '', group2: data[i][2] || '' });
  }
  return exercises;
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) return i + 1;
  }
  return -1;
}

function sortSheet(sheet, col) {
  if (sheet.getLastRow() <= 1) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column: col, ascending: true });
}

function cleanEmptyRows(sheet, dataRows) {
  const totalRows = sheet.getMaxRows();
  const neededRows = dataRows + 1; // +1 for header
  if (totalRows > neededRows + 10) {
    sheet.deleteRows(neededRows + 1, totalRows - neededRows);
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
