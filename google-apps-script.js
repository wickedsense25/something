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
  ensureDocsSheet(ss);
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
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sheet = getOrCreateSheet(ss, 'workouts', ['id', 'number', 'date', 'exercises_json']);
      const w = payload.workout;
      const rowIdx = findRowById(sheet, w.id);
      const rowData = [w.id, w.number, w.date, JSON.stringify(w.exercises)];
      if (rowIdx > 0) {
        sheet.getRange(rowIdx, 1, 1, 4).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      return jsonResponse({ ok: true });
    } finally {
      lock.releaseLock();
    }

  } else if (action === 'delete') {
    const sheet = getOrCreateSheet(ss, 'workouts', ['id', 'number', 'date', 'exercises_json']);
    const rowIdx = findRowById(sheet, payload.id);
    if (rowIdx > 0) sheet.deleteRow(rowIdx);
    return jsonResponse({ ok: true });

  } else if (action === 'sync') {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sheet = getOrCreateSheet(ss, 'workouts', ['id', 'number', 'date', 'exercises_json']);
      const lastRow = sheet.getLastRow();
      // Delete old data rows entirely
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      // Batch write all rows at once
      const rows = (payload.workouts || []).map(w =>
        [w.id, w.number, w.date, JSON.stringify(w.exercises)]
      );
      if (rows.length > 0) {
        sheet.insertRowsAfter(1, rows.length);
        sheet.getRange(2, 1, rows.length, 4).setValues(rows);
      }
      return jsonResponse({ ok: true, count: rows.length });
    } finally {
      lock.releaseLock();
    }

  // ── Exercises ──
  } else if (action === 'sync_exercises') {
    const sheet = getOrCreateSheet(ss, 'exercises', ['name', 'group', 'group2']);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    const rows = (payload.exercises || []).map(ex =>
      [ex.name, ex.group || '', ex.group2 || '']
    );
    if (rows.length > 0) {
      sheet.insertRowsAfter(1, rows.length);
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
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

function ensureDocsSheet(ss) {
  if (ss.getSheetByName('docs')) return;
  var s = ss.insertSheet('docs');
  var d = [];
  d.push(['Раздел', 'Описание']);
  d.push(['', '']);
  d.push(['ОБЩЕЕ', '']);
  d.push(['Workout Tracker', 'PWA для отслеживания силовых тренировок. Данные хранятся локально и синхронизируются с Google Таблицей.']);
  d.push(['', '']);
  d.push(['ЛИСТ workouts', '']);
  d.push(['id', 'Уникальный ID тренировки (число). Отрицательные id = демо-данные.']);
  d.push(['number', 'Порядковый номер тренировки (1, 2, 3...).']);
  d.push(['date', 'Дата тренировки YYYY-MM-DD.']);
  d.push(['exercises_json', 'JSON-массив упражнений тренировки.']);
  d.push(['', '']);
  d.push(['СТРУКТУРА УПРАЖНЕНИЯ', '']);
  d.push(['name', 'Название упражнения.']);
  d.push(['sets', 'Массив подходов {w, r}. w=вес(кг), r=повторения.']);
  d.push(['sets w', 'Вес в кг. 0 = собственный вес.']);
  d.push(['sets r', 'Повторения. Обычно число. Спецформаты: ER, составные.']);
  d.push(['er', 'true = режим эффективных повторений.']);
  d.push(['ssId', 'ID суперсета. Два упражнения с одинаковым ssId = пара.']);
  d.push(['', '']);
  d.push(['СУПЕРСЕТЫ', '']);
  d.push(['Что это', 'Два упражнения поочередно без отдыха.']);
  d.push(['Запись', 'Оба упражнения имеют одинаковый ssId.']);
  d.push(['Пример', 'ssId:1 у жима и тяги = чередуем подходы.']);
  d.push(['', '']);
  d.push(['ER (EFFECTIVE REPS)', '']);
  d.push(['Что это', 'Подход до отказа, отдых 10-15сек, еще повторения, отдых, еще. Один подход.']);
  d.push(['Формат', 'ER10+3,2,1 = 10 до отказа + 3 + 2 + 1 с паузами. Итого 16.']);
  d.push(['', '']);
  d.push(['СОСТАВНЫЕ ПОВТОРЕНИЯ', '']);
  d.push(['Формат', '15s+5 = 15 основных + 5 дополнительных (частичных).']);
  d.push(['', '']);
  d.push(['ЛИСТ exercises', '']);
  d.push(['name', 'Название упражнения.']);
  d.push(['group', 'Основная мышца: chest/back/shoulders/biceps/triceps/legs/abs.']);
  d.push(['group2', 'Вторичная мышца. Пусто = нет.']);
  d.push(['', '']);
  d.push(['АНАЛИЗ', '']);
  d.push(['Прогрессия', 'Сравни вес и повторения между тренировками.']);
  d.push(['Объем', 'Сумма подходов по всем упражнениям.']);
  d.push(['Тоннаж', 'Сумма (вес x повторения) по всем подходам.']);
  d.push(['Частота', 'Сколько раз каждая группа мышц тренировалась за период.']);
  s.getRange(1, 1, d.length, 2).setValues(d);
  s.setColumnWidth(1, 220);
  s.setColumnWidth(2, 500);
  s.getRange(1, 1, 1, 2).setFontWeight('bold');
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
