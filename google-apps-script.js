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
  const sheet = ss.insertSheet('docs');
  const docs = [
    ['Раздел', 'Описание'],
    ['', ''],
    ['═══ ОБЩЕЕ ═══', ''],
    ['Workout Tracker', 'PWA-приложение для отслеживания силовых тренировок. Данные хранятся локально на устройстве и синхронизируются с этой Google Таблицей.'],
    ['', ''],
    ['═══ ЛИСТ "workouts" ═══', ''],
    ['id', 'Уникальный идентификатор тренировки (число, например 1711900000000). Отрицательные id — это демо-данные.'],
    ['number', 'Порядковый номер тренировки (1, 2, 3...). Используется для сортировки и отображения.'],
    ['date', 'Дата тренировки в формате YYYY-MM-DD (например 2026-03-30).'],
    ['exercises_json', 'JSON-массив упражнений тренировки. Каждый элемент — объект упражнения (см. ниже).'],
    ['', ''],
    ['═══ СТРУКТУРА УПРАЖНЕНИЯ (внутри exercises_json) ═══', ''],
    ['name', 'Название упражнения (строка). Например: "Жим лежа", "Wide Pull Ups".'],
    ['sets', 'Массив подходов. Каждый подход — объект {w, r}, где w = вес (кг), r = повторения.'],
    ['sets → w', 'Вес в килограммах (число). 0 означает упражнение с собственным весом (подтягивания, отжимания и т.д.).'],
    ['sets → r', 'Повторения (строка или число). Обычно просто число: "12", "8". Может содержать спецформаты (см. ER и составные подходы).'],
    ['er', 'Флаг ER-режима (true/false). Если true — упражнение выполнялось в режиме эффективных повторений.'],
    ['ssId', 'ID суперсета (число). Если присутствует — упражнение входит в суперсет. Два упражнения с одинаковым ssId образуют пару суперсета.'],
    ['', ''],
    ['═══ СУПЕРСЕТЫ ═══', ''],
    ['Что это', 'Суперсет — два упражнения, выполняемые поочерёдно без отдыха. Например: жим лежа + тяга штанги. Подход жима → подход тяги → подход жима → ...'],
    ['Как записывается', 'Два упражнения в exercises_json имеют одинаковое поле ssId (например ssId: 1). Это значит они выполнялись в паре. Порядок подходов чередующийся.'],
    ['Пример', '{"name":"Жим лежа","ssId":1,"sets":[{w:60,r:"10"},{w:60,r:"8"}]}, {"name":"Тяга штанги","ssId":1,"sets":[{w:50,r:"12"},{w:50,r:"10"}]} — значит: жим 60x10, тяга 50x12, жим 60x8, тяга 50x10.'],
    ['', ''],
    ['═══ ЭФФЕКТИВНЫЕ ПОВТОРЕНИЯ (ER) ═══', ''],
    ['Что это', 'ER (Effective Reps) — метод повышения интенсивности. Выполняешь подход до отказа, отдыхаешь 10-15 секунд, делаешь ещё несколько повторений, снова отдых 10-15 сек, и так далее. Все мини-подходы считаются одним подходом.'],
    ['Как записывается', 'Поле r содержит строку вида "ER{начальные}+{доп1},{доп2},..." Например: "ER10+3,2,1" означает: 10 повторений до отказа, отдых, +3 повторения, отдых, +2 повторения, отдых, +1 повторение. Итого 16 повторений.'],
    ['Пример в JSON', '{"name":"Жим лежа","er":true,"sets":[{w:60,r:"ER10+3,2,1"}]} — жим 60кг: 10 повторений до отказа, затем ещё 3+2+1 с мини-паузами по 10-15 секунд.'],
    ['Зачем', 'Позволяет добирать повторения после мышечного отказа. Эффективнее классических подходов по научным данным, т.к. больше повторений выполняется вблизи отказа.'],
    ['', ''],
    ['═══ СОСТАВНЫЕ ПОВТОРЕНИЯ ═══', ''],
    ['Формат "15s+5"', 'Иногда повторения записываются как "15s+5" — это означает 15 статических/основных повторений + 5 дополнительных (например, частичных). Не путать с ER.'],
    ['', ''],
    ['═══ ЛИСТ "exercises" ═══', ''],
    ['name', 'Название упражнения (справочник всех когда-либо выполненных упражнений).'],
    ['group', 'Основная группа мышц: chest (грудь), back (спина), shoulders (плечи), biceps (бицепс), triceps (трицепс), legs (ноги), abs (пресс). Пустая строка = не назначена.'],
    ['group2', 'Дополнительная (вторичная) группа мышц. Например, жим лёжа: group=chest, group2=triceps. Пустая строка = нет вторичной группы.'],
    ['', ''],
    ['═══ КАК АНАЛИЗИРОВАТЬ ДАННЫЕ ═══', ''],
    ['Прогрессия весов', 'Для каждого упражнения сравни максимальный вес (w) и количество повторений (r) между тренировками. Рост w или r при том же w = прогресс.'],
    ['Объём тренировки', 'Суммарное количество подходов (sets.length) по всем упражнениям. Для ER: каждая запись в sets — один рабочий подход (включая доборные повторения).'],
    ['Тоннаж', 'Для каждого подхода: w × r (для ER: w × сумму всех повторений). Сумма по всем подходам = тоннаж тренировки.'],
    ['Частота по группам', 'Используй лист exercises для определения группы мышц каждого упражнения. Считай сколько раз каждая группа тренировалась за период.'],
  ];
  sheet.getRange(1, 1, docs.length, 2).setValues(docs);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 700);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  // Bold section headers
  for (let i = 0; i < docs.length; i++) {
    if (docs[i][0].startsWith('═══')) {
      sheet.getRange(i + 1, 1, 1, 2).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#4fc3f7');
    }
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
