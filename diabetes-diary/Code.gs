// Diabetes Diary — Google Apps Script backend
// All timestamps stored as UTC ISO strings; display in UTC+3 (Moscow time)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Дневник диабетика')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ── Sheet helpers ────────────────────────────────────────────────────────────

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = {
      'Инсулин':    ['ID', 'Время (UTC)', 'Тип', 'Название', 'Единицы', 'Место', 'Заметки'],
      'Еда':        ['ID', 'Время (UTC)', 'ХЕ', 'Описание', 'Сахар до', 'ID сахара после'],
      'Сахар':      ['ID', 'Время (UTC)', 'Значение', 'Тип', 'ID еды'],
      'Настройки':  ['Ключ', 'Значение']
    };
    if (headers[name]) {
      sheet.appendRow(headers[name]);
      sheet.getRange(1, 1, 1, headers[name].length)
        .setFontWeight('bold')
        .setBackground('#4A90D9')
        .setFontColor('white');
    }
  }
  return sheet;
}

function makeId() {
  return String(Date.now()) + String(Math.floor(Math.random() * 1000));
}

// ── Settings ─────────────────────────────────────────────────────────────────

function getSetting(key) {
  const sheet = getSheet('Настройки');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) return rows[i][1];
  }
  return null;
}

function setSetting(key, value) {
  const sheet = getSheet('Настройки');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ── API: Get last insulin info ───────────────────────────────────────────────

function getLastInsulinInfo() {
  try {
    return {
      lastSite:      getSetting('lastSite')      || '',
      lastLongDose:  getSetting('lastLongDose')  || 0,
      lastShortDose: getSetting('lastShortDose') || 0,
      lastLongName:  getSetting('lastLongName')  || 'Базальный',
      lastShortName: getSetting('lastShortName') || 'Болюс'
    };
  } catch(e) {
    return { lastSite: '', lastLongDose: 0, lastShortDose: 0, lastLongName: 'Базальный', lastShortName: 'Болюс' };
  }
}

// ── API: Save insulin ────────────────────────────────────────────────────────

function saveInsulin(data) {
  try {
    const id = makeId();
    getSheet('Инсулин').appendRow([
      id,
      data.time,
      data.insulinType,    // 'short' | 'long' | 'other'
      data.insulinName,    // e.g. 'Хумалог', 'Лантус', custom
      Number(data.units),
      data.site,
      data.notes || ''
    ]);

    setSetting('lastSite', data.site);
    if (data.insulinType === 'long') {
      setSetting('lastLongDose', Number(data.units));
      setSetting('lastLongName', data.insulinName);
    } else if (data.insulinType === 'short') {
      setSetting('lastShortDose', Number(data.units));
      setSetting('lastShortName', data.insulinName);
    }

    return { ok: true, id: id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Save food ───────────────────────────────────────────────────────────

function saveFood(data) {
  try {
    const id = makeId();
    getSheet('Еда').appendRow([
      id,
      data.time,
      Number(data.he),
      data.description || '',
      data.sugarBefore  || '',
      ''
    ]);
    return { ok: true, id: id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Save sugar ──────────────────────────────────────────────────────────

function saveSugar(data) {
  try {
    const id = makeId();
    getSheet('Сахар').appendRow([
      id,
      data.time,
      Number(data.value),
      data.sugarType || 'manual',
      data.foodId    || ''
    ]);

    if (data.foodId) {
      const sheet = getSheet('Еда');
      const rows  = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.foodId)) {
          sheet.getRange(i + 1, 6).setValue(id);
          break;
        }
      }
    }

    return { ok: true, id: id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Summary for a date (YYYY-MM-DD in UTC+3) ───────────────────────────

function getSummary(dateStr) {
  try {
    function belongsToDate(utcIso) {
      // Convert UTC → UTC+3 and check date
      const d = new Date(new Date(utcIso).getTime() + 3 * 3600000);
      return d.toISOString().slice(0, 10) === dateStr;
    }

    function parseRows(sheet, mapper) {
      const rows = sheet.getDataRange().getValues();
      const result = [];
      for (let i = 1; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        const iso = String(rows[i][1]);
        if (belongsToDate(iso)) result.push(mapper(rows[i]));
      }
      return result;
    }

    const insulin = parseRows(getSheet('Инсулин'), r => ({
      id: String(r[0]), time: String(r[1]), insulinType: r[2],
      insulinName: r[3], units: r[4], site: r[5], notes: r[6]
    }));

    const food = parseRows(getSheet('Еда'), r => ({
      id: String(r[0]), time: String(r[1]), he: r[2],
      description: r[3], sugarBefore: r[4], sugarAfterId: String(r[5])
    }));

    const sugar = parseRows(getSheet('Сахар'), r => ({
      id: String(r[0]), time: String(r[1]), value: r[2],
      sugarType: r[3], foodId: String(r[4])
    }));

    const totalInsulin  = insulin.reduce((s, r) => s + (Number(r.units) || 0), 0);
    const totalLong     = insulin.filter(r => r.insulinType === 'long').reduce((s, r) => s + (Number(r.units) || 0), 0);
    const totalShort    = insulin.filter(r => r.insulinType !== 'long').reduce((s, r) => s + (Number(r.units) || 0), 0);
    const totalHE       = food.reduce((s, r) => s + (Number(r.he) || 0), 0);
    const sugarValues   = sugar.map(r => Number(r.value)).filter(Boolean);
    const avgSugar      = sugarValues.length ? (sugarValues.reduce((a, b) => a + b, 0) / sugarValues.length) : null;

    return { ok: true, insulin, food, sugar, totalInsulin, totalLong, totalShort, totalHE, avgSugar };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Update record ───────────────────────────────────────────────────────

function updateRecord(type, id, patch) {
  try {
    const sheetNames = { insulin: 'Инсулин', food: 'Еда', sugar: 'Сахар' };
    const sheet = getSheet(sheetNames[type]);
    const rows  = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== String(id)) continue;

      const row = i + 1;
      if (patch.time)        sheet.getRange(row, 2).setValue(patch.time);

      if (type === 'insulin') {
        if (patch.insulinType !== undefined) sheet.getRange(row, 3).setValue(patch.insulinType);
        if (patch.insulinName !== undefined) sheet.getRange(row, 4).setValue(patch.insulinName);
        if (patch.units       !== undefined) sheet.getRange(row, 5).setValue(Number(patch.units));
        if (patch.site        !== undefined) sheet.getRange(row, 6).setValue(patch.site);
        if (patch.notes       !== undefined) sheet.getRange(row, 7).setValue(patch.notes);
      } else if (type === 'food') {
        if (patch.he          !== undefined) sheet.getRange(row, 3).setValue(Number(patch.he));
        if (patch.description !== undefined) sheet.getRange(row, 4).setValue(patch.description);
        if (patch.sugarBefore !== undefined) sheet.getRange(row, 5).setValue(patch.sugarBefore);
      } else if (type === 'sugar') {
        if (patch.value       !== undefined) sheet.getRange(row, 3).setValue(Number(patch.value));
      }

      return { ok: true };
    }
    return { ok: false, error: 'Запись не найдена' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Delete record ───────────────────────────────────────────────────────

function deleteRecord(type, id) {
  try {
    const sheetNames = { insulin: 'Инсулин', food: 'Еда', sugar: 'Сахар' };
    const sheet = getSheet(sheetNames[type]);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Запись не найдена' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Pending after-meal sugar reminders ──────────────────────────────────

function getPendingSugars() {
  try {
    const now          = Date.now();
    const twoHoursMs   = 2 * 3600000;
    const threeHoursMs = 3 * 3600000;

    const foodSheet  = getSheet('Еда');
    const foodRows   = foodSheet.getDataRange().getValues();

    const pending = [];
    for (let i = 1; i < foodRows.length; i++) {
      const r = foodRows[i];
      if (!r[0]) continue;
      const mealTime = new Date(r[1]).getTime();
      const elapsed  = now - mealTime;
      // Meal was 2–4 hours ago and has no after-meal sugar linked
      if (elapsed >= twoHoursMs && elapsed <= threeHoursMs * 1.5 && !r[5]) {
        pending.push({ id: String(r[0]), time: String(r[1]), he: r[2], description: r[3] });
      }
    }
    return { ok: true, pending };
  } catch(e) {
    return { ok: false, pending: [] };
  }
}
