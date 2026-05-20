// Diabetes Diary — Google Apps Script backend
// All timestamps stored as UTC ISO strings; displayed in device local time.

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
      'Инсулин':   ['ID', 'Время (UTC)', 'Тип', 'Название', 'Единицы', 'Место', 'Заметки'],
      'Еда':       ['ID', 'Время (UTC)', 'ХЕ', 'Описание', 'Сахар до', 'ID сахара после'],
      'Сахар':     ['ID', 'Время (UTC)', 'Значение', 'Тип', 'ID еды'],
      'Настройки': ['Ключ', 'Значение']
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
      lastShortName: getSetting('lastShortName') || 'Болюс',
      needleCount:   getSetting('needleCount')   || 0
    };
  } catch(e) {
    return { lastSite: '', lastLongDose: 0, lastShortDose: 0,
             lastLongName: 'Базальный', lastShortName: 'Болюс', needleCount: 0 };
  }
}

function resetNeedle() {
  try { setSetting('needleCount', 0); return { ok: true }; }
  catch(e) { return { ok: false, error: String(e) }; }
}

// ── API: Save insulin ────────────────────────────────────────────────────────

function saveInsulin(data) {
  try {
    const id = makeId();
    getSheet('Инсулин').appendRow([
      id,
      data.time,
      data.insulinType,
      data.insulinName,
      Number(data.units),
      data.site,
      data.notes || ''
    ]);

    setSetting('lastSite', data.site);
    setSetting('needleCount', (parseInt(getSetting('needleCount')) || 0) + 1);
    if (data.insulinType === 'long') {
      setSetting('lastLongDose', Number(data.units));
      setSetting('lastLongName', data.insulinName);
    } else if (data.insulinType === 'short') {
      setSetting('lastShortDose', Number(data.units));
      setSetting('lastShortName', data.insulinName);
    }

    return { ok: true, id: id };
  } catch(e) {
    return { ok: false, error: String(e) };
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
    return { ok: false, error: String(e) };
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
    return { ok: false, error: String(e) };
  }
}

// ── API: Summary for a date ──────────────────────────────────────────────────
// dateStr  : "YYYY-MM-DD" local date label (for reference only)
// startISO : UTC ISO string for local-day midnight  (new Date(yr,mo-1,dy).toISOString())
// endISO   : UTC ISO string for local-day 23:59:59  (new Date(yr,mo-1,dy,23,59,59,999).toISOString())
// ISO string comparison is lexicographic = chronological for UTC ISO strings.

function getSummary(dateStr, startISO, endISO) {
  try {
    if (!startISO || !endISO) {
      return { ok: false, error: 'Не переданы startISO / endISO' };
    }

    const result = {
      ok: true,
      insulin: [],
      food: [],
      sugar: [],
      totalInsulin: 0,
      totalLong: 0,
      totalShort: 0,
      totalHE: 0,
      avgSugar: null
    };

    function normalizeTime(value) {
      if (!value) return '';
      if (value instanceof Date) return value.toISOString();
      const d = new Date(value);
      if (!isNaN(d.getTime())) return d.toISOString();
      return String(value);
    }

    function inRange(iso) {
      return iso && iso >= startISO && iso <= endISO;
    }

    // Инсулин
    let rows = getSheet('Инсулин').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const iso = normalizeTime(r[1]);
      if (!inRange(iso)) continue;
      const units = Number(r[4]) || 0;
      result.insulin.push({
        id:          String(r[0]),
        time:        iso,
        insulinType: String(r[2] || ''),
        insulinName: String(r[3] || ''),
        units:       units,
        site:        String(r[5] || ''),
        notes:       String(r[6] || '')
      });
      result.totalInsulin += units;
      if (r[2] === 'long') result.totalLong += units;
      else                 result.totalShort += units;
    }

    // Еда
    rows = getSheet('Еда').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const iso = normalizeTime(r[1]);
      if (!inRange(iso)) continue;
      const he = Number(r[2]) || 0;
      result.food.push({
        id:          String(r[0]),
        time:        iso,
        he:          he,
        description: String(r[3] || ''),
        sugarBefore: r[4] || '',
        sugarAfterId: String(r[5] || '')
      });
      result.totalHE += he;
    }

    // Сахар
    rows = getSheet('Сахар').getDataRange().getValues();
    let sugarSum = 0, sugarCnt = 0;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const iso = normalizeTime(r[1]);
      if (!inRange(iso)) continue;
      const value = Number(r[2]);
      result.sugar.push({
        id:        String(r[0]),
        time:      iso,
        value:     value,
        sugarType: String(r[3] || 'manual'),
        foodId:    String(r[4] || '')
      });
      if (!isNaN(value) && value > 0) { sugarSum += value; sugarCnt++; }
    }

    result.avgSugar = sugarCnt ? sugarSum / sugarCnt : null;
    return result;

  } catch (e) {
    return { ok: false, error: e && e.stack ? e.stack : String(e) };
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
      if (patch.time) sheet.getRange(row, 2).setValue(patch.time);

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
        if (patch.value !== undefined) sheet.getRange(row, 3).setValue(Number(patch.value));
      }

      return { ok: true };
    }
    return { ok: false, error: 'Запись не найдена' };
  } catch(e) {
    return { ok: false, error: String(e) };
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
    return { ok: false, error: String(e) };
  }
}

// ── API: 30-day summary (raw records, client aggregates per day) ─────────────

function getSummary30(startISO, endISO) {
  try {
    if (!startISO || !endISO) return { ok: false, error: 'Не переданы startISO / endISO' };

    function normalizeTime(value) {
      if (!value) return '';
      if (value instanceof Date) return value.toISOString();
      const d = new Date(value);
      if (!isNaN(d.getTime())) return d.toISOString();
      return String(value);
    }
    function inRange(iso) { return iso && iso >= startISO && iso <= endISO; }

    const result = { ok: true, insulin: [], food: [], sugar: [] };

    let rows = getSheet('Инсулин').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const iso = normalizeTime(r[1]);
      if (!inRange(iso)) continue;
      result.insulin.push({ time: iso, insulinType: String(r[2] || ''), units: Number(r[4]) || 0 });
    }

    rows = getSheet('Еда').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const iso = normalizeTime(r[1]);
      if (!inRange(iso)) continue;
      result.food.push({ time: iso, he: Number(r[2]) || 0 });
    }

    rows = getSheet('Сахар').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const iso = normalizeTime(r[1]);
      if (!inRange(iso)) continue;
      const value = Number(r[2]);
      if (!isNaN(value)) result.sugar.push({ time: iso, value });
    }

    return result;
  } catch(e) {
    return { ok: false, error: e && e.stack ? e.stack : String(e) };
  }
}

// ── API: Pending after-meal sugar reminders ──────────────────────────────────

function getPendingSugars() {
  try {
    const now          = Date.now();
    const twoHoursMs   = 2 * 3600000;
    const threeHoursMs = 3 * 3600000;

    const foodRows = getSheet('Еда').getDataRange().getValues();
    const pending  = [];

    for (let i = 1; i < foodRows.length; i++) {
      const r = foodRows[i];
      if (!r[0]) continue;
      const iso = r[1] instanceof Date ? r[1].toISOString() : String(r[1]);
      const mealTime = new Date(iso).getTime();
      if (isNaN(mealTime)) continue;
      const elapsed = now - mealTime;
      // Meal was 2–4.5 hours ago and has no after-meal sugar linked
      if (elapsed >= twoHoursMs && elapsed <= threeHoursMs * 1.5 && !r[5]) {
        pending.push({ id: String(r[0]), time: iso, he: r[2], description: r[3] });
      }
    }
    return { ok: true, pending };
  } catch(e) {
    return { ok: false, pending: [] };
  }
}
