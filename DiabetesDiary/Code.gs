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
        .setFontWeight('bold').setBackground('#4A90D9').setFontColor('white');
    }
  }
  return sheet;
}

function makeId() {
  return String(Date.now()) + String(Math.floor(Math.random() * 1000));
}

// ── Cache helpers ─────────────────────────────────────────────────────────────
// A version key is stored in cache; every write bumps the version so all
// prior summary entries become unreachable (they expire on their own 5-min TTL).

function _c() { try { return CacheService.getUserCache(); } catch(e) { return null; } }

function _invalidateCache() {
  const c = _c();
  if (c) c.put('v', String(Date.now()), 21600);
}

function _getCached(suffix) {
  const c = _c();
  if (!c) return null;
  try {
    const raw = c.get('d_' + (c.get('v') || '0') + '_' + suffix);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

function _setCached(suffix, data) {
  const c = _c();
  if (!c) return;
  try { c.put('d_' + (c.get('v') || '0') + '_' + suffix, JSON.stringify(data), 300); } catch(e) {}
}

// ── Row helper: read up to `limit` most recent data rows (no header) ─────────

function _rows(sheet, numCols, limit) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const from = Math.max(2, last - limit + 1);
  return sheet.getRange(from, 1, last - from + 1, numCols).getValues();
}

// ── Settings ─────────────────────────────────────────────────────────────────

function getSetting(key) {
  const rows = getSheet('Настройки').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) return rows[i][1];
  }
  return null;
}

function setSetting(key, value) {
  const sheet = getSheet('Настройки');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ── API: Get last insulin info ───────────────────────────────────────────────

function getLastInsulinInfo() {
  const hit = _getCached('lastIns');
  if (hit) return hit;
  try {
    const result = {
      lastSite:      getSetting('lastSite')      || '',
      lastLongDose:  getSetting('lastLongDose')  || 0,
      lastShortDose: getSetting('lastShortDose') || 0,
      lastLongName:  getSetting('lastLongName')  || 'Базальный',
      lastShortName: getSetting('lastShortName') || 'Болюс',
      needleCount:   getSetting('needleCount')   || 0
    };
    _setCached('lastIns', result);
    return result;
  } catch(e) {
    return { lastSite:'', lastLongDose:0, lastShortDose:0,
             lastLongName:'Базальный', lastShortName:'Болюс', needleCount:0 };
  }
}

function resetNeedle() {
  try { setSetting('needleCount', 0); _invalidateCache(); return { ok: true }; }
  catch(e) { return { ok: false, error: String(e) }; }
}

// ── API: Save insulin ────────────────────────────────────────────────────────

function saveInsulin(data) {
  try {
    const id = makeId();
    getSheet('Инсулин').appendRow([
      id, data.time, data.insulinType, data.insulinName,
      Number(data.units), data.site, data.notes || ''
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
    _invalidateCache();
    return { ok: true, id };
  } catch(e) { return { ok: false, error: String(e) }; }
}

// ── API: Save food ───────────────────────────────────────────────────────────

function saveFood(data) {
  try {
    const id = makeId();
    getSheet('Еда').appendRow([
      id, data.time, Number(data.he), data.description || '',
      data.sugarBefore || '', ''
    ]);
    _invalidateCache();
    return { ok: true, id };
  } catch(e) { return { ok: false, error: String(e) }; }
}

// ── API: Save sugar ──────────────────────────────────────────────────────────

function saveSugar(data) {
  try {
    const id = makeId();
    getSheet('Сахар').appendRow([
      id, data.time, Number(data.value), data.sugarType || 'manual', data.foodId || ''
    ]);
    if (data.foodId) {
      const sheet = getSheet('Еда');
      const rows  = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.foodId)) {
          sheet.getRange(i + 1, 6).setValue(id); break;
        }
      }
    }
    _invalidateCache();
    return { ok: true, id };
  } catch(e) { return { ok: false, error: String(e) }; }
}

// ── Time normalizer (reused in summary functions) ─────────────────────────────

function _normTime(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString();
  const d = new Date(value);
  return !isNaN(d.getTime()) ? d.toISOString() : String(value);
}

// ── API: Summary for a date ──────────────────────────────────────────────────
// dateStr  : "YYYY-MM-DD" local date label
// startISO : UTC ISO for local-day start  (new Date(yr,mo-1,dy).toISOString())
// endISO   : UTC ISO for local-day end    (new Date(yr,mo-1,dy,23,59,59,999).toISOString())

function getSummary(dateStr, startISO, endISO) {
  try {
    if (!startISO || !endISO) return { ok:false, error:'Не переданы startISO / endISO' };

    const hit = _getCached('sum_' + dateStr);
    if (hit) return hit;

    const result = {
      ok:true, insulin:[], food:[], sugar:[],
      totalInsulin:0, totalLong:0, totalShort:0, totalHE:0, avgSugar:null
    };

    function inRange(iso) { return iso && iso >= startISO && iso <= endISO; }

    let rows = _rows(getSheet('Инсулин'), 7, 500);
    for (const r of rows) {
      if (!r[0]) continue;
      const iso = _normTime(r[1]);
      if (!inRange(iso)) continue;
      const units = Number(r[4]) || 0;
      result.insulin.push({ id:String(r[0]), time:iso, insulinType:String(r[2]||''),
        insulinName:String(r[3]||''), units, site:String(r[5]||''), notes:String(r[6]||'') });
      result.totalInsulin += units;
      if (r[2] === 'long') result.totalLong += units; else result.totalShort += units;
    }

    rows = _rows(getSheet('Еда'), 6, 500);
    for (const r of rows) {
      if (!r[0]) continue;
      const iso = _normTime(r[1]);
      if (!inRange(iso)) continue;
      const he = Number(r[2]) || 0;
      result.food.push({ id:String(r[0]), time:iso, he,
        description:String(r[3]||''), sugarBefore:r[4]||'', sugarAfterId:String(r[5]||'') });
      result.totalHE += he;
    }

    rows = _rows(getSheet('Сахар'), 5, 500);
    let sugarSum = 0, sugarCnt = 0;
    for (const r of rows) {
      if (!r[0]) continue;
      const iso = _normTime(r[1]);
      if (!inRange(iso)) continue;
      const value = Number(r[2]);
      result.sugar.push({ id:String(r[0]), time:iso, value,
        sugarType:String(r[3]||'manual'), foodId:String(r[4]||'') });
      if (!isNaN(value) && value > 0) { sugarSum += value; sugarCnt++; }
    }

    result.avgSugar = sugarCnt ? sugarSum / sugarCnt : null;
    _setCached('sum_' + dateStr, result);
    return result;
  } catch(e) { return { ok:false, error: e && e.stack ? e.stack : String(e) }; }
}

// ── Independent insights cache (NOT version-keyed, 6h TTL) ──────────────────

function _getICache(key) {
  const c = _c();
  if (!c) return null;
  try { const r = c.get(key); return r ? JSON.parse(r) : null; } catch(e) { return null; }
}
function _setICache(key, data) {
  const c = _c();
  if (!c) return;
  try { c.put(key, JSON.stringify(data), 21600); } catch(e) {}
}

// ── API: Carb-ratio + fasting sugar insights (server-side, cached 6h) ────────

function calcInsightsServer(tzOffsetMin) {
  const cKey = 'ins1_' + (tzOffsetMin | 0);
  const hit  = _getICache(cKey);
  if (hit) return hit;

  const now     = new Date();
  const endISO  = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23,59,59,999).toISOString();
  const startISO= new Date(now.getFullYear(), now.getMonth(), now.getDate()-29, 0,0,0,0).toISOString();
  const raw     = getSummary30(startISO, endISO);
  if (!raw.ok) return { morning:null, day:null, evening:null, fasting:null, total:0 };

  function lh(iso) {
    return new Date(new Date(iso).getTime() - (tzOffsetMin|0) * 60000).getUTCHours();
  }

  const bk = { m:[], d:[], e:[] };
  for (const meal of raw.food) {
    const he = Number(meal.he);
    if (!(he > 0)) continue;
    const mealMs = new Date(meal.time).getTime();
    const mealH  = lh(meal.time);
    let units = 0;
    for (const inj of raw.insulin) {
      if (inj.insulinType !== 'short') continue;
      if (Math.abs(new Date(inj.time).getTime() - mealMs) <= 90 * 60000)
        units += Number(inj.units) || 0;
    }
    if (!(units > 0)) continue;
    const r = units / he;
    if (r < 0.3 || r > 12) continue;
    if      (mealH >= 5  && mealH < 12) bk.m.push(r);
    else if (mealH >= 12 && mealH < 17) bk.d.push(r);
    else if (mealH >= 17 && mealH < 23) bk.e.push(r);
  }

  function avg(a) { return a.length ? a.reduce((s,v) => s+v, 0)/a.length : null; }
  const fv = raw.sugar
    .filter(s => { const h = lh(s.time); return h >= 5 && h < 10; })
    .map(s => Number(s.value)).filter(v => v > 0);

  const result = {
    morning: avg(bk.m), day: avg(bk.d), evening: avg(bk.e),
    fasting: avg(fv),
    total: bk.m.length + bk.d.length + bk.e.length
  };
  _setICache(cKey, result);
  return result;
}

// ── API: Combined main-screen data — one round-trip instead of two ───────────

function getMainScreenData(dateStr, startISO, endISO, tzOffsetMin) {
  return {
    summary:  getSummary(dateStr, startISO, endISO),
    pending:  getPendingSugars(),
    insights: calcInsightsServer(tzOffsetMin)
  };
}

// ── API: Update record ───────────────────────────────────────────────────────

function updateRecord(type, id, patch) {
  try {
    const sheetNames = { insulin:'Инсулин', food:'Еда', sugar:'Сахар' };
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
      _invalidateCache();
      return { ok: true };
    }
    return { ok:false, error:'Запись не найдена' };
  } catch(e) { return { ok:false, error:String(e) }; }
}

// ── API: Delete record ───────────────────────────────────────────────────────

function deleteRecord(type, id) {
  try {
    const sheetNames = { insulin:'Инсулин', food:'Еда', sugar:'Сахар' };
    const sheet = getSheet(sheetNames[type]);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1); _invalidateCache(); return { ok:true };
      }
    }
    return { ok:false, error:'Запись не найдена' };
  } catch(e) { return { ok:false, error:String(e) }; }
}

// ── API: 30-day summary (raw records, client aggregates per day) ─────────────

function getSummary30(startISO, endISO) {
  try {
    if (!startISO || !endISO) return { ok:false, error:'Не переданы startISO / endISO' };
    function inRange(iso) { return iso && iso >= startISO && iso <= endISO; }

    const result = { ok:true, insulin:[], food:[], sugar:[] };

    let rows = _rows(getSheet('Инсулин'), 7, 3000);
    for (const r of rows) {
      if (!r[0]) continue;
      const iso = _normTime(r[1]);
      if (!inRange(iso)) continue;
      result.insulin.push({ time:iso, insulinType:String(r[2]||''), units:Number(r[4])||0 });
    }
    rows = _rows(getSheet('Еда'), 6, 3000);
    for (const r of rows) {
      if (!r[0]) continue;
      const iso = _normTime(r[1]);
      if (!inRange(iso)) continue;
      result.food.push({ time:iso, he:Number(r[2])||0 });
    }
    rows = _rows(getSheet('Сахар'), 5, 3000);
    for (const r of rows) {
      if (!r[0]) continue;
      const iso = _normTime(r[1]);
      if (!inRange(iso)) continue;
      const value = Number(r[2]);
      if (!isNaN(value)) result.sugar.push({ time:iso, value });
    }
    return result;
  } catch(e) { return { ok:false, error: e && e.stack ? e.stack : String(e) }; }
}

// ── API: Pending after-meal sugar reminders ──────────────────────────────────

function getPendingSugars() {
  try {
    const now          = Date.now();
    const twoHoursMs   = 2 * 3600000;
    const threeHoursMs = 3 * 3600000;

    const pending = [];
    const rows = _rows(getSheet('Еда'), 6, 500);
    for (const r of rows) {
      if (!r[0]) continue;
      const iso      = r[1] instanceof Date ? r[1].toISOString() : String(r[1]);
      const mealTime = new Date(iso).getTime();
      if (isNaN(mealTime)) continue;
      const elapsed = now - mealTime;
      if (elapsed >= twoHoursMs && elapsed <= threeHoursMs * 1.5 && !r[5]) {
        pending.push({ id:String(r[0]), time:iso, he:r[2], description:r[3] });
      }
    }
    return { ok:true, pending };
  } catch(e) { return { ok:false, pending:[] }; }
}
