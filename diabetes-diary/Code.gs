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
      lastShortName: getSetting('lastShortName') || 'Болюс',
      needleCount:   getSetting('needleCount')   || 0
    };
  } catch(e) {
    return { lastSite: '', lastLongDose: 0, lastShortDose: 0, lastLongName: 'Базальный', lastShortName: 'Болюс', needleCount: 0 };
  }
}

function resetNeedle() {
  try { setSetting('needleCount', 0); return { ok: true }; }
  catch(e) { return { ok: false, error: e.message }; }
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

// tzOffsetMinutes: value of new Date().getTimezoneOffset() on the client device
// e.g. -180 for Moscow (UTC+3). Defaults to -180 if omitted for backwards compat.
function getSummary(dateStr, tzOffsetMinutes) {
  try {
    const tzOff = (typeof tzOffsetMinutes === 'number') ? tzOffsetMinutes : -180;
    function belongsToDate(utcIso) {
      try {
        const ms = new Date(utcIso).getTime();
        if (isNaN(ms)) return false;
        // Shift UTC ms to local time ms: local = UTC − tzOffset*60000
        const d = new Date(ms - tzOff * 60000);
        return d.toISOString().slice(0, 10) === dateStr;
      } catch(e2) { return false; }
    }

    function parseRows(sheet, mapper) {
      const rows = sheet.getDataRange().getValues();
      const result = [];
      for (let i = 1; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        const iso = rows[i][1] instanceof Date ? rows[i][1].toISOString() : String(rows[i][1]);
        if (belongsToDate(iso)) result.push(mapper(rows[i]));
      }
      return result;
    }

    function toIso(v) { return v instanceof Date ? v.toISOString() : String(v); }

    const insulin = parseRows(getSheet('Инсулин'), r => ({
      id: String(r[0]), time: toIso(r[1]), insulinType: r[2],
      insulinName: r[3], units: r[4], site: r[5], notes: r[6]
    }));

    const food = parseRows(getSheet('Еда'), r => ({
      id: String(r[0]), time: toIso(r[1]), he: r[2],
      description: r[3], sugarBefore: r[4], sugarAfterId: String(r[5])
    }));

    const sugar = parseRows(getSheet('Сахар'), r => ({
      id: String(r[0]), time: toIso(r[1]), value: r[2],
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

// ── Food database ────────────────────────────────────────────────────────────
// [Название, углеводы/100г, жиры/100г, белки/100г]
const FOOD_DB = [
  // Крупы и каши
  ["Гречка сырая",            62.1, 3.3, 12.6],
  ["Гречка варёная",          19.9, 1.1,  4.2],
  ["Рис белый сырой",         78.9, 0.7,  7.0],
  ["Рис белый варёный",       27.9, 0.3,  2.7],
  ["Рис бурый варёный",       22.8, 0.8,  2.6],
  ["Овсянка хлопья",          67.5, 6.9, 11.9],
  ["Овсяная каша на воде",    14.8, 1.7,  3.0],
  ["Перловая каша варёная",   22.2, 0.4,  2.9],
  ["Пшено варёное",           21.5, 0.7,  3.0],
  ["Манная каша на молоке",   15.3, 2.5,  3.8],
  ["Кукурузная крупа варёная",22.0, 0.6,  2.1],
  ["Макароны сырые",          74.2, 1.1, 10.4],
  ["Макароны варёные",        23.2, 0.5,  3.3],
  // Хлеб и выпечка
  ["Хлеб пшеничный белый",    48.9, 1.5,  8.1],
  ["Хлеб ржаной чёрный",      40.7, 1.1,  6.6],
  ["Хлеб Бородинский",        40.8, 1.3,  6.8],
  ["Хлеб цельнозерновой",     41.3, 2.0,  8.2],
  ["Батон нарезной",          51.4, 2.9,  7.9],
  ["Хлебцы рисовые",          81.0, 2.8,  8.0],
  ["Сушки баранки",           73.0, 1.3, 10.4],
  ["Печенье сахарное",        74.4, 9.8,  7.5],
  ["Крекеры",                 65.0,15.0,  8.0],
  ["Блины",                   26.1, 5.2,  6.1],
  ["Оладьи",                  32.3, 6.1,  6.1],
  ["Пирожок с картофелем",    35.0, 6.0,  5.5],
  // Молочные продукты
  ["Молоко 2.5%",              4.7, 2.5,  2.8],
  ["Молоко 3.2%",              4.7, 3.2,  2.9],
  ["Кефир 1%",                 4.0, 1.0,  3.0],
  ["Кефир 2.5%",               4.0, 2.5,  3.0],
  ["Ряженка 4%",               4.1, 4.0,  3.0],
  ["Сметана 15%",              3.1,15.0,  2.6],
  ["Сметана 20%",              3.3,20.0,  2.5],
  ["Творог 5%",                1.8, 5.0, 17.2],
  ["Творог 9%",                2.0, 9.0, 16.7],
  ["Творог обезжиренный",      1.5, 0.6, 18.0],
  ["Йогурт натуральный",       5.3, 3.2,  4.3],
  ["Йогурт фруктовый",        12.0, 2.5,  4.0],
  ["Сыр Российский",           0.0,29.0, 23.4],
  ["Сыр Голландский",          0.0,26.8, 26.8],
  ["Сыр Адыгейский",           0.0,15.0, 19.8],
  ["Сыр плавленый",            3.0,13.0, 14.0],
  ["Молоко сгущённое с сахаром",56.3,8.5, 7.2],
  ["Сливки 10%",               4.0,10.0,  3.0],
  ["Масло сливочное",          0.8,82.5,  0.9],
  // Мясо и птица
  ["Говядина сырая",           0.0,17.4, 18.9],
  ["Говядина варёная",         0.0,14.0, 25.8],
  ["Свинина сырая",            0.0,27.8, 16.4],
  ["Свинина варёная",          0.0,25.0, 22.0],
  ["Курица филе сырое",        0.0, 1.8, 23.6],
  ["Курица варёная",           0.0, 7.4, 25.2],
  ["Индейка филе",             0.0, 1.0, 24.0],
  ["Котлета жареная",          8.5,15.0, 15.0],
  ["Сосиски молочные",         1.3,22.0, 10.5],
  ["Сардельки",                0.6,20.0, 10.1],
  ["Колбаса Докторская",       1.5,22.8, 13.7],
  ["Пельмени",                29.0, 8.0, 10.0],
  ["Котлеты из индейки",       5.0, 6.0, 19.0],
  // Рыба и морепродукты
  ["Минтай сырой",             0.0, 0.9, 15.9],
  ["Горбуша сырая",            0.0, 6.5, 20.5],
  ["Тунец консервы",           0.0, 1.1, 22.5],
  ["Сельдь солёная",           0.0, 8.5, 14.7],
  ["Скумбрия",                 0.0,13.2, 18.0],
  ["Треска",                   0.0, 0.6, 16.0],
  ["Рыбные консервы в т/с",    0.0, 4.0, 20.0],
  ["Креветки варёные",         0.0, 1.1, 20.5],
  // Яйца
  ["Яйцо куриное (1 шт 60г)", 0.72, 5.0, 7.7],
  ["Яичница глазунья",         0.8, 8.9,  9.6],
  ["Омлет",                    2.5, 8.5, 10.0],
  // Картофель и корнеплоды
  ["Картофель сырой",         17.3, 0.1,  2.0],
  ["Картофель варёный",       16.3, 0.1,  2.0],
  ["Картофель жареный",       23.2, 9.5,  2.8],
  ["Пюре картофельное",       14.3, 3.3,  2.0],
  ["Картофель фри",           27.0,13.0,  3.2],
  ["Морковь сырая",            9.5, 0.1,  1.3],
  ["Морковь варёная",         10.0, 0.2,  0.9],
  ["Свёкла варёная",           9.9, 0.0,  1.8],
  ["Редис",                    3.4, 0.1,  1.2],
  // Овощи
  ["Капуста белокочанная свежая",4.7, 0.1,  1.8],
  ["Капуста тушёная",          5.0, 3.4,  2.0],
  ["Капуста брокколи",         7.0, 0.4,  3.0],
  ["Огурец свежий",            2.5, 0.1,  0.8],
  ["Помидор свежий",           4.2, 0.2,  1.1],
  ["Помидоры черри",           3.9, 0.2,  0.9],
  ["Перец болгарский",         5.9, 0.1,  1.3],
  ["Лук репчатый",             9.1, 0.0,  1.4],
  ["Чеснок",                  33.0, 0.5,  6.4],
  ["Кабачок",                  4.6, 0.3,  0.6],
  ["Тыква",                    7.7, 0.1,  1.0],
  ["Баклажан",                 5.5, 0.1,  1.2],
  ["Горошек зелёный консервы",13.8, 0.2,  5.0],
  ["Кукуруза варёная",        22.5, 4.1,  4.1],
  ["Листовой салат",           1.8, 0.2,  1.5],
  ["Шпинат",                   2.0, 0.4,  2.9],
  // Бобовые
  ["Фасоль варёная",          21.5, 0.5,  8.7],
  ["Горох варёный",           20.4, 0.6,  8.0],
  ["Чечевица варёная",        19.5, 0.4,  9.0],
  ["Нут варёный",             20.0, 2.6,  9.0],
  // Фрукты и ягоды
  ["Яблоко",                  13.5, 0.4,  0.4],
  ["Банан",                   23.5, 0.1,  1.5],
  ["Апельсин",                11.8, 0.1,  0.9],
  ["Мандарин",                10.6, 0.3,  0.8],
  ["Груша",                   10.3, 0.3,  0.4],
  ["Виноград",                17.5, 0.2,  0.6],
  ["Слива",                    9.9, 0.3,  0.8],
  ["Персик",                  10.4, 0.1,  0.9],
  ["Абрикос",                 10.5, 0.1,  0.9],
  ["Клубника",                 7.5, 0.3,  0.8],
  ["Вишня черешня",           12.2, 0.5,  0.8],
  ["Арбуз",                    9.2, 0.1,  0.6],
  ["Дыня",                    10.3, 0.3,  0.6],
  ["Черника",                 11.5, 0.7,  1.1],
  ["Смородина чёрная",        15.4, 0.2,  1.0],
  ["Малина",                  11.9, 0.8,  0.8],
  ["Киви",                    10.3, 0.6,  1.1],
  ["Гранат",                  18.7, 0.6,  0.9],
  ["Ананас",                  13.1, 0.1,  0.4],
  ["Манго",                   15.0, 0.4,  0.5],
  ["Хурма",                   15.9, 0.4,  0.5],
  ["Инжир свежий",            19.2, 0.3,  0.7],
  // Соки
  ["Сок яблочный",             9.9, 0.1,  0.2],
  ["Сок апельсиновый",        10.4, 0.1,  0.7],
  ["Сок томатный",             3.8, 0.1,  0.8],
  ["Сок виноградный",         16.3, 0.0,  0.3],
  // Сладкое
  ["Сахар белый",             99.9, 0.0,  0.0],
  ["Мёд",                     82.4, 0.0,  0.8],
  ["Варенье",                 65.3, 0.1,  0.4],
  ["Шоколад чёрный",          48.2,35.4,  6.9],
  ["Шоколад молочный",        56.0,30.4,  6.9],
  ["Мороженое сливочное",     22.5,11.0,  3.5],
  ["Торт бисквитный",         45.0,12.0,  4.5],
  ["Зефир",                   79.8, 0.0,  0.8],
  ["Пастила",                 80.4, 0.1,  0.5],
  ["Пряники",                 77.7, 2.8,  5.8],
  ["Карамель",                94.3, 0.1,  0.0],
  ["Вафли",                   76.8,12.8,  7.4],
  ["Конфеты шоколадные",      55.0,25.0,  4.0],
  ["Халва",                   54.0,29.0, 13.0],
  // Орехи и семена
  ["Грецкий орех",            13.7,65.2, 15.2],
  ["Фундук",                  11.0,64.5, 13.1],
  ["Арахис",                   9.9,49.2, 26.3],
  ["Миндаль",                 13.0,57.7, 21.3],
  ["Семена подсолнечника",    20.0,52.9, 20.7],
  ["Кешью",                   26.7,43.8, 18.2],
  // Готовые блюда
  ["Борщ домашний",            5.0, 1.5,  2.5],
  ["Щи",                       3.5, 1.0,  2.0],
  ["Суп куриный с лапшой",     5.0, 2.0,  4.5],
  ["Плов",                    20.0, 8.0,  8.0],
  ["Голубцы",                  9.0, 7.0,  8.0],
  ["Салат Оливье",             7.0,10.0,  5.0],
  ["Окрошка",                  4.5, 3.0,  4.0],
  ["Пицца маргарита",         28.0, 8.5, 11.0],
  ["Гамбургер",               30.0,13.0, 14.0],
  ["Роллы",                   27.0, 4.0,  6.0],
  // Напитки
  ["Квас",                     5.2, 0.0,  0.2],
  ["Компот из сухофруктов",   15.0, 0.0,  0.2],
  ["Чай с сахаром (1 ч.л.)",   5.0, 0.0,  0.0],
  ["Кофе с молоком 200мл",     4.5, 2.0,  2.0],
  // Масло и жиры
  ["Масло подсолнечное",       0.0,99.9,  0.0],
  ["Майонез",                  2.6,67.0,  2.8],
  ["Кетчуп",                  22.2, 0.6,  1.8],
];

function initFoodDb() {
  const sheet = getSheet('Продукты');
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(1, 1, 1, 5).setValues([['Название', 'Углеводы/100г', 'Жиры/100г', 'Белки/100г', 'Ккал/100г']]);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#52B881').setFontColor('white');
  const rows = FOOD_DB.map(r => [r[0], r[1], r[2], r[3], Math.round(r[1]*4 + r[2]*9 + r[3]*4)]);
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

// ── API: Search foods in local sheet ─────────────────────────────────────────

function searchFoods(query) {
  try {
    initFoodDb();
    const q = (query || '').toLowerCase().trim();
    if (!q) return { ok: true, results: [] };
    const rows = getSheet('Продукты').getDataRange().getValues();
    const results = [];
    for (let i = 1; i < rows.length && results.length < 15; i++) {
      if (String(rows[i][0]).toLowerCase().includes(q)) {
        results.push({ name: rows[i][0], carbs: rows[i][1], fat: rows[i][2], protein: rows[i][3], kcal: rows[i][4] });
      }
    }
    return { ok: true, results };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Barcode lookup via Open Food Facts ───────────────────────────────────

function getFoodByBarcode(barcode) {
  try {
    const url = 'https://world.openfoodfacts.org/api/v2/product/' + barcode +
                '.json?fields=product_name,product_name_ru,nutriments';
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(resp.getContentText());
    if (data.status !== 1) return { ok: false, error: 'Продукт не найден' };
    const p = data.product, n = p.nutriments || {};
    return { ok: true, food: {
      name:    p.product_name_ru || p.product_name || 'Неизвестный продукт',
      carbs:   +(n['carbohydrates_100g'] || 0).toFixed(1),
      fat:     +(n['fat_100g']           || 0).toFixed(1),
      protein: +(n['proteins_100g']      || 0).toFixed(1),
      kcal:    +(n['energy-kcal_100g']   || 0).toFixed(0)
    }};
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── API: Online text search via Open Food Facts ───────────────────────────────

function searchFoodsOnline(query) {
  try {
    const url = 'https://world.openfoodfacts.org/cgi/search.pl?search_terms=' +
                encodeURIComponent(query) +
                '&search_simple=1&action=process&json=1&page_size=8' +
                '&fields=product_name,product_name_ru,nutriments&lc=ru&cc=ru';
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(resp.getContentText());
    if (!data.products) return { ok: true, results: [] };
    const results = data.products
      .filter(p => p.nutriments && p.nutriments['carbohydrates_100g'] != null)
      .slice(0, 8)
      .map(p => ({
        name:    p.product_name_ru || p.product_name || '—',
        carbs:   +(p.nutriments['carbohydrates_100g'] || 0).toFixed(1),
        fat:     +(p.nutriments['fat_100g']           || 0).toFixed(1),
        protein: +(p.nutriments['proteins_100g']      || 0).toFixed(1),
        kcal:    +(p.nutriments['energy-kcal_100g']   || 0).toFixed(0)
      }));
    return { ok: true, results };
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
      const iso = r[1] instanceof Date ? r[1].toISOString() : String(r[1]);
      const mealTime = new Date(iso).getTime();
      if (isNaN(mealTime)) continue;
      const elapsed  = now - mealTime;
      // Meal was 2–4 hours ago and has no after-meal sugar linked
      if (elapsed >= twoHoursMs && elapsed <= threeHoursMs * 1.5 && !r[5]) {
        pending.push({ id: String(r[0]), time: iso, he: r[2], description: r[3] });
      }
    }
    return { ok: true, pending };
  } catch(e) {
    return { ok: false, pending: [] };
  }
}
