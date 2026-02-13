/**
 * Audience Analysis - Google Sheets add-on
 * Basic structure: custom menu, sidebar, and column setup.
 */

const HEADERS = [
  'Продукт',
  'Контекст',
  'Желаемый результат',
  'Болевые точки',
  'Текущие альтернативы',
  'Сегмент ЦА',
  'Главная задача (Main Job)',
  'Эмоциональные и социальные задачи',
  'Силы прогресса (Push & Pull)',
  'Силы сдерживания (Anxiety & Inertia)',
  'Уникальное ценностное предложение (UVP)'
];

/** Колонки вывода: F–K (6 штук). */
const OUTPUT_COLS = 6;

/** До этой колонки (вкл.) разъединяем и очищаем перед записью. Нужно шире K, иначе слияние K:L–P даёт «одинаковые» ячейки. */
const OUTPUT_UNMERGE_THRU_COL = 16;

/** Цвета для колонок: ввод A–E vs вывод F–K. */
const COLOR_INPUT = '#e3f2fd';
const COLOR_OUTPUT = '#e8f5e9';

/** Подсказки для колонок A–E (примечания к ячейкам). */
const INPUT_COLUMN_NOTES = [
  'Название продукта или услуги.',
  'Контекст или ниша (напр. «бухучёт для малого бизнеса»). Можно пусто — сегменты ЦА предложит ИИ в колонке F. Для повторного прогона: вставь сегмент из F → запусти анализ снова.',
  'Желаемый результат: какую главную цель хочет достичь аудитория?',
  'Болевые точки: что мешает, раздражает или отталкивает сейчас.',
  'Текущие альтернативы: чем решают задачу сейчас (конкуренты, старые методы).'
];

/**
 * Runs when the spreadsheet is opened. Adds the custom menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Audience Analysis')
    .addItem('Setup API Key & Model', 'showSidebar')
    .addItem('Выделить колонки и подсказки', 'applyColumnStyling')
    .addItem('Run JTBD Analysis', 'runJtbdAnalysis')
    .addToUi();
}

/** Keys for PropertiesService (user-scoped). */
const PROP_API_KEY = 'OPENROUTER_API_KEY';
const PROP_MODEL = 'OPENROUTER_MODEL';

/**
 * Saves API key and model to user properties. Call via google.script.run from the sidebar.
 * @param {string} apiKey - OpenRouter API key
 * @param {string} model - Model id (e.g. anthropic/claude-3.5-sonnet)
 */
function saveSettings(apiKey, model) {
  const props = PropertiesService.getUserProperties();
  props.setProperty(PROP_API_KEY, String(apiKey).trim());
  props.setProperty(PROP_MODEL, String(model).trim());
}

/**
 * Returns saved settings for the sidebar. Model is returned; API key is not (security).
 * @returns {{ model: string }} Object with model only.
 */
function getSettings() {
  const props = PropertiesService.getUserProperties();
  return {
    model: props.getProperty(PROP_MODEL) || 'anthropic/claude-3.5-sonnet'
  };
}

/** OpenRouter chat completions endpoint. */
const OPENROUTER_URL = 'https://openrouter.ai/api/v1/chat/completions';

/**
 * Calls OpenRouter chat completions and returns the assistant message content.
 * @param {string} promptText - User message to send.
 * @returns {string} Content of the first choice message.
 * @private
 */
function fetchAIResponse(promptText) {
  const props = PropertiesService.getUserProperties();
  const apiKey = props.getProperty(PROP_API_KEY);
  const model = (props.getProperty(PROP_MODEL) || 'anthropic/claude-3.5-sonnet').trim();

  if (!apiKey || !String(apiKey).trim()) {
    throw new Error('OpenRouter API key is not set. Use Audience Analysis > Setup API Key & Model.');
  }
  const key = apiKey.trim();

  const payload = {
    model: model,
    messages: [{ role: 'user', content: String(promptText) }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + key,
      'HTTP-Referer': 'https://docs.google.com/spreadsheets/'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(OPENROUTER_URL, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    let errMsg = 'OpenRouter API error (' + code + ')';
    try {
      const errJson = JSON.parse(text);
      if (errJson.error && errJson.error.message) {
        errMsg += ': ' + errJson.error.message;
      } else if (errJson.error && typeof errJson.error === 'string') {
        errMsg += ': ' + errJson.error;
      }
    } catch (e) { /* ignore */ }
    throw new Error(errMsg);
  }

  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error('Invalid JSON response from OpenRouter.');
  }

  const choices = data.choices;
  if (!choices || !choices.length) {
    throw new Error('OpenRouter returned no choices.');
  }
  const msg = choices[0].message;
  if (!msg || msg.content == null) {
    return '';
  }
  return String(msg.content);
}

/**
 * Opens the OpenRouter Settings sidebar (Sidebar.html).
 */
function showSidebar() {
  const html = HtmlService
    .createHtmlOutputFromFile('Sidebar')
    .setTitle('OpenRouter Settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Разделитель между сегментами в ответе ИИ. Одна строка, без пробелов по краям. */
const SEGMENT_DELIMITER = '---SEGMENT---';

/** Промпты JTBD (всё внутри проекта — без внешних библиотек; работает из любого аккаунта). */
const JTBD_SYSTEM_PROMPT = 'Ты — эксперт по маркетингу, стратег и специалист по методологии JTBD (Jobs to be Done). Твоя задача — провести JTBD-анализ и **сам предложить несколько сегментов** целевой аудитории (3–5 или все релевантные: три — значит три, пять — значит пять). Сегменты ЦА — **твой вывод**: ты их **генерируешь** по продукту и контексту. Для **каждого** сегмента делай полный анализ: Сегмент ЦА, Main Job, эм./соц. задачи, силы прогресса, силы сдерживания, UVP. Формат сегмента: тип (малый/средний/крупный бизнес, ИП), роль, характеристика — напр. «бухгалтеры, которые уже обжигались», «ИП с 1–3 компаниями». Отвечай только на русском. Тон: профессиональный, без общих фраз.';
const JTBD_USER_TEMPLATE = 'Context:\n\nПродукт: {{A}}\nКонтекст / ниша / подсказки (может быть пусто): {{B}}\nЖелаемый результат: {{C}}\nБоли (Pain Points): {{D}}\nТекущие альтернативы: {{E}}\n\nTask: Проведи JTBD-анализ. **Предложи несколько сегментов ЦА** (3–5 или все подходящие — сколько релевантных, столько и выводи). Для **каждого** сегмента — отдельный блок с полным анализом.\n\n**Формат:** между блоками сегментов — строго одна строка с текстом `---SEGMENT---` (без кавычек). Первый сегмент — сразу с ###, без вводного текста.\n\nКаждый сегмент содержит ровно 6 разделов (### заголовок, затем текст):\n\n### Сегмент ЦА (целевая аудитория):\nТип, роль + характеристика (малый/средний бизнес, бухгалтеры которые обжигались, ИП с 1–3 компаниями и т.п.).\n\n### Главная задача (Main Job):\nФормула: «Когда я [ситуация], я хочу [действие], чтобы [результат]».\n\n### Эмоциональные и социальные задачи:\nКак сегмент хочет себя чувствовать и кем казаться?\n\n### Силы прогресса (Push & Pull):\nЧто выталкивает из старого и притягивает в новое?\n\n### Силы сдерживания (Anxiety & Inertia):\nКакие страхи и привычки мешают?\n\n### Уникальное ценностное предложение (UVP):\nКонкретный оффер под этот сегмент.\n\nНе пропускай разделы. Не объединяй сегменты в один текст.';

/**
 * Собирает полный промпт (system + user) для JTBD-анализа.
 * @param {Array.<*>} rowValues - [A, B, C, D, E] из строки таблицы.
 * @returns {string}
 */
function buildJtbdPrompt(rowValues) {
  const A = rowValues[0] != null ? String(rowValues[0]).trim() : '';
  const B = rowValues[1] != null ? String(rowValues[1]).trim() : '';
  const C = rowValues[2] != null ? String(rowValues[2]).trim() : '';
  const D = rowValues[3] != null ? String(rowValues[3]).trim() : '';
  const E = rowValues[4] != null ? String(rowValues[4]).trim() : '';
  const user = JTBD_USER_TEMPLATE
    .replace(/\{\{A\}\}/g, A)
    .replace(/\{\{B\}\}/g, B)
    .replace(/\{\{C\}\}/g, C)
    .replace(/\{\{D\}\}/g, D)
    .replace(/\{\{E\}\}/g, E);
  return JTBD_SYSTEM_PROMPT + '\n\n' + user;
}

/** Сигнатуры для сопоставления разделов (индекс = колонки F–K). Хотя бы одна должна входить в заголовок. */
var JTBD_SIGNATURES = [
  ['Сегмент ЦА', 'целевая аудитория', 'малый бизнес', 'средний бизнес', 'бухгалтер', 'обжигались'],
  ['Главная задача', 'Main Job'],
  ['Эмоциональн', 'социальные задачи'],
  ['Силы прогресса', 'Push', 'Pull'],
  ['Силы сдерживания', 'Anxiety', 'Inertia'],
  ['UVP', 'УТП', 'ценностное предложение', 'уникальное ценностное']
];

/**
 * Парсит ответ ИИ по ###/##-разделам. Возвращает массив из 6 строк: [Сегмент ЦА, Main Job, эм. и соц., силы прогресса, силы сдерживания, UVP].
 * @param {string} raw - сырой ответ от fetchAIResponse
 * @returns {Array.<string>}
 */
function parseJtbdResponse(raw) {
  const text = String(raw || '').replace(/\r\n/g, '\n').trim();
  const out = ['', '', '', '', '', ''];

  function normTitle(s) {
    return String(s)
      .replace(/^\s*[\d]+[.)]\s*/, '')
      .replace(/^\s*[-–—]\s*/, '')
      .replace(/\*\*/g, '')
      .replace(/\s*\(УТП\)\s*$/gi, '')
      .replace(/\s*\(UVP\)\s*$/gi, '')
      .replace(/\s*\(Main Job\)\s*$/gi, '')
      .replace(/\s*\(Push[&\s]*Pull\)\s*$/gi, '')
      .replace(/\s*\(Anxiety[&\s]*Inertia\)\s*$/gi, '')
      .trim()
      .toLowerCase();
  }

  function sectionIndex(title) {
    const t = normTitle(title);
    for (let i = 0; i < JTBD_SIGNATURES.length; i++) {
      const sigs = JTBD_SIGNATURES[i];
      for (let j = 0; j < sigs.length; j++) {
        if (t.indexOf(sigs[j].toLowerCase()) !== -1) {
          return i;
        }
      }
    }
    return -1;
  }

  var re = /(?:^|\n)(?:###|##)\s*([^:\n]+)\s*:?\s*([\s\S]*?)(?=(?:\n(?:###|##)\s)|$)/gi;
  var m;
  while ((m = re.exec(text)) !== null) {
    var title = m[1].trim();
    var content = m[2].trim();
    var idx = sectionIndex(title);
    if (idx >= 0 && idx < OUTPUT_COLS && content) {
      out[idx] = content;
    }
  }

  if (text && out.every(function (s) { return !s; })) {
    out[0] = text;
  } else {
    var nonEmpty = out.filter(function (s) { return s.length > 0; });
    if (nonEmpty.length >= 2) {
      var first = nonEmpty[0];
      if (nonEmpty.every(function (s) { return s === first; })) {
        out[0] = text;
        for (var i = 1; i < OUTPUT_COLS; i++) { out[i] = ''; }
      }
    }
  }

  var result = [];
  for (var i = 0; i < OUTPUT_COLS; i++) {
    result.push(String(out[i] != null ? out[i] : ''));
  }
  return result;
}

/**
 * Парсит ответ ИИ с несколькими сегментами (разделитель ---SEGMENT---).
 * @param {string} raw - сырой ответ от fetchAIResponse
 * @returns {Array.<Array.<string>>} Массив строк: каждая строка — [F,G,H,I,J,K] для одного сегмента.
 */
function parseJtbdResponseMulti(raw) {
  var text = String(raw || '').replace(/\r\n/g, '\n').trim();
  var chunks = text.split(new RegExp('\\n\\s*' + SEGMENT_DELIMITER.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*\\n', 'i'));
  var rows = [];
  for (var i = 0; i < chunks.length; i++) {
    var block = chunks[i].trim();
    if (!block) { continue; }
    var parsed = parseJtbdResponse(block);
    var hasAny = parsed.some(function (s) { return s.length > 0; });
    if (hasAny) {
      rows.push(parsed);
    }
  }
  if (rows.length === 0) {
    var single = parseJtbdResponse(text);
    if (single.some(function (s) { return s.length > 0; })) {
      rows.push(single);
    }
  }
  return rows;
}

/**
 * Entry point for "Run JTBD Analysis". Reads active row or selection, runs JTBD analysis, writes to columns F–K.
 * Одна строка ввода (A–E) → несколько сегментов → несколько строк вывода (каждый сегмент — своя строка, полный анализ F–K).
 */
function runJtbdAnalysis() {
  const ui = SpreadsheetApp.getUi();
  setupColumns();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange() || sheet.getActiveCell();
  const firstRow = range.getRow();
  const lastRow = range.getLastRow();

  const dataStart = Math.max(2, firstRow);
  const dataEnd = lastRow;
  if (dataEnd < dataStart) {
    ui.alert('JTBD Analysis', 'Выделите хотя бы одну строку с данными (строка 2 или ниже) или ячейку в ней.', ui.ButtonSet.OK);
    return;
  }

  var processed = 0;
  var segmentCount = 0;
  var failed = 0;

  for (var r = dataEnd; r >= dataStart; r--) {
    var rowValues = sheet.getRange(r, 1, r, 5).getValues()[0];
    var hasData = rowValues.some(function (v) {
      return v != null && String(v).trim() !== '';
    });
    if (!hasData) {
      continue;
    }

    try {
      var unmergeRng = sheet.getRange(r, 6, r, OUTPUT_UNMERGE_THRU_COL);
      try {
        unmergeRng.breakApart();
      } catch (e) { /* ячейки не объединены */ }
      unmergeRng.clearContent();

      var prompt = buildJtbdPrompt(rowValues);
      var analysis = fetchAIResponse(prompt);
      var segmentRows = parseJtbdResponseMulti(analysis);

      if (!segmentRows.length) {
        sheet.getRange(r, 6, r, 6).setValue('[Не удалось разобрать сегменты. Проверь формат ответа ИИ.]');
        for (var c = 1; c < OUTPUT_COLS; c++) {
          sheet.getRange(r, 6 + c, r, 6 + c).setValue('');
        }
        failed++;
        continue;
      }

      for (var c = 0; c < OUTPUT_COLS; c++) {
        sheet.getRange(r, 6 + c, r, 6 + c).setValue(segmentRows[0][c]);
      }
      processed++;
      segmentCount += segmentRows.length;

      if (segmentRows.length > 1) {
        sheet.insertRowsAfter(r, segmentRows.length - 1);
        for (var i = 1; i < segmentRows.length; i++) {
          var nr = r + i;
          sheet.getRange(nr, 1, nr, 5).setValues([rowValues]);
          for (var c = 0; c < OUTPUT_COLS; c++) {
            sheet.getRange(nr, 6 + c, nr, 6 + c).setValue(segmentRows[i][c]);
          }
        }
      }
    } catch (e) {
      try {
        var errRng = sheet.getRange(r, 6, r, OUTPUT_UNMERGE_THRU_COL);
        errRng.breakApart();
        errRng.clearContent();
      } catch (err) { /* ignore */ }
      var errMsg = '[Ошибка: ' + (e.message || String(e)) + ']';
      sheet.getRange(r, 6, r, 6).setValue(String(errMsg).slice(0, 50000));
      for (var c = 1; c < OUTPUT_COLS; c++) {
        sheet.getRange(r, 6 + c, r, 6 + c).setValue('');
      }
      failed++;
    }
  }

  var message = processed > 0 || failed > 0
    ? 'Анализ выполнен.\n\nСтрок ввода: ' + processed + ', сегментов ЦА: ' + segmentCount + (failed > 0 ? '\nОшибок: ' + failed : '')
    : 'Нет строк для анализа. Заполните колонки A–E в выбранных строках.';
  ui.alert('JTBD Analysis', message, ui.ButtonSet.OK);
}

/**
 * Ensures the first row contains the required headers. Creates them if missing.
 * Headers: Продукт…Текущие альтернативы (A–E), + 6 колонок анализа (F–K).
 */
function setupColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCol = sheet.getLastColumn();
  const firstRow = lastCol >= 1
    ? sheet.getRange(1, 1, 1, Math.max(lastCol, HEADERS.length)).getValues()[0]
    : [];

  const firstRowTrimmed = firstRow.map(function (cell) {
    return typeof cell === 'string' ? cell.trim() : (cell != null ? String(cell).trim() : '');
  });

  const hasHeaders = HEADERS.every(function (h, i) {
    return firstRowTrimmed[i] === h;
  });

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  styleInputOutputColumns(sheet);
}

/**
 * Выделяет колонки ввода (A–E) и вывода (F–K) разными цветами, добавляет подсказки в примечания A1–E1.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function styleInputOutputColumns(sheet) {
  sheet.getRange(1, 1, 1, 5).setBackground(COLOR_INPUT);
  sheet.getRange(1, 6, 1, 5 + OUTPUT_COLS).setBackground(COLOR_OUTPUT);
  for (var c = 0; c < 5; c++) {
    sheet.getRange(1, 1 + c).setNote(INPUT_COLUMN_NOTES[c]);
  }
}

/** Применяет выделение и подсказки к активному листу. Вызывается из меню. */
function applyColumnStyling() {
  setupColumns();
  SpreadsheetApp.getActiveSpreadsheet().toast('Колонки A–E (ввод) — голубые, F–K (вывод) — зелёные. Подсказки в примечаниях A1–E1.', 'Готово', 5);
}
