/**
 * Audience Analysis — единый Code.gs
 * Меню, сайдбар OpenRouter, Run JTBD Analysis (читает A–E, вызывает OpenRouter, пишет F–K).
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Audience Analysis')
    .addItem('Setup API Key & Model', 'showOpenRouterSidebar')
    .addItem('Выделить колонки и подсказки', 'menuHighlightColumns')
    .addItem('Run JTBD Analysis', 'menuRunJtbdAnalysis')
    .addItem('Разбить на сегменты (ИИ)', 'menuSplitIntoSegments')
    .addSeparator()
    .addItem('Подготовить все листы', 'menuPrepareAllSheets')
    .addItem('Обновить список сегментов', 'menuRefreshSegmentList')
    .addItem('Подставить сегмент', 'menuPullSegment')
    .addItem('Сгенерировать оффер', 'generateOffer')
    .addItem('Обнулить оффер', 'menuResetOffer')
    .addSeparator()
    .addItem('Обновить HTML оффера', 'menuRefreshOfferHtml')
    .addItem('Скопировать HTML оффера', 'menuCopyOfferHtml')
    .addItem('Создать файл автозагрузки', 'menuCreateAutouploadFile')
    .addToUi();
}

// --- OpenRouter Settings (сайдбар) ---

function showOpenRouterSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('API: сервис и модель');
  SpreadsheetApp.getUi().showSidebar(html);
}

var DEFAULT_MODELS = { OpenRouter: 'anthropic/claude-3.5-sonnet', OpenAI: 'gpt-4o-mini', Gemini: 'gemini-1.5-flash' };

function getStoredApiSettings() {
  var p = PropertiesService.getUserProperties();
  var provider = p.getProperty('CURRENT_PROVIDER') || 'OpenRouter';
  var model = p.getProperty('CURRENT_MODEL') || DEFAULT_MODELS[provider] || 'anthropic/claude-3.5-sonnet';
  var apiKey = p.getProperty(provider === 'OpenAI' ? 'OPENAI_API_KEY' : provider === 'Gemini' ? 'GEMINI_API_KEY' : 'OPENROUTER_API_KEY') || '';
  return {
    provider: provider,
    model: model,
    apiKey: apiKey,
    openRouterKey: p.getProperty('OPENROUTER_API_KEY') || '',
    openAiKey: p.getProperty('OPENAI_API_KEY') || '',
    geminiKey: p.getProperty('GEMINI_API_KEY') || ''
  };
}

function saveApiSettings(apiKey, provider, model) {
  var p = PropertiesService.getUserProperties();
  p.setProperty('CURRENT_PROVIDER', String(provider || 'OpenRouter').trim());
  p.setProperty('CURRENT_MODEL', String(model || '').trim());
  if (provider === 'OpenAI') p.setProperty('OPENAI_API_KEY', String(apiKey || '').trim());
  else if (provider === 'Gemini') p.setProperty('GEMINI_API_KEY', String(apiKey || '').trim());
  else p.setProperty('OPENROUTER_API_KEY', String(apiKey || '').trim());
}

// --- Выделить колонки и подсказки ---

/** Заголовки колонок A–K для JTBD (подставляются автоматически). */
var JTBD_HEADERS_ROW1 = [
  'Продукт',                    // A
  'Сегмент ЦА / контекст',      // B
  'Желаемый результат',         // C
  'Боли (Pain Points)',         // D
  'Текущие альтернативы',       // E
  'Сегмент ЦА',                 // F — результат анализа
  'Главная задача (Main Job)',  // G
  'Эмоц. и соц. задачи',        // H
  'Силы прогресса',             // I
  'Силы сдерживания',           // J
  'Уникальное ценностное предложение (UVP)'  // K
];

/** Подставляет заголовки в первую строку (A–K), выделяет их и показывает подсказку. */
function highlightColumnsAndHints() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, 1, JTBD_HEADERS_ROW1.length).setValues([JTBD_HEADERS_ROW1]);
  sheet.getRange(1, 1, 1, JTBD_HEADERS_ROW1.length).setBackground('#e8f0fe').setFontWeight('bold');
  SpreadsheetApp.getActiveSpreadsheet().toast('Колонки A–K переименованы и выделены. Заполняйте A–E и запускайте Run JTBD Analysis.', 'Audience Analysis', 4);
}

function menuHighlightColumns() {
  highlightColumnsAndHints();
}

// --- Run JTBD Analysis (вся логика здесь, без проверки «подключён ли») ---

function menuRunJtbdAnalysis() {
  runJtbdAnalysis();
}

/**
 * Читает A–E (активная строка, выделение или первая строка с данными), вызывает OpenRouter, пишет результат в F–K.
 */
function runJtbdAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var settings = getStoredApiSettings();
  if (!settings.apiKey) {
    ss.toast('Сначала: Audience Analysis → Setup API Key & Model → вставьте ключ → Save Settings.', 'JTBD Analysis', 6);
    return;
  }

  var lastRow = Math.max(sheet.getLastRow(), 2);
  var range = sheet.getActiveRange();
  var startRow, numRows;

  if (range && range.getNumRows() >= 1 && range.getRow() >= 2) {
    startRow = range.getRow();
    numRows = range.getNumRows();
  } else {
    var activeRow = sheet.getActiveCell().getRow();
    if (activeRow >= 2) {
      startRow = activeRow;
      numRows = 1;
    } else {
      // Ничего не выделено или выделена строка 1 — ищем первую строку с данными в колонке A (со 2-й)
      startRow = null;
      for (var r = 2; r <= lastRow; r++) {
        if (sheet.getRange(r, 1).getValue()) {
          startRow = r;
          break;
        }
      }
      if (!startRow) {
        ss.toast('Заполните хотя бы колонку A в строке 2 (например «Продукт») и снова запустите Run JTBD Analysis.', 'JTBD Analysis', 5);
        return;
      }
      numRows = 1;
    }
  }

  if (startRow < 2) {
    ss.toast('Строка 1 — заголовки. Выделите строку со 2-й или заполните A2.', 'JTBD Analysis', 4);
    return;
  }

  for (var i = 0; i < numRows; i++) {
    var rowIndex = startRow + i;
    if (rowIndex > lastRow) break;
    var a = sheet.getRange(rowIndex, 1).getValue();
    var b = sheet.getRange(rowIndex, 2).getValue();
    var c = sheet.getRange(rowIndex, 3).getValue();
    var d = sheet.getRange(rowIndex, 4).getValue();
    var e = sheet.getRange(rowIndex, 5).getValue();
    var userContent = buildJtbdUserPrompt(a, b, c, d, e);
    var systemPrompt = typeof JTBD_SYSTEM_PROMPT !== 'undefined' ? JTBD_SYSTEM_PROMPT : 'Ты — эксперт по JTBD. Отвечай на русском по разделам с ###.';
    var responseText = callChatApi_(settings.provider || 'OpenRouter', settings.apiKey, settings.model, systemPrompt, userContent);
    if (responseText === null) {
      ss.toast('Ошибка API. Проверьте ключ и модель в Setup API Key & Model.', 'JTBD Analysis', 5);
      return;
    }
    var sections = parseJtbdSections_(responseText);
    var segmentTexts = splitSegmentColumnIntoRows_(sections[0]);
    if (segmentTexts.length > 1) {
      var a = sheet.getRange(rowIndex, 1).getValue();
      var b = sheet.getRange(rowIndex, 2).getValue();
      var c = sheet.getRange(rowIndex, 3).getValue();
      var d = sheet.getRange(rowIndex, 4).getValue();
      var eVal = sheet.getRange(rowIndex, 5).getValue();
      sheet.insertRowsAfter(rowIndex, segmentTexts.length - 1);
      for (var s = 0; s < segmentTexts.length; s++) {
        var r = rowIndex + s;
        sheet.getRange(r, 1).setValue(a);
        sheet.getRange(r, 2).setValue(b);
        sheet.getRange(r, 3).setValue(c);
        sheet.getRange(r, 4).setValue(d);
        sheet.getRange(r, 5).setValue(eVal);
        sheet.getRange(r, 6).setValue(segmentTexts[s]);
        for (var col = 1; col < 6 && col < sections.length; col++) sheet.getRange(r, 6 + col).setValue(sections[col]);
      }
    } else {
      for (var col = 0; col < sections.length && col < 6; col++) {
        sheet.getRange(rowIndex, 6 + col).setValue(sections[col]);
      }
    }
  }
  ss.toast('Готово: ' + numRows + ' стр. Сегменты ЦА разбиты по строкам в колонке F.', 'Audience Analysis', 4);
}

function menuSplitIntoSegments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Audience Analysis') || ss.getActiveSheet();
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var dataRowCount = 0;
  var firstDataRow = 0;
  for (var r = 2; r <= lastRow; r++) {
    var a = sheet.getRange(r, 1).getValue();
    var f = sheet.getRange(r, 6).getValue();
    if (a || f) {
      dataRowCount++;
      if (!firstDataRow) firstDataRow = r;
    }
  }
  if (dataRowCount === 0) {
    ss.toast('Заполните хотя бы строку 2 (колонки A–E или F–K), затем снова нажмите «Разбить на сегменты».', 'Разбить на сегменты', 5);
    return;
  }
  if (dataRowCount >= 2) {
    ss.toast('Уже несколько сегментов в разных строках. Для автоподбора очистите строки 2 и ниже и оставьте одну заполненную строку.', 'Разбить на сегменты', 5);
    return;
  }
  var settings = getStoredApiSettings();
  if (!settings.apiKey) {
    ss.toast('Сначала: Audience Analysis → Setup API Key & Model → вставьте ключ и сохраните.', 'Разбить на сегменты', 5);
    return;
  }
  var rowIndex = firstDataRow;
  var a1 = sheet.getRange(rowIndex, 1).getValue();
  var a2 = sheet.getRange(rowIndex, 2).getValue();
  var a3 = sheet.getRange(rowIndex, 3).getValue();
  var a4 = sheet.getRange(rowIndex, 4).getValue();
  var a5 = sheet.getRange(rowIndex, 5).getValue();
  var a6 = sheet.getRange(rowIndex, 6).getValue();
  var a7 = sheet.getRange(rowIndex, 7).getValue();
  var a8 = sheet.getRange(rowIndex, 8).getValue();
  var a9 = sheet.getRange(rowIndex, 9).getValue();
  var a10 = sheet.getRange(rowIndex, 10).getValue();
  var a11 = sheet.getRange(rowIndex, 11).getValue();
  var userContent = 'Описание одного продукта/сегмента (одна строка из таблицы):\n\n' +
    'Продукт: ' + (a1 || '') + '\n' +
    'Сегмент ЦА / контекст: ' + (a2 || '') + '\n' +
    'Желаемый результат: ' + (a3 || '') + '\n' +
    'Боли: ' + (a4 || '') + '\n' +
    'Текущие альтернативы: ' + (a5 || '') + '\n' +
    'Сегмент ЦА (подробно): ' + (a6 || '') + '\n' +
    'Главная задача: ' + (a7 || '') + '\n' +
    'Эмоц. и соц.: ' + (a8 || '') + '\n' +
    'Силы прогресса: ' + (a9 || '') + '\n' +
    'Силы сдерживания: ' + (a10 || '') + '\n' +
    'UVP: ' + (a11 || '') + '\n\n' +
    'Создай от 3 до 7 разных сегментов целевой аудитории по этому описанию. Каждый сегмент — отдельная строка таблицы с разными болями и задачами.';
  var systemPrompt = 'Ты — эксперт по сегментации ЦА. Ответь только на русском. Для каждого сегмента выведи блок, начинающийся с ### Сегмент N (N = 1, 2, 3...). Внутри блока — ровно по одной строке на поле:\nПродукт:\nСегмент ЦА (кратко):\nЖелаемый результат:\nБоли:\nГлавная задача:\nЭмоц. и соц.:\nСилы прогресса:\nСилы сдерживания:\nUVP:\nНе пропускай поля.';
  var responseText = callChatApi_(settings.provider || 'OpenRouter', settings.apiKey, settings.model, systemPrompt, userContent);
  if (responseText === null) {
    ss.toast('Ошибка API. Проверьте ключ и модель в Setup API Key & Model.', 'Разбить на сегменты', 5);
    return;
  }
  var rows = parseSplitSegments_(responseText);
  if (rows.length === 0) {
    ss.toast('ИИ не вернул сегменты в нужном формате. Попробуйте ещё раз или добавьте строки вручную.', 'Разбить на сегменты', 5);
    return;
  }
  var maxRows = Math.min(rows.length, 7);
  sheet.getRange(2, 1, 1 + maxRows, 11).clearContent();
  for (var i = 0; i < maxRows; i++) {
    sheet.getRange(2 + i, 1, 2 + i, 11).setValues([rows[i]]);
  }
  var msg = 'Готово: создано ' + maxRows + ' сегментов в строках 2–' + (1 + maxRows) + '.';
  if (sheet.getName() !== 'Audience Analysis' && sheet.getName() !== 'SEGMENTS') {
    msg += ' Чтобы оффер видел сегменты, переименуйте этот лист в «Audience Analysis» и нажмите «Обновить список сегментов».';
  } else {
    msg += ' Обновите список сегментов на листе оффера.';
  }
  ss.toast(msg, 'Разбить на сегменты', 8);
}

function buildJtbdUserPrompt(a, b, c, d, e) {
  var t = typeof JTBD_USER_TEMPLATE !== 'undefined' ? JTBD_USER_TEMPLATE : 'Продукт: {{A}}\nСегмент: {{B}}\nРезультат: {{C}}\nБоли: {{D}}\nАльтернативы: {{E}}\n\nСоздай JTBD-анализ по разделам ###.';
  return t
    .replace(/\{\{A\}\}/g, String(a || ''))
    .replace(/\{\{B\}\}/g, String(b || ''))
    .replace(/\{\{C\}\}/g, String(c || ''))
    .replace(/\{\{D\}\}/g, String(d || ''))
    .replace(/\{\{E\}\}/g, String(e || ''));
}

/** Разбивает текст блока «Сегмент ЦА» на массив строк: по одному сегменту на элемент (по нумерации 1. 2. 3. или — / •). */
function splitSegmentColumnIntoRows_(text) {
  if (!text || !String(text).trim()) return [''];
  var s = String(text).trim();
  var parts = s.split(/\n\s*\d+\.\s*/);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var t = parts[i].trim();
    if (t) out.push(t);
  }
  if (out.length > 1) return out;
  parts = s.split(/\n\s*[-•]\s*/);
  out = [];
  for (var j = 0; j < parts.length; j++) {
    t = parts[j].trim();
    if (t) out.push(t);
  }
  if (out.length > 1) return out;
  return [s];
}

function parseJtbdSections_(text) {
  var keys = typeof JTBD_SECTION_KEYS !== 'undefined' ? JTBD_SECTION_KEYS : [
    'Сегмент ЦА (целевая аудитория)',
    'Главная задача (Main Job)',
    'Эмоциональные и социальные задачи',
    'Силы прогресса (Push & Pull)',
    'Силы сдерживания (Anxiety & Inertia)',
    'Уникальное ценностное предложение (UVP)'
  ];
  var result = ['', '', '', '', '', ''];
  var blocks = String(text || '').split(/\s*###\s*/);
  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i].trim();
    if (!block) continue;
    var firstLineEnd = block.indexOf('\n');
    var firstLine = firstLineEnd >= 0 ? block.substring(0, firstLineEnd).trim() : block;
    var body = firstLineEnd >= 0 ? block.substring(firstLineEnd + 1).trim() : '';
    for (var k = 0; k < keys.length; k++) {
      if (firstLine.indexOf(keys[k]) !== -1 || keys[k].indexOf(firstLine) !== -1) {
        result[k] = body || firstLine;
        break;
      }
    }
  }
  return result;
}

/** Парсит ответ ИИ «Разбить на сегменты»: блоки ### Сегмент N с полями Продукт:, Сегмент ЦА (кратко):, … Возвращает массив строк по 11 колонок (A–K). */
function parseSplitSegments_(text) {
  var rows = [];
  var blocks = String(text || '').split(/\s*###\s*Сегмент\s*\d+/i);
  var fieldLabels = [
    'Продукт:', 'Сегмент ЦА (кратко):', 'Желаемый результат:', 'Боли:', 'Главная задача:',
    'Эмоц. и соц.:', 'Силы прогресса:', 'Силы сдерживания:', 'UVP:'
  ];
  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b].trim();
    if (!block) continue;
    var a = '', bVal = '', c = '', d = '', e = '', g = '', h = '', i = '', j = '', k = '';
    var lines = block.split(/\n/);
    for (var l = 0; l < lines.length; l++) {
      var line = lines[l].trim();
      if (line.indexOf('Продукт:') === 0) a = line.replace(/^Продукт:\s*/i, '').trim();
      else if (line.indexOf('Сегмент ЦА (кратко):') === 0) bVal = line.replace(/^Сегмент ЦА \(кратко\):\s*/i, '').trim();
      else if (line.indexOf('Желаемый результат:') === 0) c = line.replace(/^Желаемый результат:\s*/i, '').trim();
      else if (line.indexOf('Боли:') === 0) d = line.replace(/^Боли:\s*/i, '').trim();
      else if (line.indexOf('Главная задача:') === 0) g = line.replace(/^Главная задача:\s*/i, '').trim();
      else if (line.indexOf('Эмоц. и соц.:') === 0) h = line.replace(/^Эмоц\. и соц\.:\s*/i, '').trim();
      else if (line.indexOf('Силы прогресса:') === 0) i = line.replace(/^Силы прогресса:\s*/i, '').trim();
      else if (line.indexOf('Силы сдерживания:') === 0) j = line.replace(/^Силы сдерживания:\s*/i, '').trim();
      else if (line.indexOf('UVP:') === 0) k = line.replace(/^UVP:\s*/i, '').trim();
    }
    if (a || bVal || g || k) rows.push([a, bVal, c, d, e, bVal || a, g, h, i, j, k]);
  }
  return rows;
}

/**
 * Вызов чат-API: OpenRouter, OpenAI или Gemini. Возвращает текст ответа или null.
 */
function callChatApi_(provider, apiKey, model, systemPrompt, userContent) {
  if (provider === 'OpenAI') {
    return callOpenAi_(apiKey, model, systemPrompt, userContent);
  }
  if (provider === 'Gemini') {
    return callGemini_(apiKey, model, systemPrompt, userContent);
  }
  return callOpenRouter_(apiKey, model, systemPrompt, userContent);
}

function callOpenRouter_(apiKey, model, systemPrompt, userContent) {
  var url = 'https://openrouter.ai/api/v1/chat/completions';
  var payload = {
    model: model || 'anthropic/claude-3.5-sonnet',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userContent }
    ]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    var body = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300 && body.choices && body.choices[0] && body.choices[0].message) {
      return body.choices[0].message.content || '';
    }
    return null;
  } catch (e) {
    return null;
  }
}

function callOpenAi_(apiKey, model, systemPrompt, userContent) {
  var url = 'https://api.openai.com/v1/chat/completions';
  var payload = {
    model: model || 'gpt-4o-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userContent }
    ]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    var body = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300 && body.choices && body.choices[0] && body.choices[0].message) {
      return body.choices[0].message.content || '';
    }
    return null;
  } catch (e) {
    return null;
  }
}

function callGemini_(apiKey, model, systemPrompt, userContent) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (model || 'gemini-1.5-flash') + ':generateContent?key=' + encodeURIComponent(apiKey);
  var fullText = systemPrompt + '\n\n' + userContent;
  var payload = {
    contents: [{ parts: [{ text: fullText }] }],
    generationConfig: { temperature: 0.7 }
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    var body = JSON.parse(response.getContentText());
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300 && body.candidates && body.candidates[0] && body.candidates[0].content && body.candidates[0].content.parts && body.candidates[0].content.parts[0]) {
      return body.candidates[0].content.parts[0].text || '';
    }
    return null;
  } catch (e) {
    return null;
  }
}

// --- Конструктор оффера (лист SEGMENTS + лист ОФФЕР) ---

var SEGMENTS_HEADERS = ['Код сегмента', 'Название', 'Главная боль', 'Желание', 'Страх', 'Триггер', 'Силы прогресса', 'Силы сдерживания'];
var OFFER_LABELS = ['Сегмент', 'Главная боль', 'Желание', 'Страх', 'Триггер', 'Силы прогресса', 'Силы сдерживания', 'Акция', 'Спецпредложение', 'Гарантия', 'Формат работы'];

function menuSetupOfferSheets() {
  ensureSegmentsSheet_();
  ensureOfferSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Листы готовы. Заполните SEGMENTS (несколько строк — несколько ЦА), на ОФФЕР выберите сегмент из списка.', 'Оффер', 6);
}

/** Создаёт или перезаписывает структуру всех нужных листов: Audience Analysis, SEGMENTS, ОФФЕР, Оффер HTML, Автозагрузка. */
function menuPrepareAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAudienceAnalysisSheet_(ss);
  ensureSegmentsSheet_(ss);
  ensureOfferSheet_(ss);
  ensureOfferHtmlSheet_(ss);
  ensureAutouploadSheet_(ss);
  menuRefreshSegmentList();
  ss.toast('Все листы созданы/обновлены: Audience Analysis, SEGMENTS, ОФФЕР, Оффер HTML, Автозагрузка.', 'Подготовить все листы', 6);
}

/** Лист для JTBD и сегментов: заголовки A–K, строки 2–8 под данные. Всегда перезаписывает шапку и оформление. */
function ensureAudienceAnalysisSheet_(ss) {
  var sh = ss.getSheetByName('Audience Analysis');
  if (!sh) sh = ss.insertSheet('Audience Analysis', 0);
  sh.getRange(1, 1, 1, JTBD_HEADERS_ROW1.length).setValues([JTBD_HEADERS_ROW1])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff').setWrap(true);
  sh.setFrozenRows(1);
  var w = [120, 140, 160, 180, 160, 220, 260, 200, 220, 220, 260];
  for (var c = 0; c < w.length; c++) sh.setColumnWidth(c + 1, w[c]);
  for (var r = 2; r <= 8; r++) {
    sh.getRange(r, 1, r, 11).setVerticalAlignment('top').setWrap(true)
      .setBackground(r % 2 === 0 ? '#f8f9fa' : '#fff');
  }
  sh.getRange(2, 1, 8, 11).setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
}

function menuRefreshSegmentList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureOfferSheet_();
  var segSh = getSegmentsSheet_(ss);
  var count = segSh && segSh.getLastRow() >= 2 ? segSh.getLastRow() - 1 : 0;
  var sheetName = segSh ? segSh.getName() : 'SEGMENTS';
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Список сегментов обновлён: ' + count + ' шт. из листа «' + sheetName + '». Выберите сегмент в B1 и нажмите «Подставить сегмент».',
    'Оффер',
    5
  );
}

function getSegmentsSheet_(ss) {
  return ss.getSheetByName('SEGMENTS') || ss.getSheetByName('Audience Analysis');
}

function getOfferSheet_(ss) {
  return ss.getSheetByName('ОФФЕР') || ss.getSheetByName('Оффер') || ss.getSheetByName('Запуск Авито');
}

function getOfferHtmlSheet_(ss) {
  return ss.getSheetByName('Оффер HTML');
}

function ensureSegmentsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SEGMENTS');
  if (!sh) sh = ss.insertSheet('SEGMENTS');
  sh.getRange(1, 1, 1, SEGMENTS_HEADERS.length).setValues([SEGMENTS_HEADERS])
    .setFontWeight('bold').setBackground('#2e7d32').setFontColor('#fff').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  var widths = [100, 180, 220, 200, 180, 180, 220, 220];
  for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
  for (var r = 2; r <= 8; r++) {
    sh.getRange(r, 1, r, 8).setBackground(r % 2 === 0 ? '#f5f5f5' : '#fff').setVerticalAlignment('top').setWrap(true);
  }
  sh.getRange(2, 1, 8, 8).setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
}

function ensureOfferSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOfferSheet_(ss);
  if (!sh) sh = ss.insertSheet('ОФФЕР');
  for (var r = 1; r <= OFFER_LABELS.length; r++) {
    sh.getRange(r, 1).setValue(OFFER_LABELS[r - 1]);
  }
  sh.getRange(1, 1, 11, 1).setFontWeight('bold').setVerticalAlignment('middle');
  sh.getRange(1, 1, 7, 2).setBackground('#e3f2fd');
  sh.getRange(8, 1, 11, 2).setBackground('#fff8e1');
  sh.getRange(1, 1, 11, 2).setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  var segSh = getSegmentsSheet_(ss);
  if (segSh && segSh.getLastRow() >= 2) {
    try {
      var nameCol = (segSh.getName() === 'Audience Analysis') ? 1 : 2;
      var range = segSh.getRange(2, nameCol, segSh.getLastRow(), nameCol);
      var dv = SpreadsheetApp.newDataValidation().requireValueInRange(range).setAllowInvalid(false).build();
      sh.getRange(1, 2).clearDataValidations().setDataValidation(dv);
    } catch (e) {}
  }
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 380);
  sh.getRange(13, 1).setValue('Оффер').setFontWeight('bold').setBackground('#fff8e1');
  sh.getRange(13, 2).setWrap(true).setVerticalAlignment('top').setBackground('#fffde7');
  sh.setRowHeight(13, 120);
  sh.getRange(13, 1, 13, 2).setBorder(true, true, true, true, true, true, '#ffc107', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/** Лист «Оффер HTML»: текст оффера в HTML для копирования в объявление. */
function ensureOfferHtmlSheet_(ss) {
  var sh = ss.getSheetByName('Оффер HTML');
  if (!sh) sh = ss.insertSheet('Оффер HTML');
  sh.getRange(1, 1).setValue('HTML оффера (для вставки в объявление)').setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(2, 1).setValue('Содержимое (обновить из меню «Обновить HTML оффера»)').setFontWeight('bold');
  sh.getRange(1, 1, 2, 2).setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 420);
  sh.getRange(3, 2).setWrap(true).setVerticalAlignment('top').setBackground('#f8f9fa');
  sh.setRowHeight(3, 200);
}

/**
 * Лист «Автозагрузка» — файл для автозагрузки на Авито.
 * Структура: Заголовок до 50 симв. (без форматирования и смайлов, ВЧ/НЧ/СЧ по запросу),
 * текст объявления в HTML из двух частей: оффер/спецофер + универсальный блок (гарантия и т.д.), смайлы можно.
 * Итог: из одной задачи (напр. услуги бухгалтера) — много объявлений (заголовки × спецоферы × вариации).
 */
var AUTOUPLOAD_HEADERS = [
  'Заголовок (до 50 симв, без смайлов)',  // A — ВЧ/НЧ/СЧ в зависимости от запроса
  'Оффер_спецофер_HTML',                  // B — часть 1 текста объявления (HTML, смайлы можно)
  'Универсальный_блок_HTML',              // C — часть 2: гарантия, формат работы и т.д.
  'Цена', 'Категория', 'Регион', 'Фото_URL', 'Контакты'  // D–H
];
var AUTOUPLOAD_TITLE_MAX_LEN = 50;

function ensureAutouploadSheet_(ss) {
  var sh = ss.getSheetByName('Автозагрузка');
  if (!sh) sh = ss.insertSheet('Автозагрузка');
  sh.getRange(1, 1, 1, AUTOUPLOAD_HEADERS.length).setValues([AUTOUPLOAD_HEADERS])
    .setFontWeight('bold').setBackground('#0d9488').setFontColor('#fff');
  sh.setFrozenRows(1);
  var w = [220, 280, 280, 80, 120, 100, 180, 120];
  for (var c = 0; c < w.length; c++) sh.setColumnWidth(c + 1, w[c]);
  for (var r = 2; r <= 101; r++) {
    sh.getRange(r, 1, r, AUTOUPLOAD_HEADERS.length).setVerticalAlignment('top').setWrap(true)
      .setBackground(r % 2 === 0 ? '#f0fdfa' : '#fff');
  }
  sh.getRange(2, 1, 101, AUTOUPLOAD_HEADERS.length).setBorder(true, true, true, true, true, true, '#99f6e4', SpreadsheetApp.BorderStyle.SOLID);
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (name !== 'ОФФЕР' && name !== 'Оффер' && name !== 'Запуск Авито') return;
  if (e.range.getRow() !== 1 || e.range.getColumn() !== 2) return;
  pullSegmentToOffer_(e.source);
}

/** По одной строке листа (SEGMENTS 8 колонок или Audience Analysis 11 колонок) возвращает [code, name, pain, desire, fear, trigger, progress, hindrance]. */
function mapSegmentRowToOffer_(row, isAudienceAnalysis) {
  if (isAudienceAnalysis && row.length >= 11) {
    return [
      row[0],  // A Продукт → код
      (row[5] || row[0] || '').toString().trim() || (row[0] || '').toString().trim(), // F или A → название
      row[3],  // D Боли → главная боль
      row[2],  // C Желаемый результат → желание
      row[7],  // H Эмоц. и соц. → страх (близко по смыслу)
      row[6],  // G Главная задача → триггер
      row[8],  // I Силы прогресса
      row[9]   // J Силы сдерживания
    ];
  }
  if (row.length >= 8) {
    return [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]];
  }
  return null;
}

/** Подставляет характеристики выбранного в B1 сегмента в B2:B7. Читает из SEGMENTS (8 кол.) или Audience Analysis (11 кол. A–K). */
function pullSegmentToOffer_(ss) {
  var offerSh = getOfferSheet_(ss);
  var segSh = getSegmentsSheet_(ss);
  if (!offerSh || !segSh || segSh.getLastRow() < 2) return false;
  var selectedName = offerSh.getRange(1, 2).getValue();
  if (!selectedName) return false;
  selectedName = String(selectedName).trim();
  var isJtbd = segSh.getName() === 'Audience Analysis';
  var numCols = isJtbd ? 11 : 8;
  var data = segSh.getRange(2, 1, segSh.getLastRow(), numCols).getValues();
  for (var i = 0; i < data.length; i++) {
    var mapped = mapSegmentRowToOffer_(data[i], isJtbd);
    if (!mapped) continue;
    var rowCode = String(mapped[0] || '').trim();
    var rowName = String(mapped[1] || '').trim();
    if (rowName === selectedName || rowCode === selectedName) {
      for (var c = 0; c < 6; c++) offerSh.getRange(2 + c, 2).setValue(mapped[2 + c] != null ? mapped[2 + c] : '');
      return true;
    }
  }
  return false;
}

function menuPullSegment() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ok = pullSegmentToOffer_(ss);
  if (ok) {
    ss.toast('Характеристики сегмента подставлены в B2:B7.', 'Оффер', 4);
  } else {
    ss.toast('В B1 выберите сегмент из выпадающего списка (или введите название из листа SEGMENTS) и нажмите снова.', 'Оффер', 5);
  }
}

function generateOffer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOfferSheet_(ss);
  if (!sh) {
    ss.toast('Сначала: Audience Analysis → Подготовить листы оффера.', 'Оффер', 4);
    return;
  }
  var segment = sh.getRange(1, 2).getValue();
  var pain = sh.getRange(2, 2).getValue();
  var desire = sh.getRange(3, 2).getValue();
  var fear = sh.getRange(4, 2).getValue();
  var trigger = sh.getRange(5, 2).getValue();
  var progress = sh.getRange(6, 2).getValue();
  var restrain = sh.getRange(7, 2).getValue();
  var action = sh.getRange(8, 2).getValue();
  var special = sh.getRange(9, 2).getValue();
  var guarantee = sh.getRange(10, 2).getValue();
  var format = sh.getRange(11, 2).getValue();
  var parts = [];
  if (segment) parts.push('Сегмент: ' + segment);
  if (pain) parts.push('Главная боль: ' + pain);
  if (desire) parts.push('Желание: ' + desire);
  if (fear) parts.push('Страх: ' + fear);
  if (trigger) parts.push('Триггер: ' + trigger);
  if (progress) parts.push('Силы прогресса: ' + progress);
  if (restrain) parts.push('Силы сдерживания: ' + restrain);
  if (action) parts.push('Акция: ' + action);
  if (special) parts.push('Спецпредложение: ' + special);
  if (guarantee) parts.push('Гарантия: ' + guarantee);
  if (format) parts.push('Формат работы: ' + format);
  var text = parts.length ? parts.join('\n\n') : 'Выберите сегмент и заполните поля.';
  sh.getRange(13, 2).setValue(text);
  var htmlSh = getOfferHtmlSheet_(ss);
  if (htmlSh) {
    var html = String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
    htmlSh.getRange(3, 2).setValue(html);
  }
  ss.toast('Оффер сформирован в B13 и на листе «Оффер HTML».', 'Оффер', 3);
}

function menuResetOffer() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Обнулить оффер',
    'Всё сбросится: сегмент, поля и текст оффера. Не забудьте сохранить нужные данные перед сбросом.\n\nПродолжить?',
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) resetOffer_();
}

function resetOffer_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOfferSheet_(ss);
  if (!sh) return;
  sh.getRange(1, 2, 13, 2).clearContent();
  sh.getRange(13, 2).setValue('Выберите сегмент и заполните поля.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Оффер обнулён.', 'Оффер', 3);
}

// --- Оффер HTML (копирование и автозагрузка) ---

/** Возвращает текст оффера в виде HTML (переносы → &lt;br&gt;). Для диалога «Скопировать». */
function getOfferHtmlContent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var htmlSh = getOfferHtmlSheet_(ss);
  var offerSh = getOfferSheet_(ss);
  var raw = '';
  if (htmlSh && htmlSh.getRange(3, 2).getValue()) raw = htmlSh.getRange(3, 2).getValue();
  else if (offerSh) raw = offerSh.getRange(13, 2).getValue() || '';
  return String(raw).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

/** Берёт оффер с листа ОФФЕР B13, конвертирует в HTML и записывает на лист «Оффер HTML» в B3. */
function menuRefreshOfferHtml() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var offerSh = getOfferSheet_(ss);
  var htmlSh = getOfferHtmlSheet_(ss);
  if (!offerSh) {
    ss.toast('Сначала: Подготовить все листы.', 'Оффер HTML', 4);
    return;
  }
  if (!htmlSh) ensureOfferHtmlSheet_(ss);
  htmlSh = getOfferHtmlSheet_(ss);
  var text = offerSh.getRange(13, 2).getValue() || '';
  var html = String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  htmlSh.getRange(3, 2).setValue(html);
  ss.toast('HTML оффера обновлён на листе «Оффер HTML».', 'Оффер HTML', 3);
}

/** Открывает диалог с HTML оффера и кнопкой «Копировать». */
function menuCopyOfferHtml() {
  var html = getOfferHtmlContent();
  if (!html) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Сначала сгенерируйте оффер и нажмите «Обновить HTML оффера».', 'Скопировать HTML', 4);
    return;
  }
  var escaped = html.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  var template = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>HTML оффера</title></head><body style="font-family:sans-serif;padding:12px;">' +
    '<p><strong>HTML оффера</strong> — нажмите «Копировать», затем вставьте в объявление (Ctrl+V).</p>' +
    '<textarea id="t" rows="12" style="width:100%;box-sizing:border-box;font-size:12px;">' + escaped + '</textarea>' +
    '<p><button onclick="copy()">Копировать</button> <span id="msg"></span></p>' +
    '<script>function copy(){var t=document.getElementById("t");t.select();t.setSelectionRange(0,99999);try{document.execCommand("copy");document.getElementById("msg").textContent="Скопировано.";}catch(e){document.getElementById("msg").textContent="Выделите текст и Ctrl+C.";}}<\/script></body></html>';
  var output = HtmlService.createHtmlOutput(template).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, 'Скопировать HTML оффера');
}

/** Обрезает заголовок до 50 символов (для Авито), без переносов. */
function trimTitleForAvito_(title) {
  var s = String(title == null ? '' : title).replace(/\s+/g, ' ').trim();
  return s.length > AUTOUPLOAD_TITLE_MAX_LEN ? s.substring(0, AUTOUPLOAD_TITLE_MAX_LEN) : s;
}

/** Собирает текст объявления: часть 1 (оффер/спецофер) + часть 2 (универсальный блок) в одном HTML. */
function buildAvitoDescriptionHtml_(part1, part2) {
  var p1 = String(part1 == null ? '' : part1).trim();
  var p2 = String(part2 == null ? '' : part2).trim();
  if (!p1 && !p2) return '';
  if (!p1) return p2;
  if (!p2) return p1;
  return p1 + '<br><br>' + p2;
}

/** Создаёт CSV для автозагрузки на Авито: Заголовок (до 50 симв), Текст_HTML (часть1+часть2), Цена, Категория, Регион, Фото_URL, Контакты. */
function menuCreateAutouploadFile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Автозагрузка');
  if (!sh) {
    ensureAutouploadSheet_(ss);
    sh = ss.getSheetByName('Автозагрузка');
  }
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    ss.toast('Заполните строки 2 и ниже на листе «Автозагрузка».', 'Автозагрузка', 4);
    return;
  }
  var data = sh.getRange(2, 1, lastRow, AUTOUPLOAD_HEADERS.length).getValues();
  var csvRows = [];
  var exportHeaders = ['Заголовок', 'Текст_HTML', 'Цена', 'Категория', 'Регион', 'Фото_URL', 'Контакты'];
  csvRows.push(exportHeaders.join(';'));
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var title = trimTitleForAvito_(row[0]);
    var descHtml = buildAvitoDescriptionHtml_(row[1], row[2]);
    var rest = [row[3], row[4], row[5], row[6], row[7]];
    var outRow = [title, descHtml].concat(rest);
    csvRows.push(outRow.map(function(cell) {
      var s = String(cell == null ? '' : cell);
      if (/[";\n\r]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    }).join(';'));
  }
  var csv = '\uFEFF' + csvRows.join('\r\n');
  var blob = Utilities.newBlob(csv, 'text/csv;charset=utf-8', 'avito_autoupload.csv');
  var folder;
  try {
    var parentIt = DriveApp.getFileById(ss.getId()).getParents();
    folder = parentIt.hasNext() ? parentIt.next() : DriveApp.getRootFolder();
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
  var file = folder.createFile(blob);
  var folderName = file.getParents().hasNext() ? file.getParents().next().getName() : 'Диск';
  ss.toast('Файл для Авито: ' + file.getName() + '\nОбъявлений: ' + data.length + ' (заголовок до 50 симв., текст = оффер + универсальный блок)\nПапка: ' + folderName, 'Автозагрузка', 10);
}
