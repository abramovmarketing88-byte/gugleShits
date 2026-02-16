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
    for (var col = 0; col < sections.length && col < 6; col++) {
      sheet.getRange(rowIndex, 6 + col).setValue(sections[col]);
    }
  }
  ss.toast('Готово: ' + numRows + ' стр. Результат в колонках F–K.', 'Audience Analysis', 4);
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
