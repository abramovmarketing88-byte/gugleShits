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
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate().setTitle('OpenRouter Settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getStoredApiSettings() {
  var p = PropertiesService.getUserProperties();
  return {
    apiKey: p.getProperty('OPENROUTER_API_KEY') || '',
    model: p.getProperty('OPENROUTER_MODEL') || 'anthropic/claude-3.5-sonnet'
  };
}

function saveApiSettings(apiKey, model) {
  var p = PropertiesService.getUserProperties();
  p.setProperty('OPENROUTER_API_KEY', String(apiKey || '').trim());
  p.setProperty('OPENROUTER_MODEL', String(model || 'anthropic/claude-3.5-sonnet').trim());
}

// --- Выделить колонки и подсказки ---

function menuHighlightColumns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastCol = Math.max(11, sheet.getLastColumn());
  if (sheet.getLastRow() < 1) return;
  sheet.getRange(1, 1, 1, lastCol).setBackground('#e8f0fe').setFontWeight('bold');
  SpreadsheetApp.getActiveSpreadsheet().toast('Колонки A–K: заголовок выделен. Заполняйте A–E и запускайте Run JTBD Analysis.', 'Audience Analysis', 4);
}

// --- Run JTBD Analysis (вся логика здесь, без проверки «подключён ли») ---

function menuRunJtbdAnalysis() {
  runJtbdAnalysis();
}

/**
 * Читает A–E (активная строка или выделение), вызывает OpenRouter с JTBD-промптами, пишет результат в F–K.
 */
function runJtbdAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var settings = getStoredApiSettings();
  if (!settings.apiKey) {
    ss.toast('Сначала: Audience Analysis → Setup API Key & Model → вставьте ключ → Save Settings.', 'JTBD Analysis', 6);
    return;
  }

  var range = sheet.getActiveRange();
  var startRow, numRows;
  if (range && range.getNumRows() >= 1) {
    startRow = range.getRow();
    numRows = range.getNumRows();
  } else {
    var row = sheet.getActiveCell().getRow();
    if (row < 2) {
      ss.toast('Выделите строку с данными (со 2-й) или одну ячейку в ней.', 'JTBD Analysis', 4);
      return;
    }
    startRow = row;
    numRows = 1;
  }

  var lastRow = sheet.getLastRow();
  if (startRow < 2) {
    ss.toast('Строка 1 — заголовки. Выделите строки с данными (со 2-й).', 'JTBD Analysis', 4);
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
    var responseText = callOpenRouterChat_(settings.apiKey, settings.model, systemPrompt, userContent);
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

function callOpenRouterChat_(apiKey, model, systemPrompt, userContent) {
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
    var code = response.getResponseCode();
    var body = JSON.parse(response.getContentText());
    if (code >= 200 && code < 300 && body.choices && body.choices[0] && body.choices[0].message) {
      return body.choices[0].message.content || '';
    }
    return null;
  } catch (e) {
    return null;
  }
}
