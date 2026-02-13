// INPUT column indices (1-based)
var INPUT_COL_ID = 1;
var INPUT_COL_PRODUCT = 2;
var INPUT_COL_CITY = 3;
var INPUT_COL_SOURCE_TEXT = 4;
var INPUT_COL_TONE = 5;
var INPUT_COL_CONSTRAINTS = 6;
var INPUT_COL_MODE = 7;
var INPUT_COL_STATUS = 8;
var INPUT_COL_LAST_ERROR = 9;
var INPUT_COL_UPDATED_AT = 10;
var INPUT_COL_LOCKED_AT = 11;
var INPUT_COL_LOCKED_BY = 12;
var INPUT_COL_PROCESSING = 13;

var LOCK_DEDUP_MINUTES = 10;
var CHECKPOINT_SECONDS_LEFT = 25;

var VALID_MODES = ['FULL', 'OFFER_ONLY', 'SPIN_ONLY', 'HTML_ONLY'];
var VALID_STATUSES = ['NEW', 'QUEUED', 'PROCESSING', 'OFFER_OK', 'FORMAT_OK', 'SPIN_OK', 'HTML_OK', 'QA_OK', 'DONE', 'ERROR'];

function loadSettings_(ss) {
  var sh = ss.getSheetByName('SETTINGS');
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0] || '').trim();
    var v = String(values[i][1] || '').trim();
    if (k) map[k] = v;
  }
  var fallback = map.MODEL_FALLBACK || map.MODEL_IMPROVE || 'openai/gpt-4o-mini';
  return {
    OPENROUTER_API_KEY: map.OPENROUTER_API_KEY,
    MODEL_OFFER: map.MODEL_OFFER || map.MODEL_IMPROVE || fallback,
    MODEL_FORMAT: map.MODEL_FORMAT || map.MODEL_AVITO || fallback,
    MODEL_FALLBACK: fallback,
    PROMPT_VERSION_OFFER: map.PROMPT_VERSION_OFFER || '1',
    PROMPT_VERSION_FORMAT: map.PROMPT_VERSION_FORMAT || '1',
    MODEL_IMPROVE: map.MODEL_IMPROVE || fallback,
    MODEL_SPIN: map.MODEL_SPIN || fallback,
    MODEL_AVITO: map.MODEL_AVITO || fallback,
    MAX_ROWS_PER_RUN: Number(map.MAX_ROWS_PER_RUN || 10),
    BATCH_SIZE: Number(map.BATCH_SIZE || map.MAX_ROWS_PER_RUN || 50),
    MAX_RUNTIME_SECONDS: Number(map.MAX_RUNTIME_SECONDS || 300),
    TEMPERATURE_IMPROVE: Number(map.TEMPERATURE_IMPROVE || 0.4),
    TEMPERATURE_SPIN: Number(map.TEMPERATURE_SPIN || 0.6),
    TEMPERATURE_AVITO: Number(map.TEMPERATURE_AVITO || 0.5),
    TEMPERATURE_OFFER: Number(map.TEMPERATURE_OFFER || map.TEMPERATURE_IMPROVE || 0.4),
    TEMPERATURE_FORMAT: Number(map.TEMPERATURE_FORMAT || map.TEMPERATURE_AVITO || 0.5),
    AVITO_STYLE: map.AVITO_STYLE || 'bold_emojis_safe',
    LANGUAGE: map.LANGUAGE || 'ru',
    LIMIT_TITLE_CHARS: Number(map.LIMIT_TITLE_CHARS || 60),
    LIMIT_DESC_CHARS: Number(map.LIMIT_DESC_CHARS || 4000),
    EMOJI_MAX: Number(map.EMOJI_MAX || 15),
    CAPS_MAX_PERCENT: Number(map.CAPS_MAX_PERCENT || 0.35),
    AUTO_FIX: String(map.AUTO_FIX || '').toLowerCase() === 'true'
  };
}

function normalizeMode_(mode) {
  var m = String(mode || 'FULL').trim().toUpperCase();
  if (VALID_MODES.indexOf(m) !== -1) return m;
  return 'FULL';
}

function validateInputRow_(data, mode) {
  mode = normalizeMode_(data.mode || mode || 'FULL');
  var st = String(data.source_text || '').trim();
  var product = String(data.product || '').trim();
  var city = String(data.city || '').trim();
  if (mode === 'FULL' || mode === 'OFFER_ONLY') {
    if (!st) return { valid: false, message: 'source_text обязателен для режима ' + mode };
    if (!product) return { valid: false, message: 'product обязателен' };
    if (!city) return { valid: false, message: 'city обязателен' };
  }
  return { valid: true };
}

function getProcessableRows_(inputSheet, maxRows) {
  var lastRow = inputSheet.getLastRow();
  if (lastRow < 2) return [];
  var values = inputSheet.getRange(2, 1, lastRow, INPUT_COL_PROCESSING).getValues();
  var nowMs = Date.now();
  var dedupMs = LOCK_DEDUP_MINUTES * 60 * 1000;
  var res = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = String(row[INPUT_COL_STATUS - 1] || '').trim();
    if (status !== 'QUEUED' && status !== 'OFFER_OK') continue;
    var processing = row[INPUT_COL_PROCESSING - 1] === true || row[INPUT_COL_PROCESSING - 1] === 'TRUE';
    var lockedAt = row[INPUT_COL_LOCKED_AT - 1];
    if (processing && lockedAt) {
      var lockedMs = lockedAt instanceof Date ? lockedAt.getTime() : (new Date(lockedAt)).getTime();
      if (isFinite(lockedMs) && (nowMs - lockedMs) < dedupMs) continue;
    }
    var obj = {
      id: row[0],
      product: row[1],
      city: row[2],
      source_text: row[3],
      tone: row[4],
      constraints: row[5],
      mode: normalizeMode_(row[6]),
      status: status,
      last_error: row[8],
      updated_at: row[9],
      locked_at: row[10],
      locked_by: row[11],
      processing: processing
    };
    res.push({ rowIndex: i + 2, data: obj });
    if (res.length >= maxRows) break;
  }
  return res;
}

function setStatus_(sheet, rowIndex, status, err) {
  sheet.getRange(rowIndex, INPUT_COL_STATUS).setValue(status);
  sheet.getRange(rowIndex, INPUT_COL_LAST_ERROR).setValue(err || '');
  sheet.getRange(rowIndex, INPUT_COL_UPDATED_AT).setValue(new Date());
  if (status === 'PROCESSING') {
    sheet.getRange(rowIndex, INPUT_COL_LOCKED_AT).setValue(new Date());
    sheet.getRange(rowIndex, INPUT_COL_LOCKED_BY).setValue('batch');
    sheet.getRange(rowIndex, INPUT_COL_PROCESSING).setValue(true);
  } else if (status === 'QUEUED' || status === 'DONE' || status === 'ERROR' || status === 'QA_OK' || status === 'FORMAT_OK') {
    clearProcessingLock_(sheet, rowIndex);
  }
}

function clearProcessingLock_(sheet, rowIndex) {
  sheet.getRange(rowIndex, INPUT_COL_LOCKED_AT).setValue('');
  sheet.getRange(rowIndex, INPUT_COL_LOCKED_BY).setValue('');
  sheet.getRange(rowIndex, INPUT_COL_PROCESSING).setValue(false);
}

function touchUpdatedAt_(sheet, rowIndex) {
  sheet.getRange(rowIndex, INPUT_COL_UPDATED_AT).setValue(new Date());
  sheet.getRange(rowIndex, INPUT_COL_LAST_ERROR).setValue('');
}

function createId_() {
  return Utilities.getUuid();
}

function log_(logSheet, id, step, status, message) {
  if (!logSheet) return;
  logSheet.appendRow([new Date(), id, step, status, message]);
}

function ensureLogHeader_(logSheet) {
  if (!logSheet) return;
  var lastRow = logSheet.getLastRow();
  if (lastRow >= 1) return;
  logSheet.appendRow([
    'timestamp',
    'row_id',
    'step',
    'status_before',
    'status_after',
    'duration_ms',
    'cache_hit',
    'error'
  ]);
}

function logAction_(logSheet, row_id, step, status_before, status_after, duration_ms, cache_hit, error) {
  if (!logSheet) return;
  ensureLogHeader_(logSheet);
  logSheet.appendRow([
    new Date(),
    String(row_id || ''),
    String(step || ''),
    String(status_before != null ? status_before : ''),
    String(status_after != null ? status_after : ''),
    Number(duration_ms) || 0,
    cache_hit === true || cache_hit === 'TRUE',
    String(error != null ? error : '').slice(0, 500)
  ]);
}

function safeJsonParse_(text) {
  var result = parseStrictJson_(text);
  if (result.ok) return result.data;
  throw result.error;
}

/**
 * Строгий парсинг JSON: убирает префиксы/суффиксы, извлекает объект.
 * @return {{ ok: boolean, data: Object | null, error: Error | null, raw: string }}
 */
function parseStrictJson_(text) {
  var raw = String(text || '').trim();
  if (!raw) {
    return { ok: false, data: null, error: new Error('Пустой ответ модели'), raw: '' };
  }
  var cleaned = raw
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```\s*$/i, '')
    .trim();
  var first = cleaned.indexOf('{');
  var last = cleaned.lastIndexOf('}');
  if (first !== -1 && last !== -1 && last > first) {
    cleaned = cleaned.substring(first, last + 1);
  }
  try {
    var data = JSON.parse(cleaned);
    if (data === null) {
      return { ok: false, data: null, error: new Error('Ответ не является JSON-объектом'), raw: raw.slice(0, 500) };
    }
    if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object' && data[0] !== null) {
      data = data[0];
    }
    if (typeof data !== 'object') {
      return { ok: false, data: null, error: new Error('Ответ не является JSON-объектом'), raw: raw.slice(0, 500) };
    }
    return { ok: true, data: data, error: null, raw: raw.slice(0, 500) };
  } catch (e) {
    return {
      ok: false,
      data: null,
      error: new Error('Невалидный JSON: ' + (e.message || e) + '. Фрагмент: ' + cleaned.slice(0, 300)),
      raw: raw.slice(0, 500)
    };
  }
}

/** Сообщение для повторного запроса при невалидном JSON */
var JSON_RETRY_USER_MESSAGE = 'Исправь формат. Верни только валидный JSON без префиксов, суффиксов и пояснений.';

function validateKeys_(obj, keys) {
  keys.forEach(function (k) {
    if (!(k in obj)) {
      throw new Error('В ответе модели нет ключа: ' + k);
    }
  });
}

function ensureArray_(val) {
  if (Array.isArray(val)) return val;
  if (val == null || val === '') return [];
  return [String(val)];
}

function ensureString_(val) {
  return val == null ? '' : String(val);
}

/**
 * QA-валидация текста под лимиты Avito (без ИИ).
 * @param {string} text - текст (заголовок или описание/HTML)
 * @param {Object} options - { limitChars: number, emojiMax: number, capsMaxPercent: number }
 * @returns {{ valid: boolean, reasons: string[] }}
 */
function validateAvitoText_(text, options) {
  var reasons = [];
  var s = String(text || '');
  var limitChars = options && options.limitChars != null ? options.limitChars : 4000;
  var emojiMax = options && options.emojiMax != null ? options.emojiMax : 15;
  var capsMaxPercent = options && options.capsMaxPercent != null ? options.capsMaxPercent : 0.35;

  if (s.length > limitChars) {
    reasons.push('length_' + s.length + '_max_' + limitChars);
  }

  var emojiRanges = /[\u{1F300}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F600}-\u{1F64F}\u{1F1E0}-\u{1F1FF}]/gu;
  var emojiMatches = s.match(emojiRanges);
  var emojiCount = emojiMatches ? emojiMatches.length : 0;
  if (emojiCount > emojiMax) {
    reasons.push('emoji_' + emojiCount + '_max_' + emojiMax);
  }

  var letters = s.replace(/\s/g, '').match(/[A-Za-zА-Яа-яЁё]/g);
  if (letters && letters.length > 0) {
    var upper = s.match(/[A-ZА-ЯЁ]/g);
    var upperCount = upper ? upper.length : 0;
    var ratio = upperCount / letters.length;
    if (ratio > capsMaxPercent) {
      reasons.push('caps_' + Math.round(ratio * 100) + 'pct_max_' + Math.round(capsMaxPercent * 100));
    }
  }

  if (/!{3,}|\?{3,}|\.{4,}/.test(s)) {
    reasons.push('repeated_chars');
  }

  var stripped = s.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  if (stripped.length === 0) {
    reasons.push('empty_blocks');
  }

  var lower = s.toLowerCase();
  var ctaPattern = /звони|пиши|оставь|свяж|напиши|позвони|заявк|обращайся|обратитесь|напишите|позвоните|свяжитесь/;
  if (!ctaPattern.test(lower)) {
    reasons.push('no_cta');
  }

  return {
    valid: reasons.length === 0,
    reasons: reasons
  };
}

/**
 * Нужно ли запускать автоисправление (AUTO_FIX включён и есть замечания).
 * Само исправление вызывается из main через stepQAFix_ (prompts.gs).
 */
function shouldApplyAutoFix_(settings, reasons) {
  return settings.AUTO_FIX && reasons && reasons.length > 0;
}

function writeOut_(outSheet, id, patch) {
  var rowIndex = findOrCreateOutRow_(outSheet, id);
  var headers = getOutHeaders_(outSheet);
  var map = {
    offer_text: 'improved_text',
    title_1: 'title',
    spintax_text: 'spintext',
    avito_html: 'avito_html'
  };
  Object.keys(patch).forEach(function (key) {
    var col = headers[key] || (map[key] ? headers[map[key]] : null);
    if (col) outSheet.getRange(rowIndex, col).setValue(patch[key]);
  });
}

var OUT_HEADERS_NEW = ['id', 'offer_text', 'title_1', 'title_2', 'title_3', 'spintax_text', 'avito_html', 'qa_checks', 'qa_status', 'qa_reasons', 'qa_fixed', 'model_used', 'tokens_est'];

function readOutById_(outSheet, id) {
  var lastRow = outSheet.getLastRow();
  if (lastRow < 2) return {};
  var lastCol = Math.max(outSheet.getLastColumn(), 10);
  var values = outSheet.getRange(2, 1, lastRow, lastCol).getValues();
  var h = getOutHeaders_(outSheet);
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      var row = values[i];
      var offer = (h.offer_text ? row[h.offer_text - 1] : '') || (h.improved_text ? row[h.improved_text - 1] : '') || '';
      var spin = (h.spintax_text ? row[h.spintax_text - 1] : '') || (h.spintext ? row[h.spintext - 1] : '') || '';
      var html = (h.avito_html ? row[h.avito_html - 1] : '') || '';
      var t1 = (h.title_1 ? row[h.title_1 - 1] : '') || (h.title ? row[h.title - 1] : '');
      return {
        offer_text: offer,
        title_1: t1,
        title_2: h.title_2 ? row[h.title_2 - 1] : '',
        title_3: h.title_3 ? row[h.title_3 - 1] : '',
        spintax_text: spin,
        avito_html: html,
        qa_checks: h.qa_checks ? row[h.qa_checks - 1] : '',
        qa_status: h.qa_status ? row[h.qa_status - 1] : '',
        qa_reasons: h.qa_reasons ? row[h.qa_reasons - 1] : '',
        qa_fixed: h.qa_fixed ? row[h.qa_fixed - 1] : '',
        model_used: h.model_used ? row[h.model_used - 1] : '',
        tokens_est: h.tokens_est ? row[h.tokens_est - 1] : '',
        improved_text: offer,
        spintext: spin,
        title: t1,
        bullets: h.bullets ? row[h.bullets - 1] : ''
      };
    }
  }
  return {};
}

function clearOutFromOffer_(outSheet, id) {
  var rowIndex = findOutRowIndex_(outSheet, id);
  if (!rowIndex) return;
  var h = getOutHeaders_(outSheet);
  var cols = ['offer_text', 'title_1', 'title_2', 'title_3', 'spintax_text', 'avito_html', 'qa_checks', 'qa_status', 'qa_reasons', 'qa_fixed', 'improved_text', 'title', 'bullets', 'spintext'];
  for (var c = 0; c < cols.length; c++) {
    if (h[cols[c]]) outSheet.getRange(rowIndex, h[cols[c]]).setValue('');
  }
}

function clearOutFromSpin_(outSheet, id) {
  var rowIndex = findOutRowIndex_(outSheet, id);
  if (!rowIndex) return;
  var h = getOutHeaders_(outSheet);
  ['spintax_text', 'avito_html', 'qa_checks', 'spintext'].forEach(function (col) {
    if (h[col]) outSheet.getRange(rowIndex, h[col]).setValue('');
  });
}

function clearOutFromHtml_(outSheet, id) {
  var rowIndex = findOutRowIndex_(outSheet, id);
  if (!rowIndex) return;
  var h = getOutHeaders_(outSheet);
  ['avito_html', 'qa_checks'].forEach(function (col) {
    if (h[col]) outSheet.getRange(rowIndex, h[col]).setValue('');
  });
}

function findOutRowIndex_(outSheet, id) {
  var lastRow = outSheet.getLastRow();
  if (lastRow < 2) return null;
  var values = outSheet.getRange(2, 1, lastRow, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 2;
  }
  return null;
}

function findOrCreateOutRow_(outSheet, id) {
  var lastRow = outSheet.getLastRow();
  if (lastRow < 1) {
    outSheet.appendRow(OUT_HEADERS_NEW);
  }

  var lr = outSheet.getLastRow();
  if (lr < 2) {
    var empty = [];
    for (var k = 0; k < OUT_HEADERS_NEW.length; k++) empty.push(k === 0 ? id : '');
    outSheet.appendRow(empty);
    return outSheet.getLastRow();
  }

  var values = outSheet.getRange(2, 1, lr, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return i + 2;
    }
  }

  var empty2 = [];
  for (var k2 = 0; k2 < OUT_HEADERS_NEW.length; k2++) empty2.push(k2 === 0 ? id : '');
  outSheet.appendRow(empty2);
  return outSheet.getLastRow();
}

function getOutHeaders_(outSheet) {
  const lastCol = Math.max(9, outSheet.getLastColumn());
  const header = outSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  for (let i = 0; i < header.length; i++) {
    map[String(header[i])] = i + 1;
  }
  return map;
}

function estimateTokens_(text) {
  return Math.ceil(String(text || '').length / 4);
}

// --- CONTROL dashboard ---

var CONTROL_TITLE_ROW = 1;
var CONTROL_STATUS_HEADER_ROW = 3;
var CONTROL_STATUS_ROWS = 10;  // NEW, QUEUED, PROCESSING, OFFER_OK, FORMAT_OK, SPIN_OK, HTML_OK, QA_OK, DONE, ERROR
var CONTROL_PROGRESS_HEADER_ROW = 13;
var CONTROL_PROGRESS_TEXT_ROW = 14;
var CONTROL_ACTIONS_HEADER_ROW = 16;
var CONTROL_ACTIONS_DATA_START = 17;
var CONTROL_ACTIONS_ROWS = 15;
var CONTROL_ERRORS_HEADER_ROW = 34;
var CONTROL_ERRORS_COLS_HEADER_ROW = 35;
var CONTROL_ERRORS_DATA_START = 36;
var CONTROL_ERRORS_ROWS = 20;

function ensureControlSheet_(ss) {
  var sh = ss.getSheetByName('CONTROL');
  if (sh) {
    return sh;
  }
  sh = ss.insertSheet('CONTROL', 0);
  sh.getRange(1, 1, 1, 5).merge().setValue('Avito AI — CONTROL')
    .setFontSize(16).setFontWeight('bold');
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 80);
  sh.setColumnWidth(4, 280);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 120);
  refreshControlDashboard_(ss);
  return sh;
}

function refreshControlDashboard_(ss) {
  var controlSheet = ss.getSheetByName('CONTROL');
  if (!controlSheet) {
    return;
  }
  var inputSheet = ss.getSheetByName('INPUT');
  var logSheet = ss.getSheetByName('LOG');
  var lastInputRow = inputSheet ? inputSheet.getLastRow() : 0;
  var lastLogRow = logSheet ? logSheet.getLastRow() : 0;

  // Block: Статусы (NEW, QUEUED, PROCESSING, OFFER_OK, SPIN_OK, HTML_OK, QA_OK, DONE, ERROR)
  var statusLabels = ['NEW', 'QUEUED', 'PROCESSING', 'OFFER_OK', 'FORMAT_OK', 'SPIN_OK', 'HTML_OK', 'QA_OK', 'DONE', 'ERROR'];
  var counts = {};
  statusLabels.forEach(function (s) { counts[s] = 0; });
  if (inputSheet && lastInputRow >= 2) {
    var statusCol = INPUT_COL_STATUS;
    var statusValues = inputSheet.getRange(2, statusCol, lastInputRow, statusCol).getValues();
    for (var i = 0; i < statusValues.length; i++) {
      var s = String(statusValues[i][0] || '').trim();
      if (counts.hasOwnProperty(s)) counts[s]++;
    }
  }
  controlSheet.getRange(CONTROL_STATUS_HEADER_ROW, 1).setValue('——— Статусы ———').setFontWeight('bold');
  for (var s = 0; s < statusLabels.length; s++) {
    controlSheet.getRange(CONTROL_STATUS_HEADER_ROW + 1 + s, 1).setValue(statusLabels[s]);
    controlSheet.getRange(CONTROL_STATUS_HEADER_ROW + 1 + s, 2).setValue(counts[statusLabels[s]]);
  }

  // Block: Прогресс
  var total = 0;
  for (var k in counts) total += counts[k];
  var inProgress = counts.QUEUED + counts.PROCESSING + counts.OFFER_OK + counts.FORMAT_OK + counts.SPIN_OK + counts.HTML_OK + counts.QA_OK;
  var progressText = 'Всего: ' + total + '  |  Готово: ' + counts.DONE + '  |  Ошибки: ' + counts.ERROR + '  |  В очереди/в работе: ' + inProgress;
  controlSheet.getRange(CONTROL_PROGRESS_HEADER_ROW, 1).setValue('——— Прогресс ———').setFontWeight('bold');
  controlSheet.getRange(CONTROL_PROGRESS_TEXT_ROW, 1).setValue(progressText);

  // Block: Последние действия (из LOG, 8 колонок)
  var logCols = 8;
  controlSheet.getRange(CONTROL_ACTIONS_HEADER_ROW, 1).setValue('——— Последние действия ———').setFontWeight('bold');
  controlSheet.getRange(CONTROL_ACTIONS_HEADER_ROW + 1, 1, CONTROL_ACTIONS_HEADER_ROW + 1, logCols)
    .setValues([['timestamp', 'row_id', 'step', 'status_before', 'status_after', 'duration_ms', 'cache_hit', 'error']]).setFontWeight('bold');
  var actionRows = [];
  if (logSheet && lastLogRow >= 2) {
    var lastCol = logSheet.getLastColumn();
    var cols = Math.min(logCols, Math.max(5, lastCol));
    var startLog = Math.max(2, lastLogRow - CONTROL_ACTIONS_ROWS + 1);
    var logValues = logSheet.getRange(startLog, 1, lastLogRow, cols).getValues();
    for (var j = logValues.length - 1; j >= 0; j--) {
      var r = logValues[j];
      while (r.length < logCols) r.push('');
      actionRows.push(r.slice(0, logCols));
    }
  }
  while (actionRows.length < CONTROL_ACTIONS_ROWS) {
    actionRows.push(['', '', '', '', '', '', '', '']);
  }
  actionRows = actionRows.slice(0, CONTROL_ACTIONS_ROWS);
  if (actionRows.length) {
    controlSheet.getRange(CONTROL_ACTIONS_DATA_START, 1, CONTROL_ACTIONS_DATA_START + actionRows.length - 1, logCols).setValues(actionRows);
  }

  // Block: Последние 20 ошибок (из INPUT, status=ERROR)
  controlSheet.getRange(CONTROL_ERRORS_HEADER_ROW, 1).setValue('——— Последние 20 ошибок ———').setFontWeight('bold');
  controlSheet.getRange(CONTROL_ERRORS_COLS_HEADER_ROW, 1, CONTROL_ERRORS_COLS_HEADER_ROW, 5)
    .setValues([['ID', 'Статус', 'Ошибка', 'Время', 'Ссылка']]).setFontWeight('bold');
  var errors = [];
  if (inputSheet && lastInputRow >= 2) {
    var allInput = inputSheet.getRange(2, 1, lastInputRow, 12).getValues();
    for (var r = 0; r < allInput.length; r++) {
      if (String(allInput[r][INPUT_COL_STATUS - 1] || '') === 'ERROR') {
        errors.push({
          rowIndex: r + 2,
          id: allInput[r][0],
          status: allInput[r][INPUT_COL_STATUS - 1],
          last_error: allInput[r][INPUT_COL_LAST_ERROR - 1],
          updated_at: allInput[r][INPUT_COL_UPDATED_AT - 1]
        });
      }
    }
    errors = errors.slice(-CONTROL_ERRORS_ROWS);
  }
  var sheetUrl = ss.getUrl();
  var inputGid = inputSheet ? inputSheet.getSheetId() : 0;
  for (var e = 0; e < CONTROL_ERRORS_ROWS; e++) {
    var r = controlSheet.getRange(CONTROL_ERRORS_DATA_START + e, 1, CONTROL_ERRORS_DATA_START + e, 5);
    r.clearContent();
    if (e < errors.length) {
      var err = errors[e];
      controlSheet.getRange(CONTROL_ERRORS_DATA_START + e, 1).setValue(err.id);
      controlSheet.getRange(CONTROL_ERRORS_DATA_START + e, 2).setValue(err.status);
      controlSheet.getRange(CONTROL_ERRORS_DATA_START + e, 3).setValue(String(err.last_error || '').slice(0, 200));
      controlSheet.getRange(CONTROL_ERRORS_DATA_START + e, 4).setValue(err.updated_at);
      var linkUrl = sheetUrl + '#gid=' + inputGid + '&range=A' + err.rowIndex;
      controlSheet.getRange(CONTROL_ERRORS_DATA_START + e, 5).setFormula('=HYPERLINK("' + linkUrl + '";"Перейти к строке")');
    }
  }
}

function getStopFlag_(ss) {
  var props = PropertiesService.getDocumentProperties();
  return props.getProperty('AVITO_AI_STOP') === '1';
}

function setStopFlag_(ss, value) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('AVITO_AI_STOP', value ? '1' : '0');
}
