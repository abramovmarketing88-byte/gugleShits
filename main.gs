function onOpen() {
  var ss = SpreadsheetApp.getActive();
  ensureControlSheet_(ss);
  refreshControlDashboard_(ss);
  SpreadsheetApp.getUi()
    .createMenu('Avito AI')
    .addItem('Сгенерировать (выделенные)', 'menuGenerateSelected')
    .addItem('Сгенерировать (очередь 50)', 'menuGenerateQueue50')
    .addItem('Повторить ошибки', 'menuRetryErrors')
    .addSeparator()
    .addItem('Перегенерировать с шага OFFER', 'menuRegenFromOffer')
    .addItem('Перегенерировать с шага SPIN', 'menuRegenFromSpin')
    .addItem('Перегенерировать с шага HTML', 'menuRegenFromHtml')
    .addSeparator()
    .addItem('Остановить', 'menuStop')
    .addItem('Логи', 'menuLogs')
    .addToUi();
}

function menuGenerateSelected() {
  processSelectedRows();
}

function menuGenerateQueue50() {
  processQueue(null);
}

function menuRetryErrors() {
  retryErrors();
}

function menuRegenFromOffer() {
  regenFromStep_('OFFER');
}

function menuRegenFromSpin() {
  regenFromStep_('SPIN');
}

function menuRegenFromHtml() {
  regenFromStep_('HTML');
}

function menuStop() {
  var ss = SpreadsheetApp.getActive();
  setStopFlag_(ss, true);
  SpreadsheetApp.getActiveSpreadsheet().toast('Запрошена остановка. Текущая строка завершится, затем обработка остановится.', 'Остановить', 4);
}

function menuLogs() {
  showLastLogs();
}

function regenFromStep_(step) {
  var ss = SpreadsheetApp.getActive();
  var inputSheet = ss.getSheetByName('INPUT');
  var outSheet = ss.getSheetByName('OUT');
  var range = inputSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Выделите строки на листе INPUT.', 'Avito AI', 4);
    return;
  }
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var count = 0;
  for (var i = 0; i < numRows; i++) {
    var rowIndex = startRow + i;
    if (rowIndex === 1) continue;
    var id = inputSheet.getRange(rowIndex, INPUT_COL_ID).getValue();
    if (!id) continue;
    if (step === 'OFFER') {
      clearOutFromOffer_(outSheet, id);
      setStatus_(inputSheet, rowIndex, 'QUEUED', '');
      count++;
    } else if (step === 'SPIN') {
      clearOutFromSpin_(outSheet, id);
      setStatus_(inputSheet, rowIndex, 'OFFER_OK', '');
      count++;
    } else if (step === 'HTML') {
      clearOutFromHtml_(outSheet, id);
      setStatus_(inputSheet, rowIndex, 'OFFER_OK', '');
      count++;
    }
  }
  refreshControlDashboard_(ss);
  SpreadsheetApp.getActiveSpreadsheet().toast('Сброшено на шаг ' + step + ': ' + count + ' строк. Запустите «Сгенерировать (очередь 50)».', 'Avito AI', 5);
}

function processQueue(optionalMaxRows) {
  var ss = SpreadsheetApp.getActive();
  setStopFlag_(ss, false);
  var settings = loadSettings_(ss);
  var maxRows = (optionalMaxRows != null && optionalMaxRows > 0)
    ? optionalMaxRows
    : (settings.BATCH_SIZE || settings.MAX_ROWS_PER_RUN || 50);
  var maxRuntimeMs = (settings.MAX_RUNTIME_SECONDS || 300) * 1000;
  var checkpointMarginMs = CHECKPOINT_SECONDS_LEFT * 1000;
  var inputSheet = ss.getSheetByName('INPUT');
  var outSheet = ss.getSheetByName('OUT');
  var logSheet = ss.getSheetByName('LOG');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Другой батч уже выполняется. Подождите или повторите позже.', 'Avito AI', 5);
    return;
  }

  try {
    var rows = getProcessableRows_(inputSheet, maxRows);
    if (rows.length === 0) {
      logAction_(logSheet, '-', 'QUEUE', '', 'OK', 0, false, '');
      refreshControlDashboard_(ss);
      SpreadsheetApp.getActiveSpreadsheet().toast('Нет строк в очереди (QUEUED/OFFER_OK/SPIN_OK) или все заблокированы.', 'Avito AI', 3);
      return;
    }

    logAction_(logSheet, '-', 'BATCH_START', '', '', 0, false, 'rows=' + rows.length);
    SpreadsheetApp.getActiveSpreadsheet().toast('Запуск: ' + rows.length + ' строк.', 'Avito AI', 3);
    var doneCount = 0;
    var errorCount = 0;
    var startTime = Date.now();

    for (var i = 0; i < rows.length; i++) {
      if (Date.now() - startTime >= maxRuntimeMs - checkpointMarginMs) {
        logAction_(logSheet, '-', 'CHECKPOINT', '', '', Date.now() - startTime, false, 'Сохранили прогресс, до лимита < 25 сек');
        refreshControlDashboard_(ss);
        SpreadsheetApp.getActiveSpreadsheet().toast('Чекпоинт: прогресс сохранён. Обработано: ' + doneCount + ', ошибок: ' + errorCount + '. Запустите снова для продолжения.', 'Avito AI', 6);
        return;
      }
      if (getStopFlag_(ss)) {
        logAction_(logSheet, '-', 'QUEUE', 'STOP', '', Date.now() - startTime, false, 'Остановлено пользователем');
        refreshControlDashboard_(ss);
        SpreadsheetApp.getActiveSpreadsheet().toast('Остановлено. Успешно: ' + doneCount + ', ошибок: ' + errorCount, 'Avito AI', 5);
        return;
      }

      var r = rows[i];
      var rowIndex = r.rowIndex;
      var data = r.data;
      var id = String(data.id || createId_());
      var mode = data.mode || 'FULL';
      var statusBefore = String(data.status || 'QUEUED').trim();

      if (!data.id) {
        inputSheet.getRange(rowIndex, INPUT_COL_ID).setValue(id);
      }

      setStatus_(inputSheet, rowIndex, 'PROCESSING', '');
      logAction_(logSheet, id, 'LOCK', statusBefore, 'PROCESSING', 0, false, '');
      refreshControlDashboard_(ss);
      if (i > 0 && i % 5 === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Прогресс: ' + (i + 1) + ' / ' + rows.length, 'Avito AI', 2);
      }

      try {
        runStateMachine_(ss, settings, inputSheet, outSheet, logSheet, rowIndex, id, data, mode);
        touchUpdatedAt_(inputSheet, rowIndex);
        doneCount++;
      } catch (e) {
        var errMsg = String(e.message || e);
        setStatus_(inputSheet, rowIndex, 'ERROR', errMsg);
        touchUpdatedAt_(inputSheet, rowIndex);
        logAction_(logSheet, id, 'ERROR', 'PROCESSING', 'ERROR', 0, false, errMsg);
        errorCount++;
      }
    }

    logAction_(logSheet, '-', 'BATCH_END', '', '', Date.now() - startTime, false, 'done=' + doneCount + ' errors=' + errorCount);
    refreshControlDashboard_(ss);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Готово. Успешно: ' + doneCount + ', ошибок: ' + errorCount,
      errorCount > 0 ? 'Avito AI (есть ошибки)' : 'Avito AI',
      5
    );
  } finally {
    lock.releaseLock();
  }
}

function runStateMachine_(ss, settings, inputSheet, outSheet, logSheet, rowIndex, id, data, mode) {
  var currentStatus = String(data.status || 'QUEUED').trim();
  var t0;

  // FULL: OfferWriter -> AvitoFormatter. QUEUED -> OFFER -> OFFER_OK -> FORMAT -> FORMAT_OK -> QA -> DONE
  if (mode === 'FULL') {
    if (currentStatus === 'QUEUED') {
      t0 = Date.now();
      var offer = stepOfferWriter_(settings, data);
      writeOut_(outSheet, id, {
        offer_text: offer.offer_text,
        title_1: offer.title_1,
        title_2: offer.title_2,
        title_3: offer.title_3,
        model_used: offer.model,
        tokens_est: estimateTokens_(offer.offer_text)
      });
      setStatus_(inputSheet, rowIndex, 'OFFER_OK', '');
      logAction_(logSheet, id, 'OfferWriter', 'PROCESSING', 'OFFER_OK', Date.now() - t0, false, '');
      currentStatus = 'OFFER_OK';
    }
    if (currentStatus === 'OFFER_OK') {
      var out1 = readOutById_(outSheet, id);
      var baseText = out1.offer_text || out1.improved_text;
      if (!baseText) throw new Error('Нет offer_text в OUT');
      t0 = Date.now();
      var fmt = stepAvitoFormatter_(settings, data, baseText);
      writeOut_(outSheet, id, {
        spintax_text: fmt.spintax_text,
        avito_html: fmt.avito_html,
        model_used: fmt.model,
        tokens_est: estimateTokens_(fmt.avito_html)
      });
      setStatus_(inputSheet, rowIndex, 'FORMAT_OK', '');
      logAction_(logSheet, id, 'AvitoFormatter', 'OFFER_OK', 'FORMAT_OK', Date.now() - t0, false, '');
      currentStatus = 'FORMAT_OK';
    }
    if (currentStatus === 'FORMAT_OK') {
      runQALayer_(ss, settings, outSheet, id, logSheet);
      setStatus_(inputSheet, rowIndex, 'QA_OK', '');
      setStatus_(inputSheet, rowIndex, 'DONE', '');
      logAction_(logSheet, id, 'QA', 'FORMAT_OK', 'DONE', 0, false, '');
    }
    return;
  }

  // OFFER_ONLY: только OfferWriter
  if (mode === 'OFFER_ONLY') {
    if (currentStatus === 'QUEUED') {
      t0 = Date.now();
      var offer2 = stepOfferWriter_(settings, data);
      writeOut_(outSheet, id, {
        offer_text: offer2.offer_text,
        title_1: offer2.title_1,
        title_2: offer2.title_2,
        title_3: offer2.title_3,
        model_used: offer2.model,
        tokens_est: estimateTokens_(offer2.offer_text)
      });
      setStatus_(inputSheet, rowIndex, 'OFFER_OK', '');
      logAction_(logSheet, id, 'OfferWriter', 'PROCESSING', 'OFFER_OK', Date.now() - t0, false, '');
      currentStatus = 'OFFER_OK';
    }
    if (currentStatus === 'OFFER_OK') {
      runQALayer_(ss, settings, outSheet, id, logSheet);
      setStatus_(inputSheet, rowIndex, 'QA_OK', '');
      setStatus_(inputSheet, rowIndex, 'DONE', '');
      logAction_(logSheet, id, 'QA', 'OFFER_OK', 'DONE', 0, false, '');
    }
    return;
  }

  // SPIN_ONLY: только AvitoFormatter (база = offer_text / improved_text)
  if (mode === 'SPIN_ONLY') {
    if (currentStatus === 'QUEUED') {
      var outSpin = readOutById_(outSheet, id);
      var baseSpin = outSpin.offer_text || outSpin.improved_text;
      if (!baseSpin) throw new Error('Для SPIN_ONLY нужен offer_text в OUT. Сначала выполните OFFER.');
      t0 = Date.now();
      var fmtSpin = stepAvitoFormatter_(settings, data, baseSpin);
      writeOut_(outSheet, id, {
        spintax_text: fmtSpin.spintax_text,
        avito_html: fmtSpin.avito_html,
        model_used: fmtSpin.model,
        tokens_est: estimateTokens_(fmtSpin.avito_html)
      });
      setStatus_(inputSheet, rowIndex, 'FORMAT_OK', '');
      logAction_(logSheet, id, 'AvitoFormatter', 'PROCESSING', 'FORMAT_OK', Date.now() - t0, false, '');
      currentStatus = 'FORMAT_OK';
    }
    if (currentStatus === 'FORMAT_OK') {
      runQALayer_(ss, settings, outSheet, id, logSheet);
      setStatus_(inputSheet, rowIndex, 'QA_OK', '');
      setStatus_(inputSheet, rowIndex, 'DONE', '');
      logAction_(logSheet, id, 'QA', 'FORMAT_OK', 'DONE', 0, false, '');
    }
    return;
  }

  // HTML_ONLY: только AvitoFormatter (база = offer_text / improved_text)
  if (mode === 'HTML_ONLY') {
    if (currentStatus === 'QUEUED') {
      var outHtml = readOutById_(outSheet, id);
      var baseHtml = outHtml.offer_text || outHtml.improved_text;
      if (!baseHtml) throw new Error('Для HTML_ONLY нужен offer_text в OUT. Сначала выполните OFFER.');
      t0 = Date.now();
      var fmtHtml = stepAvitoFormatter_(settings, data, baseHtml);
      writeOut_(outSheet, id, {
        spintax_text: fmtHtml.spintax_text,
        avito_html: fmtHtml.avito_html,
        model_used: fmtHtml.model,
        tokens_est: estimateTokens_(fmtHtml.avito_html)
      });
      setStatus_(inputSheet, rowIndex, 'FORMAT_OK', '');
      logAction_(logSheet, id, 'AvitoFormatter', 'PROCESSING', 'FORMAT_OK', Date.now() - t0, false, '');
      currentStatus = 'FORMAT_OK';
    }
    if (currentStatus === 'FORMAT_OK') {
      runQALayer_(ss, settings, outSheet, id, logSheet);
      setStatus_(inputSheet, rowIndex, 'QA_OK', '');
      setStatus_(inputSheet, rowIndex, 'DONE', '');
      logAction_(logSheet, id, 'QA', 'FORMAT_OK', 'DONE', 0, false, '');
    }
    return;
  }

  throw new Error('Неизвестный режим: ' + mode);
}

/**
 * QA-слой: валидация заголовка и описания, при AUTO_FIX — автоисправление через модель.
 * Пишет в OUT: qa_status (OK | QA_ERR), qa_reasons, qa_fixed (true/false).
 */
function runQALayer_(ss, settings, outSheet, id, logSheet) {
  var out = readOutById_(outSheet, id);
  var title = out.title_1 || out.title || '';
  var desc = out.avito_html || out.improved_text || '';
  var optsTitle = {
    limitChars: settings.LIMIT_TITLE_CHARS || 60,
    emojiMax: settings.EMOJI_MAX || 15,
    capsMaxPercent: settings.CAPS_MAX_PERCENT || 0.35
  };
  var optsDesc = {
    limitChars: settings.LIMIT_DESC_CHARS || 4000,
    emojiMax: settings.EMOJI_MAX || 15,
    capsMaxPercent: settings.CAPS_MAX_PERCENT || 0.35
  };
  var rTitle = validateAvitoText_(title, optsTitle);
  var rDesc = validateAvitoText_(desc, optsDesc);
  var allReasons = [];
  rTitle.reasons.forEach(function (r) { allReasons.push('title:' + r); });
  rDesc.reasons.forEach(function (r) { allReasons.push('desc:' + r); });

  var qaFixed = false;
  if (shouldApplyAutoFix_(settings, allReasons)) {
    var fixedTitle = stepQAFix_(settings, title, rTitle.reasons, 'title', { limitChars: optsTitle.limitChars });
    if (fixedTitle) {
      writeOut_(outSheet, id, { title_1: fixedTitle });
      qaFixed = true;
    }
    var fixedDesc = stepQAFix_(settings, desc, rDesc.reasons, 'desc', { limitChars: optsDesc.limitChars });
    if (fixedDesc) {
      writeOut_(outSheet, id, { avito_html: fixedDesc });
      qaFixed = true;
    }
  }

  var qaStatus = allReasons.length === 0 ? 'OK' : 'QA_ERR';
  var qaReasons = allReasons.join('; ');
  writeOut_(outSheet, id, { qa_status: qaStatus, qa_reasons: qaReasons, qa_fixed: qaFixed });
}

/**
 * Автоисправление через модель по замечаниям QA (если включено AUTO_FIX).
 * @returns {string|null} исправленный текст или null
 */
function applyAutoFixIfEnabled_(settings, text, reasons, context, limits) {
  if (!shouldApplyAutoFix_(settings, reasons)) return null;
  return stepQAFix_(settings, text, reasons, context || 'text', limits);
}

function processSelectedRows() {
  var ss = SpreadsheetApp.getActive();
  var inputSheet = ss.getSheetByName('INPUT');
  var range = inputSheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Выделите диапазон на листе INPUT.', 'Avito AI', 4);
    return;
  }

  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var validCount = 0;
  var invalidCount = 0;

  for (var i = 0; i < numRows; i++) {
    var rowIndex = startRow + i;
    if (rowIndex === 1) continue;
    var row = inputSheet.getRange(rowIndex, 1, rowIndex, 12).getValues()[0];
    var data = {
      id: row[0],
      product: row[1],
      city: row[2],
      source_text: row[3],
      tone: row[4],
      constraints: row[5],
      mode: normalizeMode_(row[6]),
      status: row[7],
      last_error: row[8],
      updated_at: row[9],
      locked_at: row[10],
      locked_by: row[11]
    };
    var v = validateInputRow_(data);
    if (!v.valid) {
      invalidCount++;
      continue;
    }
    inputSheet.getRange(rowIndex, INPUT_COL_STATUS).setValue('QUEUED');
    inputSheet.getRange(rowIndex, INPUT_COL_LAST_ERROR).setValue('');
    inputSheet.getRange(rowIndex, INPUT_COL_UPDATED_AT).setValue(new Date());
    inputSheet.getRange(rowIndex, INPUT_COL_LOCKED_AT).setValue('');
    inputSheet.getRange(rowIndex, INPUT_COL_LOCKED_BY).setValue('');
    inputSheet.getRange(rowIndex, INPUT_COL_PROCESSING).setValue(false);
    validCount++;
  }

  refreshControlDashboard_(ss);
  if (validCount === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      invalidCount > 0 ? 'Нет валидных строк. Заполните product, city, source_text. Невалидных: ' + invalidCount : 'Выделите строки с данными.',
      'Avito AI',
      5
    );
    return;
  }
  if (invalidCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('В очередь добавлено: ' + validCount + ', пропущено (невалидно): ' + invalidCount, 'Avito AI', 4);
  }
  var settings = loadSettings_(ss);
  processQueue(settings.BATCH_SIZE || settings.MAX_ROWS_PER_RUN);
}

function retryErrors() {
  var ss = SpreadsheetApp.getActive();
  var inputSheet = ss.getSheetByName('INPUT');
  var lastRow = inputSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Нет данных на INPUT.', 'Avito AI', 3);
    return;
  }
  var count = 0;
  for (var rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    var status = String(inputSheet.getRange(rowIndex, INPUT_COL_STATUS).getValue() || '').trim();
    if (status === 'ERROR') {
      inputSheet.getRange(rowIndex, INPUT_COL_STATUS).setValue('QUEUED');
      inputSheet.getRange(rowIndex, INPUT_COL_LAST_ERROR).setValue('');
      inputSheet.getRange(rowIndex, INPUT_COL_UPDATED_AT).setValue(new Date());
      inputSheet.getRange(rowIndex, INPUT_COL_LOCKED_AT).setValue('');
      inputSheet.getRange(rowIndex, INPUT_COL_LOCKED_BY).setValue('');
      inputSheet.getRange(rowIndex, INPUT_COL_PROCESSING).setValue(false);
      count++;
    }
  }
  if (count === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Нет строк со статусом ERROR.', 'Avito AI', 3);
    refreshControlDashboard_(ss);
    return;
  }
  var settings = loadSettings_(ss);
  SpreadsheetApp.getActiveSpreadsheet().toast('Сброшено в очередь: ' + count + ' строк. Запуск обработки.', 'Avito AI', 4);
  processQueue(settings.BATCH_SIZE || 50);
}

function showLastLogs() {
  var ss = SpreadsheetApp.getActive();
  var logSheet = ss.getSheetByName('LOG');
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Логов пока нет');
    return;
  }
  var start = Math.max(2, lastRow - 14);
  var logCols = Math.min(8, logSheet.getLastColumn());
  var values = logSheet.getRange(start, 1, lastRow, logCols).getValues();
  var text = values.map(function (row) { return row.join(' | '); }).join('\n');
  SpreadsheetApp.getUi().alert(text || 'Логов пока нет');
}
