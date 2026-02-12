function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Avito AI')
    .addItem('Запустить обработку очереди', 'processQueue')
    .addItem('Обработать выделенные строки', 'processSelectedRows')
    .addSeparator()
    .addItem('Показать последние логи', 'showLastLogs')
    .addToUi();
}

function processQueue() {
  const ss = SpreadsheetApp.getActive();
  const settings = loadSettings_(ss);
  const inputSheet = ss.getSheetByName('INPUT');
  const outSheet = ss.getSheetByName('OUT');
  const logSheet = ss.getSheetByName('LOG');

  const rows = getProcessableRows_(inputSheet, settings.MAX_ROWS_PER_RUN);
  if (rows.length === 0) {
    log_(logSheet, '-', 'QUEUE', 'OK', 'Нет строк для обработки');
    return;
  }

  rows.forEach((r) => {
    const { rowIndex, data } = r;
    const id = String(data.id || createId_());
    let currentStatus = String(data.status || 'NEW');

    try {
      if (!data.id) {
        inputSheet.getRange(rowIndex, 1).setValue(id);
      }

      if (currentStatus === 'NEW' || !currentStatus) {
        log_(logSheet, id, 'STEP1', 'START', 'Улучшение текста');
        const improved = stepImprove_(settings, data);
        writeOut_(outSheet, id, {
          improved_text: improved.text,
          title: improved.title,
          bullets: improved.bullets,
          model_used: improved.model,
          tokens_est: estimateTokens_(improved.text)
        });
        setStatus_(inputSheet, rowIndex, 'STEP1_DONE', '');
        currentStatus = 'STEP1_DONE';
        log_(logSheet, id, 'STEP1', 'OK', 'Готово');
      }

      const outAfterStep1 = readOutById_(outSheet, id);
      if (!outAfterStep1.improved_text) {
        throw new Error('Нет improved_text в OUT после STEP1');
      }

      if (currentStatus === 'STEP1_DONE') {
        log_(logSheet, id, 'STEP2', 'START', 'Генерация спин-текста');
        const spin = stepSpin_(settings, data, outAfterStep1.improved_text);
        writeOut_(outSheet, id, {
          spintext: spin.spintext,
          model_used: spin.model,
          tokens_est: estimateTokens_(spin.spintext)
        });
        setStatus_(inputSheet, rowIndex, 'STEP2_DONE', '');
        currentStatus = 'STEP2_DONE';
        log_(logSheet, id, 'STEP2', 'OK', 'Готово');
      }

      const outAfterStep2 = readOutById_(outSheet, id);
      if (!outAfterStep2.spintext) {
        throw new Error('Нет spintext в OUT после STEP2');
      }

      if (currentStatus === 'STEP2_DONE') {
        log_(logSheet, id, 'STEP3', 'START', 'Сборка Avito HTML');
        const avito = stepAvitoHtml_(settings, data, outAfterStep2.spintext);
        writeOut_(outSheet, id, {
          avito_html: avito.html,
          qa_checks: avito.qa_checks,
          model_used: avito.model,
          tokens_est: estimateTokens_(avito.html)
        });
        setStatus_(inputSheet, rowIndex, 'DONE', '');
        currentStatus = 'DONE';
        log_(logSheet, id, 'STEP3', 'OK', 'DONE');
      }

      touchUpdatedAt_(inputSheet, rowIndex);
    } catch (e) {
      setStatus_(inputSheet, rowIndex, 'ERROR', String(e.message || e));
      touchUpdatedAt_(inputSheet, rowIndex);
      log_(logSheet, id, 'ERROR', 'FAIL', String(e.message || e));
    }
  });
}

function processSelectedRows() {
  const ss = SpreadsheetApp.getActive();
  const inputSheet = ss.getSheetByName('INPUT');
  const range = inputSheet.getActiveRange();
  if (!range) {
    return;
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();

  for (let i = 0; i < numRows; i++) {
    const rowIndex = startRow + i;
    if (rowIndex === 1) {
      continue;
    }
    const statusCell = inputSheet.getRange(rowIndex, 7);
    const status = String(statusCell.getValue() || '');
    if (!status || status === 'ERROR') {
      statusCell.setValue('NEW');
    }
  }

  processQueue();
}

function showLastLogs() {
  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName('LOG');
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Логов пока нет');
    return;
  }
  const start = Math.max(2, lastRow - 15);
  const values = logSheet.getRange(start, 1, lastRow - start + 1, 5).getValues();
  const text = values.map((r) => r.join(' | ')).join('\n');
  SpreadsheetApp.getUi().alert(text || 'Логов пока нет');
}
