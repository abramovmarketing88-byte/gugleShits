function loadSettings_(ss) {
  const sh = ss.getSheetByName('SETTINGS');
  const values = sh.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const k = String(values[i][0] || '').trim();
    const v = String(values[i][1] || '').trim();
    if (k) {
      map[k] = v;
    }
  }

  return {
    OPENROUTER_API_KEY: map.OPENROUTER_API_KEY,
    MODEL_IMPROVE: map.MODEL_IMPROVE || 'openai/gpt-4o-mini',
    MODEL_SPIN: map.MODEL_SPIN || 'openai/gpt-4o-mini',
    MODEL_AVITO: map.MODEL_AVITO || 'openai/gpt-4o-mini',
    MAX_ROWS_PER_RUN: Number(map.MAX_ROWS_PER_RUN || 10),
    TEMPERATURE_IMPROVE: Number(map.TEMPERATURE_IMPROVE || 0.4),
    TEMPERATURE_SPIN: Number(map.TEMPERATURE_SPIN || 0.6),
    TEMPERATURE_AVITO: Number(map.TEMPERATURE_AVITO || 0.5),
    AVITO_STYLE: map.AVITO_STYLE || 'bold_emojis_safe',
    LANGUAGE: map.LANGUAGE || 'ru'
  };
}

function getProcessableRows_(inputSheet, maxRows) {
  const lastRow = inputSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = inputSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const res = [];

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2;
    const row = values[i];
    const obj = {
      id: row[0],
      product: row[1],
      city: row[2],
      source_text: row[3],
      tone: row[4],
      constraints: row[5],
      status: row[6],
      last_error: row[7],
      updated_at: row[8]
    };

    const status = String(obj.status || '');
    if (status === 'NEW' || status === 'STEP1_DONE' || status === 'STEP2_DONE') {
      res.push({ rowIndex: rowIndex, data: obj });
      if (res.length >= maxRows) {
        break;
      }
    }
  }
  return res;
}

function setStatus_(sheet, rowIndex, status, err) {
  sheet.getRange(rowIndex, 7).setValue(status);
  sheet.getRange(rowIndex, 8).setValue(err || '');
}

function touchUpdatedAt_(sheet, rowIndex) {
  sheet.getRange(rowIndex, 9).setValue(new Date());
}

function createId_() {
  return Utilities.getUuid();
}

function log_(logSheet, id, step, status, message) {
  logSheet.appendRow([new Date(), id, step, status, message]);
}

function safeJsonParse_(text) {
  const cleaned = String(text || '')
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```$/i, '')
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    throw new Error('Не смог распарсить JSON. Ответ модели: ' + cleaned.slice(0, 400));
  }
}

function validateKeys_(obj, keys) {
  keys.forEach((k) => {
    if (!(k in obj)) {
      throw new Error('В ответе модели нет ключа: ' + k);
    }
  });
}

function writeOut_(outSheet, id, patch) {
  const rowIndex = findOrCreateOutRow_(outSheet, id);
  const headers = getOutHeaders_(outSheet);

  Object.keys(patch).forEach((key) => {
    if (!headers[key]) {
      return;
    }
    outSheet.getRange(rowIndex, headers[key]).setValue(patch[key]);
  });
}

function readOutById_(outSheet, id) {
  const lastRow = outSheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  const values = outSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return {
        improved_text: values[i][1],
        spintext: values[i][2],
        avito_html: values[i][3],
        title: values[i][4],
        bullets: values[i][5],
        qa_checks: values[i][6],
        model_used: values[i][7],
        tokens_est: values[i][8]
      };
    }
  }
  return {};
}

function findOrCreateOutRow_(outSheet, id) {
  const lastRow = outSheet.getLastRow();
  if (lastRow < 1) {
    outSheet.appendRow([
      'id',
      'improved_text',
      'spintext',
      'avito_html',
      'title',
      'bullets',
      'qa_checks',
      'model_used',
      'tokens_est'
    ]);
  }

  const lr = outSheet.getLastRow();
  if (lr < 2) {
    outSheet.appendRow([id, '', '', '', '', '', '', '', '']);
    return outSheet.getLastRow();
  }

  const values = outSheet.getRange(2, 1, lr - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      return i + 2;
    }
  }

  outSheet.appendRow([id, '', '', '', '', '', '', '', '']);
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
