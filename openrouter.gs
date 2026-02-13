/**
 * Вызов OpenRouter API с ретраями, backoff и опциональным fallback на другую модель.
 * @param {string} apiKey
 * @param {string} model
 * @param {Array} messages
 * @param {number} temperature
 * @param {boolean} useJsonFormat - ответ как JSON (default true)
 * @param {string} [fallbackModel] - при 429 / rate limit / overloaded переключиться на эту модель
 * @returns {{ ok: boolean, content?: string, error?: Error, raw?: object, model_used: string, attempt: number, latency_ms: number, error_code?: number|string }}
 */
function callOpenRouter_(apiKey, model, messages, temperature, useJsonFormat, fallbackModel) {
  if (!apiKey) {
    return {
      ok: false,
      error: new Error('Не задан OPENROUTER_API_KEY в SETTINGS'),
      model_used: model || '',
      attempt: 0,
      latency_ms: 0,
      error_code: 'no_api_key'
    };
  }
  if (useJsonFormat === undefined) useJsonFormat = true;

  var url = 'https://openrouter.ai/api/v1/chat/completions';
  var payload = {
    model: model,
    messages: messages,
    temperature: temperature
  };
  if (useJsonFormat) {
    payload.response_format = { type: 'json_object' };
  }

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://docs.google.com',
      'X-Title': 'Avito Sheets AI Pipeline'
    }
  };

  var BACKOFF_429_MS = [2000, 5000, 10000];
  var BACKOFF_5XX_MS = [1000, 3000];
  var MAX_RETRIES_429 = 3;
  var MAX_RETRIES_5XX = 2;

  function doOneRequest(currentModel) {
    var p = { model: currentModel, messages: messages, temperature: temperature };
    if (useJsonFormat) p.response_format = { type: 'json_object' };
    options.payload = JSON.stringify(p);

    var t0 = Date.now();
    var code;
    var body = '';
    var content = '';
    var raw = null;
    var err = null;
    var errorCode = null;

    try {
      var resp = UrlFetchApp.fetch(url, options);
      code = resp.getResponseCode();
      body = resp.getContentText();
      var latency = Date.now() - t0;

      if (code >= 200 && code < 300) {
        raw = JSON.parse(body);
        content = raw.choices && raw.choices[0] && raw.choices[0].message && raw.choices[0].message.content
          ? raw.choices[0].message.content
          : '';
        return { ok: true, content: content, raw: raw, model_used: currentModel, latency_ms: latency, error_code: null };
      }

      err = new Error('OpenRouter ' + code + ': ' + body.slice(0, 400));
      errorCode = code;
      return { ok: false, error: err, model_used: currentModel, latency_ms: latency, error_code: errorCode, body: body };
    } catch (e) {
      var latency = Date.now() - t0;
      err = e && e.message ? e : new Error(String(e));
      errorCode = 'timeout';
      if (e && e.message && /ECONNRESET|ETIMEDOUT|timeout/i.test(e.message)) {
        errorCode = 'timeout';
      }
      return { ok: false, error: err, model_used: currentModel, latency_ms: latency, error_code: errorCode, body: body };
    }
  }

  function isRateLimitOrOverloaded(res) {
    if (res.error_code === 429) return true;
    var b = String(res.body || '').toLowerCase();
    return /rate limit|overloaded|too many requests/i.test(b);
  }

  var totalAttempts = 0;
  var retries429 = 0;
  var retries5xx = 0;
  var currentModel = model;
  var lastResult = null;

  while (true) {
    totalAttempts++;
    lastResult = doOneRequest(currentModel);

    if (lastResult.ok) {
      lastResult.attempt = totalAttempts;
      return lastResult;
    }

    var code = lastResult.error_code;
    var body = lastResult.body || '';

    if (code === 429) {
      retries429++;
      if (retries429 > MAX_RETRIES_429) {
        if (fallbackModel && fallbackModel !== currentModel) {
          currentModel = fallbackModel;
          retries429 = 0;
          retries5xx = 0;
          continue;
        }
        lastResult.attempt = totalAttempts;
        return lastResult;
      }
      Utilities.sleep(BACKOFF_429_MS[retries429 - 1]);
      continue;
    }

    if ((code >= 500 && code <= 599) || code === 'timeout') {
      retries5xx++;
      if (retries5xx > MAX_RETRIES_5XX) {
        if (fallbackModel && fallbackModel !== currentModel && isRateLimitOrOverloaded(lastResult)) {
          currentModel = fallbackModel;
          retries429 = 0;
          retries5xx = 0;
          continue;
        }
        lastResult.attempt = totalAttempts;
        return lastResult;
      }
      Utilities.sleep(BACKOFF_5XX_MS[Math.min(retries5xx - 1, BACKOFF_5XX_MS.length - 1)]);
      continue;
    }

    if (fallbackModel && fallbackModel !== currentModel && isRateLimitOrOverloaded(lastResult)) {
      currentModel = fallbackModel;
      retries429 = 0;
      retries5xx = 0;
      continue;
    }

    lastResult.attempt = totalAttempts;
    return lastResult;
  }
}
