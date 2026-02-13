function callOpenRouter_(apiKey, model, messages, temperature, useJsonFormat) {
  if (!apiKey) {
    return { ok: false, error: new Error('Не задан OPENROUTER_API_KEY в SETTINGS') };
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

  var maxAttempts = 4;
  var lastErr = null;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    var body = resp.getContentText();

    if (code >= 200 && code < 300) {
      var json = JSON.parse(body);
      var content = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content
        ? json.choices[0].message.content
        : '';
      return { ok: true, content: content, raw: json };
    }

    lastErr = new Error('OpenRouter error ' + code + ': ' + body.slice(0, 500));
    if (code === 429 || (code >= 500 && code <= 599)) {
      Utilities.sleep(700 * attempt);
      continue;
    }
    break;
  }

  return { ok: false, error: lastErr || new Error('OpenRouter request failed') };
}
