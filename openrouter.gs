function callOpenRouter_(apiKey, model, messages, temperature) {
  if (!apiKey) {
    return { ok: false, error: new Error('Не задан OPENROUTER_API_KEY в SETTINGS') };
  }

  const url = 'https://openrouter.ai/api/v1/chat/completions';
  const payload = {
    model: model,
    messages: messages,
    temperature: temperature,
    response_format: { type: 'json_object' }
  };

  const options = {
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

  const maxAttempts = 4;
  let lastErr = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    const body = resp.getContentText();

    if (code >= 200 && code < 300) {
      const json = JSON.parse(body);
      const content = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content
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
