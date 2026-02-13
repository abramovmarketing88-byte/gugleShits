// --- OfferWriter: строгий JSON — title_variants, offer_text, bullets, guarantees, cta, warnings ---

var OFFER_WRITER_REQUIRED_KEYS = ['title_variants', 'offer_text', 'bullets', 'guarantees', 'cta', 'warnings'];

function buildOfferWriterPrompt_(data, version) {
  version = version || '1';
  var system = {
    role: 'system',
    content: 'Ты OfferWriter. Создаёшь продающий текст объявления и варианты заголовков. Факты не выдумывай. Без запрещённых обещаний (100%, гарантия результата). Ответ — только один валидный JSON-объект, без markdown и пояснений.'
  };
  var user = {
    role: 'user',
    content: JSON.stringify({
      product: String(data.product || ''),
      city: String(data.city || ''),
      tone: String(data.tone || 'нейтральный'),
      constraints: String(data.constraints || ''),
      source_text: String(data.source_text || ''),
      output_format: {
        title_variants: 'массив из 3 строк — заголовки до 60 символов',
        offer_text: 'строка — продающий текст объявления',
        bullets: 'массив строк — УТП/преимущества',
        guarantees: 'массив строк — гарантии и доверие',
        cta: 'строка — призыв к действию',
        warnings: 'массив строк — замечания по тексту (или пустой массив)'
      }
    })
  };
  return [system, user];
}

function validateOfferWriterResponse_(data) {
  validateKeys_(data, OFFER_WRITER_REQUIRED_KEYS);
  var titles = ensureArray_(data.title_variants);
  if (titles.length < 1) {
    throw new Error('В ответе OfferWriter title_variants должен содержать минимум 1 элемент');
  }
}

function stepOfferWriter_(settings, data) {
  var model = settings.MODEL_OFFER || settings.MODEL_FALLBACK;
  var ttlHours = settings.cache_ttl_hours != null ? settings.cache_ttl_hours : 72;
  var normalizedInput = JSON.stringify({
    product: String(data.product || '').trim(),
    city: String(data.city || '').trim(),
    tone: String(data.tone || '').trim(),
    constraints: String(data.constraints || '').trim(),
    source_text: String(data.source_text || '').trim()
  });
  var cacheKey = buildCacheKey_('OfferWriter', settings.PROMPT_VERSION_OFFER, model, normalizedInput);
  var cached = getCache_(cacheKey, ttlHours);
  if (cached) {
    var parsed = parseStrictJson_(cached);
    if (parsed.ok) {
      var p = parsed.data;
      try {
        validateOfferWriterResponse_(p);
        var titles = ensureArray_(p.title_variants);
        return {
          offer_text: ensureString_(p.offer_text),
          title_1: ensureString_(titles[0]).slice(0, 120),
          title_2: ensureString_(titles[1]).slice(0, 120),
          title_3: ensureString_(titles[2]).slice(0, 120),
          bullets: ensureArray_(p.bullets).join(' • '),
          guarantees: ensureArray_(p.guarantees),
          cta: ensureString_(p.cta),
          warnings: ensureArray_(p.warnings),
          model: model,
          cache_hit: true
        };
      } catch (e) {}
    }
  }

  var temp = settings.TEMPERATURE_OFFER != null ? settings.TEMPERATURE_OFFER : 0.4;
  var messages = buildOfferWriterPrompt_(data, settings.PROMPT_VERSION_OFFER);
  var res = callOpenRouter_(settings.OPENROUTER_API_KEY, model, messages, temp, true, settings.MODEL_FALLBACK);
  if (!res.ok) {
    var meta = ' [model_used=' + (res.model_used || model) + ' attempt=' + (res.attempt || 0) + ' latency_ms=' + (res.latency_ms || 0) + ' error_code=' + (res.error_code || '') + ']';
    throw new Error((res.error && res.error.message) || 'API error' + meta);
  }
  model = res.model_used || model;
  setCache_(cacheKey, res.content, ttlHours);

  var parsed = parseStrictJson_(res.content);
  if (!parsed.ok) {
    messages.push({ role: 'user', content: JSON_RETRY_USER_MESSAGE });
    res = callOpenRouter_(settings.OPENROUTER_API_KEY, model, messages, temp, true, settings.MODEL_FALLBACK);
    if (!res.ok) {
      meta = ' [model_used=' + (res.model_used || model) + ' attempt=' + (res.attempt || 0) + ' latency_ms=' + (res.latency_ms || 0) + ' error_code=' + (res.error_code || '') + ']';
      throw new Error((res.error && res.error.message) || 'API error' + meta);
    }
    model = res.model_used || model;
    parsed = parseStrictJson_(res.content);
  }
  if (!parsed.ok) throw parsed.error;

  var p = parsed.data;
  validateOfferWriterResponse_(p);

  var titles = ensureArray_(p.title_variants);
  return {
    offer_text: ensureString_(p.offer_text),
    title_1: ensureString_(titles[0]).slice(0, 120),
    title_2: ensureString_(titles[1]).slice(0, 120),
    title_3: ensureString_(titles[2]).slice(0, 120),
    bullets: ensureArray_(p.bullets).join(' • '),
    guarantees: ensureArray_(p.guarantees),
    cta: ensureString_(p.cta),
    warnings: ensureArray_(p.warnings),
    model: model,
    cache_hit: false
  };
}

// --- AvitoFormatter: строгий JSON — spintax_text, avito_html, warnings ---

var AVITO_FORMATTER_REQUIRED_KEYS = ['spintax_text', 'avito_html', 'warnings'];

function buildAvitoFormatterPrompt_(settings, data, baseText, version) {
  version = version || '1';
  var system = {
    role: 'system',
    content: 'Ты AvitoFormatter. Делаешь уникализацию (spintax в формате {вариант1|вариант2}) и готовый HTML для Avito. Смысл сохраняй. Только <p>, <b>, эмодзи умеренно. Ответ — только один валидный JSON-объект, без markdown и пояснений.'
  };
  var user = {
    role: 'user',
    content: JSON.stringify({
      product: String(data.product || ''),
      city: String(data.city || ''),
      base_text: String(baseText || ''),
      style: settings.AVITO_STYLE || 'bold_emojis_safe',
      output_format: {
        spintax_text: 'строка — текст с spintax',
        avito_html: 'строка — HTML для вставки в Avito',
        warnings: 'массив строк — замечания (или пустой массив)'
      }
    })
  };
  return [system, user];
}

function validateAvitoFormatterResponse_(data) {
  validateKeys_(data, AVITO_FORMATTER_REQUIRED_KEYS);
}

function stepAvitoFormatter_(settings, data, baseText) {
  var model = settings.MODEL_FORMAT || settings.MODEL_FALLBACK;
  var ttlHours = settings.cache_ttl_hours != null ? settings.cache_ttl_hours : 72;
  var normalizedInput = JSON.stringify({
    product: String(data.product || '').trim(),
    city: String(data.city || '').trim(),
    base_text: String(baseText || '').trim(),
    style: String(settings.AVITO_STYLE || '').trim()
  });
  var cacheKey = buildCacheKey_('AvitoFormatter', settings.PROMPT_VERSION_FORMAT, model, normalizedInput);
  var cached = getCache_(cacheKey, ttlHours);
  if (cached) {
    var parsed = parseStrictJson_(cached);
    if (parsed.ok) {
      var p = parsed.data;
      try {
        validateAvitoFormatterResponse_(p);
        return {
          spintax_text: ensureString_(p.spintax_text),
          avito_html: ensureString_(p.avito_html),
          warnings: ensureArray_(p.warnings),
          model: model,
          cache_hit: true
        };
      } catch (e) {}
    }
  }

  var temp = settings.TEMPERATURE_FORMAT != null ? settings.TEMPERATURE_FORMAT : 0.5;
  var messages = buildAvitoFormatterPrompt_(settings, data, baseText, settings.PROMPT_VERSION_FORMAT);
  var res = callOpenRouter_(settings.OPENROUTER_API_KEY, model, messages, temp, true, settings.MODEL_FALLBACK);
  if (!res.ok) {
    var metaFmt = ' [model_used=' + (res.model_used || model) + ' attempt=' + (res.attempt || 0) + ' latency_ms=' + (res.latency_ms || 0) + ' error_code=' + (res.error_code || '') + ']';
    throw new Error((res.error && res.error.message) || 'API error' + metaFmt);
  }
  model = res.model_used || model;
  setCache_(cacheKey, res.content, ttlHours);

  var parsed = parseStrictJson_(res.content);
  if (!parsed.ok) {
    messages.push({ role: 'user', content: JSON_RETRY_USER_MESSAGE });
    res = callOpenRouter_(settings.OPENROUTER_API_KEY, model, messages, temp, true, settings.MODEL_FALLBACK);
    if (!res.ok) {
      metaFmt = ' [model_used=' + (res.model_used || model) + ' attempt=' + (res.attempt || 0) + ' latency_ms=' + (res.latency_ms || 0) + ' error_code=' + (res.error_code || '') + ']';
      throw new Error((res.error && res.error.message) || 'API error' + metaFmt);
    }
    model = res.model_used || model;
    parsed = parseStrictJson_(res.content);
  }
  if (!parsed.ok) throw parsed.error;

  var p = parsed.data;
  validateAvitoFormatterResponse_(p);

  return {
    spintax_text: ensureString_(p.spintax_text),
    avito_html: ensureString_(p.avito_html),
    warnings: ensureArray_(p.warnings),
    model: model,
    cache_hit: false
  };
}

// --- Legacy steps (совместимость) ---

function stepImprove_(settings, data) {
  var r = stepOfferWriter_(settings, data);
  return {
    title: r.title_1,
    text: r.offer_text,
    bullets: r.bullets || '',
    model: r.model
  };
}

function stepSpin_(settings, data, improvedText) {
  var r = stepAvitoFormatter_(settings, data, improvedText);
  return { spintext: r.spintax_text, model: r.model };
}

function stepAvitoHtml_(settings, data, spintext) {
  var r = stepAvitoFormatter_(settings, data, spintext);
  return {
    html: r.avito_html,
    qa_checks: (r.warnings || []).join(', '),
    model: r.model
  };
}

function buildImprovePrompt_(data) {
  return buildOfferWriterPrompt_(data, '1');
}

function buildSpinPrompt_(data, improvedText) {
  return buildAvitoFormatterPrompt_({ AVITO_STYLE: 'bold_emojis_safe' }, data, improvedText, '1');
}

function buildAvitoHtmlPrompt_(settings, data, spintext) {
  return buildAvitoFormatterPrompt_(settings, data, spintext, '1');
}

// --- QA автоисправление (AUTO_FIX): один запрос — вернуть только исправленный текст ---

function buildQAFixPrompt_(text, reasons, context, limits) {
  var limitStr = limits && limits.limitChars != null ? ' Лимит символов: ' + limits.limitChars + '.' : '';
  var system = {
    role: 'system',
    content: 'Ты редактор объявлений Avito. Исправляй текст по замечаниям, укладывайся в лимиты, сохраняй смысл. Ответ — только исправленный текст, без пояснений и markdown.'
  };
  var user = {
    role: 'user',
    content: 'Контекст: ' + (context === 'title' ? 'заголовок объявления' : 'описание объявления (HTML допускается).') + limitStr + '\n\nЗамечания QA: ' + (reasons.join(', ')) + '\n\nИсходный текст:\n' + String(text || '')
  };
  return [system, user];
}

/**
 * Один запрос к модели: исправить текст по замечаниям QA.
 * @param {Object} settings
 * @param {string} text - исходный текст
 * @param {string[]} reasons - список причин (из validateAvitoText_)
 * @param {string} context - 'title' или 'desc'
 * @param {Object} limits - { limitChars: number }
 * @returns {string|null} исправленный текст или null при ошибке
 */
function stepQAFix_(settings, text, reasons, context, limits) {
  if (!text || !reasons || reasons.length === 0) return null;
  var model = settings.MODEL_FORMAT || settings.MODEL_FALLBACK;
  var temp = 0.3;
  var messages = buildQAFixPrompt_(text, reasons, context || 'desc', limits);
  var res = callOpenRouter_(settings.OPENROUTER_API_KEY, model, messages, temp, false, settings.MODEL_FALLBACK);
  if (!res.ok) return null;
  var content = String(res.content || '').trim();
  content = content.replace(/^```\w*\s*/i, '').replace(/\s*```$/i, '').trim();
  return content || null;
}
