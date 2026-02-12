function stepImprove_(settings, data) {
  const prompt = buildImprovePrompt_(data);
  const res = callOpenRouter_(
    settings.OPENROUTER_API_KEY,
    settings.MODEL_IMPROVE,
    prompt,
    settings.TEMPERATURE_IMPROVE
  );
  if (!res.ok) {
    throw res.error;
  }

  const parsed = safeJsonParse_(res.content);
  validateKeys_(parsed, ['title', 'improved_text', 'bullets']);

  return {
    title: parsed.title,
    text: parsed.improved_text,
    bullets: Array.isArray(parsed.bullets) ? parsed.bullets.join(' • ') : String(parsed.bullets || ''),
    model: settings.MODEL_IMPROVE
  };
}

function stepSpin_(settings, data, improvedText) {
  const prompt = buildSpinPrompt_(data, improvedText);
  const res = callOpenRouter_(
    settings.OPENROUTER_API_KEY,
    settings.MODEL_SPIN,
    prompt,
    settings.TEMPERATURE_SPIN
  );
  if (!res.ok) {
    throw res.error;
  }

  const parsed = safeJsonParse_(res.content);
  validateKeys_(parsed, ['spintext']);

  return { spintext: parsed.spintext, model: settings.MODEL_SPIN };
}

function stepAvitoHtml_(settings, data, spintext) {
  const prompt = buildAvitoHtmlPrompt_(settings, data, spintext);
  const res = callOpenRouter_(
    settings.OPENROUTER_API_KEY,
    settings.MODEL_AVITO,
    prompt,
    settings.TEMPERATURE_AVITO
  );
  if (!res.ok) {
    throw res.error;
  }

  const parsed = safeJsonParse_(res.content);
  validateKeys_(parsed, ['html', 'qa_checks']);

  return {
    html: parsed.html,
    qa_checks: Array.isArray(parsed.qa_checks)
      ? parsed.qa_checks.join(', ')
      : String(parsed.qa_checks || ''),
    model: settings.MODEL_AVITO
  };
}

function buildImprovePrompt_(data) {
  const system = {
    role: 'system',
    content:
      'Ты редактор коммерческих текстов для объявлений Avito. ' +
      'Сохраняй факты, не выдумывай. Убирай воду. Добавляй доверие и структуру. ' +
      'Не используй запрещенные обещания (гарантии результата, 100% и т.п.). ' +
      'Отвечай строго валидным JSON без лишнего текста.'
  };

  const user = {
    role: 'user',
    content: JSON.stringify({
      task: 'improve',
      product: String(data.product || ''),
      city: String(data.city || ''),
      tone: String(data.tone || 'нейтральный'),
      constraints: String(data.constraints || ''),
      source_text: String(data.source_text || ''),
      output_format: {
        title: 'строка до 60 символов',
        improved_text: 'строка',
        bullets: 'массив из 5-8 коротких преимуществ'
      }
    })
  };

  return [system, user];
}

function buildSpinPrompt_(data, improvedText) {
  const system = {
    role: 'system',
    content:
      'Ты создаешь спин-текст для уникализации объявлений Avito. ' +
      'Делай спин блоками: фразы/предложения, а не по одному слову. ' +
      'Сохраняй смысл и факты, не добавляй новые характеристики. ' +
      'Формат: {вариант 1|вариант 2|вариант 3}. ' +
      'Отвечай строго валидным JSON.'
  };

  const user = {
    role: 'user',
    content: JSON.stringify({
      task: 'spin',
      product: String(data.product || ''),
      improved_text: improvedText,
      requirements: {
        variants_per_block: 3,
        keep_facts: true,
        avoid_synonym_trash: true,
        length: 'примерно как improved_text'
      },
      output_format: { spintext: 'строка' }
    })
  };

  return [system, user];
}

function buildAvitoHtmlPrompt_(settings, data, spintext) {
  const system = {
    role: 'system',
    content:
      'Ты создаешь готовое объявление для Avito в HTML. ' +
      'Используй <p> и <b>. Можно эмодзи, но умеренно. ' +
      'Никаких запрещенных обещаний. Не выдумывай факты. ' +
      'Сделай читабельную структуру: лид → выгоды → условия → призыв. ' +
      'Отвечай строго валидным JSON.'
  };

  const user = {
    role: 'user',
    content: JSON.stringify({
      task: 'avito_html',
      product: String(data.product || ''),
      city: String(data.city || ''),
      constraints: String(data.constraints || ''),
      spintext: spintext,
      style: settings.AVITO_STYLE,
      output_format: {
        html: 'строка html',
        qa_checks: 'массив строк (короткие проверки качества)'
      }
    })
  };

  return [system, user];
}
