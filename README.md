# Avito AI Pipeline for Google Sheets (Apps Script)

Этот репозиторий содержит Google Apps Script-конвейер для пакетной обработки объявлений Avito:

Два ассистента (роли) через API:
1. **OfferWriter** — создаёт продающий текст (оффер) и 3 варианта заголовка.
2. **AvitoFormatter** — уникализация (spintax) и HTML под Avito с сохранением смысла.

## Файлы

- `main.gs` — меню и оркестрация pipeline (маршрутизация по режиму).
- `openrouter.gs` — вызов OpenRouter API, retry/backoff.
- `prompts.gs` — инструкции и вызовы для OfferWriter и AvitoFormatter.
- `utils.gs` — листы, статусы, лог, OUT (колонки и маппинг).

## Структура листов и колонки

### INPUT (13 колонок)

| № | Колонка      | Описание |
|---|--------------|----------|
| 1 | id           | Уникальный ID (подставляется при первом запуске) |
| 2 | product      | Продукт/услуга (обязателен для FULL, OFFER_ONLY) |
| 3 | city         | Город (обязателен для FULL, OFFER_ONLY) |
| 4 | source_text  | Исходный текст объявления (обязателен для FULL, OFFER_ONLY) |
| 5 | tone         | Тон (например нейтральный) |
| 6 | constraints  | Ограничения/пожелания |
| 7 | mode         | Режим: `FULL`, `OFFER_ONLY`, `SPIN_ONLY`, `HTML_ONLY` |
| 8 | status       | Статус (см. стейт-машину ниже) |
| 9 | last_error   | Текст последней ошибки (очищается при постановке в очередь и при успехе) |
| 10 | updated_at   | Время последнего обновления (проставляется автоматически) |
| 11 | locked_at    | Время захвата строки обработкой (при PROCESSING) |
| 12 | locked_by    | Кто захватил (например `batch`; очищается при DONE/ERROR/QUEUED) |
| 13 | processing   | TRUE — строка в обработке; FALSE по завершении (для дедупа и блокировок) |

**Статусы:** `NEW`, `QUEUED`, `PROCESSING`, `OFFER_OK`, `FORMAT_OK`, `SPIN_OK`, `HTML_OK`, `QA_OK`, `DONE`, `ERROR`.

**Маршрутизация по режиму (два ассистента):**
- **FULL:** сначала **OfferWriter** (оффер + 3 заголовка) → OFFER_OK; затем **AvitoFormatter** (spintax + HTML) → FORMAT_OK → QA → DONE.
- **OFFER_ONLY:** только **OfferWriter** → OFFER_OK → QA → DONE.
- **SPIN_ONLY** / **HTML_ONLY:** только **AvitoFormatter** (база = уже имеющийся `offer_text` / `improved_text` в OUT) → FORMAT_OK → QA → DONE.

### OUT (10 колонок, новый формат)

| № | Колонка       | Описание |
|---|---------------|----------|
| 1 | id            | ID из INPUT |
| 2 | offer_text    | Продающий текст (OfferWriter) |
| 3 | title_1       | Заголовок 1 (OfferWriter) |
| 4 | title_2       | Заголовок 2 (OfferWriter) |
| 5 | title_3       | Заголовок 3 (OfferWriter) |
| 6 | spintax_text  | Текст с spintax (AvitoFormatter) |
| 7 | avito_html    | Готовый HTML для Avito (AvitoFormatter) |
| 8 | qa_checks     | Проверки качества |
| 9 | model_used    | Модель (для отладки) |
| 10 | tokens_est    | Оценка токенов |

Поддерживается совместимость со старыми колонками: `improved_text` ↔ `offer_text`, `spintext` ↔ `spintax_text`, `title` ↔ `title_1`.

### SETTINGS (2 колонки: key, value)

- **Модели и промпты:** `MODEL_OFFER`, `MODEL_FORMAT`, `MODEL_FALLBACK`, `PROMPT_VERSION_OFFER`, `PROMPT_VERSION_FORMAT`.
- Остальное: `OPENROUTER_API_KEY`, `BATCH_SIZE`, `MAX_RUNTIME_SECONDS`, `TEMPERATURE_OFFER`, `TEMPERATURE_FORMAT`, `AVITO_STYLE`, `LANGUAGE`; по желанию — `MODEL_IMPROVE`, `MODEL_SPIN`, `MODEL_AVITO` (используются как fallback).

### LOG (8 колонок)

Каждое действие пишется одной строкой.

| № | Колонка        | Описание |
|---|----------------|----------|
| 1 | timestamp      | Время события |
| 2 | row_id         | ID строки INPUT (или `-` для батча) |
| 3 | step           | Шаг: `LOCK`, `OfferWriter`, `AvitoFormatter`, `QA`, `ERROR`, `BATCH_START`, `BATCH_END`, `CHECKPOINT`, `QUEUE` |
| 4 | status_before  | Статус строки до действия |
| 5 | status_after   | Статус строки после действия |
| 6 | duration_ms    | Длительность шага в мс (0 для не-AI шагов) |
| 7 | cache_hit      | Попадание в кэш (пока не используется) |
| 8 | error          | Текст ошибки или пусто |

Лист **CONTROL** создаётся автоматически при первом открытии меню Avito AI (дашборд).

### CONTROL — структура дашборда

| Область | Строки | Содержимое |
|--------|--------|------------|
| Заголовок | 1 | A1: «Avito AI — CONTROL» |
| **Статусы** | 3–13 | Заголовок «——— Статусы ———», затем: NEW, QUEUED, PROCESSING, OFFER_OK, FORMAT_OK, SPIN_OK, HTML_OK, QA_OK, DONE, ERROR и количество по каждому |
| **Прогресс** | 13–14 | Заголовок «——— Прогресс ———», сводный текст (всего / готово / ошибки / в очереди) |
| **Последние действия** | 16+ | Заголовок, заголовки колонок, последние 15 записей из LOG |
| **Последние 20 ошибок** | 34+ | Заголовок, колонки ID, Статус, Ошибка, Время, Ссылка (HYPERLINK на строку INPUT) |

## SETTINGS

Обязательные и рекомендуемые ключи:

- `OPENROUTER_API_KEY`
- `MODEL_OFFER` (OfferWriter), `MODEL_FORMAT` (AvitoFormatter), `MODEL_FALLBACK`
- `PROMPT_VERSION_OFFER`, `PROMPT_VERSION_FORMAT`
- `MAX_ROWS_PER_RUN`
- `TEMPERATURE_IMPROVE`
- `TEMPERATURE_SPIN`
- `TEMPERATURE_AVITO`
- `AVITO_STYLE`
- `LANGUAGE`

## Меню Avito AI

- **Сгенерировать (выделенные)** — проверяет обязательные поля (product, city, source_text для FULL/OFFER_ONLY); валидные строки переводит в **QUEUED**, очищает last_error и updated_at/locked_* и сразу запускает батч (до MAX_ROWS_PER_RUN).
- **Сгенерировать (очередь 50)** — обрабатывает до 50 строк со статусом **QUEUED**, **OFFER_OK** или **SPIN_OK** (ручная очередь, без триггеров).
- **Повторить ошибки** — строки со статусом ERROR переводятся в QUEUED, last_error очищается, запускается батч до 50.
- **Перегенерировать с шага OFFER** — по выделенным строкам: очищает в OUT improved_text, title, bullets, spintext, avito_html, qa_checks; ставит status = QUEUED.
- **Перегенерировать с шага SPIN** — очищает в OUT spintext, avito_html, qa_checks; ставит status = OFFER_OK (следующий шаг — SPIN).
- **Перегенерировать с шага HTML** — очищает в OUT avito_html, qa_checks; ставит status = SPIN_OK (следующий шаг — HTML).
- **Остановить** — запрос остановки: текущая строка завершится, затем обработка прекратится.
- **Логи** — диалог с последними 15 записями из LOG.

Запуск только вручную из меню. Триггеры по времени не используются. При старте, в процессе и по завершении — toast-уведомления. UPDATED_AT и LOCKED_AT/LOCKED_BY проставляются автоматически; LAST_ERROR очищается при постановке в очередь и при успешном шаге.

**Блокировки и устойчивость:**
- **ScriptLock:** на время батча берётся `LockService.getScriptLock()` (ожидание до 30 с); при занятости показывается toast и выход.
- **На строке:** при старте обработки выставляются `PROCESSING=true`, `LOCKED_AT`, `LOCKED_BY`; по завершении (DONE/ERROR/QUEUED) — `PROCESSING=false`, lock очищается.
- **Дедуп:** если строка уже в статусе PROCESSING и `LOCKED_AT` свежее 10 минут — строка не попадает в батч (пропуск).
- **Чекпоинт:** если до лимита времени выполнения (`MAX_RUNTIME_SECONDS`) осталось меньше 25 секунд — прогресс сохраняется, батч завершается без ошибки; можно снова запустить «Сгенерировать (очередь 50)» для продолжения.
- **Батч из SETTINGS:** `BATCH_SIZE` (сколько строк за один запуск), `MAX_RUNTIME_SECONDS` (макс. время выполнения в секундах).

## Запуск

1. Откройте Google Таблицу.
2. Перейдите в **Расширения → Apps Script**.
3. Скопируйте код из `.gs` файлов.
4. Обновите таблицу.
5. Используйте меню **Avito AI** (лист CONTROL создаётся при первом открытии).

---

## Список колонок (кратко)

- **INPUT:** id, product, city, source_text, tone, constraints, mode, status, last_error, updated_at, locked_at, locked_by, processing
- **OUT:** id, offer_text, title_1, title_2, title_3, spintax_text, avito_html, qa_checks, model_used, tokens_est
- **SETTINGS:** key, value (MODEL_OFFER, MODEL_FORMAT, MODEL_FALLBACK, PROMPT_VERSION_OFFER, PROMPT_VERSION_FORMAT, BATCH_SIZE, MAX_RUNTIME_SECONDS, …)
- **LOG:** timestamp, row_id, step, status_before, status_after, duration_ms, cache_hit, error
