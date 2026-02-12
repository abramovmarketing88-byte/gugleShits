# Avito AI Pipeline for Google Sheets (Apps Script)

Этот репозиторий содержит Google Apps Script-конвейер для пакетной обработки объявлений Avito:

1. Улучшение исходного текста.
2. Генерация spin-текста.
3. Сборка итогового HTML объявления для Avito.

## Файлы

- `main.gs` — меню Google Sheets и оркестрация pipeline.
- `openrouter.gs` — интеграция с OpenRouter API и retry/backoff.
- `prompts.gs` — промты и шаги STEP1/STEP2/STEP3.
- `utils.gs` — работа с листами, статусами, логами и утилитами.

## Структура листов

Создайте листы:

- `INPUT`: `id, product, city, source_text, tone, constraints, status, last_error, updated_at`
- `OUT`: `id, improved_text, spintext, avito_html, title, bullets, qa_checks, model_used, tokens_est`
- `SETTINGS`: пары `key/value`
- `LOG`: `timestamp, id, step, status, message`

## SETTINGS

Обязательные и рекомендуемые ключи:

- `OPENROUTER_API_KEY`
- `MODEL_IMPROVE`
- `MODEL_SPIN`
- `MODEL_AVITO`
- `MAX_ROWS_PER_RUN`
- `TEMPERATURE_IMPROVE`
- `TEMPERATURE_SPIN`
- `TEMPERATURE_AVITO`
- `AVITO_STYLE`
- `LANGUAGE`

## Запуск

1. Откройте Google Таблицу.
2. Перейдите в **Расширения → Apps Script**.
3. Скопируйте код из `.gs` файлов.
4. Обновите таблицу.
5. Используйте меню **Avito AI**.
