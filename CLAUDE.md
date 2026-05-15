# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Команды

```bash
# Активировать виртуальное окружение (Windows)
venv\Scripts\activate.bat

# Запустить dev-сервер
uvicorn app.main:app --reload

# Установить зависимости
pip install -r requirements.txt
```

Тестового фреймворка, линтера и шага сборки нет. Проверка — вручную через браузер.

## Архитектура

**Бэкенд** — FastAPI-приложение (`app/main.py`), 5 POST-эндпоинтов:
- `/upload` — сохраняет файлы в `uploads/raw_<name>`, возвращает список листов
- `/process_sheets` — читает выбранные листы, запускает `clean_dataframe()`, сохраняет `uploads/clean_<name>.parquet`, возвращает метаданные столбцов
- `/analyze` — вызывает `generate_report_data()`: читает parquet-файлы, считает частоты ответов по каждому вопросу, возвращает данные в форме `window.appData`
- `/ai_group_answers` — нормализует свободные ответы через Mistral API или локальный Ollama (модель `mistral`)
- `/export_docx` — генерирует Word-документ через `python-docx`

Ключевые модули бэкенда:
- `app/data_logic.py` — `clean_dataframe()` убирает нумерацию из ответов/столбцов, обрабатывает столбцы с возрастом, унифицирует нумерованные ответы; `generate_report_data()` считает частоты по файлам; `get_column_groups()` группирует столбцы вида `"Вопрос / Подвопрос"` по префиксу
- `app/docx_gen.py` — одна функция `generate_docx()`, всё форматирование документа (поля, шрифт, стили) находится здесь
- `app/ai_report.py` — `group_answers_local()` использует Ollama, `group_answers_api()` использует Mistral API; для API-варианта нужен `MISTRAL_API_KEY` в `.env`

**Фронтенд** — 6-шаговый визард, без шага сборки, без ES-модулей. Весь JS работает в глобальной области видимости через `<script>`-теги в порядке загрузки:

| Порядок | Файл | Ответственность |
|---------|------|-----------------|
| 1 | `modules/utils.js` | Инициализация глобального состояния, `showToast`, `_escHtml`, `_escAttr`, `randomColor` |
| 2 | `modules/wizard.js` | `goToStep()`, обновление состояния кнопок, делегирование кнопки «Назад» |
| 3 | `modules/upload.js` | Drag-and-drop, отправка формы → `/upload` |
| 4 | `modules/sheets.js` | Форма выбора листов → `/process_sheets`, `renderLegendSettings()` |
| 5 | `modules/questions.js` | Рендер списка вопросов, `addQuestionToSortable`, `openMappingModal` |
| 6 | `modules/fuzzy.js` | Логика модалок: нечёткое/диапазонное/AI-группирование ответов |
| 7 | `modules/charts.js` | `drawChart`, `drawStackedChart`, `drawPieChart`; модалки редактирования/скрытия столбцов |
| 8 | `modules/step4.js` | CRUD разделов отчёта, перетаскивание вопросов в разделы |
| 9 | `modules/step5.js` | Обработчик `analyzeBtn` — вызывает `/analyze`, рендерит таблицы и графики |
| 10 | `modules/step6.js` | Вызов `/export_docx`, финальная инициализация тултипов |

HTML-шаблоны используют Jinja2 `{% include 'partials/stepN.html' %}`. Партиалы лежат в `app/templates/partials/` (step0–step5, wizard_nav, modals).

**Глобальное состояние** (все переменные на `window.*`, инициализируются в `utils.js`):
- `processedFiles` — массив объектов, возвращённых `/process_sheets` (содержит `clean_filename`, `columns`)
- `questionMapping` — `{ qName: { clean_filename: mappedQName } }` — соответствие вопросов между файлами
- `questionSourceFile` — `{ qName: fileIndex }` — из какого файла вопрос был изначально выбран
- `appData` — `{ "q_N": { data, file_keys, file_labels, options, ... } }` — заполняется при вызове анализа на шаге 5
- `charts` — кэш экземпляров Chart.js; ключи: `id` (столбчатая), `"stacked_"+id` (накопленная), `"pie_"+id+"_"+fi` (круговая)

**Шаги визарда** соответствуют DOM-идентификаторам `wizardStep0`–`wizardStep5`. Навигация реализована через CSS `translateX` на `#wizardTrack`.

## Важные ограничения

- **Динамические значения в HTML всегда экранировать**: `_escAttr(s)` — для значений атрибутов, `_escHtml(s)` — для текстового содержимого. Никогда не вставлять имена вопросов или тексты ответов «сырыми» в шаблонные строки, генерирующие HTML. Не использовать inline `onclick="fn('${value}')"` — вместо этого применять `data-*`-атрибуты и делегирование событий.
- `style="display:none"` на `#legendSettingsBlock` и `#fileSelectContainer` — намеренно: JS переключает `.style.display` напрямую; заменять CSS-классами нельзя.
- Загруженные файлы сохраняются в `uploads/` и не удаляются между перезапусками сервера.
- Для AI-функций нужен либо `MISTRAL_API_KEY` в `.env` (API-режим), либо запущенный Ollama с загруженной моделью `mistral` (локальный режим).
