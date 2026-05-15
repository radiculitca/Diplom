# Аналитика опросов — МГУ им. Огарева

Веб-приложение для автоматизированной обработки результатов студенческих опросов: построение таблиц, диаграмм и экспорт отчёта в Word.

## Требования

- Python 3.10+

## Установка и запуск

```bash
# 1. Клонировать репозиторий
git clone <url>
cd analytics_reports

# 2. Создать виртуальное окружение и установить зависимости
python -m venv venv

зайти и нажать:
venv\Scripts\activate.bat        # Windows
# source venv/bin/activate   # Linux / macOS

pip install -r requirements.txt

# 3. Запустить сервер
uvicorn app.main:app --reload
```

Открыть в браузере: [http://localhost:8000](http://localhost:8000)

## Поддерживаемые форматы файлов

- `.xlsx` — Excel (все листы)
- `.csv` — CSV с автоопределением разделителя
