# 📦 Аналитика доставки ТК

Веб-приложение для анализа данных доставки с разбивкой по транспортным компаниям
и сегментам городов по численности населения.

## Что умеет

- Загрузить Excel-файл (листы «Чеки ЗДР 26», «Чеки увел 26»)
- Показать % доставки, возвратов, средний чек по каждой ТК
- Разбить по сегментам города: до 20 тыс. / 20–50 / 50–100 / 100 тыс.+
- Тепловая карта: ТК × сегмент города
- История загрузок (через Supabase)
- Графики: Chart.js, без зависимостей

## Быстрый старт (локально)

```bash
# 1. Установить зависимости
pip install -r requirements.txt

# 2. Настроить Supabase (опционально — без него история не сохраняется)
cp .env.example .env
# Заполните SUPABASE_URL и SUPABASE_KEY

# 3. Создать таблицу в Supabase (один раз)
# Откройте Supabase Dashboard → SQL Editor, вставьте:
# CREATE TABLE IF NOT EXISTS uploads (
#     id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),
#     filename    TEXT NOT NULL,
#     uploaded_at TIMESTAMPTZ DEFAULT NOW(),
#     rows_count  INTEGER,
#     analytics   JSONB
# );

# 4. Запустить
uvicorn app.main:app --reload --port 8000

# 5. Открыть в браузере
# http://localhost:8000
```

## Деплой на Railway / Render

```bash
# Railway
railway init
railway up

# Render — добавить в render.yaml:
# startCommand: uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

## Структура проекта

```
delivery_app/
├── app/
│   ├── main.py          # FastAPI роуты
│   ├── analytics.py     # Парсинг Excel + расчёт метрик
│   └── database.py      # Supabase (сохранение истории)
├── data/
│   └── cities_population.py  # База населения городов РФ
├── templates/
│   └── index.html       # Фронтенд (HTML + Chart.js)
├── static/              # CSS/JS (если понадобятся)
├── requirements.txt
├── .env.example
└── README.md
```

## Сегменты городов

| Сегмент | Численность |
|---------|------------|
| до 20 тыс. | сёла, деревни, малые города |
| 20–50 тыс. | малые города |
| 50–100 тыс. | средние города |
| 100 тыс.+ | крупные города и мегаполисы |

Сопоставление города из файла с населением происходит через встроенную
базу ~500 городов РФ (данные Росстат 2023). Неизвестные населённые пункты
попадают в сегмент «до 20 тыс.» как сёла/деревни.
