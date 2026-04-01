# БизнесМатика — Resume Formatter Bot

Сервис автоматически переформатирует любое резюме в корпоративный стиль БизнесМатика.

## Как это работает
1. Telegram-бот получает резюме (PDF или DOCX)
2. n8n отправляет файл на этот сервер
3. Сервер извлекает текст → Claude структурирует данные → генерируется DOCX с корпоративным оформлением
4. Готовый файл возвращается в Telegram

## Деплой на Railway (бесплатно, 5 минут)

### Шаг 1 — Подготовь файлы
Убедись что в папке есть:
- app.py
- requirements.txt
- Procfile
- header_logo.png (логотип для шапки)
- footer_logo.png (логотип для подвала)

### Шаг 2 — Загрузи на GitHub
1. Зайди на github.com → New repository → назови "resume-bot"
2. Загрузи все файлы из этой папки

### Шаг 3 — Деплой на Railway
1. Зайди на railway.app
2. New Project → Deploy from GitHub repo
3. Выбери свой репозиторий
4. В разделе Variables добавь:
   - ANTHROPIC_API_KEY = твой ключ от console.anthropic.com

Railway автоматически запустит сервер и даст тебе URL вида:
https://resume-bot-production.up.railway.app

### Шаг 4 — Настрой n8n

В n8n создай workflow:

1. **Telegram Trigger** — получает файл от пользователя
2. **HTTP Request** — скачивает файл из Telegram
3. **HTTP Request** — отправляет файл на твой сервер:
   - Method: POST
   - URL: https://твой-адрес.railway.app/format-resume
   - Body: Form Data
   - Field name: file
   - Value: {{ бинарные данные файла }}
4. **Telegram** — отправляет готовый DOCX обратно пользователю

## API

POST /format-resume
- Content-Type: multipart/form-data
- Поле: file (PDF или DOCX)
- Возвращает: DOCX файл

GET /health
- Проверка работоспособности сервера
