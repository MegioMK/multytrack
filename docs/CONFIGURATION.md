# Настройка и доступы

## 1. Google Таблица и Apps Script

1. Создайте новую Google Таблицу или сделайте ее копию.
2. В `Расширения -> Apps Script` замените файлы содержимым `src/google_apps_script`.
3. Сохраните проект и выполните функцию `installTaskTrackerAutomation`, подтвердив запрошенные Google-разрешения.
4. Откройте таблицу заново и запустите `Таск-трекер -> Восстановить автоматизацию`.

Для привязанного скрипта ID таблицы не требуется. Для standalone-скрипта добавьте Script Property `SPREADSHEET_ID`.

Apps Script Script Properties:

| Ключ | Назначение |
| --- | --- |
| `YANDEX_FUNCTION_URL` | URL Yandex Function для служебных запросов к боту. |
| `RELAY_SECRET` | Общий секрет Apps Script, Function и Worker. |
| `SPREADSHEET_ID` | Необязателен для привязанного Apps Script. |

## 2. Google service account

1. В Google Cloud создайте service account с доступом к Google Sheets API.
2. Создайте JSON-ключ и закодируйте его в base64 локально.
3. Дайте email сервисного аккаунта доступ редактора к пользовательской таблице.
4. Передайте base64-строку только в секрет `GOOGLE_SERVICE_ACCOUNT_JSON_B64` Yandex Function.

JSON-ключ нельзя класть в файл проекта, в таблицу или в GitHub.

## 3. Yandex Cloud Function

Загрузите `handler.py` и `requirements.txt` в Python Function. Задайте переменные окружения:

| Переменная | Назначение |
| --- | --- |
| `TELEGRAM_BOT_TOKEN` | Токен BotFather. |
| `ALLOWED_CHAT_ID` | Разрешенный Telegram chat для MVP. |
| `SPREADSHEET_ID` | ID таблицы этого пользователя. |
| `GOOGLE_SERVICE_ACCOUNT_JSON_B64` | Base64 JSON-ключ сервисного аккаунта. |
| `RELAY_SECRET` | Случайный общий секрет relay-цепочки. |
| `TELEGRAM_OUTBOUND_RELAY_URL` | URL Cloudflare Worker. |

Перед Function рекомендуется поставить лимит конкурентности и логирование ошибок. Function не должна быть Telegram webhook напрямую: публичной точкой служит Worker.

## 4. Cloudflare Worker

1. Скопируйте `src/cloudflare_worker/wrangler.toml.example` в локальный `wrangler.toml`.
2. Укажите URL Yandex Function в `YANDEX_FUNCTION_URL`.
3. Добавьте Worker secrets: `TELEGRAM_BOT_TOKEN`, `TELEGRAM_WEBHOOK_SECRET`, `RELAY_SECRET`.
4. Разверните Worker.
5. Вызовите Telegram `setWebhook` с URL Worker и `secret_token`, совпадающим с `TELEGRAM_WEBHOOK_SECRET`.

## Проверка доступов

- Сообщение боту появляется во `Входящие`.
- `Таск-трекер -> Обновить всё сейчас` отрабатывает без ошибок.
- `/today` возвращает ответ.
- В логах Worker и Function нет секретов и полного содержимого личных сообщений.
