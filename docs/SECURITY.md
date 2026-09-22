# Безопасность и данные

## Что никогда не попадает в Git

- Telegram bot token и webhook secret.
- `RELAY_SECRET`.
- JSON-ключ service account и его base64-версия.
- ID и URL конкретных таблиц, функций, Workers, API Gateway.
- Экспорт листов, задачи, сообщения, имена авторов, chat ID и логи с персональными данными.

## Где хранить настройки

| Данные | Место |
| --- | --- |
| Токен Telegram, Google key, chat ID, spreadsheet ID | Yandex Cloud secrets / environment variables. |
| Webhook secret и relay secret | Cloudflare Worker secrets. |
| URL Function и relay secret для расписания | Apps Script Script Properties. |
| Пользовательские задачи и настройки работы | Конкретная Google Таблица. |

## Минимальные права

- Service account получает редакторский доступ только к нужной таблице.
- Worker принимает Telegram запросы только с корректным webhook secret.
- Исходящие служебные вызовы в Worker требуют `X-Relay-Secret`.
- Function принимает рабочие Telegram updates только от разрешенного chat в MVP.

При подозрении на утечку немедленно отозвать Telegram token, заменить relay/webhook secrets и удалить/перевыпустить Google service account key.
