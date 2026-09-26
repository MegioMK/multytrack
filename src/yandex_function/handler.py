import base64
import hmac
import json
import logging
import os
import re
import socket
import ssl
from datetime import date, datetime, timedelta, timezone
from html import escape
from typing import Any

import requests
from google.auth.transport.requests import AuthorizedSession
from google.oauth2 import service_account


logger = logging.getLogger()
logger.setLevel(logging.INFO)

TELEGRAM_API_BASE = "https://api.telegram.org"
SHEETS_API_BASE = "https://sheets.googleapis.com/v4/spreadsheets"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SAMARA_TZ = timezone(timedelta(hours=4))


SHEET_SCHEMAS: dict[str, list[str]] = {
    "Пульт": [
        "ID",
        "Тип",
        "Приоритет",
        "Проект / область",
        "Статус",
        "Текущая неделя",
        "Следующая неделя",
        "Дедлайн",
        "Размер",
        "Цвет",
        "Последнее обновление",
        "Заметки",
    ],
    "Задачи": [
        "ID",
        "Заголовок",
        "Тип",
        "Проект / область",
        "Родительский ID",
        "Статус",
        "Приоритет",
        "Дедлайн",
        "Дата планирования",
        "Следующий шаг",
        "Источник",
        "Ссылка на входящее",
        "Создано",
        "Обновлено",
        "Заметки",
    ],
    "Подзадачи": [
        "ID",
        "Задача ID",
        "Название",
        "Статус",
        "Дедлайн",
        "Дата планирования",
        "Оценка времени, мин",
        "Тип мышления",
        "Готовность",
        "Недельный исход",
        "Порядок",
        "Обновлено",
        "Заметки",
    ],
    "План дня": [
        "ID",
        "Дата",
        "Начало",
        "Конец",
        "Тип слота",
        "Задача ID",
        "Подзадача ID",
        "Название",
        "Статус",
        "Источник",
        "Заметки",
    ],
    "Входящие": [
        "ID",
        "Дата захвата",
        "Источник",
        "Telegram chat_id",
        "Telegram message_id",
        "Автор",
        "Исходный текст",
        "Предложенный тип",
        "Подсказка проекта",
        "Срочность",
        "Формулировка задачи",
        "Статус разбора",
        "Созданная задача ID",
        "Заметки",
    ],
    "Рутины": [
        "ID",
        "Название",
        "Заметки",
        "Категория",
        "Статус",
        "Повторение",
        "Следующая дата",
        "Последнее выполнение",
        "Интервал",
        "Оценка, мин",
        "Текущая подзадача ID",
        "Пауза до",
    ],
    "Импорт календаря": [
        "ID",
        "Дата",
        "Начало",
        "Конец",
        "Название",
        "Календарь",
        "Источник",
        "Заметки",
    ],
    "Итоги недели": [
        "ID",
        "Неделя",
        "Главный результат",
        "Энергия и фокус",
        "Что сработало",
        "Напряжение и откладывание",
        "Урок недели",
        "Коммуникация",
        "Система работы",
        "Рост",
        "Следующий цикл",
        "Закрыто подзадач",
        "Закрыто рутин",
        "Закрыто задач",
        "Куплено",
        "Закрыто, мин",
        "Запланировано, мин",
        "Просрочено на конец недели",
        "По проектам",
        "Создано",
    ],
    "Настройки": [
        "Ключ",
        "Значение",
        "Описание",
        "Обновлено",
    ],
    "Сводки": [
        "ID",
        "Дата",
        "Тип",
        "Период",
        "Сводка",
        "Создано",
    ],
}


DEFAULT_SETTINGS = [
    ["часовой_пояс", "Europe/Samara", "Часовой пояс для дат и ежедневных сводок"],
    ["telegram_chat_id", "", "Разрешенный личный Telegram chat_id"],
    ["telegram_last_update_id", "0", "Последний обработанный Telegram update_id для polling"],
    ["утренний_обзор_включен", "да", "Отправлять ли обзор по cron"],
    ["утренний_обзор_час", "8", "Локальный час отправки утреннего обзора"],
    ["утренний_обзор_минута", "0", "Локальная минута отправки утреннего обзора"],
    ["входящие_напоминание_дней", "1", "Через сколько дней напоминать о неразобранном входящем"],
    ["wip_большие_камни", "2", "Ориентир по числу крупных активных проектов"],
    ["wip_маленькие_камни", "3", "Ориентир по числу небольших активных проектов"],
    ["горячая_задача_дней_до_дедлайна", "2", "Порог для подсветки близких дедлайнов"],
    ["рефлексия_недели_состояние", "", "Служебное состояние диалога недельной рефлексии"],
]


WEEKLY_REFLECTION_QUESTIONS = [
    ("Главный результат", "Где на этой неделе ты реально продвинулась, а не просто была занята?"),
    ("Энергия и фокус", "Что заряжало? Что съедало энергию или распыляло внимание?"),
    ("Что сработало", "Какие действия, решения или элементы системы работы хочется повторить?"),
    ("Напряжение и откладывание", "Что получилось частично, сорвалось или откладывалось? Почему на самом деле?"),
    ("Урок недели", "Какое решение было особенно удачным? Какой урок или кейс забираешь с собой?"),
    ("Коммуникация", "Где люди особенно помогли? Где не хватило ясности в договоренностях и ожиданиях?"),
    ("Система работы", "Что в планировании, трекинге и приоритизации стоит упростить, убрать или изменить?"),
    ("Рост", "Какие навыки ты прокачала? Какой маленький, но конкретный шаг роста выбираешь на следующий цикл?"),
    ("Следующий цикл", "Что переносишь, какой один главный фокус выбираешь, что точно повторишь или не повторишь и какая поддержка нужна?"),
]


class ConfigError(RuntimeError):
    pass


class ForbiddenError(RuntimeError):
    pass


def handler(event: dict[str, Any], context: Any) -> dict[str, Any]:
    try:
        control = _parse_control_request(event)
        if control:
            if control.get("mode") == "send_notification":
                text = str(control.get("text") or "").strip()
                if not text:
                    raise ConfigError("Не передан текст уведомления")
                send_telegram_message(_configured_chat_id(), text)
                return _json_response({"ok": True, "mode": "send_notification"})
            if control.get("mode") == "send_notifications":
                texts = [str(text).strip() for text in control.get("texts", []) if str(text).strip()]
                if not texts:
                    raise ConfigError("Не переданы тексты уведомлений")
                for text in texts:
                    send_telegram_message(_configured_chat_id(), text)
                return _json_response({"ok": True, "mode": "send_notifications", "sent": len(texts)})
            if control.get("mode") == "weekly_reflection_prompt":
                return _json_response(start_weekly_reflection_prompt(
                    str(control.get("week_start") or ""),
                    str(control.get("cutoff") or ""),
                ))

        if event.get("mode") == "send_test":
            chat_id = str(event.get("chat_id") or _configured_chat_id())
            send_telegram_message(chat_id, "Self-test: Yandex Function умеет отправлять сообщения в Telegram.")
            return _json_response({"ok": True, "mode": "send_test"})

        if event.get("mode") == "net_diag":
            return _json_response(run_network_diagnostics())

        if event.get("mode") == "poll":
            return _json_response(poll_telegram_updates())

        if _is_timer_event(event):
            if _timer_payload_mode(event) == "poll":
                return _json_response(poll_telegram_updates())
            return _json_response({"ok": True, "mode": "timer_ignored"})

        update = _parse_update(event)
        if not update:
            return _json_response({"ok": True, "ignored": True})

        callback = update.get("callback_query")
        if callback:
            return handle_callback_query(callback)

        message = update.get("message") or update.get("edited_message")
        if not message:
            return _json_response({"ok": True, "ignored": True})

        _enforce_allowed_chat(message)
        reply_text = process_message(message)
        if reply_text is None:
            # Telegram может повторно доставить уже обработанное сообщение.
            return _json_response({"ok": True, "duplicate": True})
        if isinstance(reply_text, dict):
            return _telegram_reply(
                str(message["chat"]["id"]),
                str(reply_text["text"]),
                reply_text.get("reply_markup"),
            )
        return _telegram_reply(str(message["chat"]["id"]), reply_text)
    except ForbiddenError as exc:
        logger.warning("forbidden request: %s", exc)
        return _json_response({"ok": False, "error": _public_error_message(exc)}, status_code=403)
    except Exception as exc:
        logger.exception("handler failed")
        return _json_response({"ok": False, "error": _public_error_message(exc)}, status_code=500)


def process_message(message: dict[str, Any]) -> str | dict[str, Any] | None:
    text = (message.get("text") or message.get("caption") or "").strip()
    if text.startswith("/"):
        return handle_command(message, text)
    reflection_handled, reflection_reply = process_weekly_reflection_reply(message, text)
    if reflection_handled:
        return reflection_reply
    return capture_incoming(message, force_task=False)


def handle_command(message: dict[str, Any], text: str) -> str | dict[str, Any]:
    command, _ = _split_command(text)

    if command in {"/start", "/help"}:
        return (
            "Любой текст или пересланное сообщение сохраню во <a href=\""
            + _incoming_sheet_url() + "\">Входящие</a>.\n\n"
            "/today - актуальные подзадачи на сегодня\n"
            "/hot - просроченные подзадачи\n\n"
            "/shopping - активный список покупок\n\n"
            "Кнопки Done и Куплено сразу обновляют таблицу."
        )
    if command == "/today":
        return build_actionable_subtasks_digest("today")
    if command == "/hot":
        return build_actionable_subtasks_digest("hot")
    if command == "/shopping":
        return build_actionable_shopping_digest()
    return "Не знаю такую команду. Доступно: /today, /hot, /shopping, /help."


def capture_incoming(
    message: dict[str, Any],
    force_task: bool,
    force_routine: bool = False,
    override_text: str | None = None,
) -> str | None:
    raw_text = (override_text or message.get("text") or message.get("caption") or "").strip()
    if not raw_text:
        raw_text = _non_text_message_title(message)
    if not raw_text:
        return "Не смогла превратить это сообщение в задачу. Добавь к нему короткую подпись."

    now = _now()
    entry_id = f"tg_{message['chat']['id']}_{message['message_id']}"
    if _incoming_entry_exists(entry_id):
        # Telegram повторяет webhook, если предыдущий ответ задержался или оборвался.
        return None
    source = "Telegram (переслано)" if _is_forwarded_message(message) else "Telegram"
    project = _project_hint(raw_text)
    deadline = _extract_deadline(raw_text)
    notes = _capture_notes(message, project, deadline)
    suggested_type = "рутина" if force_routine else (
        "задача" if force_task else _suggest_incoming_type(raw_text)
    )
    row = [
        entry_id,
        now.isoformat(timespec="seconds"),
        source,
        str(message["chat"]["id"]),
        str(message["message_id"]),
        _author_name(message.get("from") or {}),
        raw_text,
        suggested_type,
        project,
        "срочно" if _looks_urgent(raw_text) else "",
        "",
        "new",
        "",
        notes,
    ]
    sheets_append_values("Входящие", [row])
    return _incoming_confirmation(suggested_type)


def _incoming_entry_exists(entry_id: str) -> bool:
    rows = sheets_get_values("Входящие", "A12:A2000")
    return any(row and str(row[0]).strip() == entry_id for row in rows)


def _incoming_confirmation(suggested_type: str) -> str:
    labels = {"задача": "задачу", "рутина": "рутину", "покупка": "покупку"}
    label = labels.get(suggested_type, "задачу")
    new_count = _new_incoming_count()
    return (
        f'Записал во <a href="{_incoming_sheet_url()}">Входящие</a> как {label}. '
        f"Неразобранных входящих: {new_count}."
    )


def _new_incoming_count() -> int:
    # В первых десяти строках расположен дашборд, а шапка списка - в строке 11.
    values = sheets_get_values("Входящие", "A11:N1000")
    if not values:
        return 0
    headers = values[0]
    try:
        status_index = headers.index("Статус разбора")
    except ValueError:
        return 0
    return sum(
        len(row) > status_index and str(row[status_index]).strip().lower() == "new"
        for row in values[1:]
    )


def _suggest_incoming_type(text: str) -> str:
    """Recognize purchases before routines and ordinary tasks."""
    normalized = text.lower()
    if re.search(r"\bкупить\b", normalized):
        return "покупка"
    routine_pattern = re.compile(
        r"\b(кажд(?:ый|ая|ое|ые|ую)|ежеднев\w*|еженедел\w*|ежемесяч\w*|"
        r"ежегодн\w*|раз\s+в\s+(?:день|недел\w*|месяц|год)|"
        r"раз\s+в\s+\d+\s+(?:дн\w*|недел\w*|месяц\w*|год\w*)|"
        r"кажд\w*\s+\d+\s+(?:дн\w*|недел\w*|месяц\w*|год\w*)|"
        r"через\s+\d+\s+(?:дн\w*|недел\w*|месяц\w*)|"
        r"регулярн\w*|периодич\w*|повторя\w*|на\s+постоянной\s+основе|"
        r"по\s+(?:понедельникам|вторникам|средам|четвергам|пятницам|субботам|воскресеньям))"
    )
    habitual_time_pattern = re.compile(r"^\s*(?:утром|вечером|дн[её]м)\b")
    return "рутина" if routine_pattern.search(normalized) or habitual_time_pattern.search(normalized) else "задача"


def promote_ready_incoming() -> int:
    values = sheets_get_values("Входящие", "A1:Z500")
    if len(values) < 2:
        return 0
    headers = values[0]
    tasks = _rows_as_dicts("Задачи")
    existing_task_ids = {row.get("ID") for row in tasks}
    promoted = 0
    terminal_statuses = {"done", "cancelled", "skipped"}

    for row_index, raw_row in enumerate(values[1:], start=2):
        row = raw_row + [""] * (len(headers) - len(raw_row))
        item = {headers[index]: str(row[index]).strip() for index in range(len(headers))}
        title = item.get("Формулировка задачи", "")
        entry_id = item.get("ID", "")
        if not title or not entry_id or item.get("Статус разбора", "").lower() in terminal_statuses:
            continue

        task_id = f"task_{entry_id}"
        source = item.get("Источник") or "Ручной ввод"
        project = item.get("Подсказка проекта") or _project_hint(title)
        deadline = _extract_deadline(title) or _extract_deadline(item.get("Исходный текст", ""))
        priority = "высокий" if item.get("Срочность") else "обычный"
        now = _now().isoformat(timespec="seconds")
        if task_id not in existing_task_ids:
            sheets_append_values("Задачи", [[
                task_id,
                title,
                "задача",
                project,
                "",
                "new",
                priority,
                deadline,
                "",
                "",
                source,
                entry_id,
                now,
                now,
                item.get("Заметки", ""),
            ]])
            existing_task_ids.add(task_id)
        sheets_update_values("Входящие", f"L{row_index}:M{row_index}", [["done", task_id]])
        promoted += 1
    return promoted


def build_morning_digest() -> str:
    lines = [build_today_digest()]
    stale = stale_incoming()
    if stale:
        lines.extend([
            "",
            f"Неразобранных входящих старше {_incoming_reminder_days()} дн.: {len(stale)}",
            f"Разобрать в таблице: {_spreadsheet_url()}",
        ])
    return "\n".join(lines)


def stale_incoming() -> list[dict[str, str]]:
    today = _now().date()
    threshold = today - timedelta(days=_incoming_reminder_days())
    terminal_statuses = {"done", "cancelled", "skipped"}
    return [
        row for row in _rows_as_dicts("Входящие")
        if row.get("Статус разбора", "").lower() not in terminal_statuses
        and not row.get("Формулировка задачи", "")
        and (_parse_date(row.get("Дата захвата", "")) or today) <= threshold
    ]


def _is_forwarded_message(message: dict[str, Any]) -> bool:
    return bool(
        message.get("forward_origin")
        or message.get("forward_from")
        or message.get("forward_sender_name")
        or message.get("is_automatic_forward")
    )


def _forward_note(message: dict[str, Any]) -> str:
    if not _is_forwarded_message(message):
        return ""
    origin = message.get("forward_origin") or {}
    sender = origin.get("sender_user")
    sender_name = ""
    if isinstance(sender, dict):
        sender_name = " ".join(
            part for part in [sender.get("first_name", ""), sender.get("last_name", "")]
            if part
        ).strip()
    if not sender_name and isinstance(origin.get("sender_chat"), dict):
        sender_name = origin["sender_chat"].get("title", "")
    if not sender_name:
        sender_name = str(origin.get("sender_user_name") or "")
    sender_name = sender_name or str(message.get("forward_sender_name") or "")
    return f"Переслано сообщение{f' от {sender_name}' if sender_name else ''}."


def _capture_notes(message: dict[str, Any], project: str, deadline: str) -> str:
    notes = [_forward_note(message)]
    tags = []
    if project:
        tags.append(f"проект: {project}")
    if deadline:
        tags.append(f"дедлайн: {deadline}")
    if tags:
        notes.append("Авторазбор: " + "; ".join(tags) + ".")
    return " ".join(note for note in notes if note)


def _non_text_message_title(message: dict[str, Any]) -> str:
    kinds = [
        ("photo", "Фото"),
        ("video", "Видео"),
        ("document", "Документ"),
        ("audio", "Аудио"),
        ("voice", "Голосовое сообщение"),
        ("location", "Геолокация"),
        ("contact", "Контакт"),
    ]
    for key, title in kinds:
        if message.get(key):
            prefix = "Пересланное " if _is_forwarded_message(message) else ""
            return f"{prefix}{title.lower()} без подписи"
    return ""


PAGE_SIZE = 5


def build_actionable_subtasks_digest(mode: str, page: int = 0) -> dict[str, Any]:
    today = _now().date()
    selected: list[dict[str, str]] = []
    for item in _sheet_rows_with_numbers("Подзадачи", "A11:O1000"):
        status = item.get("Статус", "").lower()
        due = _parse_date(item.get("Дата", ""))
        if not item.get("ID") or status in {"done", "cancelled", "skipped"} or not due:
            continue
        if mode == "today" and due == today:
            selected.append(item)
        if mode == "hot" and due < today:
            selected.append(item)

    selected.sort(key=lambda item: (item.get("Дата", ""), item.get("Название", "")))
    if mode == "today":
        heading = f"<b>Сегодня, {today.strftime('%d.%m')}</b>"
        empty = "На сегодня открытых подзадач нет."
    else:
        heading = "<b>🔥 Просрочено</b>"
        empty = "Просроченных подзадач нет. Отличный темп!"
    if not selected:
        return {"text": heading + "\n\n" + empty}

    page_count = max(1, (len(selected) + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(0, min(page, page_count - 1))
    start = page * PAGE_SIZE
    visible = selected[start:start + PAGE_SIZE]
    lines = [heading, f"Подзадачи {start + 1}-{start + len(visible)} из {len(selected)}"]
    keyboard = []
    for item in visible:
        title = item.get("Название") or "(без названия)"
        is_routine = str(item.get("Источник ID", "")).startswith("routine_")
        prefix = "🌿 " if is_routine else "• "
        lines.append(prefix + escape(title))
        button_title = title if len(title) <= 28 else title[:25] + "..."
        keyboard.append([{
            "text": "Done: " + button_title,
            "callback_data": _done_callback_data(item),
        }])
    navigation = []
    if page > 0:
        navigation.append({"text": "◀️ Назад", "callback_data": f"page:{mode}:{page - 1}"})
    if page + 1 < page_count:
        count = min(PAGE_SIZE, len(selected) - start - PAGE_SIZE)
        navigation.append({"text": f"Следующие {count} ▶️", "callback_data": f"page:{mode}:{page + 1}"})
    if navigation:
        keyboard.append(navigation)
    return {"text": "\n".join(lines), "reply_markup": {"inline_keyboard": keyboard}}


def build_actionable_shopping_digest(page: int = 0) -> dict[str, Any]:
    closed_statuses = {"done", "cancelled", "skipped"}
    selected = [
        item for item in _sheet_rows_with_numbers("Список покупок", "A11:H1000")
        if item.get("ID") and item.get("Покупка")
        and item.get("Статус", "").lower() not in closed_statuses
    ]
    heading = "<b>🛒 Список покупок</b>"
    if not selected:
        return {"text": heading + "\n\nСписок пуст. Всё куплено!"}

    page_count = max(1, (len(selected) + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(0, min(page, page_count - 1))
    start = page * PAGE_SIZE
    visible = selected[start:start + PAGE_SIZE]
    lines = [heading, f"Покупки {start + 1}-{start + len(visible)} из {len(selected)}"]
    keyboard = []
    for item in visible:
        title = item["Покупка"]
        lines.append("• " + escape(title))
        button_title = title if len(title) <= 28 else title[:25] + "..."
        keyboard.append([{
            "text": "Куплено: " + button_title,
            "callback_data": _purchase_callback_data(item),
        }])
    navigation = []
    if page > 0:
        navigation.append({"text": "◀️ Назад", "callback_data": f"page:shopping:{page - 1}"})
    if page + 1 < page_count:
        count = min(PAGE_SIZE, len(selected) - start - PAGE_SIZE)
        navigation.append({"text": f"Следующие {count} ▶️", "callback_data": f"page:shopping:{page + 1}"})
    if navigation:
        keyboard.append(navigation)
    return {"text": "\n".join(lines), "reply_markup": {"inline_keyboard": keyboard}}


def handle_callback_query(callback: dict[str, Any]) -> dict[str, Any]:
    callback_id = str(callback.get("id") or "")
    message = callback.get("message") or {}
    sender_id = str((callback.get("from") or {}).get("id") or "")
    if sender_id != _configured_chat_id() or not message:
        return _telegram_callback_reply(callback_id, "Недостаточно прав.", show_alert=True)
    _enforce_allowed_chat(message)
    reflection_action = str(callback.get("data") or "")
    if reflection_action in {"reflection:start", "reflection:later", "reflection:skip", "reflection:finish"}:
        return handle_weekly_reflection_callback(callback_id, str(message["chat"]["id"]), reflection_action)
    page_match = re.fullmatch(r"page:(today|hot|shopping):(\d+)", str(callback.get("data") or ""))
    if page_match:
        mode = page_match.group(1)
        digest = (
            build_actionable_shopping_digest(int(page_match.group(2)))
            if mode == "shopping"
            else build_actionable_subtasks_digest(mode, int(page_match.group(2)))
        )
        edit_telegram_message(
            str(message["chat"]["id"]),
            int(message["message_id"]),
            str(digest["text"]),
            digest.get("reply_markup"),
        )
        return _telegram_callback_reply(callback_id, "Показала следующую страницу.")
    purchase_match = re.fullmatch(r"purchase:(\d+):([0-9a-f]{8})", str(callback.get("data") or ""))
    if purchase_match:
        row_number = int(purchase_match.group(1))
        item = _purchase_by_row_number(row_number)
        if not item or not hmac.compare_digest(purchase_match.group(2), _purchase_callback_signature(row_number, item["ID"])):
            return _telegram_callback_reply(callback_id, "Покупка изменилась. Обновите список.", show_alert=True)
        if item.get("Статус", "").lower() in {"done", "cancelled", "skipped"}:
            return _telegram_callback_reply(callback_id, "Уже отмечена.")
        sheets_update_values("Список покупок", f"C{row_number}:C{row_number}", [["done"]])
        sheets_update_values("Список покупок", f"G{row_number}:G{row_number}", [[_now().isoformat(timespec="seconds")]])
        sheets_update_value_by_header("Список покупок", row_number, "Дата закрытия", _now().isoformat(timespec="seconds"))
        title = str(item.get("Покупка") or "Покупка")
        send_telegram_message(str(message["chat"]["id"]), f"✅ Отметил купленным: <b>{escape(title)}</b>.")
        return _telegram_callback_reply(callback_id, "Готово, отметил купленным.")
    match = re.fullmatch(r"done:(\d+):([0-9a-f]{8})", str(callback.get("data") or ""))
    if not match:
        return _telegram_callback_reply(callback_id, "Эта кнопка больше не действует.", show_alert=True)
    row_number = int(match.group(1))
    item = _subtask_by_row_number(row_number)
    if not item or not hmac.compare_digest(match.group(2), _done_callback_signature(row_number, item["ID"])):
        return _telegram_callback_reply(callback_id, "Подзадача изменилась. Обновите список.", show_alert=True)
    if item.get("Статус", "").lower() in {"done", "cancelled", "skipped"}:
        return _telegram_callback_reply(callback_id, "Уже закрыта.")
    sheets_update_values("Подзадачи", f"E{row_number}:E{row_number}", [["done"]])
    sheets_update_values("Подзадачи", f"L{row_number}:L{row_number}", [[_now().isoformat(timespec="seconds")]])
    sheets_update_value_by_header("Подзадачи", row_number, "Дата закрытия", _now().isoformat(timespec="seconds"))
    title = str(item.get("Название") or item.get("Задача") or "Подзадача")
    send_telegram_message(str(message["chat"]["id"]), f"✅ Обновил: <b>{escape(title)}</b>.")
    return _telegram_callback_reply(callback_id, "Готово, обновил.")


def start_weekly_reflection_prompt(week_start: str, cutoff: str) -> dict[str, Any]:
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", week_start) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", cutoff):
        raise ConfigError("Для недельной рефлексии нужны даты начала и среза.")
    state = _weekly_reflection_state()
    if state.get("cutoff") == cutoff and state.get("status") == "completed":
        return {"ok": True, "mode": "weekly_reflection_prompt", "skipped": "completed"}
    row = _weekly_reflection_row(cutoff)
    if not row:
        raise ConfigError("Не найдена строка итогов недели для рефлексии.")
    if state.get("cutoff") == cutoff and state.get("status") in {"waiting", "active"}:
        return {"ok": True, "mode": "weekly_reflection_prompt", "skipped": "already_started"}
    state = {"week_start": week_start, "cutoff": cutoff, "status": "waiting", "step": 0, "question_message_id": 0}
    _set_weekly_reflection_state(state)
    send_telegram_message(
        _configured_chat_id(),
        "<b>Недельная рефлексия</b> 🌿\n"
        f"Цикл: {week_start} 12:00 — {cutoff} 12:00.\n"
        "Соберем главное за 5–10 минут? Ответы попадут в Итоги недели.",
        {"inline_keyboard": [[
            {"text": "Начать", "callback_data": "reflection:start"},
            {"text": "Не сейчас", "callback_data": "reflection:later"},
        ]]},
    )
    return {"ok": True, "mode": "weekly_reflection_prompt"}


def handle_weekly_reflection_callback(callback_id: str, chat_id: str, action: str) -> dict[str, Any]:
    state = _weekly_reflection_state()
    if not state or state.get("status") == "completed":
        return _telegram_callback_reply(callback_id, "Эта рефлексия уже завершена.", show_alert=True)
    if action == "reflection:later":
        state["status"] = "waiting"
        _set_weekly_reflection_state(state)
        return _telegram_callback_reply(callback_id, "Хорошо. Кнопка «Начать» останется здесь.")
    if action == "reflection:start":
        _send_next_weekly_reflection_question(chat_id, state)
        return _telegram_callback_reply(callback_id, "Начинаем.")
    if action == "reflection:skip":
        _advance_weekly_reflection(chat_id, state, "")
        return _telegram_callback_reply(callback_id, "Пропустила вопрос.")
    state["status"] = "completed"
    _set_weekly_reflection_state(state)
    send_telegram_message(chat_id, "Рефлексию сохранила. Ее всегда можно дополнить в <b>Итоги недели</b>.")
    return _telegram_callback_reply(callback_id, "Рефлексия завершена.")


def process_weekly_reflection_reply(message: dict[str, Any], text: str) -> tuple[bool, str | None]:
    state = _weekly_reflection_state()
    if state.get("status") != "active":
        return False, None
    reply_to = message.get("reply_to_message") or {}
    if int(reply_to.get("message_id") or 0) != int(state.get("question_message_id") or 0):
        # Только явный ответ на вопрос бота относится к рефлексии. Остальной
        # поток остается Входящими, даже пока диалог еще не завершен.
        return False, None
    if not text:
        return True, "Ответ не вижу. Напиши его текстом ответом на мой вопрос или нажми «Пропустить вопрос»."
    _advance_weekly_reflection(str(message["chat"]["id"]), state, text)
    return True, None


def _advance_weekly_reflection(chat_id: str, state: dict[str, Any], answer: str) -> None:
    step = int(state.get("step") or 0)
    if step >= len(WEEKLY_REFLECTION_QUESTIONS):
        state["status"] = "completed"
        _set_weekly_reflection_state(state)
        return
    header, _ = WEEKLY_REFLECTION_QUESTIONS[step]
    row = _weekly_reflection_row(str(state.get("cutoff") or ""))
    if not row:
        raise ConfigError("Строка итогов недели больше не найдена.")
    if answer:
        sheets_update_value_by_header("Итоги недели", int(row["__row_number"]), header, answer)
    state["step"] = step + 1
    _send_next_weekly_reflection_question(chat_id, state)


def _send_next_weekly_reflection_question(chat_id: str, state: dict[str, Any]) -> None:
    step = int(state.get("step") or 0)
    if step >= len(WEEKLY_REFLECTION_QUESTIONS):
        state["status"] = "completed"
        state["question_message_id"] = 0
        _set_weekly_reflection_state(state)
        send_telegram_message(chat_id, "✨ Рефлексия сохранена в <b>Итоги недели</b>. Спасибо, что остановилась и посмотрела на неделю целиком.")
        return
    _, question = WEEKLY_REFLECTION_QUESTIONS[step]
    response = send_telegram_message(
        chat_id,
        f"<b>{step + 1}/{len(WEEKLY_REFLECTION_QUESTIONS)}</b>\n{escape(question)}\n\n"
        "Ответь именно на это сообщение: тогда текст точно пойдет в рефлексию, а не во Входящие.",
        {"force_reply": True, "input_field_placeholder": "Напиши ответ"},
    )
    message_id = int(((response.get("result") or {}).get("message_id")) or 0)
    if not message_id:
        raise ConfigError("Telegram не вернул ID вопроса рефлексии.")
    state["status"] = "active"
    state["question_message_id"] = message_id
    _set_weekly_reflection_state(state)
    send_telegram_message(chat_id, "Можно пропустить вопрос или завершить рефлексию в любой момент.", {
        "inline_keyboard": [[
            {"text": "Пропустить вопрос", "callback_data": "reflection:skip"},
            {"text": "Завершить", "callback_data": "reflection:finish"},
        ]]
    })


def _weekly_reflection_row(cutoff: str) -> dict[str, str] | None:
    return next((row for row in _sheet_rows_with_numbers("Итоги недели", "A11:Z1000") if row.get("ID") == f"week_{cutoff}"), None)


def _weekly_reflection_state() -> dict[str, Any]:
    raw = _setting_value("рефлексия_недели_состояние", "")
    try:
        value = json.loads(raw) if raw else {}
    except ValueError:
        return {}
    return value if isinstance(value, dict) else {}


def _set_weekly_reflection_state(state: dict[str, Any]) -> None:
    _set_setting("рефлексия_недели_состояние", json.dumps(state, ensure_ascii=False, separators=(",", ":")))


def _sheet_rows_with_numbers(sheet_name: str, range_a1: str) -> list[dict[str, str]]:
    values = sheets_get_values(sheet_name, range_a1)
    if len(values) < 2:
        return []
    headers = [str(value).strip() for value in values[0]]
    first_data_row = 12
    rows = []
    for offset, row in enumerate(values[1:]):
        padded = row + [""] * (len(headers) - len(row))
        if not any(str(cell).strip() for cell in padded):
            continue
        item = {headers[index]: str(padded[index]).strip() for index in range(len(headers))}
        item["__row_number"] = str(first_data_row + offset)
        rows.append(item)
    return rows


def _subtask_by_row_number(row_number: int) -> dict[str, str] | None:
    if row_number < 12 or row_number > 1000:
        return None
    rows = _sheet_rows_with_numbers("Подзадачи", f"A11:O{row_number}")
    return next((row for row in rows if int(row["__row_number"]) == row_number), None)


def _purchase_by_row_number(row_number: int) -> dict[str, str] | None:
    if row_number < 12 or row_number > 1000:
        return None
    rows = _sheet_rows_with_numbers("Список покупок", f"A11:H{row_number}")
    return next((row for row in rows if int(row["__row_number"]) == row_number), None)


def _done_callback_data(item: dict[str, str]) -> str:
    row_number = int(item["__row_number"])
    return f"done:{row_number}:{_done_callback_signature(row_number, item['ID'])}"


def _done_callback_signature(row_number: int, subtask_id: str) -> str:
    payload = f"done:{row_number}:{subtask_id}".encode("utf-8")
    return hmac.new(_env("RELAY_SECRET").encode("utf-8"), payload, "sha256").hexdigest()[:8]


def _purchase_callback_data(item: dict[str, str]) -> str:
    row_number = int(item["__row_number"])
    return f"purchase:{row_number}:{_purchase_callback_signature(row_number, item['ID'])}"


def _purchase_callback_signature(row_number: int, purchase_id: str) -> str:
    payload = f"purchase:{row_number}:{purchase_id}".encode("utf-8")
    return hmac.new(_env("RELAY_SECRET").encode("utf-8"), payload, "sha256").hexdigest()[:8]


def _task_suffix(item: dict[str, str]) -> str:
    task = item.get("Задача", "")
    title = item.get("Название", "")
    return f" <i>- {escape(task)}</i>" if task and task != title else ""


def build_today_digest() -> str:
    today = _today_iso()
    tasks = _rows_as_dicts("Задачи")
    subtasks = _rows_as_dicts("Подзадачи")
    plan = _rows_as_dicts("План дня")
    pult = _rows_as_dicts("Пульт")

    terminal_statuses = {"done", "cancelled", "закрыта", "закрыто", "неактуальна", "неактуально", "отменено", "архив"}
    planned_tasks = [
        row for row in tasks
        if row.get("Статус", "").lower() not in terminal_statuses
        and (row.get("Дата планирования") == today or row.get("Дедлайн") == today)
    ]
    planned_subtasks = [
        row for row in subtasks
        if row.get("Статус", "").lower() not in terminal_statuses
        and (row.get("Дата планирования", "") <= today or row.get("Дедлайн", "") == today)
    ]
    today_slots = [
        row for row in plan
        if row.get("Статус", "").lower() not in terminal_statuses
        and row.get("Дата", "") <= today
    ]
    active_focus = [
        row for row in pult
        if row.get("Статус", "").lower() in {"активно", "в работе", "желтый", "красный"}
    ][:5]

    lines = [f"План на сегодня, {today}"]
    if active_focus:
        lines.append("")
        lines.append("Фокусы:")
        for row in active_focus:
            focus = row.get("Текущая неделя") or row.get("Проект / область") or "(без названия)"
            lines.append(f"- {focus}")

    lines.append("")
    lines.append("Задачи:")
    _extend_limited(lines, planned_tasks, "Заголовок", empty_text="На сегодня явных задач нет.")

    if planned_subtasks:
        lines.append("")
        lines.append("Подзадачи:")
        _extend_limited(lines, planned_subtasks, "Название")

    if today_slots:
        lines.append("")
        lines.append("Слоты дня, включая незакрытое с прошлых дней:")
        for row in today_slots[:6]:
            title = row.get("Название") or row.get("Тип слота") or "(без названия)"
            time_part = " ".join(part for part in [row.get("Начало"), row.get("Конец")] if part)
            lines.append(f"- {time_part} {title}".strip())

    return "\n".join(lines)


def build_hot_digest() -> str:
    today = date.fromisoformat(_today_iso())
    threshold = today + timedelta(days=_hot_threshold_days())
    tasks = _rows_as_dicts("Задачи")
    pult = _rows_as_dicts("Пульт")

    hot_tasks = []
    for row in tasks:
        status = row.get("Статус", "").lower()
        due_raw = row.get("Дедлайн", "")
        if status in {"done", "cancelled", "закрыто", "отменено", "архив"} or not due_raw:
            continue
        due = _parse_date(due_raw)
        if due and due <= threshold:
            hot_tasks.append(row)

    red_focuses = [
        row for row in pult
        if row.get("Цвет", "").lower() in {"красный", "желтый"}
        or row.get("Статус", "").lower() in {"красный", "желтый", "горит"}
    ]

    lines = ["Горячее"]
    lines.append("")
    lines.append("Дедлайны:")
    _extend_limited(lines, hot_tasks, "Заголовок", empty_text="Близких дедлайнов не вижу.")

    if red_focuses:
        lines.append("")
        lines.append("Красное/желтое в пульте:")
        for row in red_focuses[:8]:
            title = row.get("Проект / область") or row.get("Текущая неделя") or "(без названия)"
            color = row.get("Цвет") or row.get("Статус")
            lines.append(f"- {color}: {title}")

    return "\n".join(lines)


def poll_telegram_updates() -> dict[str, Any]:
    ensure_workbook()
    last_update_id = _setting_int("telegram_last_update_id", 0)
    updates = telegram_api_post("getUpdates", {"offset": last_update_id + 1, "timeout": 0}).get("result", [])
    processed = 0
    skipped = 0

    for update in updates:
        update_id = int(update.get("update_id", 0) or 0)
        try:
            message = update.get("message") or update.get("edited_message")
            if not message:
                skipped += 1
                continue
            _enforce_allowed_chat(message)
            reply_text = process_message(message)
            if reply_text is None:
                set_telegram_reaction(message)
            elif isinstance(reply_text, dict):
                send_telegram_message(
                    str(message["chat"]["id"]),
                    str(reply_text["text"]),
                    reply_text.get("reply_markup"),
                )
            else:
                send_telegram_message(str(message["chat"]["id"]), reply_text)
            processed += 1
        except Exception:
            logger.exception("poll update failed update_id=%s", update_id)
        finally:
            if update_id:
                _set_setting("telegram_last_update_id", str(update_id))

    return {"ok": True, "mode": "poll", "processed": processed, "skipped": skipped}


def run_network_diagnostics() -> dict[str, Any]:
    hosts = ["api.telegram.org", "sheets.googleapis.com", "ya.ru"]
    results = {}
    for host in hosts:
        results[host] = _diagnose_host(host)
    return {"ok": True, "mode": "net_diag", "results": results}


def _diagnose_host(host: str) -> dict[str, Any]:
    result: dict[str, Any] = {"dns": [], "tcp_443": None, "tls": None}
    try:
        result["dns"] = sorted({
            item[4][0]
            for item in socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
        })
    except OSError as exc:
        result["dns_error"] = f"{type(exc).__name__}: {exc}"
        return result

    address = result["dns"][0] if result["dns"] else host
    try:
        with socket.create_connection((address, 443), timeout=4):
            result["tcp_443"] = "ok"
    except OSError as exc:
        result["tcp_443"] = f"{type(exc).__name__}: {exc}"
        return result

    try:
        context = ssl.create_default_context()
        with socket.create_connection((address, 443), timeout=4) as raw_socket:
            with context.wrap_socket(raw_socket, server_hostname=host):
                result["tls"] = "ok"
    except OSError as exc:
        result["tls"] = f"{type(exc).__name__}: {exc}"

    return result


def ensure_workbook() -> None:
    metadata = sheets_get_metadata()
    existing = {sheet["properties"]["title"]: sheet["properties"]["sheetId"] for sheet in metadata.get("sheets", [])}
    missing = [title for title in SHEET_SCHEMAS if title not in existing]

    if missing:
        batch_update([{"addSheet": {"properties": {"title": title}}} for title in missing])
        metadata = sheets_get_metadata()
        existing = {sheet["properties"]["title"]: sheet["properties"]["sheetId"] for sheet in metadata.get("sheets", [])}

    for title, headers in SHEET_SCHEMAS.items():
        current = sheets_get_values(title, "1:1")
        current_headers = current[0] if current else []
        if title == "Входящие":
            legacy_headers = [header for header in headers if header != "Формулировка задачи"]
            if current_headers == legacy_headers:
                legacy_values = sheets_get_values(title, "A1:M500")
                migrated_values = [headers]
                for row in legacy_values[1:]:
                    padded = row + [""] * (len(legacy_headers) - len(row))
                    migrated_values.append(padded[:10] + [""] + padded[10:])
                sheets_update_values(title, f"A1:N{len(migrated_values)}", migrated_values)
                current_headers = headers
        if current_headers[: len(headers)] != headers:
            sheets_update_values(title, "1:1", [headers])
        _freeze_header(existing[title])

    _ensure_default_settings()


def sheets_get_metadata() -> dict[str, Any]:
    response = _authorized_session().get(f"{SHEETS_API_BASE}/{_spreadsheet_id()}")
    response.raise_for_status()
    return response.json()


def sheets_get_values(sheet_name: str, range_a1: str) -> list[list[str]]:
    encoded_range = requests.utils.quote(f"{sheet_name}!{range_a1}", safe="")
    response = _authorized_session().get(
        f"{SHEETS_API_BASE}/{_spreadsheet_id()}/values/{encoded_range}",
        params={"valueRenderOption": "FORMATTED_VALUE"},
    )
    response.raise_for_status()
    return response.json().get("values", [])


def sheets_update_values(sheet_name: str, range_a1: str, values: list[list[Any]]) -> None:
    encoded_range = requests.utils.quote(f"{sheet_name}!{range_a1}", safe="")
    response = _authorized_session().put(
        f"{SHEETS_API_BASE}/{_spreadsheet_id()}/values/{encoded_range}",
        params={"valueInputOption": "USER_ENTERED"},
        json={"values": values},
    )
    response.raise_for_status()


def sheets_update_value_by_header(sheet_name: str, row_number: int, header: str, value: Any) -> None:
    headers = sheets_get_values(sheet_name, "A11:Z11")
    if not headers or header not in headers[0]:
        raise ConfigError(f"Не найден столбец {header} на листе {sheet_name}.")
    column_index = headers[0].index(header) + 1
    column = ""
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        column = chr(65 + remainder) + column
    sheets_update_values(sheet_name, f"{column}{row_number}:{column}{row_number}", [[value]])


def sheets_append_values(sheet_name: str, values: list[list[Any]]) -> None:
    encoded_range = requests.utils.quote(f"{sheet_name}!A:Z", safe="")
    response = _authorized_session().post(
        f"{SHEETS_API_BASE}/{_spreadsheet_id()}/values/{encoded_range}:append",
        params={"valueInputOption": "USER_ENTERED", "insertDataOption": "INSERT_ROWS"},
        json={"values": values},
    )
    response.raise_for_status()


def batch_update(requests_payload: list[dict[str, Any]]) -> None:
    response = _authorized_session().post(
        f"{SHEETS_API_BASE}/{_spreadsheet_id()}:batchUpdate",
        json={"requests": requests_payload},
    )
    response.raise_for_status()


def _ensure_default_settings() -> None:
    existing = {row.get("Ключ") for row in _rows_as_dicts("Настройки")}
    now = _now().isoformat(timespec="seconds")
    missing = [row + [now] for row in DEFAULT_SETTINGS if row[0] not in existing]
    if missing:
        sheets_append_values("Настройки", missing)


def send_telegram_message(chat_id: str, text: str, reply_markup: dict[str, Any] | None = None) -> dict[str, Any]:
    outbound_relay_url = os.getenv("TELEGRAM_OUTBOUND_RELAY_URL", "")
    if outbound_relay_url:
        return _send_via_outbound_relay(outbound_relay_url, chat_id, text, reply_markup)
    payload: dict[str, Any] = {"chat_id": chat_id, "text": text[:3900], "parse_mode": "HTML"}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    return telegram_api_post("sendMessage", payload)


def edit_telegram_message(
    chat_id: str,
    message_id: int,
    text: str,
    reply_markup: dict[str, Any] | None = None,
) -> None:
    outbound_relay_url = os.getenv("TELEGRAM_OUTBOUND_RELAY_URL", "")
    if outbound_relay_url:
        _send_edit_via_outbound_relay(outbound_relay_url, chat_id, message_id, text, reply_markup)
        return
    payload: dict[str, Any] = {
        "chat_id": chat_id,
        "message_id": message_id,
        "text": text[:3900],
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    if reply_markup:
        payload["reply_markup"] = reply_markup
    telegram_api_post("editMessageText", payload)


def set_telegram_reaction(message: dict[str, Any]) -> None:
    payload = {
        "chat_id": str(message["chat"]["id"]),
        "message_id": int(message["message_id"]),
        "reaction": [{"type": "emoji", "emoji": "✅"}],
    }
    outbound_relay_url = os.getenv("TELEGRAM_OUTBOUND_RELAY_URL", "")
    if outbound_relay_url:
        _send_reaction_via_outbound_relay(outbound_relay_url, payload)
        return
    telegram_api_post("setMessageReaction", payload)


def _send_via_outbound_relay(
    relay_url: str,
    chat_id: str,
    text: str,
    reply_markup: dict[str, Any] | None = None,
) -> dict[str, Any]:
    try:
        response = requests.post(
            relay_url,
            headers={"X-Relay-Secret": _env("RELAY_SECRET")},
            json={
                "mode": "send_notification",
                "chat_id": chat_id,
                "text": text[:3900],
                "reply_markup": reply_markup,
            },
            timeout=(4, 12),
        )
    except requests.RequestException as exc:
        raise ConfigError("Исходящий Telegram relay недоступен") from exc
    try:
        data = response.json() if response.text else {}
    except ValueError:
        data = {}
    if response.status_code >= 400 or data.get("ok") is False:
        raise ConfigError("Исходящий Telegram relay отклонил сообщение")
    return data


def _send_edit_via_outbound_relay(
    relay_url: str,
    chat_id: str,
    message_id: int,
    text: str,
    reply_markup: dict[str, Any] | None,
) -> None:
    try:
        response = requests.post(
            relay_url,
            headers={"X-Relay-Secret": _env("RELAY_SECRET")},
            json={
                "mode": "edit_message",
                "chat_id": chat_id,
                "message_id": message_id,
                "text": text[:3900],
                "reply_markup": reply_markup,
            },
            timeout=(4, 12),
        )
    except requests.RequestException as exc:
        raise ConfigError("Исходящий Telegram relay недоступен") from exc
    try:
        data = response.json() if response.text else {}
    except ValueError:
        data = {}
    if response.status_code >= 400 or data.get("ok") is False:
        raise ConfigError("Исходящий Telegram relay не смог обновить сообщение")


def _send_reaction_via_outbound_relay(relay_url: str, payload: dict[str, Any]) -> None:
    try:
        response = requests.post(
            relay_url,
            headers={"X-Relay-Secret": _env("RELAY_SECRET")},
            json={"mode": "set_reaction", **payload},
            timeout=(4, 12),
        )
    except requests.RequestException as exc:
        raise ConfigError("Исходящий Telegram relay недоступен") from exc
    try:
        data = response.json() if response.text else {}
    except ValueError:
        data = {}
    if response.status_code >= 400 or data.get("ok") is False:
        raise ConfigError("Исходящий Telegram relay не смог поставить реакцию")


def telegram_api_post(method: str, payload: dict[str, Any]) -> dict[str, Any]:
    token = _env("TELEGRAM_BOT_TOKEN")
    try:
        response = requests.post(
            f"{TELEGRAM_API_BASE}/bot{token}/{method}",
            json=payload,
            timeout=(4, 8),
        )
    except requests.RequestException as exc:
        raise ConfigError(f"Telegram API request failed for {method}") from exc
    try:
        data = response.json() if response.text else {}
    except ValueError:
        data = {}
    if response.status_code >= 400 or data.get("ok") is False:
        description = data.get("description") or f"HTTP {response.status_code}"
        raise ConfigError(f"Telegram API error for {method}: {description}")
    return data


def _authorized_session() -> AuthorizedSession:
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    raw_b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_B64")
    if raw_b64:
        raw = base64.b64decode(raw_b64).decode("utf-8")
    if not raw:
        raise ConfigError("Нужен GOOGLE_SERVICE_ACCOUNT_JSON или GOOGLE_SERVICE_ACCOUNT_JSON_B64")
    info = json.loads(raw)
    credentials = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return AuthorizedSession(credentials)


def _spreadsheet_id() -> str:
    return _env("SPREADSHEET_ID")


def _env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise ConfigError(f"Нужна переменная окружения {name}")
    return value


def _configured_chat_id() -> str:
    return os.getenv("ALLOWED_CHAT_ID", "") or _setting_value("telegram_chat_id", "")


def _enforce_allowed_chat(message: dict[str, Any]) -> None:
    allowed = _configured_chat_id()
    actual = str(message["chat"]["id"])
    # Разрешаем узнать ID группы через /chat, не открывая прием обычных
    # сообщений из посторонних групп. Команда возвращает только ID этой группы.
    text = str(message.get("text") or "").strip()
    is_group_lookup = (
        str((message.get("chat") or {}).get("type") or "") in {"group", "supergroup"}
        and text.split(" ", 1)[0].split("@", 1)[0] == "/chat"
    )
    if allowed and actual != allowed and not is_group_lookup:
        raise ConfigError(f"Неразрешенный Telegram chat_id: {actual}")


def _parse_update(event: dict[str, Any]) -> dict[str, Any] | None:
    body = event.get("body")
    if body is None and "messages" in event:
        return None
    _enforce_relay_secret(event, body)
    if event.get("isBase64Encoded"):
        body = base64.b64decode(body).decode("utf-8")
    if not body:
        return None
    return json.loads(body)


def _parse_control_request(event: dict[str, Any]) -> dict[str, Any] | None:
    """Accept only a relay-authenticated service request from Apps Script."""
    body = event.get("body")
    if not body:
        return None
    _enforce_relay_secret(event, body)
    if event.get("isBase64Encoded"):
        body = base64.b64decode(body).decode("utf-8")
    try:
        payload = json.loads(body)
    except (TypeError, ValueError):
        return None
    if isinstance(payload, dict) and payload.get("mode") in {"send_notification", "send_notifications", "weekly_reflection_prompt"}:
        return payload
    return None


def _enforce_relay_secret(event: dict[str, Any], body: Any) -> None:
    secret = os.getenv("RELAY_SECRET", "")
    if not secret or body is None:
        return
    headers = event.get("headers") or {}
    provided = ""
    for key, value in headers.items():
        if str(key).lower() == "x-relay-secret":
            provided = str(value)
            break
    if not hmac.compare_digest(provided, secret):
        raise ForbiddenError("Неверный relay secret")


def _timer_payload_mode(event: dict[str, Any]) -> str:
    for message in event.get("messages") or []:
        details = message.get("details") or {}
        payload = details.get("payload") or message.get("payload")
        if isinstance(payload, str) and payload.strip():
            try:
                parsed = json.loads(payload)
            except ValueError:
                continue
            if isinstance(parsed, dict):
                return str(parsed.get("mode") or "")
    return ""


def _is_timer_event(event: dict[str, Any]) -> bool:
    messages = event.get("messages")
    if isinstance(messages, list) and messages:
        return True
    return event.get("event_metadata", {}).get("event_type") == "yandex.cloud.events.serverless.triggers.TimerMessage"


def _split_command(text: str) -> tuple[str, str]:
    command, _, payload = text.partition(" ")
    return command.split("@", 1)[0], payload.strip()


def _rows_as_dicts(sheet_name: str) -> list[dict[str, str]]:
    values = sheets_get_values(sheet_name, "A11:Z500")
    if len(values) < 2:
        return []
    headers = values[0]
    rows = []
    for row in values[1:]:
        padded = row + [""] * (len(headers) - len(row))
        if any(str(cell).strip() for cell in padded):
            rows.append({headers[index]: str(padded[index]).strip() for index in range(len(headers))})
    return rows


def _setting_value(key: str, default: str = "") -> str:
    for row in _rows_as_dicts("Настройки"):
        if row.get("Ключ") == key:
            return row.get("Значение") or default
    return default


def _setting_int(key: str, default: int = 0) -> int:
    try:
        return int(_setting_value(key, str(default)) or default)
    except ValueError:
        return default


def _incoming_reminder_days() -> int:
    return max(1, _setting_int("входящие_напоминание_дней", 1))


def _spreadsheet_url() -> str:
    return f"https://docs.google.com/spreadsheets/d/{_spreadsheet_id()}/edit"


def _incoming_sheet_url() -> str:
    return f"{_spreadsheet_url()}#gid={_sheet_id_by_title('Входящие')}"


def _sheet_id_by_title(title: str) -> int:
    metadata = sheets_get_metadata()
    for sheet in metadata.get("sheets", []):
        properties = sheet.get("properties") or {}
        if properties.get("title") == title:
            return int(properties["sheetId"])
    raise ConfigError(f"Не найден лист {title}.")


def _set_setting(key: str, value: str) -> None:
    values = sheets_get_values("Настройки", "A11:D500")
    now = _now().isoformat(timespec="seconds")
    for index, row in enumerate(values[1:], start=12):
        if row and row[0] == key:
            sheets_update_values("Настройки", f"B{index}:D{index}", [[value, _setting_description(key), now]])
            return
    sheets_append_values("Настройки", [[key, value, _setting_description(key), now]])


def _setting_description(key: str) -> str:
    descriptions = {row[0]: row[2] for row in DEFAULT_SETTINGS}
    return descriptions.get(key, "")


def _freeze_header(sheet_id: int) -> None:
    try:
        batch_update([
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount",
                }
            }
        ])
    except requests.HTTPError:
        logger.warning("could not freeze header for sheet_id=%s", sheet_id)


def _extend_limited(
    lines: list[str],
    rows: list[dict[str, str]],
    title_key: str,
    empty_text: str = "Нет записей.",
) -> None:
    if not rows:
        lines.append(empty_text)
        return
    for row in rows[:8]:
        title = row.get(title_key) or "(без названия)"
        due = row.get("Дедлайн")
        suffix = f" до {due}" if due else ""
        lines.append(f"- {title}{suffix}")


def _author_name(user: dict[str, Any]) -> str:
    parts = [user.get("first_name", ""), user.get("last_name", "")]
    name = " ".join(part for part in parts if part).strip()
    return name or user.get("username", "") or str(user.get("id", ""))


def _looks_like_task(text: str) -> bool:
    lowered = text.lower()
    markers = ["надо", "нужно", "сделать", "проверить", "написать", "созвониться", "оплатить", "выставить"]
    return any(marker in lowered for marker in markers)


def _looks_urgent(text: str) -> bool:
    lowered = text.lower()
    markers = ["срочно", "сегодня", "до вечера", "горит", "asap", "urgent"]
    return any(marker in lowered for marker in markers)


def _project_hint(text: str) -> str:
    explicit = re.search(
        r"(?:проект|по проекту|для проекта)\s*[:\-]\s*([^,.;\n]+?)(?:\s+[—-]\s+|$)",
        text,
        re.IGNORECASE,
    )
    if explicit:
        return explicit.group(1).strip()
    lowered = text.lower()
    hints = {
        "карьер": "Карьера",
        "резюме": "Карьера",
        "бот": "Порядок",
        "трекер": "Порядок",
        "telegram": "Порядок",
        "тг": "Порядок",
        "финанс": "Финансы",
        "счет": "Финансы",
        "бюджет": "Финансы",
        "расход": "Финансы",
        "доход": "Финансы",
        "клиент": "Клиенты",
    }
    for marker, value in hints.items():
        if marker in lowered:
            return value
    return ""


def _extract_deadline(text: str) -> str:
    lowered = text.lower()
    today = _now().date()

    relative_days = re.search(r"через\s+(\d+)\s+(?:день|дня|дней)", lowered)
    if relative_days:
        return (today + timedelta(days=int(relative_days.group(1)))).isoformat()

    if re.search(r"\bпослезавтра\b", lowered):
        return (today + timedelta(days=2)).isoformat()
    if re.search(r"\bзавтра\b", lowered):
        return (today + timedelta(days=1)).isoformat()
    if re.search(r"\bсегодня\b", lowered):
        return today.isoformat()

    next_week = re.search(r"\b(?:на\s+)?следующей\s+неделе\b", lowered)
    if next_week:
        return (today + timedelta(days=(7 - today.weekday()))).isoformat()

    weekdays = {
        "понедельник": 0,
        "понедельника": 0,
        "вторник": 1,
        "вторника": 1,
        "среду": 2,
        "среда": 2,
        "среды": 2,
        "четверг": 3,
        "четверга": 3,
        "пятницу": 4,
        "пятница": 4,
        "пятницы": 4,
        "субботу": 5,
        "суббота": 5,
        "субботы": 5,
        "воскресенье": 6,
        "воскресенья": 6,
    }
    for name, weekday in weekdays.items():
        if re.search(rf"\b(?:в|до)\s+{name}\b", lowered):
            delta = (weekday - today.weekday()) % 7
            return (today + timedelta(days=delta or 7)).isoformat()

    numeric = re.search(r"\b(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\b", text)
    if numeric:
        day = int(numeric.group(1))
        month = int(numeric.group(2))
        year = int(numeric.group(3)) if numeric.group(3) else today.year
        if year < 100:
            year += 2000
        try:
            parsed = date(year, month, day)
        except ValueError:
            return ""
        if not numeric.group(3) and parsed < today:
            parsed = date(year + 1, month, day)
        return parsed.isoformat()
    return ""


def _morning_digest_enabled_now() -> bool:
    settings = {row.get("Ключ"): row.get("Значение") for row in _rows_as_dicts("Настройки")}
    if settings.get("утренний_обзор_включен", "да").lower() not in {"да", "yes", "true", "1"}:
        return False
    now = _now()
    hour = int(settings.get("утренний_обзор_час") or "9")
    minute = int(settings.get("утренний_обзор_минута") or "0")
    return now.hour == hour and now.minute == minute


def _hot_threshold_days() -> int:
    settings = {row.get("Ключ"): row.get("Значение") for row in _rows_as_dicts("Настройки")}
    try:
        return int(settings.get("горячая_задача_дней_до_дедлайна") or "2")
    except ValueError:
        return 2


def _parse_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        pass
    match = re.fullmatch(r"\s*(\d{1,2})\s+([а-яё]+)\.?\s*(\d{4})?\s*", str(value).lower())
    if not match:
        return None
    months = {
        "янв": 1, "января": 1, "фев": 2, "февраля": 2, "мар": 3, "марта": 3,
        "апр": 4, "апреля": 4, "мая": 5, "июн": 6, "июня": 6, "июл": 7,
        "июля": 7, "авг": 8, "августа": 8, "сен": 9, "сент": 9, "сентября": 9,
        "окт": 10, "октября": 10, "ноя": 11, "ноября": 11, "дек": 12, "декабря": 12,
    }
    month = months.get(match.group(2).rstrip("."))
    if not month:
        return None
    try:
        return date(int(match.group(3) or _now().year), month, int(match.group(1)))
    except ValueError:
        return None


def _now() -> datetime:
    return datetime.now(SAMARA_TZ)


def _today_iso() -> str:
    return _now().date().isoformat()


def _json_response(payload: dict[str, Any], status_code: int = 200) -> dict[str, Any]:
    return {
        "statusCode": status_code,
        "headers": {"Content-Type": "application/json; charset=utf-8"},
        "body": json.dumps(payload, ensure_ascii=False),
    }


def _telegram_reply(chat_id: str, text: str, reply_markup: dict[str, Any] | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "method": "sendMessage",
        "chat_id": chat_id,
        "text": text[:3900],
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    if reply_markup:
        payload["reply_markup"] = reply_markup
    return _json_response(payload)


def _telegram_callback_reply(callback_query_id: str, text: str, show_alert: bool = False) -> dict[str, Any]:
    return _json_response({
        "method": "answerCallbackQuery",
        "callback_query_id": callback_query_id,
        "text": text[:200],
        "show_alert": show_alert,
    })


def _public_error_message(exc: Exception) -> str:
    if isinstance(exc, (ConfigError, ForbiddenError)):
        return str(exc)
    return "Внутренняя ошибка функции"
