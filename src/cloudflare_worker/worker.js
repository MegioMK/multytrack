export default {
  async fetch(request, env) {
    if (request.method !== "POST") {
      return new Response("Telegram task tracker relay is running", { status: 200 });
    }

    const providedRelaySecret = request.headers.get("X-Relay-Secret") || "";
    if (env.RELAY_SECRET && providedRelaySecret === env.RELAY_SECRET) {
      return sendTelegramAction(request, env);
    }

    const expectedTelegramSecret = env.TELEGRAM_WEBHOOK_SECRET || "";
    if (expectedTelegramSecret) {
      const providedTelegramSecret = request.headers.get("X-Telegram-Bot-Api-Secret-Token") || "";
      if (providedTelegramSecret !== expectedTelegramSecret) {
        return new Response("forbidden", { status: 403 });
      }
    }

    const endpoint = env.YANDEX_FUNCTION_URL || env.YANDEX_GATEWAY_URL;
    if (!endpoint) {
      return json({ ok: false, error: "missing Yandex endpoint" }, 500);
    }

    const body = await request.text();
    const headers = {
      "Content-Type": request.headers.get("Content-Type") || "application/json",
    };

    if (env.RELAY_SECRET) {
      headers["X-Relay-Secret"] = env.RELAY_SECRET;
    }

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers,
        body,
      });
      const responseBody = await response.text();
      return new Response(responseBody, {
        status: response.status,
        headers: {
          "Content-Type": response.headers.get("Content-Type") || "application/json; charset=utf-8",
        },
      });
    } catch (error) {
      return json({ ok: false, error: "relay failed" }, 502);
    }
  },
};

async function sendTelegramAction(request, env) {
  let payload;
  try {
    payload = await request.json();
  } catch (error) {
    return json({ ok: false, error: "invalid notification payload" }, 400);
  }
  if (!env.TELEGRAM_BOT_TOKEN) {
    return json({ ok: false, error: "missing Telegram token" }, 500);
  }

  const isNotification = payload?.mode === "send_notification" && payload.chat_id && payload.text;
  const isReaction = payload?.mode === "set_reaction" && payload.chat_id && payload.message_id;
  const isEdit = payload?.mode === "edit_message" && payload.chat_id && payload.message_id && payload.text;
  if (!isNotification && !isReaction && !isEdit) {
    return json({ ok: false, error: "invalid Telegram action" }, 400);
  }

  const method = isReaction ? "setMessageReaction" : (isEdit ? "editMessageText" : "sendMessage");
  const body = isReaction
    ? {
        chat_id: String(payload.chat_id),
        message_id: Number(payload.message_id),
        reaction: [{ type: "emoji", emoji: "✅" }],
      }
    : {
        chat_id: String(payload.chat_id),
        ...(isEdit ? { message_id: Number(payload.message_id) } : {}),
        text: String(payload.text).slice(0, 3900),
        parse_mode: "HTML",
        disable_web_page_preview: true,
        ...(payload.reply_markup ? { reply_markup: payload.reply_markup } : {}),
      };

  try {
    const response = await fetch(`https://api.telegram.org/bot${env.TELEGRAM_BOT_TOKEN}/${method}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const responseBody = await response.text();
    return new Response(responseBody, {
      status: response.status,
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  } catch (error) {
    return json({ ok: false, error: "Telegram relay failed" }, 502);
  }
}

function json(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}
