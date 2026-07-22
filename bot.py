# -*- coding: utf-8 -*-
"""
Telegram → Claude → Excel invoice bot for Gerner LTD.

Flow:
  1. Partner writes a free-text message (Georgian or English) describing the invoice.
  2. Claude parses it into the JSON payload that invoice_generator expects.
  3. The .xlsx is generated and sent back in the same chat.

Environment variables (required):
  TELEGRAM_TOKEN     - token from @BotFather
  ANTHROPIC_API_KEY  - key from console.anthropic.com
Optional:
  WEBHOOK_SECRET     - secret_token you pass to setWebhook (recommended)
  ALLOWED_USER_IDS   - comma-separated Telegram user IDs allowed to use the bot.
                       Empty = everyone (not recommended).

Run locally:   python bot.py          (then use a tunnel like ngrok for the webhook)
Run on Render: gunicorn bot:app
"""
import json
import os
import re

import requests
from flask import Flask, request, jsonify

from invoice_generator import build_invoice_xlsx, PRODUCT_LIBRARY

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "")
ALLOWED_USER_IDS = {
    int(x) for x in os.environ.get("ALLOWED_USER_IDS", "").split(",") if x.strip().isdigit()
}

TG_API = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
CLAUDE_MODEL = "claude-sonnet-4-6"

app = Flask(__name__)

# In-memory conversation buffer per chat (survives until restart; enough for
# "the bot asked me a follow-up question" flows).
chat_history = {}   # chat_id -> [{"role": "...", "content": "..."}]
MAX_TURNS = 12

PRODUCT_LIST_TEXT = "\n".join(
    f'- id "{pid}": {p["name"]} (default cost {p["cost_usd"]} USD)'
    for pid, p in PRODUCT_LIBRARY.items()
)

SYSTEM_PROMPT = f"""You convert free-text invoice requests (usually in Georgian) from employees of
a fire-safety company into a strict JSON payload for an Excel invoice generator.

KNOWN PRODUCT LIBRARY (addressable fire alarm equipment):
{PRODUCT_LIST_TEXT}

OUTPUT FORMAT — respond with ONLY a JSON object, no markdown fences, no commentary.
One of two shapes:

1) When you have enough information:
{{
  "status": "ok",
  "invoice": {{
    "client_name": "<string>",
    "invoice_number": "<string>",
    "consumables_pct": <number, default 0>,
    "items": [
      {{
        "product_id": "<library id, only if the item matches the library>",
        "name": "<Georgian name; required if no product_id>",
        "unit": "<default 'ცალი'; use 'მ' or 'გრძ/მ' for cable/pipe>",
        "qty": <number>,
        "manual": <true|false>,
        "cost_usd": <number; USD supplier cost; include to override library default>,
        "exchange_rate": <number, USD->GEL>,
        "transport_pct": <number>,
        "profit_pct": <number>,
        "price": <number; final unit price in GEL, only when manual=true>,
        "install_price": <number; installation price per unit in GEL, default 0>
      }}
    ]
  }}
}}

2) When something essential is missing or ambiguous:
{{"status": "need_info", "message": "<short question in Georgian>"}}

RULES:
- Match products to the library by meaning, not exact wording ("კვამლის დეტექტორი", "smoke detector", "დეტექტორი" -> smoke_det, etc.). If matched, set product_id and omit name.
- "manual": false items are calculated from USD cost x exchange rate x transport x profit. Use this whenever a USD cost is known (from the library or from the message).
- "manual": true items have a fixed final GEL price given directly (typical for cable per meter, pipe, custom services). Then set "price" and omit cost_usd/exchange_rate/transport_pct/profit_pct.
- If the user gives one exchange rate / transport % / profit % (a global markup), apply it to every non-manual item.
- If the user says a markup like "მოგება 25%" or "25% margin", that is profit_pct. "ტრანსპორტი" is transport_pct. "კურსი" is exchange_rate.
- Installation prices: "მონტაჟი X ლარი" per item, or a global installation price if stated for all. Cable installation price goes to install_price of the cable line with manual pricing for the material if a GEL price is given.
- "სახარჯი მასალები" percentage -> consumables_pct.
- Ask (need_info) only for truly essential missing data: client name, invoice number, or exchange rate when non-manual items exist and no rate was given. Do NOT ask about optional things — default transport_pct/profit_pct/install_price/consumables_pct to 0 if unstated but mention nothing.
- Quantities and prices must be numbers, never strings.
- Never invent products, quantities or prices that were not stated."""

HELP_TEXT = (
    "გამარჯობა! მე ვარ გერნერის ინვოისების ბოტი. 🧾\n\n"
    "მომწერეთ ერთი შეტყობინებით:\n"
    "• დამკვეთი და ინვოისის ნომერი\n"
    "• პროდუქტები რაოდენობებით\n"
    "• კურსი, ტრანსპორტის % და მოგების %\n"
    "• მონტაჟის ფასები და სახარჯი მასალების %\n\n"
    "მაგალითი:\n"
    "დამკვეთი შპს ალფა, ინვოისი 2026-014.\n"
    "25 კვამლის დეტექტორი, 3 ღილაკი, 4 სირენა, 1 პანელი.\n"
    "კურსი 2.71, ტრანსპორტი 8%, მოგება 25%.\n"
    "მონტაჟი დეტექტორზე 15 ლარი, დანარჩენზე 20 ლარი.\n"
    "კაბელი JE-H(St)H 1x2x0.8 — 800 მეტრი, 2.5 ლარი მეტრი, მონტაჟი 1.5 ლარი.\n"
    "სახარჯი მასალები 5%.\n\n"
    "ახალი ინვოისის დასაწყებად სუფთა ფურცლიდან გამოიყენეთ /clear"
)


# ---------------------------------------------------------------- Telegram I/O

def tg_send_text(chat_id, text):
    requests.post(f"{TG_API}/sendMessage", json={"chat_id": chat_id, "text": text}, timeout=30)


def tg_send_document(chat_id, file_obj, filename):
    requests.post(
        f"{TG_API}/sendDocument",
        data={"chat_id": chat_id},
        files={"document": (
            filename, file_obj,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        timeout=60,
    )


# ---------------------------------------------------------------- Claude parse

def parse_with_claude(messages):
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": CLAUDE_MODEL,
            "max_tokens": 4000,
            "system": SYSTEM_PROMPT,
            "messages": messages,
        },
        timeout=120,
    )
    resp.raise_for_status()
    text = "".join(
        block.get("text", "") for block in resp.json().get("content", [])
        if block.get("type") == "text"
    )
    # strip accidental code fences and grab the JSON object
    text = re.sub(r"^```(json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    start, end = text.find("{"), text.rfind("}")
    return json.loads(text[start:end + 1])


# ---------------------------------------------------------------- Webhook

@app.route("/webhook", methods=["POST"])
def webhook():
    if WEBHOOK_SECRET and request.headers.get(
            "X-Telegram-Bot-Api-Secret-Token") != WEBHOOK_SECRET:
        return "forbidden", 403

    update = request.get_json(silent=True) or {}
    message = update.get("message") or update.get("edited_message")
    if not message:
        return jsonify(ok=True)

    chat_id = message["chat"]["id"]
    user_id = (message.get("from") or {}).get("id")
    text = (message.get("text") or "").strip()

    if ALLOWED_USER_IDS and user_id not in ALLOWED_USER_IDS:
        tg_send_text(chat_id,
                     f"ამ ბოტზე წვდომა შეზღუდულია. თქვენი ID: {user_id} — "
                     "გადაუგზავნეთ ადმინისტრატორს დასამატებლად.")
        return jsonify(ok=True)

    if not text:
        tg_send_text(chat_id, "გთხოვთ გამომიგზავნოთ ტექსტური შეტყობინება.")
        return jsonify(ok=True)

    if text.startswith("/start") or text.startswith("/help"):
        tg_send_text(chat_id, HELP_TEXT)
        return jsonify(ok=True)

    if text.startswith("/clear"):
        chat_history.pop(chat_id, None)
        tg_send_text(chat_id, "ისტორია გასუფთავდა. შეგიძლიათ ახალი ინვოისის აღწერა. ✅")
        return jsonify(ok=True)

    history = chat_history.setdefault(chat_id, [])
    history.append({"role": "user", "content": text})
    del history[:-MAX_TURNS]

    try:
        result = parse_with_claude(history)
    except Exception as e:
        app.logger.exception("Claude parse failed")
        tg_send_text(chat_id, f"ვერ დავამუშავე მოთხოვნა ({type(e).__name__}). "
                              "სცადეთ თავიდან ან /clear.")
        return jsonify(ok=True)

    if result.get("status") == "need_info":
        question = result.get("message", "დამატებითი ინფორმაცია მჭირდება.")
        history.append({"role": "assistant", "content": json.dumps(result, ensure_ascii=False)})
        tg_send_text(chat_id, question)
        return jsonify(ok=True)

    if result.get("status") == "ok" and result.get("invoice"):
        try:
            file_obj, filename = build_invoice_xlsx(result["invoice"])
        except Exception as e:
            app.logger.exception("Excel generation failed")
            tg_send_text(chat_id, f"Excel-ის გენერაცია ჩაიშალა ({type(e).__name__}).")
            return jsonify(ok=True)

        inv = result["invoice"]
        summary = (f"✅ ინვოისი {inv.get('invoice_number', '')} — "
                   f"{inv.get('client_name', '')}, "
                   f"{len(inv.get('items', []))} პოზიცია.")
        tg_send_text(chat_id, summary)
        tg_send_document(chat_id, file_obj, filename)
        chat_history.pop(chat_id, None)   # invoice done — fresh start
        return jsonify(ok=True)

    tg_send_text(chat_id, "გაუგებარი პასუხი მივიღე. სცადეთ თავიდან ან /clear.")
    return jsonify(ok=True)


@app.route("/", methods=["GET"])
def health():
    return "Gerner invoice bot is running."


if __name__ == "__main__":
    app.run(port=8080, debug=True)
