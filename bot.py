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
import base64
import json
import os
import re
from io import BytesIO

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
    "project_title": "<string, default 'სახანძრო სიგნალიზაციის მიწოდება და მონტაჟი'>",
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
- "manual": false items are calculated from cost x coefficient x transport x profit. The coefficient is just a multiplier the user gives (e.g. 2.75) — it converts cost to GEL. NEVER ask what currency the cost is in or what the coefficient represents. Just multiply.
- "manual": true items have a fixed final GEL price given directly (typical for cable per meter, pipe, custom services). Then set "price" and omit cost_usd/exchange_rate/transport_pct/profit_pct.
- If the user gives one exchange rate / transport % / profit % (a global markup), apply it to every non-manual item.
- If the user says a markup like "მოგება 25%" or "25% margin", that is profit_pct. "ტრანსპორტი" is transport_pct. "კურსი" is exchange_rate.
- Installation prices: "მონტაჟი X ლარი" per item, or a global installation price if stated for all. Cable installation price goes to install_price of the cable line with manual pricing for the material if a GEL price is given.
- "სახარჯი მასალები" percentage -> consumables_pct.
- Ask (need_info) only for truly essential missing data: client name and invoice number. Do NOT ask about optional things — default transport_pct/profit_pct/install_price/consumables_pct/exchange_rate to 0 if unstated.
- NEVER ask about currency (USD, EUR, etc.) or what the coefficient means. The user gives a number — use it as exchange_rate directly.
- If the user provides a project title or description (e.g. "სათაური: ...", "პროექტი: ..."), put it in project_title. Otherwise use the default.
- CRITICAL: when asking need_info, list ALL missing items in ONE single message. Never ask questions one at a time across multiple turns.
- CRITICAL: before asking anything, re-read the ENTIRE conversation. Never ask for information the user has already provided in any earlier message or file caption. If the user already answered, use their answer.
- If after one need_info round something minor is still unclear, make a reasonable assumption and proceed rather than asking again.
- Quantities and prices must be numbers, never strings.
- Never invent products, quantities or prices that were not stated.

SUPPLIER PRICE LISTS (PDF, Excel or CSV):
- The user may upload a supplier price list as a PDF document or as extracted spreadsheet text (Excel/CSV, tab-separated). Extract only the products the user asks for, or all listed products if they say to include everything, with quantities from the user's message.
- If supplier prices are in USD, use them as cost_usd (manual=false) with the exchange_rate/transport/profit the user gives.
- If supplier prices are in GEL and the user gives a markup %, compute the final unit price yourself: price = supplier_price * (1 + markup/100), rounded to 2 decimals, and output the item as manual=true with that price.
- Per-product markups override the global one ("კაბელებზე 15%, დანარჩენზე 10%").
- Installation prices come from the user's message, per product or per group.
- Copy product names from the PDF accurately (keep type/size markings like JE-H(St)H 1x2x0.8).
- After receiving a PDF, if the user has not yet said quantities, markup or installation prices, reply need_info asking for them in one short Georgian question."""

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
    "📄 ასევე შეგიძლიათ ატვირთოთ მომწოდებლის პრაისი PDF, Excel (.xlsx) ან CSV ფორმატში — "
    "შემდეგ მომწერეთ რომელი პროდუქტები გჭირდებათ, რაოდენობები, "
    "მარკაპი (მაგ. 10% ან 15%, ან სხვადასხვა პროდუქტზე სხვადასხვა) "
    "და მონტაჟის ფასები.\n\n"
    "ნებისმიერი პროდუქტი, რომელიც სიაში არ არის, უბრალოდ აღწერეთ ტექსტში "
    "ფასთან ერთად და დაემატება.\n\n"
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


def tg_download_file(file_id):
    """Download a file sent to the bot; returns raw bytes or None."""
    r = requests.get(f"{TG_API}/getFile", params={"file_id": file_id}, timeout=30)
    r.raise_for_status()
    file_path = r.json()["result"]["file_path"]
    f = requests.get(
        f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}", timeout=60)
    f.raise_for_status()
    return f.content


def excel_to_text(file_bytes):
    """Convert an .xlsx price list into plain text (tab-separated) for Claude."""
    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    lines = []
    for ws in wb.worksheets:
        lines.append(f"=== Sheet: {ws.title} ===")
        for row in ws.iter_rows(values_only=True):
            if any(v is not None and str(v).strip() for v in row):
                lines.append("\t".join("" if v is None else str(v) for v in row))
    wb.close()
    text = "\n".join(lines)
    return text[:60000]  # safety cap


# ---------------------------------------------------------------- Claude parse

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = "gemini-2.5-flash"


def extract_pdf_text_local(pdf_bytes):
    """FREE: pull the text layer out of a digital PDF. Returns '' for scanned PDFs."""
    import pdfplumber
    parts = []
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages[:40]:
            parts.append(page.extract_text() or "")
    return "\n".join(parts).strip()[:60000]


def extract_pricelist_gemini(pdf_bytes):
    """FREE tier: Gemini reads scanned/complex PDFs at no cost."""
    resp = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent",
        params={"key": GEMINI_API_KEY},
        json={
            "contents": [{
                "parts": [
                    {"inline_data": {
                        "mime_type": "application/pdf",
                        "data": base64.b64encode(pdf_bytes).decode()}},
                    {"text": "This is a supplier price list. Extract EVERY product as a "
                             "tab-separated table: name<TAB>unit<TAB>price<TAB>currency. "
                             "Copy names exactly, including type/size markings. "
                             "Output ONLY the table, no commentary."},
                ],
            }],
        },
        timeout=180,
    )
    resp.raise_for_status()
    data = resp.json()
    return "".join(
        p.get("text", "")
        for p in data["candidates"][0]["content"]["parts"]
    ).strip()


def extract_pricelist_from_pdf(pdf_bytes):
    """One-time extraction: PDF -> compact text table. Keeps follow-up turns fast."""
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        json={
            "model": CLAUDE_MODEL,
            "max_tokens": 8000,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "document",
                     "source": {"type": "base64",
                                "media_type": "application/pdf",
                                "data": base64.b64encode(pdf_bytes).decode()}},
                    {"type": "text",
                     "text": "This is a supplier price list. Extract EVERY product as a "
                             "tab-separated table: name<TAB>unit<TAB>price<TAB>currency. "
                             "Copy names exactly, including type/size markings. "
                             "Output ONLY the table, no commentary."},
                ],
            }],
        },
        timeout=180,
    )
    resp.raise_for_status()
    return "".join(
        b.get("text", "") for b in resp.json().get("content", [])
        if b.get("type") == "text"
    ).strip()


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
            "system": [{"type": "text", "text": SYSTEM_PROMPT,
                        "cache_control": {"type": "ephemeral"}}],
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

    document = message.get("document")
    if document:
        mime = document.get("mime_type", "")
        fname = (document.get("file_name") or "").lower()
        size = document.get("file_size", 0)

        is_pdf = mime == "application/pdf" or fname.endswith(".pdf")
        is_xlsx = (mime == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                   or fname.endswith(".xlsx"))
        is_csv = mime in ("text/csv", "text/plain") or fname.endswith((".csv", ".txt"))

        if fname.endswith(".xls") and not is_xlsx:
            tg_send_text(chat_id, "ძველი .xls ფორმატს ვერ ვკითხულობ — "
                                  "გთხოვთ შეინახოთ როგორც .xlsx და ისე ატვირთოთ.")
            return jsonify(ok=True)
        if not (is_pdf or is_xlsx or is_csv):
            tg_send_text(chat_id, "მიღებული ფორმატები: PDF, Excel (.xlsx) და CSV.")
            return jsonify(ok=True)
        if size > 15 * 1024 * 1024:
            tg_send_text(chat_id, "ფაილი ძალიან დიდია (მაქს. 15 MB).")
            return jsonify(ok=True)

        try:
            file_bytes = tg_download_file(document["file_id"])
        except Exception:
            app.logger.exception("File download failed")
            tg_send_text(chat_id, "ფაილის ჩამოტვირთვა ვერ მოხერხდა, სცადეთ თავიდან.")
            return jsonify(ok=True)

        caption = (message.get("caption") or "").strip()
        default_note = "მომწოდებლის პრაისი ავტვირთე. (supplier price list uploaded)"

        if is_pdf:
            tg_send_text(chat_id, "📄 ფაილი მივიღე, ვკითხულობ პრაისს...")
            sheet_text = ""
            # 1) FREE: local text layer (works for most digital price lists)
            try:
                sheet_text = extract_pdf_text_local(file_bytes)
            except Exception:
                app.logger.exception("Local PDF extraction failed")
            # 2) FREE: Gemini for scanned / image-only PDFs
            if len(sheet_text) < 200 and GEMINI_API_KEY:
                try:
                    sheet_text = extract_pricelist_gemini(file_bytes)
                except Exception:
                    app.logger.exception("Gemini extraction failed")
            # 3) PAID fallback: Claude reads the PDF
            if len(sheet_text) < 200:
                try:
                    sheet_text = extract_pricelist_from_pdf(file_bytes)
                except Exception:
                    app.logger.exception("Claude PDF extraction failed")
            if len(sheet_text) < 20:
                tg_send_text(chat_id, "PDF-ის წაკითხვა ვერ მოხერხდა, სცადეთ თავიდან "
                                      "ან ატვირთეთ Excel ვერსია.")
                return jsonify(ok=True)
            content = (f"Supplier price list ({fname}), extracted contents:\n"
                       f"{sheet_text}\n\n"
                       f"User note: {caption or default_note}")
        else:
            try:
                if is_xlsx:
                    sheet_text = excel_to_text(file_bytes)
                else:
                    sheet_text = file_bytes.decode("utf-8-sig", errors="replace")[:60000]
            except Exception:
                app.logger.exception("Spreadsheet read failed")
                tg_send_text(chat_id, "ფაილის წაკითხვა ვერ მოხერხდა — "
                                      "დარწმუნდით რომ სწორი Excel/CSV ფაილია.")
                return jsonify(ok=True)
            content = (f"Supplier price list ({fname}), extracted contents:\n"
                       f"{sheet_text}\n\n"
                       f"User note: {caption or default_note}")

        history = chat_history.setdefault(chat_id, [])
        history.append({"role": "user", "content": content})
        del history[:-MAX_TURNS]
        if not is_pdf:
            tg_send_text(chat_id, "📄 ფაილი მივიღე, ვამუშავებ...")
        _process_and_reply(chat_id, history)
        return jsonify(ok=True)

    if not text:
        tg_send_text(chat_id, "გთხოვთ გამომიგზავნოთ ტექსტური შეტყობინება ან PDF ფაილი.")
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
    _process_and_reply(chat_id, history)
    return jsonify(ok=True)


def _process_and_reply(chat_id, history):
    try:
        result = parse_with_claude(history)
    except Exception as e:
        app.logger.exception("Claude parse failed")
        tg_send_text(chat_id, f"ვერ დავამუშავე მოთხოვნა ({type(e).__name__}). "
                              "სცადეთ თავიდან ან /clear.")
        return

    if result.get("status") == "need_info":
        question = result.get("message", "დამატებითი ინფორმაცია მჭირდება.")
        history.append({"role": "assistant", "content": json.dumps(result, ensure_ascii=False)})
        tg_send_text(chat_id, question)
        return

    if result.get("status") == "ok" and result.get("invoice"):
        try:
            file_obj, filename = build_invoice_xlsx(result["invoice"])
        except Exception as e:
            app.logger.exception("Excel generation failed")
            tg_send_text(chat_id, f"Excel-ის გენერაცია ჩაიშალა ({type(e).__name__}).")
            return

        inv = result["invoice"]
        summary = (f"✅ ინვოისი {inv.get('invoice_number', '')} — "
                   f"{inv.get('client_name', '')}, "
                   f"{len(inv.get('items', []))} პოზიცია.")
        tg_send_text(chat_id, summary)
        tg_send_document(chat_id, file_obj, filename)
        chat_history.pop(chat_id, None)   # invoice done — fresh start
        return

    tg_send_text(chat_id, "გაუგებარი პასუხი მივიღე. სცადეთ თავიდან ან /clear.")


@app.route("/", methods=["GET"])
def health():
    return "Gerner invoice bot is running."


if __name__ == "__main__":
    app.run(port=8080, debug=True)
