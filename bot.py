#!/usr/bin/env python3
"""Bot Telegram: carica un export CSV della banca o aggiunge una spesa a mano.

Uso:
    python bot.py

- Invia al bot il file CSV: lo carica sul foglio, come `report.py`.
- /spesa (oppure un messaggio come "12,50 kebab") aggiunge una spesa non
  collegata al conto, ad esempio in contanti.
"""
from __future__ import annotations

import asyncio
import logging
import re
import sys
import tempfile
import uuid
import warnings
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    ApplicationHandlerStop,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    TypeHandler,
    filters,
)
from telegram.warnings import PTBUserWarning

import report

log = logging.getLogger("bot")

AMOUNT, NAME, CATEGORY, ACCOUNT, DATE = range(5)

# "12,50 kebab", "€ 12.5 kebab", "12 kebab": importo seguito dal nome della spesa.
AMOUNT_RE = r"€?\s*(\d+(?:[.,]\d{1,2})?)\s*€?"
QUICK_RE = re.compile(rf"^\s*{AMOUNT_RE}\s+(\S.*)$", re.DOTALL)
ONLY_AMOUNT_RE = re.compile(rf"^\s*{AMOUNT_RE}\s*$")
DATE_RE = re.compile(r"^\s*(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2,4}))?(?:\s+(\d{1,2})[:.](\d{2}))?\s*$")

MAX_MESSAGE = 4000  # Telegram accetta al massimo 4096 caratteri per messaggio

HELP = (
    "Cosa posso fare:\n"
    "• Inviami il file CSV esportato dalla banca: carico le transazioni sul foglio.\n"
    "• /spesa per aggiungere una spesa a mano (es. in contanti).\n"
    '• Più veloce: scrivi importo e nome, ad esempio "12,50 kebab".\n'
    "• /annulla per interrompere l'inserimento."
)

# Una sola scrittura sul foglio alla volta.
sheet_lock = asyncio.Lock()


# --- Lettura di importi e date -----------------------------------------------


def parse_amount(text: str) -> float:
    return float(text.replace(",", "."))


def parse_date(text: str, now: datetime) -> datetime | None:
    """ "14/11", "14/11/2026" o "14/11 13:30". Senza anno vale l'anno in corso, senza ora le 12:00."""
    match = DATE_RE.match(text)
    if not match:
        return None
    day, month, year, hour, minute = match.groups()
    year = now.year if year is None else int(year) + (2000 if len(year) == 2 else 0)
    try:
        return datetime(year, int(month), int(day), int(hour or 12), int(minute or 0))
    except ValueError:
        return None


def now_local(cfg: dict) -> datetime:
    return datetime.now(ZoneInfo(cfg.get("fuso_orario", "Europe/Rome"))).replace(tzinfo=None, microsecond=0)


def euro(amount: float) -> str:
    return f"{amount:.2f}".replace(".", ",") + " €"


def buttons(prefix: str, labels: list[str], per_row: int = 2) -> InlineKeyboardMarkup:
    keys = [InlineKeyboardButton(label, callback_data=f"{prefix}:{i}") for i, label in enumerate(labels)]
    return InlineKeyboardMarkup([keys[i : i + per_row] for i in range(0, len(keys), per_row)])


# --- Caricamento del CSV -----------------------------------------------------


async def on_csv(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    document = update.message.document
    await update.message.reply_text("Ricevuto, carico le transazioni sul foglio…")
    lines: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "export.csv"
        try:
            await (await document.get_file()).download_to_drive(path)
            cfg, rules = report.load_settings()
            async with sheet_lock:
                await asyncio.to_thread(report.sync_csv, path, cfg, rules, False, lines.append)
        except report.ReportError as error:
            await update.message.reply_text(f"Non caricato: {error}")
            return
        except (KeyError, ValueError, UnicodeDecodeError) as error:
            log.exception("CSV non valido")
            await update.message.reply_text(
                f"Non caricato: il file non sembra un export della banca nel formato atteso ({error!r})."
            )
            return
        except Exception as error:
            log.exception("Caricamento fallito")
            await update.message.reply_text(f"Qualcosa è andato storto durante il caricamento: {error!r}")
            return
    await update.message.reply_text("\n".join(lines).strip()[:MAX_MESSAGE] or "Fatto.")


async def on_other_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("Mi serve il file .csv esportato dalla banca.")


# --- Inserimento di una spesa a mano -----------------------------------------


async def ask_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    _cfg, rules = report.load_settings()
    categories = rules.get("categorie_spese", [])
    context.user_data["categories"] = categories
    expense = context.user_data["expense"]
    await update.effective_message.reply_text(
        f"{euro(expense['amount'])} · {expense['name']}\nCategoria?", reply_markup=buttons("cat", categories)
    )
    return CATEGORY


async def start_expense(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["expense"] = {}
    await update.message.reply_text("Quanto hai speso? (es. 12,50)")
    return AMOUNT


async def quick_expense(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Messaggio "12,50 kebab": importo e nome in un colpo solo."""
    amount, name = QUICK_RE.match(update.message.text).groups()
    context.user_data["expense"] = {"amount": parse_amount(amount), "name": name.strip()}
    return await ask_category(update, context)


async def on_amount(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text
    if match := QUICK_RE.match(text):
        return await quick_expense(update, context)
    match = ONLY_AMOUNT_RE.match(text)
    if not match or parse_amount(match.group(1)) <= 0:
        await update.message.reply_text("Non ho capito l'importo. Scrivi solo il numero, ad esempio 12,50")
        return AMOUNT
    context.user_data["expense"]["amount"] = parse_amount(match.group(1))
    await update.message.reply_text("Che spesa è? (il nome che vuoi vedere sul foglio)")
    return NAME


async def on_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["expense"]["name"] = update.message.text.strip()
    return await ask_category(update, context)


async def on_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    category = context.user_data["categories"][int(query.data.split(":")[1])]
    context.user_data["expense"]["category"] = category
    await query.edit_message_text(f"Categoria: {category}")

    cfg, _rules = report.load_settings()
    accounts = cfg.get("conti_manuali", ["Contanti", "N26"])
    context.user_data["accounts"] = accounts
    if len(accounts) == 1:
        context.user_data["expense"]["account"] = accounts[0]
        return await ask_date(update, context)
    await query.message.reply_text("Con cosa hai pagato?", reply_markup=buttons("acc", accounts))
    return ACCOUNT


async def on_account(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    account = context.user_data["accounts"][int(query.data.split(":")[1])]
    context.user_data["expense"]["account"] = account
    await query.edit_message_text(f"Pagato con: {account}")
    return await ask_date(update, context)


async def ask_date(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.effective_message.reply_text(
        "Quando? Scegli qui sotto oppure scrivi la data (es. 14/11 oppure 14/11 13:30).",
        reply_markup=buttons("day", ["Adesso", "Ieri"]),
    )
    return DATE


async def on_day_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    cfg, _rules = report.load_settings()
    now = now_local(cfg)
    when = now if query.data == "day:0" else (now - timedelta(days=1)).replace(hour=12, minute=0, second=0)
    await query.edit_message_text("Quando: " + ("adesso" if query.data == "day:0" else "ieri"))
    return await save_expense(update, context, when)


async def on_date_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    cfg, _rules = report.load_settings()
    when = parse_date(update.message.text, now_local(cfg))
    if when is None:
        await update.message.reply_text("Non ho capito la data. Scrivila come 14/11, 14/11/2026 oppure 14/11 13:30")
        return DATE
    return await save_expense(update, context, when)


async def save_expense(update: Update, context: ContextTypes.DEFAULT_TYPE, when: datetime) -> int:
    expense = context.user_data.pop("expense")
    message = update.effective_message
    try:
        cfg, rules = report.load_settings()
        async with sheet_lock:
            tab = await asyncio.to_thread(
                report.add_manual_expense,
                cfg,
                rules,
                when,
                expense["amount"],
                expense["name"],
                expense["category"],
                expense["account"],
                f"telegram-{uuid.uuid4().hex[:12]}",
            )
    except report.ReportError as error:
        await message.reply_text(f"Non salvata: {error}")
        return ConversationHandler.END
    except Exception as error:
        log.exception("Salvataggio fallito")
        await message.reply_text(f"Non salvata, qualcosa è andato storto: {error!r}")
        return ConversationHandler.END
    await message.reply_text(
        f"✅ Salvata in «{tab}»\n"
        f"{euro(expense['amount'])} · {expense['name']} · {expense['category']} · "
        f"{expense['account']} · {when:%d/%m/%Y %H:%M}"
    )
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.pop("expense", None)
    await update.message.reply_text("Annullato, non ho salvato nulla.")
    return ConversationHandler.END


# --- Resto -------------------------------------------------------------------


async def help_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(HELP)


async def expired_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.callback_query.answer("Questo bottone non è più valido: ricomincia con /spesa", show_alert=True)


def make_gate(allowed_users: set[int]):
    """Controllo eseguito prima di ogni altro: fa passare solo gli utenti autorizzati."""

    async def gate(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        user = update.effective_user
        if user is not None and user.id in allowed_users:
            return
        if user is not None:
            # Chi non è autorizzato riceve solo il proprio ID, utile per la prima configurazione.
            log.warning("Messaggio da utente non autorizzato: %s", user.id)
            text = (
                f"Non sei autorizzato a usare questo bot.\nIl tuo ID Telegram è {user.id}: "
                "se il bot è tuo, aggiungilo a telegram_utenti in config.toml e riavvialo."
            )
            if update.callback_query is not None:
                await update.callback_query.answer(text, show_alert=True)
            elif update.effective_message is not None:
                await update.effective_message.reply_text(text)
        raise ApplicationHandlerStop

    return gate


def build_application(token: str, allowed_users: list[int]) -> Application:
    application = Application.builder().token(token).build()
    text = filters.TEXT & ~filters.COMMAND

    # Gruppo -1: viene prima di tutti gli altri gestori e può fermare l'aggiornamento.
    application.add_handler(TypeHandler(Update, make_gate(set(allowed_users))), group=-1)

    with warnings.catch_warnings():
        # Avviso noto: mescoliamo bottoni e messaggi di testo nella stessa conversazione.
        warnings.simplefilter("ignore", PTBUserWarning)
        conversation = ConversationHandler(
            entry_points=[
                CommandHandler("spesa", start_expense),
                MessageHandler(text & filters.Regex(QUICK_RE), quick_expense),
            ],
            states={
                AMOUNT: [MessageHandler(text, on_amount)],
                NAME: [MessageHandler(text, on_name)],
                CATEGORY: [CallbackQueryHandler(on_category, pattern=r"^cat:\d+$")],
                ACCOUNT: [CallbackQueryHandler(on_account, pattern=r"^acc:\d+$")],
                DATE: [
                    CallbackQueryHandler(on_day_button, pattern=r"^day:\d+$"),
                    MessageHandler(text, on_date_text),
                ],
            },
            fallbacks=[CommandHandler("annulla", cancel)],
        )
    application.add_handler(conversation)
    application.add_handler(CallbackQueryHandler(expired_button))
    application.add_handler(MessageHandler(filters.Document.FileExtension("csv"), on_csv))
    application.add_handler(MessageHandler(filters.Document.ALL, on_other_document))
    application.add_handler(CommandHandler("annulla", cancel))
    application.add_handler(MessageHandler(filters.ALL, help_message))
    return application


def main() -> None:
    logging.basicConfig(format="%(asctime)s %(levelname)s %(name)s: %(message)s", level=logging.INFO)
    logging.getLogger("httpx").setLevel(logging.WARNING)  # altrimenti il token finisce nei log
    try:
        cfg, _rules = report.load_settings()
    except report.ReportError as error:
        sys.exit(str(error))
    token = cfg.get("telegram_token")
    if not token:
        sys.exit("Manca telegram_token in config.toml (vedi README, sezione Bot Telegram).")
    allowed_users = [int(user) for user in cfg.get("telegram_utenti", [])]
    if not allowed_users:
        log.warning("telegram_utenti è vuoto: scrivi al bot per farti dire il tuo ID, poi aggiungilo a config.toml.")
    build_application(token, allowed_users).run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
