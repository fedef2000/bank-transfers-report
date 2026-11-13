#!/usr/bin/env python3
"""Carica le transazioni di un export CSV della banca su un foglio Google.

Uso:
    python report.py "Transaction export.csv" --dry-run   # anteprima, non scrive
    python report.py "Transaction export.csv"             # carica sul foglio
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import tomllib
from dataclasses import dataclass
from datetime import datetime, timedelta
from itertools import groupby
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent

HEADER = ["Data", "Importo", "Nome", "Categoria", "Conto", "Descrizione banca", "ID transazione"]
DATE_COL = HEADER.index("Data")
AMOUNT_COL = HEADER.index("Importo")
CATEGORY_COL = HEADER.index("Categoria")
DESCRIPTION_COL = HEADER.index("Descrizione banca")
ID_COL = HEADER.index("ID transazione")
COLUMN_WIDTHS = [130, 100, 170, 160, 130, 420, 280]

# Colonne compilate a mano per correggere quanto pesa una transazione nel riepilogo.
ADJUST_HEADER = ["Importo contato", "Mesi"]
COUNTED_COL = len(HEADER)
SPREAD_COL = COUNTED_COL + 1
DATA_WIDTH = len(HEADER) + len(ADJUST_HEADER)
# Colonne di servizio nascoste, calcolate da formule: il programma non le
# rilegge e non le riscrive mai. Eventuali colonne aggiunte a mano stanno dopo.
HELPER_HEADER = ["Quota mensile", "Mesi contati", "Mesi letti"]
UNREADABLE = "?"  # nella colonna "Mesi letti": il testo scritto in "Mesi" non è stato capito
EXTRA_START = DATA_WIDTH + len(HELPER_HEADER)

SUMMARY_FIXED = ["Mese", "Entrate", "Spese", "Saldo", "% risparmiata", "Patrimonio"]
DEFAULT_ACCOUNTS = ["Investimenti", "Trade Republic", "N26", "Contanti"]
NOTES = "Note"
# Righe lasciate vuote sotto ogni anno del riepilogo, dove stanno i grafici.
CHART_ROWS = 20
CHART_HEIGHT = 380

MONTHS = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre",
]  # fmt: skip
MONTH_TITLE_RE = re.compile(rf"^({'|'.join(MONTHS)}) \d{{4}}$")

# Le date sul foglio sono numeri seriali (giorni dal 30/12/1899): così non
# dipendono dalla lingua del foglio.
SHEETS_EPOCH = datetime(1899, 12, 30)

# Separatore degli argomenti nelle formule: ";" è accettato in ogni lingua.
SEP = ";"

HEADER_BG = {"red": 0.16, "green": 0.22, "blue": 0.33}
WHITE = {"red": 1, "green": 1, "blue": 1}
MONTH_BG = {"red": 0.82, "green": 0.87, "blue": 0.95}
HIGHLIGHT_BG = {"red": 0.99, "green": 0.88, "blue": 0.88}
INPUT_BG = {"red": 1, "green": 0.97, "blue": 0.84}
ERROR_BG = {"red": 0.92, "green": 0.3, "blue": 0.3}
BLACK = {"red": 0, "green": 0, "blue": 0}
GREEN_TEXT = {"red": 0.11, "green": 0.53, "blue": 0.2}
RED_TEXT = {"red": 0.8, "green": 0.15, "blue": 0.15}
EURO_FORMAT = {"type": "NUMBER", "pattern": '#,##0.00 "€"'}
# Come sopra, ma gli zeri restano vuoti: il riepilogo si legge meglio.
EURO_FORMAT_NO_ZERO = {"type": "NUMBER", "pattern": '#,##0.00 "€";-#,##0.00 "€";'}

# Entrate la cui natura è già chiara dal tipo di transazione.
INCOME_BY_TYPE = {
    "INTEREST_PAYMENT": ("Interessi", "Interessi"),
    "DIVIDEND": (None, "Dividendi"),  # il nome è quello del fondo
}


@dataclass
class Record:
    tab: str  # titolo della tab di destinazione
    tx_id: str
    when: datetime
    row: list


@dataclass
class Parsed:
    records: list[Record]
    transfers_skipped: int = 0  # giroconti tra i propri conti
    rules_skipped: int = 0  # regole con ignora = true
    investments_skipped: int = 0  # compravendite di titoli


def load_toml(path: Path) -> dict:
    with path.open("rb") as f:
        return tomllib.load(f)


def to_float(value: str) -> float:
    return float(value) if value.strip() else 0.0


def to_serial(when: datetime) -> float:
    return (when - SHEETS_EPOCH).total_seconds() / 86400


def from_serial(serial: float) -> datetime:
    return SHEETS_EPOCH + timedelta(days=serial)


def col_letter(index: int) -> str:
    """Lettera della colonna con indice (da zero) `index`: 0 -> A, 26 -> AA."""
    from gspread.utils import rowcol_to_a1

    return re.sub(r"\d", "", rowcol_to_a1(1, index + 1))


def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def safe_text(text: str) -> str:
    """Evita che Sheets interpreti come formula un testo che inizia con = + - @."""
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


def year_tab(cfg: dict, year: int) -> str:
    return f"{cfg['foglio_spese']} {year}"


def apply_rules(tx: dict, rules: dict) -> tuple[str, str, bool]:
    """Restituisce (nome, categoria, ignora) per una transazione."""
    text = " ".join(tx[k] for k in ("name", "description", "counterparty_name")).casefold()
    nome = categoria = ""
    for rule in rules.get("regola", []):
        if rule["contiene"].casefold() in text:
            if rule.get("ignora"):
                return "", "", True
            nome, categoria = rule.get("nome", ""), rule.get("categoria", "")
            break
    if not categoria:
        categoria = rules.get("mcc", {}).get(tx["mcc_code"].strip(), "")
    return nome, categoria, False


def parse_csv(path: Path, cfg: dict, rules: dict) -> Parsed:
    own_ibans = {iban.replace(" ", "").upper() for iban in cfg.get("miei_iban", [])}
    timezone = ZoneInfo(cfg.get("fuso_orario", "Europe/Rome"))
    parsed = Parsed(records=[])

    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = sorted(csv.DictReader(f), key=lambda r: r["datetime"])

    for tx in rows:
        tx = {k: (v or "") for k, v in tx.items()}
        tx_type = tx["type"]

        # Compravendite di titoli: non sono né spese né entrate.
        if tx["shares"].strip() and tx_type != "DIVIDEND":
            parsed.investments_skipped += 1
            continue

        # Giroconti. Negli addebiti diretti la controparte è il proprio conto,
        # quindi lì il controllo non vale.
        if (
            "TRANSFER" in tx_type
            and "DIRECT_DEBIT" not in tx_type
            and tx["counterparty_iban"].replace(" ", "").upper() in own_ibans
        ):
            parsed.transfers_skipped += 1
            continue

        nome, categoria, ignore = apply_rules(tx, rules)
        if ignore:
            parsed.rules_skipped += 1
            continue

        # Il CSV riporta l'ora in UTC: la convertiamo in ora locale.
        when = (
            datetime.fromisoformat(tx["datetime"].replace("Z", "+00:00"))
            .astimezone(timezone)
            .replace(tzinfo=None, microsecond=0)
        )

        # Importo netto: quello che è davvero entrato/uscito dal conto.
        net = round(to_float(tx["amount"]) + to_float(tx["tax"]) + to_float(tx["fee"]), 2)

        # I movimenti con carta vanno sempre tra le spese (uno storno compare
        # come spesa negativa); il resto dipende dal segno.
        if tx_type.startswith("CARD") or net < 0:
            tab, amount = year_tab(cfg, when.year), -net
        else:
            tab, amount = cfg["foglio_entrate"], net
            default_nome, default_cat = INCOME_BY_TYPE.get(tx_type, ("", ""))
            if default_nome is None:
                default_nome = tx["name"]
            nome = nome or default_nome
            categoria = categoria or default_cat

        parsed.records.append(
            Record(
                tab=tab,
                tx_id=tx["transaction_id"],
                when=when,
                row=[
                    to_serial(when),
                    amount,
                    safe_text(nome),
                    categoria,
                    cfg["conto"],
                    safe_text(tx["description"] or tx["name"]),
                    tx["transaction_id"],
                ],
            )
        )
    return parsed


# --- Google Sheets: tab con le transazioni -----------------------------------


def open_spreadsheet(cfg: dict):
    import gspread

    client = gspread.service_account(filename=str(ROOT / cfg["credenziali"]))
    return client.open_by_key(cfg["spreadsheet_id"])


def find_tab(spreadsheet, title: str, first_header: list[str]):
    """La tab con quel titolo, oppure None. Si ferma se l'intestazione non è la nostra."""
    import gspread

    try:
        ws = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return None
    header = ws.row_values(1)
    if header[: len(first_header)] != first_header:
        sys.exit(
            f'La tab "{title}" esiste già ma la prima riga non è quella attesa:\n'
            f"  attesa:  {first_header}\n  trovata: {header}\n"
            "Rinominala, oppure cambia i nomi delle tab in config.toml."
        )
    return ws


def read_rows(ws) -> list[list]:
    """Le righe di transazioni già sul foglio, senza le righe-titolo dei mesi.

    Ogni riga contiene le colonne dei dati seguite dalle eventuali colonne
    aggiunte a mano; le colonne di servizio sono escluse.
    """
    from gspread.utils import ValueRenderOption

    last_col = col_letter(ws.col_count - 1)
    values = ws.get_values(f"A2:{last_col}", value_render_option=ValueRenderOption.formula)
    extra_width = max([0] + [len(v) - EXTRA_START for v in values])
    rows = []
    for number, raw in enumerate(values, start=2):
        raw = list(raw) + [""] * (EXTRA_START + extra_width - len(raw))
        row = raw[:DATA_WIDTH] + raw[EXTRA_START:]
        if not any(str(cell).strip() for cell in row):
            continue
        date = row[DATE_COL]
        if is_number(date):
            # Il testo riletto ha perso l'eventuale apostrofo iniziale.
            row = [safe_text(c) if isinstance(c, str) and c[:1] != "=" else c for c in row]
            # "1,10" deve restare un elenco di mesi, non diventare il numero 1,1.
            if str(row[SPREAD_COL]).strip():
                row[SPREAD_COL] = "'" + str(row[SPREAD_COL]).lstrip("'")
            rows.append(row)
        elif isinstance(date, str) and MONTH_TITLE_RE.match(date.strip()):
            continue
        else:
            sys.exit(
                f'Tab "{ws.title}", riga {number}: la data "{date}" non è una data valida. '
                "Correggila sul foglio e rilancia: non è stato scritto nulla su questa tab."
            )
    return rows


def build_layout(rows: list[list]) -> list[list]:
    """Righe ordinate per data, con una riga-titolo (e il totale) prima di ogni mese."""
    width = max([DATA_WIDTH] + [len(r) for r in rows])
    rows = sorted((r + [""] * (width - len(r)) for r in rows), key=lambda r: r[DATE_COL])

    def month_of(row):
        when = from_serial(row[DATE_COL])
        return when.year, when.month

    layout = []
    for (year, month), group in groupby(rows, key=month_of):
        group = list(group)
        first = len(layout) + 3  # +1 intestazione, +1 riga-titolo, +1 perché le righe partono da 1
        last = first + len(group) - 1
        # L'apostrofo impedisce a Sheets di leggere "Settembre 2026" come una data.
        title = [f"'{MONTHS[month - 1]} {year}", f"=SUM(B{first}:B{last})"]
        layout.append(title + [""] * (width - len(title)))
        layout.extend(group)
    return layout


def write_data_tab(ws, layout: list[list]) -> None:
    """Riscrive le transazioni dalla riga 2 in un colpo solo, saltando le colonne di servizio."""
    end_row = len(layout) + 1
    extra_width = len(layout[0]) - DATA_WIDTH
    need_cols = EXTRA_START + extra_width
    if end_row > ws.row_count or need_cols > ws.col_count:
        ws.resize(rows=max(end_row + 200, ws.row_count), cols=max(need_cols, ws.col_count))

    data = [{"range": "A2", "values": [r[:DATA_WIDTH] for r in layout]}]
    tails = [f"A{end_row + 1}:{col_letter(DATA_WIDTH - 1)}{ws.row_count}"]
    if extra_width:
        first_extra = col_letter(EXTRA_START)
        data.append({"range": f"{first_extra}2", "values": [r[DATA_WIDTH:] for r in layout]})
        tails.append(f"{first_extra}{end_row + 1}:{col_letter(ws.col_count - 1)}{ws.row_count}")
    ws.batch_update(data, value_input_option="USER_ENTERED")
    if end_row < ws.row_count:
        ws.batch_clear(tails)


def repeat_cell(sheet_id: int, cell_format: dict, **grid_range) -> dict:
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, **grid_range},
            "cell": {"userEnteredFormat": cell_format},
            "fields": ",".join(f"userEnteredFormat.{key}" for key in cell_format),
        }
    }


def header_requests(sheet_id: int) -> list[dict]:
    return [
        repeat_cell(
            sheet_id,
            {
                "backgroundColor": HEADER_BG,
                "textFormat": {"bold": True, "foregroundColor": WHITE},
                "verticalAlignment": "MIDDLE",
            },
            startRowIndex=0,
            endRowIndex=1,
        ),
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                "properties": {"pixelSize": 32},
                "fields": "pixelSize",
            }
        },
    ]


def conditional_rule(sheet_id: int, formula: str, cell_format: dict, index: int, **columns) -> dict:
    return {
        "addConditionalFormatRule": {
            "index": index,
            "rule": {
                "ranges": [{"sheetId": sheet_id, "startRowIndex": 1, **columns}],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                    "format": cell_format,
                },
            },
        }
    }


def column_width(sheet_id: int, start: int, end: int, pixels: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": start, "endIndex": end},
            "properties": {"pixelSize": pixels},
            "fields": "pixelSize",
        }
    }


def create_data_tab(spreadsheet, title: str, categories: list[str], highlight_above: int | None):
    ws = spreadsheet.add_worksheet(title=title, rows=1000, cols=EXTRA_START, index=1)
    ws.update([HEADER], "A1")
    sid = ws.id
    body_rows = {"startRowIndex": 1}

    requests = header_requests(sid)
    requests += [
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        repeat_cell(sid, {"verticalAlignment": "MIDDLE"}, **body_rows),
        repeat_cell(
            sid,
            {"numberFormat": {"type": "DATE_TIME", "pattern": "d mmm hh:mm"}, "horizontalAlignment": "LEFT"},
            startColumnIndex=DATE_COL,
            endColumnIndex=DATE_COL + 1,
            **body_rows,
        ),
        repeat_cell(
            sid, {"numberFormat": EURO_FORMAT}, startColumnIndex=AMOUNT_COL, endColumnIndex=AMOUNT_COL + 1, **body_rows
        ),
        repeat_cell(
            sid,
            {"wrapStrategy": "WRAP"},
            startColumnIndex=DESCRIPTION_COL,
            endColumnIndex=DESCRIPTION_COL + 1,
            **body_rows,
        ),
        # L'ID serve solo al programma per riconoscere i doppioni.
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": ID_COL, "endIndex": ID_COL + 1},
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        },
        # Righe-titolo dei mesi: sono le uniche con un testo nella colonna Data.
        conditional_rule(sid, "=ISTEXT($A2)", {"backgroundColor": MONTH_BG, "textFormat": {"bold": True}}, 0),
    ]
    requests += [column_width(sid, i, i + 1, px) for i, px in enumerate(COLUMN_WIDTHS)]
    if highlight_above is not None:
        formula = f"=ISNUMBER($A2)*($B2>{highlight_above})=1"
        requests.append(conditional_rule(sid, formula, {"backgroundColor": HIGHLIGHT_BG}, 1))
    if categories:
        requests.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sid,
                        "startColumnIndex": CATEGORY_COL,
                        "endColumnIndex": CATEGORY_COL + 1,
                        **body_rows,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": [{"userEnteredValue": c} for c in categories],
                        },
                        "showCustomUi": True,
                        "strict": False,
                    },
                }
            }
        )
    spreadsheet.batch_update({"requests": requests})
    ensure_adjust_columns(spreadsheet, ws)
    return ws


def month_key(date: str) -> str:
    """Formula che trasforma una data in una chiave "2026-09"."""
    return f'YEAR({date})&"-"&RIGHT("0"&MONTH({date}){SEP}2)'


def helper_formulas() -> list[str]:
    """Formule (una per colonna, ancorate alla riga 1) delle colonne di servizio.

    - Mesi letti: i mesi scritti nella colonna "Mesi" tradotti in chiavi
      ";2026-10;2026-11;" ("?" se il testo non si capisce). Un mese senza anno
      è il primo con quel nome dal mese della transazione in poi.
    - Mesi contati: come sopra; senza indicazioni valide, il mese della transazione.
    - Quota mensile: importo contato diviso per il numero di mesi.
    """
    s = SEP
    date, amount, counted, spread = (
        f"{col_letter(c)}:{col_letter(c)}" for c in (DATE_COL, AMOUNT_COL, COUNTED_COL, SPREAD_COL)
    )
    date_body, spread_body = (f"{col_letter(c)}2:{col_letter(c)}" for c in (DATE_COL, SPREAD_COL))
    keys = f"{col_letter(DATA_WIDTH + 1)}:{col_letter(DATA_WIDTH + 1)}"
    parsed = f"{col_letter(DATA_WIDTH + 2)}:{col_letter(DATA_WIDTH + 2)}"
    names = "{" + ";".join(f'"{m[:3].lower()}"' for m in MONTHS) + "}"

    def column(title: str, expression: str) -> str:
        return f'=ARRAYFORMULA(IF(ROW({date})=1{s}"{title}"{s}IF(ISNUMBER({date}){s}{expression}{s}"")))'

    # Un elemento dell'elenco ("10", "ottobre", "1/2027", "gennaio 2027") -> "2026-10".
    # Ogni elemento inizia con "#": senza, SPLIT trasformerebbe "8/2026" in una data.
    shape = '"^# *([a-zA-Z]+|[0-9]{1,2})( *[/ ] *[0-9]{4})? *$"'
    token = (
        f'LET(name{s}REGEXEXTRACT(t{s}"^# *([a-zA-Z]+|[0-9]+)")'
        f"{s}m{s}IFERROR(MATCH(LEFT(LOWER(name){s}3){s}{names}{s}0){s}VALUE(name))"
        f'{s}y{s}IFERROR(VALUE(REGEXEXTRACT(t{s}"([0-9]{{4}}) *$")){s}IF(m>=MONTH(d){s}YEAR(d){s}YEAR(d)+1))'
        f'{s}IF(REGEXMATCH(t{s}{shape})*(m>=1)*(m<=12){s}y&"-"&RIGHT("0"&m{s}2){s}NA()))'
    )
    tokens = f'SPLIT("#"&REGEXREPLACE(ms&""{s}" *[,;] *"{s}",#"){s}",")'
    months_list = (
        f'IFERROR(";"&TEXTJOIN(";"{s}TRUE{s}UNIQUE(MAP({tokens}{s}LAMBDA(t{s}{token})){s}TRUE))&";"'
        f'{s}"{UNREADABLE}")'
    )
    own_month = f'";"&{month_key("d")}&";"'
    read = (
        f'=VSTACK("{HELPER_HEADER[2]}"{s}MAP({date_body}{s}{spread_body}{s}LAMBDA(d{s}ms{s}'
        f'IF(NOT(ISNUMBER(d)){s}""{s}IF(TRIM(ms&"")=""{s}{own_month}{s}{months_list})))))'
    )
    n_months = f'(LEN({keys})-LEN(SUBSTITUTE({keys}{s}";"{s}""))-1)'
    return [
        column(HELPER_HEADER[0], f'IF({counted}=""{s}{amount}{s}{counted})/{n_months}'),
        column(HELPER_HEADER[1], f'IF({parsed}="{UNREADABLE}"{s}";"&{month_key(date)}&";"{s}{parsed})'),
        read,
    ]


def ensure_adjust_columns(spreadsheet, ws) -> None:
    """Aggiunge o aggiorna le colonne di correzione e quelle di servizio."""
    if ws.row_values(1)[COUNTED_COL:EXTRA_START] == ADJUST_HEADER + HELPER_HEADER:
        return
    if ws.col_count < EXTRA_START:
        ws.resize(cols=EXTRA_START)
    # Svuota le colonne di servizio: le formule in riga 1 devono potersi espandere.
    ws.batch_clear([f"{col_letter(DATA_WIDTH)}1:{col_letter(EXTRA_START - 1)}{ws.row_count}"])
    ws.update([ADJUST_HEADER + helper_formulas()], f"{col_letter(COUNTED_COL)}1", value_input_option="USER_ENTERED")
    sid = ws.id
    spread_column = {"startColumnIndex": SPREAD_COL, "endColumnIndex": SPREAD_COL + 1}
    unreadable = f'=${col_letter(EXTRA_START - 1)}2="{UNREADABLE}"'
    requests = header_requests(sid) + [
        repeat_cell(
            sid, {"numberFormat": EURO_FORMAT}, startRowIndex=1, startColumnIndex=COUNTED_COL, endColumnIndex=COUNTED_COL + 1
        ),
        # Testo semplice: "1,10" non deve essere letto come un numero decimale.
        repeat_cell(sid, {"numberFormat": {"type": "TEXT"}}, startRowIndex=1, **spread_column),
        column_width(sid, COUNTED_COL, COUNTED_COL + 1, 140),
        column_width(sid, SPREAD_COL, SPREAD_COL + 1, 170),
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": DATA_WIDTH, "endIndex": EXTRA_START},
                "properties": {"hiddenByUser": True},
                "fields": "hiddenByUser",
            }
        },
        # Mesi scritti in modo non riconoscibile: la cella diventa rossa.
        conditional_rule(
            sid, unreadable, {"backgroundColor": ERROR_BG, "textFormat": {"foregroundColor": WHITE}}, 0, **spread_column
        ),
    ]
    spreadsheet.batch_update({"requests": requests})


def write_grid(ws, values: list[list], first_row: int = 2) -> None:
    """Scrive `values` da `first_row` in giù in un colpo solo e svuota ciò che resta sotto."""
    from gspread.utils import rowcol_to_a1

    width = max(len(v) for v in values)
    end_row = first_row + len(values) - 1
    if end_row > ws.row_count or width > ws.col_count:
        ws.resize(rows=max(end_row + 200, ws.row_count), cols=max(width, ws.col_count))
    ws.update(values, f"A{first_row}", value_input_option="USER_ENTERED")
    if end_row < ws.row_count:
        last_col = re.sub(r"\d", "", rowcol_to_a1(1, ws.col_count))
        ws.batch_clear([f"A{end_row + 1}:{last_col}{ws.row_count}"])


# --- Google Sheets: riepilogo mensile ----------------------------------------


def quoted(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def read_manual_cells(ws, columns: list[str]) -> dict[int, dict[str, object]]:
    """Le celle scritte a mano nel riepilogo (saldi dei conti e note), per mese: vanno conservate."""
    from gspread.utils import ValueRenderOption

    values = ws.get_values(value_render_option=ValueRenderOption.formula)
    if not values:
        return {}
    position = {name: values[0].index(name) for name in columns if name in values[0]}
    saved = {}
    for row in values[1:]:
        if row and is_number(row[0]):
            saved[round(row[0])] = {name: row[i] for name, i in position.items() if i < len(row) and row[i] != ""}
    return saved


def build_summary(
    years: list[int],
    cfg: dict,
    categories: list[str],
    accounts: list[str],
    has_income: bool,
    manual: dict[int, dict[str, object]],
) -> tuple[list[list], list[dict]]:
    """Una riga per mese. Tutto formule, tranne i saldi dei conti e le note che si scrivono a mano.

    Restituisce le righe (dalla riga 2 del foglio) e, per ogni anno, i numeri
    di riga del suo blocco: titolo, primo e ultimo mese, totale, media, grafici.
    """
    n_fixed = len(SUMMARY_FIXED)
    first_account = n_fixed
    notes = first_account + len(accounts)
    uncategorized = notes + 1
    first_category = uncategorized + 1
    n_cols = first_category + len(categories)
    quota, month_keys = col_letter(DATA_WIDTH), col_letter(DATA_WIDTH + 1)
    category = col_letter(CATEGORY_COL)

    def month_sum(tabs: list[str], r: int, category_cell: str = "") -> str:
        """Somma delle quote che pesano sul mese della riga `r`, su una o più tab."""
        terms = []
        for tab in tabs:
            extra = f"{SEP}{tab}!${category}:${category}{SEP}{category_cell}" if category_cell else ""
            terms.append(
                f"SUMIFS({tab}!${quota}:${quota}"
                f'{SEP}{tab}!${month_keys}:${month_keys}{SEP}"*;"&{month_key(f"$A{r}")}&";*"{extra})'
            )
        return "=" + "+".join(terms) if terms else ""

    def saved_share(r: int) -> str:
        return f'=IF(B{r}>0{SEP}D{r}/B{r}{SEP}"")'

    income_tabs = [quoted(cfg["foglio_entrate"])] if has_income else []
    money_cols = [1, 2, 3, uncategorized] + list(range(first_category, n_cols))
    grid, blocks = [], []
    for year in years:
        # Una spesa può essere spostata su mesi di un altro anno.
        spese_tabs = [quoted(year_tab(cfg, y)) for y in years]
        block = {"year": year, "title": len(grid) + 2}
        grid.append([f"'{year}"] + [""] * (n_cols - 1))
        start = len(grid) + 2
        for month in range(1, 13):
            r = len(grid) + 2
            month_start = round(to_serial(datetime(year, month, 1)))
            row = [""] * n_cols
            row[0] = month_start
            row[1] = month_sum(income_tabs, r)
            row[2] = month_sum(spese_tabs, r)
            row[3] = f"=B{r}-C{r}"
            row[4] = saved_share(r)
            if accounts:
                span = f"{col_letter(first_account)}{r}:{col_letter(notes - 1)}{r}"
                row[5] = f'=IF(COUNT({span})=0{SEP}""{SEP}SUM({span}))'
                for i, name in enumerate(accounts):
                    row[first_account + i] = manual.get(month_start, {}).get(name, "")
            note = manual.get(month_start, {}).get(NOTES, "")
            row[notes] = safe_text(note) if isinstance(note, str) and note[:1] != "=" else note
            if categories:
                span = f"{col_letter(first_category)}{r}:{col_letter(n_cols - 1)}{r}"
                row[uncategorized] = f"=ROUND(C{r}-SUM({span}){SEP}2)"
            for c in range(first_category, n_cols):
                row[c] = month_sum(spese_tabs, r, f"{col_letter(c)}$1")
            grid.append(row)
        end = len(grid) + 1

        r = len(grid) + 2
        total = [f"Totale {year}"] + [""] * (n_cols - 1)
        for c in money_cols:
            total[c] = f"=SUM({col_letter(c)}{start}:{col_letter(c)}{end})"
        total[4] = saved_share(r)
        grid.append(total)

        # Media sui soli mesi in cui c'è stata almeno una spesa.
        average = [f"Media mensile {year}"] + [""] * (n_cols - 1)
        for c in money_cols:
            span = f"{col_letter(c)}{start}:{col_letter(c)}{end}"
            average[c] = f'=IFERROR(AVERAGEIF($C{start}:$C{end}{SEP}">0"{SEP}{span}){SEP}"")'
        grid.append(average)
        block |= {"first": start, "last": end, "total": end + 1, "average": end + 2, "charts": end + 4}
        blocks.append(block)
        grid.extend([""] * n_cols for _ in range(CHART_ROWS))
    return grid, blocks


def create_summary_tab(spreadsheet, title: str, n_cols: int):
    return spreadsheet.add_worksheet(title=title, rows=200, cols=n_cols, index=0)


def summary_charts(sid: int, block: dict, first_slice: int, n_cols: int) -> list[dict]:
    """I due grafici di un anno: torta delle spese per categoria e andamento mensile."""

    def cells(first_row: int, last_row: int, first_col: int, last_col: int) -> dict:
        """Intervallo con righe numerate come sul foglio e colonne da zero, estremi inclusi."""
        return {
            "sourceRange": {
                "sources": [
                    {
                        "sheetId": sid,
                        "startRowIndex": first_row - 1,
                        "endRowIndex": last_row,
                        "startColumnIndex": first_col,
                        "endColumnIndex": last_col + 1,
                    }
                ]
            }
        }

    def chart(spec: dict, left: int, width: int) -> dict:
        """Grafico largo `width` pixel, a `left` pixel dal bordo sinistro del foglio."""
        position = {
            "overlayPosition": {
                "anchorCell": {"sheetId": sid, "rowIndex": block["charts"] - 1, "columnIndex": 0},
                "offsetXPixels": left,
                "widthPixels": width,
                "heightPixels": CHART_HEIGHT,
            }
        }
        return {"addChart": {"chart": {"spec": spec, "position": position}}}

    year = block["year"]
    pie = {
        "title": f"Spese {year} per categoria",
        "pieChart": {
            "legendPosition": "RIGHT_LEGEND",
            "domain": cells(1, 1, first_slice, n_cols - 1),
            "series": cells(block["total"], block["total"], first_slice, n_cols - 1),
        },
    }
    line = {
        "title": f"Spese {year} mese per mese",
        "basicChart": {
            "chartType": "LINE",
            "legendPosition": "NO_LEGEND",
            "headerCount": 0,
            "axis": [{"position": "LEFT_AXIS", "title": "€"}],
            "domains": [{"domain": cells(block["first"], block["last"], 0, 0)}],
            "series": [
                {"series": cells(block["first"], block["last"], 2, 2), "targetAxis": "LEFT_AXIS", "color": RED_TEXT}
            ],
        },
    }
    # La torta parte dal bordo sinistro e la linea inizia esattamente dove finisce la torta.
    pie_width = 560
    return [chart(pie, 0, pie_width), chart(line, pie_width, 760)]


def format_summary(spreadsheet, ws, n_cols: int, n_accounts: int, blocks: list[dict]) -> None:
    """Formati e grafici della tab di riepilogo: rifatti a ogni aggiornamento."""
    sid = ws.id
    body_rows = {"startRowIndex": 1}

    # La tab è tutta del programma: regole di formattazione condizionale e
    # grafici vengono tolti e ricreati, così non si accumulano a ogni caricamento.
    fields = "sheets(properties.sheetId,conditionalFormats,charts.chartId)"
    metadata = spreadsheet.fetch_sheet_metadata(params={"fields": fields})
    current = next(sh for sh in metadata["sheets"] if sh["properties"]["sheetId"] == sid)
    old_rules = current.get("conditionalFormats", [])
    requests = [{"deleteConditionalFormatRule": {"sheetId": sid, "index": i}} for i in reversed(range(len(old_rules)))]
    requests += [{"deleteEmbeddedObject": {"objectId": c["chartId"]}} for c in current.get("charts", [])]

    def balance_rule(condition: str, color: dict, index: int) -> dict:
        return {
            "addConditionalFormatRule": {
                "index": index,
                "rule": {
                    "ranges": [{"sheetId": sid, "startColumnIndex": 3, "endColumnIndex": 4, **body_rows}],
                    "booleanRule": {
                        "condition": {"type": condition, "values": [{"userEnteredValue": "0"}]},
                        "format": {"textFormat": {"foregroundColor": color}},
                    },
                },
            }
        }

    requests += [
        # Saldo: verde se positivo, rosso se negativo.
        balance_rule("NUMBER_GREATER", GREEN_TEXT, 0),
        balance_rule("NUMBER_LESS", RED_TEXT, 1),
    ]
    requests += header_requests(sid)
    requests += [
        {
            "updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 1},
                "properties": {"pixelSize": 21},
                "fields": "pixelSize",
            }
        },
        repeat_cell(sid, {"textFormat": {"foregroundColor": BLACK}}, **body_rows),
        repeat_cell(sid, {"textFormat": {"foregroundColor": GREEN_TEXT}}, startColumnIndex=1, endColumnIndex=2, **body_rows),
        repeat_cell(sid, {"textFormat": {"foregroundColor": RED_TEXT}}, startColumnIndex=2, endColumnIndex=3, **body_rows),
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1, "frozenColumnCount": 1}},
                "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
            }
        },
        repeat_cell(sid, {"wrapStrategy": "WRAP"}, startRowIndex=0, endRowIndex=1),
        repeat_cell(
            sid,
            {"numberFormat": {"type": "DATE", "pattern": "mmmm yyyy"}, "horizontalAlignment": "LEFT"},
            startColumnIndex=0,
            endColumnIndex=1,
            **body_rows,
        ),
        repeat_cell(sid, {"numberFormat": EURO_FORMAT_NO_ZERO}, startColumnIndex=1, **body_rows),
        repeat_cell(
            sid, {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}, startColumnIndex=4, endColumnIndex=5, **body_rows
        ),
        column_width(sid, 0, 1, 190),
        column_width(sid, 1, n_cols, 115),
        # Le colonne da compilare a mano (saldi dei conti) hanno uno sfondo diverso.
        repeat_cell(sid, {"backgroundColor": WHITE}, **body_rows),
    ]
    # Saldi dei conti e note.
    first, notes = len(SUMMARY_FIXED), len(SUMMARY_FIXED) + n_accounts
    requests += [
        repeat_cell(sid, {"backgroundColor": INPUT_BG}, startColumnIndex=first, endColumnIndex=notes + 1, **body_rows),
        repeat_cell(
            sid,
            {"numberFormat": {"type": "TEXT"}, "wrapStrategy": "WRAP"},
            startColumnIndex=notes,
            endColumnIndex=notes + 1,
            **body_rows,
        ),
        column_width(sid, notes, notes + 1, 320),
    ]

    # Per ultimi, perché vincano sui formati di colonna: la fascia con l'anno
    # e le righe "Totale" e "Media" di ogni blocco.
    first_slice = notes + 1  # "Da categorizzare", seguita dalle categorie
    for block in blocks:
        title = {"startRowIndex": block["title"] - 1, "endRowIndex": block["title"]}
        requests += [
            repeat_cell(
                sid,
                {"backgroundColor": HEADER_BG, "textFormat": {"bold": True, "fontSize": 16, "foregroundColor": WHITE}},
                **title,
            ),
            {
                "updateDimensionProperties": {
                    "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": title["startRowIndex"], "endIndex": title["endRowIndex"]},
                    "properties": {"pixelSize": 40},
                    "fields": "pixelSize",
                }
            },
            {
                "repeatCell": {
                    "range": {"sheetId": sid, "startRowIndex": block["total"] - 1, "endRowIndex": block["average"]},
                    "cell": {"userEnteredFormat": {"backgroundColor": MONTH_BG, "textFormat": {"bold": True}}},
                    "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold",
                }
            },
        ]
        requests += summary_charts(sid, block, first_slice, n_cols)
    spreadsheet.batch_update({"requests": requests})


def update_summary(spreadsheet, cfg: dict, categories: list[str]) -> None:
    title = cfg.get("foglio_riepilogo", "Riepilogo")
    accounts = cfg.get("patrimonio", DEFAULT_ACCOUNTS)
    header = SUMMARY_FIXED + accounts + [NOTES, "Da categorizzare"] + categories
    pattern = re.compile(rf"^{re.escape(cfg['foglio_spese'])} (\d{{4}})$")

    years, has_income = [], False
    for tab in spreadsheet.worksheets():
        if match := pattern.match(tab.title):
            years.append(int(match.group(1)))
        elif tab.title == cfg["foglio_entrate"]:
            has_income = True
        else:
            continue
        ensure_adjust_columns(spreadsheet, tab)
    if not years:
        return

    ws = find_tab(spreadsheet, title, SUMMARY_FIXED[:1])
    manual = read_manual_cells(ws, accounts + [NOTES]) if ws is not None else {}
    if ws is None:
        ws = create_summary_tab(spreadsheet, title, len(header))
    grid, blocks = build_summary(sorted(years), cfg, categories, accounts, has_income, manual)
    write_grid(ws, [header] + grid, first_row=1)
    format_summary(spreadsheet, ws, len(header), len(accounts), blocks)
    print(f"{title}: aggiornato ({', '.join(map(str, sorted(years)))})")


# --- CLI ---------------------------------------------------------------------


def print_preview(title: str, records: list[Record]) -> None:
    print(f"\n{title} ({len(records)})")
    for r in records:
        _data, importo, nome, categoria, _conto, descrizione, _id = r.row
        print(f"  {r.when:%d/%m %H:%M}  {importo:>9.2f}  {nome[:18]:<18}  {categoria[:18]:<18}  {descrizione[:50]}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("csv", type=Path, help="export CSV della banca")
    parser.add_argument("--dry-run", action="store_true", help="mostra cosa verrebbe caricato senza scrivere")
    args = parser.parse_args()

    config_path = ROOT / "config.toml"
    if not config_path.exists():
        sys.exit("Manca config.toml: copia config.example.toml in config.toml e compilalo.")
    cfg = load_toml(config_path)
    rules = load_toml(ROOT / "regole.toml")
    spese_categories = rules.get("categorie_spese", [])
    highlight_above = int(cfg.get("soglia_evidenza", 30))

    parsed = parse_csv(args.csv, cfg, rules)

    online = bool(cfg.get("spreadsheet_id")) and (ROOT / cfg["credenziali"]).exists()
    if not online and not args.dry_run:
        sys.exit("Imposta spreadsheet_id e il file delle credenziali in config.toml (vedi README).")

    spreadsheet = open_spreadsheet(cfg) if online else None
    duplicates = 0
    # Prima le tab delle spese in ordine di anno, poi le entrate.
    titles = sorted({r.tab for r in parsed.records}, key=lambda t: (t == cfg["foglio_entrate"], t))
    for title in titles:
        is_income = title == cfg["foglio_entrate"]
        records = [r for r in parsed.records if r.tab == title]
        ws = find_tab(spreadsheet, title, HEADER) if online else None
        if ws is not None and not args.dry_run:
            ensure_adjust_columns(spreadsheet, ws)
        existing = read_rows(ws) if ws is not None else []
        known_ids = {row[ID_COL] for row in existing}
        new = [r for r in records if r.tx_id not in known_ids]
        duplicates += len(records) - len(new)

        if args.dry_run:
            print_preview(title, new)
        elif new:
            if ws is None:
                ws = create_data_tab(
                    spreadsheet,
                    title,
                    rules.get("categorie_entrate", []) if is_income else spese_categories,
                    None if is_income else highlight_above,
                )
            write_data_tab(ws, build_layout(existing + [r.row for r in new]))
            print(f"{title}: {len(new)} righe aggiunte")
        else:
            print(f"{title}: niente di nuovo")

        if new:
            missing = sum(1 for r in new if not r.row[CATEGORY_COL])
            print(f"  {missing} su {len(new)} senza categoria, da compilare a mano")

    if online and not args.dry_run:
        update_summary(spreadsheet, cfg, spese_categories)

    print()
    if online:
        print(f"Già presenti sul foglio, saltate: {duplicates}")
    else:
        print("Foglio Google non configurato: anteprima senza controllo dei duplicati.")
    print(f"Giroconti tra i tuoi conti, saltati: {parsed.transfers_skipped}")
    if parsed.rules_skipped:
        print(f"Ignorate da una regola: {parsed.rules_skipped}")
    if parsed.investments_skipped:
        print(f"Compravendite di titoli, saltate: {parsed.investments_skipped}")
    if args.dry_run:
        print("Dry-run: non è stato scritto nulla.")


if __name__ == "__main__":
    main()
