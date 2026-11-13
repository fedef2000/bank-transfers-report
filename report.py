#!/usr/bin/env python3
"""Carica le transazioni di un export CSV della banca su un foglio Google.

Uso:
    python report.py "Transaction export.csv" --dry-run   # anteprima, non scrive
    python report.py "Transaction export.csv"             # carica sul foglio
"""
from __future__ import annotations

import argparse
import csv
import sys
import tomllib
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent

HEADER = ["Data", "Importo", "Nome", "Categoria", "Conto", "Descrizione banca", "ID transazione"]
CATEGORY_COL = HEADER.index("Categoria")
ID_COL = HEADER.index("ID transazione")

SPESE, ENTRATE = "spese", "entrate"

# Entrate la cui natura è già chiara dal tipo di transazione.
INCOME_BY_TYPE = {
    "INTEREST_PAYMENT": ("Interessi", "Interessi"),
    "DIVIDEND": (None, "Dividendi"),  # il nome è quello del fondo
}


@dataclass
class Record:
    tab: str  # SPESE | ENTRATE
    tx_id: str
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


def safe_text(text: str) -> str:
    """Evita che Sheets interpreti come formula un testo che inizia con = + - @."""
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


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

        # Importo netto: quello che è davvero entrato/uscito dal conto.
        net = round(to_float(tx["amount"]) + to_float(tx["tax"]) + to_float(tx["fee"]), 2)

        # I movimenti con carta vanno sempre tra le spese (uno storno compare
        # come spesa negativa); il resto dipende dal segno.
        if tx_type.startswith("CARD") or net < 0:
            tab, amount = SPESE, -net
        else:
            tab, amount = ENTRATE, net
            default_nome, default_cat = INCOME_BY_TYPE.get(tx_type, ("", ""))
            if default_nome is None:
                default_nome = tx["name"]
            nome = nome or default_nome
            categoria = categoria or default_cat

        parsed.records.append(
            Record(
                tab=tab,
                tx_id=tx["transaction_id"],
                row=[
                    tx["date"],
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


# --- Google Sheets -----------------------------------------------------------


def open_spreadsheet(cfg: dict):
    import gspread

    client = gspread.service_account(filename=str(ROOT / cfg["credenziali"]))
    return client.open_by_key(cfg["spreadsheet_id"])


def find_tab(spreadsheet, title: str):
    """La tab con quel titolo, oppure None. Si ferma se l'intestazione non è la nostra."""
    import gspread

    try:
        ws = spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return None
    header = ws.row_values(1)
    if header[: len(HEADER)] != HEADER:
        sys.exit(
            f'La tab "{title}" esiste già ma la prima riga non è quella attesa:\n'
            f"  attesa:  {HEADER}\n  trovata: {header}\n"
            "Rinominala o cambia foglio_spese / foglio_entrate in config.toml."
        )
    return ws


def create_tab(spreadsheet, title: str, categories: list[str]):
    ws = spreadsheet.add_worksheet(title=title, rows=1000, cols=len(HEADER))
    ws.update([HEADER], "A1")
    ws.freeze(rows=1)
    ws.format("1:1", {"textFormat": {"bold": True}})
    if categories:
        spreadsheet.batch_update(
            {
                "requests": [
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": ws.id,
                                "startRowIndex": 1,
                                "startColumnIndex": CATEGORY_COL,
                                "endColumnIndex": CATEGORY_COL + 1,
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
                ]
            }
        )
    return ws


# --- CLI ---------------------------------------------------------------------


def print_preview(title: str, records: list[Record]) -> None:
    print(f"\n{title} ({len(records)})")
    for r in records:
        data, importo, nome, categoria, _conto, descrizione, _id = r.row
        print(f"  {data}  {importo:>9.2f}  {nome[:18]:<18}  {categoria[:18]:<18}  {descrizione[:50]}")


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

    parsed = parse_csv(args.csv, cfg, rules)
    tabs = {
        SPESE: (cfg["foglio_spese"], rules.get("categorie_spese", [])),
        ENTRATE: (cfg["foglio_entrate"], rules.get("categorie_entrate", [])),
    }

    online = bool(cfg.get("spreadsheet_id")) and (ROOT / cfg["credenziali"]).exists()
    if not online and not args.dry_run:
        sys.exit("Imposta spreadsheet_id e il file delle credenziali in config.toml (vedi README).")

    spreadsheet = open_spreadsheet(cfg) if online else None
    duplicates = 0
    for tab, (title, categories) in tabs.items():
        records = [r for r in parsed.records if r.tab == tab]
        ws = find_tab(spreadsheet, title) if online else None
        if ws is not None:
            existing = set(ws.col_values(ID_COL + 1))
            new = [r for r in records if r.tx_id not in existing]
            duplicates += len(records) - len(new)
            records = new

        if args.dry_run:
            print_preview(title, records)
        elif records:
            if ws is None:
                ws = create_tab(spreadsheet, title, categories)
            ws.append_rows(
                [r.row for r in records],
                value_input_option="USER_ENTERED",
                insert_data_option="INSERT_ROWS",
                table_range="A1",
            )
            print(f"{title}: {len(records)} righe aggiunte")
        else:
            print(f"{title}: niente di nuovo")

        if records:
            missing = sum(1 for r in records if not r.row[CATEGORY_COL])
            print(f"  {missing} su {len(records)} senza categoria, da compilare a mano")

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
