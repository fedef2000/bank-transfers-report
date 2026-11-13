# Bank transaction report

Legge l'export CSV delle transazioni della banca, assegna nome e categoria alle
spese ricorrenti e aggiunge le righe a un foglio Google.

## Installazione

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp config.example.toml config.toml
```

## Collegare il foglio Google (una volta sola)

1. Su <https://console.cloud.google.com> crea un progetto e abilita la **Google Sheets API**.
2. In *IAM e amministrazione → Service account* crea un service account, poi
   *Chiavi → Aggiungi chiave → JSON*. Salva il file come `credentials.json`
   nella cartella del progetto (è ignorato da git).
3. Apri il tuo foglio Google e condividilo come **Editor** con l'indirizzo email
   del service account (`...@...iam.gserviceaccount.com`).
4. Copia in `config.toml` l'ID del foglio (la parte dell'URL tra `/d/` e `/edit`).

## Uso

```sh
.venv/bin/python report.py "Transaction export.csv" --dry-run   # anteprima
.venv/bin/python report.py "Transaction export.csv"             # carica
```

Il programma aggiunge soltanto righe in fondo alle tab `Spese` ed `Entrate`
(create alla prima esecuzione) e non tocca le altre tab. Le transazioni già
presenti, riconosciute dalla colonna `ID transazione`, vengono saltate: lo
stesso CSV si può ricaricare senza creare doppioni.

Colonne: `Data | Importo | Nome | Categoria | Conto | Descrizione banca | ID transazione`

- **Spese**: importo positivo = soldi usciti; uno storno compare negativo.
- **Entrate**: importo al netto delle tasse (interessi, dividendi, bonifici ricevuti).
- I bonifici da/verso gli IBAN elencati in `miei_iban` (`config.toml`) sono
  giroconti e non vengono caricati.
- Le spese dell'altro conto si aggiungono a mano nelle stesse tab, lasciando
  vuoto `ID transazione`.

## Regole

`regole.toml` contiene le regole "testo → nome, categoria", la tabella dei
codici MCC usata come ripiego per la categoria e l'elenco delle categorie.
Per le spese che non ricordi c'è la categoria `Da identificare`.
