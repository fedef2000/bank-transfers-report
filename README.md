# Bank transaction report

Legge l'export CSV delle transazioni della banca, assegna nome e categoria alle
spese ricorrenti e le carica su un foglio Google, con un riepilogo mensile.

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
5. Per avere i mesi in italiano: *File → Impostazioni → Impostazioni
   internazionali → Italia*.

## Uso

```sh
.venv/bin/python report.py "Transaction export.csv" --dry-run   # anteprima
.venv/bin/python report.py "Transaction export.csv"             # carica
```

Le transazioni già presenti, riconosciute dalla colonna nascosta
`ID transazione`, vengono saltate: lo stesso CSV si può ricaricare senza creare
doppioni. Le altre tab del foglio non vengono toccate.

## Struttura del foglio

- **`Spese <anno>`** (una tab per anno): una riga per spesa, ordinate per data.
  Ogni mese inizia con una riga-titolo ("Settembre 2026") che ne riporta il
  totale. Le spese sopra i 30 € hanno lo sfondo rosso chiaro. Importo positivo
  = soldi usciti; uno storno compare negativo.
- **`Entrate`** (tab unica): stessa struttura, importi al netto delle tasse.
- **`Riepilogo`**: una riga per mese con entrate, spese, saldo, percentuale
  risparmiata, patrimonio, quanto resta da categorizzare e il totale per
  categoria, più totale e media mensile dell'anno. Sono formule, quindi si
  aggiorna da solo quando modifichi le transazioni. Viene riscritta a ogni
  caricamento: le uniche celle da compilare a mano sono i saldi dei conti e
  la colonna `Note` (sfondo giallo), che vengono conservati.

Colonne delle tab di transazioni:
`Data | Importo | Nome | Categoria | Conto | Descrizione banca | Importo contato | Mesi`
(più `ID transazione` e tre colonne di servizio, nascoste).

### Patrimonio

Nel `Riepilogo`, a fine mese, scrivi nelle colonne gialle quanto hai su ogni
conto (`Investimenti`, `Trade Republic`, `N26`, `Contanti`): la colonna
`Patrimonio` ne fa la somma. L'elenco dei conti si cambia con `patrimonio` in
`config.toml`.

### Correggere quanto conta una transazione

Il riepilogo usa due colonne facoltative, da compilare sulla riga della
transazione (vale sia per le spese sia per le entrate):

- **`Importo contato`**: l'importo da contare al posto di quello reale. `0` =
  la transazione non viene contata (es. soldi anticipati per amici: metti `0`
  anche sul bonifico con cui te li ridanno, nella tab `Entrate`). Vuoto =
  conta l'importo reale.
- **`Mesi`**: in quali mesi far risultare l'importo, diviso in parti uguali.
  Elenco separato da virgole, con numeri o nomi: `10, 11` oppure
  `ottobre, novembre` (basta l'inizio: `ott, nov`). Vuoto = il mese della
  transazione.
  - Senza anno, un mese è il primo con quel nome dal mese della transazione in
    poi: su una spesa di novembre, `1` è gennaio dell'anno dopo.
  - Per indicare l'anno: `1/2027` oppure `gennaio 2027`. Serve per spostare
    una spesa a un mese precedente (`8/2026` su una spesa di settembre).
  - Se il testo non viene capito la cella diventa rossa e la spesa resta nel
    suo mese.
  - Un mese compare nel riepilogo solo se esiste la tab `Spese` di quell'anno.

Il totale nella riga-titolo del mese resta quello dei soldi realmente usciti.

### Modifiche a mano

- `Nome` e `Categoria` si compilano direttamente sul foglio.
- Le spese dell'altro conto si aggiungono scrivendo una riga nuova in fondo
  alla tab (data, importo, nome, categoria, conto; `ID transazione` vuoto): al
  caricamento successivo viene spostata nel mese giusto.
- Non serve toccare le righe-titolo dei mesi: il programma le rigenera.
- Dopo un caricamento che inserisce righe in mezzo, colori e note messi a mano
  sulle singole celle possono non seguire la riga.

## Regole

`regole.toml` contiene le regole "testo → nome, categoria", la tabella dei
codici MCC usata come ripiego per la categoria e l'elenco delle categorie.
Per le spese che non ricordi c'è la categoria `Da identificare`.

I bonifici da/verso gli IBAN elencati in `miei_iban` (`config.toml`) sono
giroconti e non vengono caricati.
