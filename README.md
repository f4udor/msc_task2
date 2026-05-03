# Packaging PDF Extraction Workflow

Workflow Python-first per estrarre dati strutturati da PDF packaging e produrre:
- un Excel finale con i dati raccolti
- un Excel di log con affidabilita`, review e dettagli di estrazione

## Obiettivo

Dato un batch di PDF packaging:
- analizzare i file
- estrarre i campi utili
- normalizzare l'output in una tabella business
- segnalare campi mancanti, inferiti o poco affidabili
- generare un output ripetibile e schedulabile

Il focus del progetto e` il parsing affidabile dei PDF.  
L'orchestrazione esterna puo` essere aggiunta dopo con cron, watcher o altri strumenti.

## Stack

- Python
- `PyMuPDF` (`fitz`) per parsing PDF, testo e metadati pagina
- `pytesseract` + `Tesseract OCR` per OCR fallback
- `Pillow` per gestione immagini OCR
- `openpyxl` per generazione degli Excel finali

## Struttura Finale del Repo

- [packaging_pdf_parser.py](./packaging_pdf_parser.py): parser core
- [run_packaging_workflow.py](./run_packaging_workflow.py): entrypoint finale
- [images](./images): dataset PDF
- [README.md](./README.md): documentazione finale

## Processo End-to-End

1. Scansione di una cartella di PDF.
2. Parsing del filename per recuperare campi strutturati come EAN, nome prodotto e dimensioni pack.
3. Parsing del testo PDF selezionabile.
4. OCR fallback sulle pagine renderizzate.
5. Anchor detection sul testo OCR.
6. Parsing guidato da zone logiche:
   - technical specs
   - compliance
   - QR / environmental info
   - marketing
7. Normalizzazione dei campi nello schema finale.
8. Calcolo di review e affidabilita`.
9. Generazione di:
   - `pack_data.xlsx`
   - `pack_log.xlsx`

## Output

### 1. Excel dati

File predefinito:

```text
pack_data.xlsx
```

Contiene:
- `Data`: una riga per PDF con i campi finali
- `Trace`: dettaglio per campo con `source` e `confidence`

### 2. Excel log

File predefinito:

```text
pack_log.xlsx
```

Contiene:
- `Run Summary`: metriche del batch
- `File Log`: una riga per PDF con stato, tempi e review
- `Review Queue`: campi da ricontrollare
- `Coverage`: copertura per campo

## Come Eseguire

Dalla root del progetto:

```bash
.venv/bin/python run_packaging_workflow.py
```

Con output custom:

```bash
.venv/bin/python run_packaging_workflow.py \
  --data-output /percorso/pack_data.xlsx \
  --log-output /percorso/pack_log.xlsx
```

Con directory diversa:

```bash
.venv/bin/python run_packaging_workflow.py /percorso/cartella_pdf
```

## Gestione Errori

Il workflow non si ferma al primo errore.

Per ogni PDF:
- se il parsing riesce, il record viene marcato come `perfect` o `partial`
- se il parsing fallisce, il record viene marcato come `failed`
- il batch continua comunque

Segnali usati:
- `source=missing`: campo non trovato
- `source=inference`: campo inferito, non evidenza diretta
- `confidence=low`: campo debole, da ricontrollare
- `review_needed=true`: il record richiede controllo umano

Gli errori di record finiscono nel foglio `Review Queue` e nel foglio `File Log`.

## Accuratezza sui 50 Pack

Risultati dell'ultimo run sul dataset attuale:

- PDF analizzati: `50`
- perfect: `0`
- partial: `50`
- failed: `0`
- average missing fields: `5.82`

Metriche di copertura gia` misurate nel progetto:

- asserted fill: `82.88%`
- evidence fill: `43.76%`
- strong evidence fill: `40.0%`

Interpretazione corretta:
- il sistema produce una riga utile per tutti i 50 pack
- nessun file e` fallito tecnicamente
- tutti i record richiedono almeno un minimo di review, quindi oggi nessuno e` da considerare "perfetto"

## Tempi di Elaborazione

Ultimo run sul dataset da 50 PDF:

- tempo totale: `72.277s`
- tempo medio per PDF: `1.445s`

Stima lineare a regime:

- 10 PDF: ~`14-15s`
- 50 PDF: ~`72s`
- 100 PDF: ~`145s`

Il costo computazionale e` dominato dall'OCR.

## Costo Stimato per Pack

Con l'architettura attuale:
- nessun modello AI a pagamento
- OCR locale
- parsing locale

Quindi il costo marginale per pack e` molto basso.

Il costo reale e` dato da:
- tempo macchina
- manutenzione delle regole di parsing
- eventuale review manuale

## Limiti Attuali

- alcuni simboli grafici non sono ancora estratti in modo robusto
- diversi campi sono valorizzati via `inference` o `default_false`
- la review umana e` ancora necessaria
- il parser e` ottimizzato per PDF simili a quelli del dataset attuale

## Miglioramenti Possibili

- template matching per simboli grafici
- fonte dati esterna per campi non affidabili da PDF
- riduzione dei campi `inference`
- trigger automatico via cron o watcher directory
- scrittura diretta su Google Sheets come step successivo

## Architettura Consigliata

La scelta progettuale finale e`:

- parsing e logica critica in Python
- orchestrazione opzionale esterna

Questa scelta evita di forzare tool di automazione nel punto sbagliato e mantiene:
- controllo
- trasparenza
- testabilita`
- costi prevedibili

## Conclusione

Il progetto raggiunge l'obiettivo principale:
- trasformare PDF packaging in dati strutturati
- produrre output tabellare utilizzabile
- segnalare errori, affidabilita` e bisogno di review

La parte piu` difficile era il parsing.  
Quella parte e` stata isolata, resa batchabile e documentata.
