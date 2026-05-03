# n8n Integration

## Obiettivo

Usare il parser come step batch dentro n8n, senza introdurre ancora una API HTTP.

Nel repository trovi anche uno scaffold importabile:

```text
n8n_packaging_workflow.json
```

E` un workflow base con:
- `Manual Trigger`
- `Set Config`
- `Execute Command`
- `Code` per espandere i record
- `IF` per separare `parsed_ready` e `needs_review`
- due nodi Google Sheets finali

Il punto di ingresso da usare e`:

```bash
.venv/bin/python run_packaging_batch.py images --ocr --output-json /private/tmp/packaging_batch.json
```

Oppure per un PDF singolo:

```bash
.venv/bin/python run_packaging_batch.py /path/to/file.pdf --ocr
```

## Struttura del payload

Il comando restituisce un JSON con questa forma:

```json
{
  "meta": {
    "input_path": "...",
    "total_pdfs": 50,
    "ocr_enabled": true,
    "ocr_dpi": 200
  },
  "records": [
    {
      "file": "8055712770125_100x55x55_Clitofono.pdf",
      "missing_fields_count": 6,
      "analysis": { "...": "..." },
      "anchors": { "...": true },
      "zones": { "...": true },
      "sheet_row": { "...": "..." },
      "fields": {
        "tipo_o_modello": {
          "column": "E",
          "label": "Tipo o modello",
          "mode": "value",
          "value": "Clitofono",
          "raw_value": "Clitofono",
          "source": "filename",
          "confidence": "medium"
        }
      },
      "extra_fields": {
        "ean": "8055712770125"
      },
      "review": {
        "review_needed": true,
        "review_fields_count": 9,
        "source_counts": { "...": 0 },
        "confidence_counts": { "...": 0 },
        "review_fields": [
          {
            "field_name": "simbolo_ce",
            "column": "H",
            "label": "Simbolo CE",
            "value": "✅",
            "raw_value": "✅",
            "source": "inference",
            "confidence": "low"
          }
        ]
      }
    }
  ]
}
```

## Campo da usare per Google Sheets

Per scrivere una riga nel foglio usa:

```text
records[*].sheet_row
```

`sheet_row` e` gia` allineato alle colonne del file Excel business.

## Campo da usare per review manuale

Per capire quali record richiedono controllo umano usa:

```text
records[*].review.review_needed
```

Per vedere perche` un record e` da controllare usa:

```text
records[*].review.review_fields
```

Regole attuali:
- `source=inference` -> review
- `source=missing` -> review
- `confidence=low` -> review

## Workflow n8n consigliato

### 1. Trigger

Uno di questi:
- Google Drive Trigger
- Manual Trigger
- Webhook

Per partire subito, importa [n8n_packaging_workflow.json](/Users/faudor/Desktop/progetti/msc_task2/n8n_packaging_workflow.json:1) e fai prima un test con `Manual Trigger`.

### 2. Recupero file

Se il file arriva da Drive o upload:
- scaricalo in una directory locale accessibile
- passa il path assoluto al comando

### 3. Execute Command

Esempio:

```bash
.venv/bin/python /Users/faudor/Desktop/progetti/msc_task2/run_packaging_batch.py {{$json.input_path}} --ocr
```

Se vuoi evitare payload troppo grandi su stdout:

```bash
.venv/bin/python /Users/faudor/Desktop/progetti/msc_task2/run_packaging_batch.py {{$json.input_path}} --ocr --output-json /private/tmp/packaging_batch.json
cat /private/tmp/packaging_batch.json
```

### 4. Code Node

Nel Code node puoi trasformare il payload in item n8n cosi`:

```javascript
const payload = JSON.parse($json.stdout || $json.data || $json);

return payload.records.map((record) => ({
  json: {
    file: record.file,
    review_needed: record.review.review_needed,
    review_fields_count: record.review.review_fields_count,
    ...record.sheet_row,
  }
}));
```

### 5. IF Node

Smista i record:
- `review_needed = false` -> scrittura diretta su Google Sheets
- `review_needed = true` -> tab di review o coda manuale

### 6. Google Sheets

Scrivi i campi provenienti da:

```text
record.sheet_row
```

Nel workflow scaffold i due tab previsti sono:
- `parsed_ready`
- `needs_review`

I nodi Google Sheets hanno ancora placeholder da sostituire:
- `YOUR_SPREADSHEET_ID`
- eventuali credenziali n8n del tuo account Google

## Adattamenti Minimi Dopo Import

Dopo aver importato il workflow:

1. apri `Set Config`
2. cambia:
   - `repo_dir`
   - `input_path`
   - `output_json`
3. apri i due nodi Google Sheets
4. imposta:
   - credenziali Google
   - `Spreadsheet ID`
   - tab `parsed_ready`
   - tab `needs_review`
5. esegui con `Manual Trigger`

Se il parser girera` su una macchina diversa, il punto da cambiare quasi sicuramente e`:

```text
repo_dir
```

## Strategia operativa consigliata

- primo tab: `parsed_ready`
- secondo tab: `needs_review`
- colonna aggiuntiva: `source_summary` opzionale
- colonna aggiuntiva: `review_fields_count`

Questo ti permette di non bloccare il batch per i casi sporchi.

## Perche` questa integrazione e` la scelta giusta adesso

- niente servizio HTTP da mantenere
- payload JSON stabile
- facile da testare da terminale
- stesso entrypoint per sviluppo locale e automazione
- review manuale esplicita, non implicita
