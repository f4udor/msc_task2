# Delivery Roadmap

## Obiettivo

Portare il prototipo attuale a una consegna chiara, difendibile e integrabile in n8n.

Il focus non e` aggiungere feature a caso, ma:
- consolidare quello che gia` funziona
- documentare bene cosa e` affidabile e cosa no
- chiudere il flusso end-to-end dal PDF alla riga Google Sheets

## Stato Attuale

Oggi esiste gia`:
- parser core in `packaging_pdf_parser.py`
- wrapper CLI legacy in `test_vectorial_pdf.py`
- entrypoint batch per automazione in `run_packaging_batch.py`
- export Excel in `export_pack_excel.py`
- benchmark in `evaluate_pipeline.py`
- strumenti esplorativi per missing fields e symbol detection
- guida integrazione n8n in `n8n_integration.md`

## Cosa Documentare

### 1. Stack usato

Da documentare in modo esplicito:
- Python
- PyMuPDF per parsing PDF
- Tesseract + pytesseract per OCR fallback
- Pillow per immagini OCR
- openpyxl per export Excel
- n8n come orchestratore previsto
- Google Sheets come destinazione finale

Nota importante:
- il workflow attuale non usa modelli AI come soluzione principale
- l'approccio e` deterministic-first con OCR fallback

### 2. Processo end-to-end

Da spiegare in sequenza:
1. input PDF singolo o cartella
2. parsing filename
3. parsing testo PDF selezionabile
4. OCR fallback se attivato
5. anchor detection
6. zone-based parsing
7. normalizzazione campi
8. generazione `sheet_row`
9. flag `review_needed`
10. output JSON / scrittura futura su Google Sheets

### 3. Gestione errori

Da documentare chiaramente:
- file non PDF -> errore esplicito
- path non trovato -> errore esplicito
- OCR non disponibile -> parser continua senza OCR, ma segnala il problema
- campo mancante -> `source=missing`
- campo inferito -> `source=inference`
- campo debole -> `confidence=low`
- record dubbio -> `review.review_needed = true`

### 4. Accuratezza sui 50 pack

Va documentata con una metrica onesta, non solo con la copertura "apparente".

Numeri attuali:
- `asserted_fill_pct`: 82.88%
- `evidence_fill_pct`: 43.76%
- `strong_evidence_fill_pct`: 40.0%
- `avg_missing_fields`: 5.82

Da aggiungere nel report finale:
- quanti record sono quasi completi
- quanti sono parziali ma utili
- quanti sono deboli / da review

La classificazione consigliata e`:
- `perfetti`: nessun `missing`, nessun `inference`, nessun `low`
- `parziali`: riga utile ma con review necessaria
- `falliti`: troppi campi mancanti o OCR quasi nullo

### 5. Costo per pack a regime

Oggi il costo software e` vicino a zero se:
- usi parsing locale
- usi Tesseract locale
- non usi modelli AI a pagamento

Quindi:
- costo marginale per pack: molto basso
- costo principale: tempo macchina + manutenzione parser

Se in futuro aggiungi AI fallback:
- va documentato separatamente
- il costo non va mescolato con il flusso deterministic-first

### 6. Tempi di elaborazione batch

Numeri attuali:
- senza OCR: `0.099s/PDF`
- con OCR: `1.487s/PDF`

Stima lineare utile da documentare:
- 10 pack con OCR: circa `15s`
- 50 pack con OCR: circa `74s`
- 100 pack con OCR: circa `149s`

Va specificato che:
- il tempo OCR domina il costo computazionale
- senza OCR il parser e` molto veloce

### 7. Limiti e miglioramenti

Limiti attuali:
- alcuni simboli grafici non sono ancora estratti bene
- alcuni campi non dovrebbero venire dal PDF ma da una fonte esterna
- `default_false` migliora la copertura apparente ma non l'evidenza reale
- i layout sono simili ma non identici, quindi esistono edge case

Miglioramenti possibili:
- template matching per simboli
- ROI visuali piu` stabili
- catalogo prodotto esterno per campi non affidabili da PDF
- distinzione piu` netta tra parsing certo e review manuale

## Roadmap Operativa

### Fase 1. Consolidamento del parser

Obiettivo:
- congelare una baseline affidabile prima di toccare n8n

Da fare:
- definire il contratto dei campi
- etichettare ogni campo come:
  - `pdf_native`
  - `ocr_fallback`
  - `inference_allowed`
  - `external_source_only`
- classificare i 50 PDF in:
  - perfetti
  - parziali
  - falliti

Perche`:
- senza questo, automatizzi dati ambigui

### Fase 2. Chiusura del gap principale

Obiettivo:
- migliorare i campi che oggi bloccano l'affidabilita`

Priorita`:
1. simboli grafici
2. campi da fonte esterna
3. review logic

Da fare:
- usare `test_symbol_detection.py` per costruire detection migliore
- decidere quali campi vanno fuori dal parser PDF
- abbassare il numero di record che richiedono review manuale

Perche`:
- questo e` il punto dove migliori davvero la qualita`, non la cosmetica

### Fase 3. Integrazione n8n

Obiettivo:
- rendere il sistema dimostrabile end-to-end

Da fare:
- workflow n8n con `Execute Command`
- parsing del JSON batch
- split `parsed_ready` / `needs_review`
- scrittura su Google Sheets

Perche`:
- e` il primo punto in cui il prototipo diventa una pipeline reale

### Fase 4. Packaging della consegna

Obiettivo:
- rendere il progetto leggibile da chi valuta

Da fare:
- README finale
- architettura sintetica
- metriche finali
- limiti dichiarati in modo onesto
- screenshot o export esempio

Perche`:
- una buona soluzione tecnica senza una buona spiegazione perde molto valore

## Ordine Giusto di Esecuzione

1. fissare contratto dati
2. classificare risultati sui 50 PDF
3. migliorare solo i gap ad alto impatto
4. chiudere il workflow n8n
5. scrivere la documentazione finale

## Messaggio Chiave da Portare in Consegna

La cosa forte di questa soluzione non e` "estrae tutto perfettamente".

La cosa forte e`:
- approccio deterministic-first
- OCR solo come fallback
- output strutturato e spiegabile
- review esplicita per i casi dubbi
- integrazione batch pronta per n8n

Questo rende il progetto credibile dal punto di vista:
- tecnico
- operativo
- economico
