# Pipeline Evaluation

- PDF analizzati: 50
- Cartella: `/Users/faudor/Desktop/progetti/msc_task2/images`

## Efficienza

- Senza OCR: 4.947s totali, 0.099s/PDF
- Con OCR: 74.348s totali, 1.487s/PDF

## Efficacia

- Fill asserted senza OCR: 75.47%
- Fill asserted con OCR: 82.88%
- Fill con evidenza con OCR: 43.76%
- Fill con evidenza forte con OCR: 40.0%
- Missing medi con OCR: 5.82

## Note

- `asserted_fill_pct` conta anche i campi riempiti via `default_false`.
- `evidence_fill_pct` esclude `default_false`.
- `strong_evidence_fill_pct` esclude sia `default_false` sia `inference`.

## Top Campi OCR/Inference

- A nome_fabbricante: asserted=100.0% evidence=100.0% strong=100.0%
- B indirizzo_fabbricante: asserted=100.0% evidence=100.0% strong=100.0%
- C nome_importatore: asserted=100.0% evidence=100.0% strong=100.0%
- D indirizzo_importatore: asserted=100.0% evidence=100.0% strong=100.0%
- E tipo_o_modello: asserted=100.0% evidence=100.0% strong=100.0%
- M simboli_materiali_smaltimento: asserted=92.0% evidence=92.0% strong=92.0%
- AD codice_smaltimento_scatola: asserted=92.0% evidence=92.0% strong=92.0%
- AG contenuto_triman_corretto: asserted=92.0% evidence=92.0% strong=92.0%
- AE codice_smaltimento_sacchetto: asserted=90.0% evidence=90.0% strong=90.0%
- N qr_code_junker: asserted=100.0% evidence=76.0% strong=76.0%
- O simbolo_garanzia_2_anni: asserted=100.0% evidence=70.0% strong=70.0%
- S materiale: asserted=64.0% evidence=64.0% strong=58.0%
- AH sexy_ideas: asserted=100.0% evidence=62.0% strong=62.0%
- K simbolo_triman: asserted=100.0% evidence=56.0% strong=0.0%
- T modalita_ricarica: asserted=46.0% evidence=46.0% strong=26.0%
