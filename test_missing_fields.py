from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from packaging_pdf_parser import EXCEL_FIELDS, IMAGES_DIR, build_structured_record, iter_pdfs


FIELD_LABELS = {field_name: label for _, field_name, label in EXCEL_FIELDS}

FIELD_PATTERNS = {
    "simbolo_ce": [r"\bCE\b"],
    "simbolo_raee": [r"\bRAEE\b", r"\bWEEE\b", r"wheeled bin"],
    "simbolo_ukca": [r"\bUKCA\b"],
    "simbolo_triman": [r"\bTRIMAN\b"],
    "simbolo_smaltimento_spagnolo": [r"smaltimento spagnolo", r"spagna", r"spain"],
    "simbolo_libretto_informativo": [r"istruzioni", r"manuale", r"read.*instructions"],
    "numero_velocita": [r"(\d+)\s*speed", r"(\d+)\s*velocit"],
    "numero_modalita_suzione": [r"(\d+)\s*modalit[aà]\s*suz", r"(\d+)\s*suction"],
    "numero_modalita_tapping": [r"(\d+)\s*modalit[aà]\s*tapp", r"(\d+)\s*tapping"],
    "numero_modalita_rotazione": [r"(\d+)\s*modalit[aà]\s*rot", r"(\d+)\s*rotation"],
    "funzione_riscaldante": [r"riscald", r"warming", r"heating"],
    "codice_asin": [r"\bB0[A-Z0-9]{8}\b"],
    "codice_smaltimento_doypack": [r"\bD[O0]YPACK\b", r"\bP[AO]UCH\b"],
    "sexy_ideas": [r"SEXY IDEAS"],
}

DEFAULT_FOCUS_FIELDS = [
    "simbolo_raee",
    "simbolo_ukca",
    "simbolo_triman",
    "simbolo_smaltimento_spagnolo",
    "simbolo_libretto_informativo",
    "numero_velocita",
    "numero_modalita_suzione",
    "numero_modalita_tapping",
    "numero_modalita_rotazione",
    "funzione_riscaldante",
    "codice_asin",
    "codice_smaltimento_doypack",
    "sexy_ideas",
]


def compact_text(text: str, limit: int = 260) -> str:
    normalized = re.sub(r"\s+", " ", text).strip()
    return normalized[:limit]


def search_patterns(text: str, field_name: str) -> list[str]:
    matches: list[str] = []
    for pattern in FIELD_PATTERNS.get(field_name, []):
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            matches.append(match.group(0))
    return matches


def build_missing_fields_report(
    records: list[dict[str, object]],
    focus_fields: list[str],
    sample_limit: int,
) -> dict[str, object]:
    report_fields: list[dict[str, object]] = []

    for field_name in focus_fields:
        missing_records = []
        present_records = []
        keyword_hits = []

        for record in records:
            field = record["fields"][field_name]
            ocr_text = str(record.get("ocr", {}).get("text", ""))
            pattern_hits = search_patterns(ocr_text, field_name)

            if field["source"] == "missing":
                if len(missing_records) < sample_limit:
                    missing_records.append(
                        {
                            "file": record["file"],
                            "missing_fields_count": record["missing_fields_count"],
                            "ocr_preview": compact_text(ocr_text),
                        }
                    )
            else:
                if len(present_records) < sample_limit:
                    present_records.append(
                        {
                            "file": record["file"],
                            "value": field["value"],
                            "source": field["source"],
                        }
                    )

            if pattern_hits and len(keyword_hits) < sample_limit:
                keyword_hits.append(
                    {
                        "file": record["file"],
                        "hits": pattern_hits[:5],
                        "ocr_preview": compact_text(ocr_text),
                    }
                )

        report_fields.append(
            {
                "field_name": field_name,
                "label": FIELD_LABELS[field_name],
                "filled_count": sum(
                    1
                    for record in records
                    if record["fields"][field_name]["source"] != "missing"
                ),
                "missing_count": sum(
                    1
                    for record in records
                    if record["fields"][field_name]["source"] == "missing"
                ),
                "patterns": FIELD_PATTERNS.get(field_name, []),
                "present_examples": present_records,
                "missing_examples": missing_records,
                "keyword_hits": keyword_hits,
            }
        )

    return {
        "total_pdfs": len(records),
        "focus_fields": report_fields,
    }


def print_missing_fields_report(report: dict[str, object]) -> None:
    print(f"PDF analizzati: {report['total_pdfs']}")
    for field in report["focus_fields"]:
        print("=" * 80)
        print(
            f"{field['field_name']} | "
            f"filled={field['filled_count']} "
            f"missing={field['missing_count']}"
        )
        if field["patterns"]:
            print(f"Pattern OCR candidati: {', '.join(field['patterns'])}")

        if field["present_examples"]:
            print("ESEMPI PRESENTI:")
            for item in field["present_examples"]:
                print(
                    f"- {item['file']}: {item['value']} "
                    f"(source={item['source']})"
                )

        if field["keyword_hits"]:
            print("KEYWORD HITS OCR:")
            for item in field["keyword_hits"]:
                print(
                    f"- {item['file']}: hits={item['hits']} | "
                    f"ocr={item['ocr_preview']}"
                )

        if field["missing_examples"]:
            print("ESEMPI MANCANTI:")
            for item in field["missing_examples"]:
                print(
                    f"- {item['file']} "
                    f"(missing_total={item['missing_fields_count']}): "
                    f"{item['ocr_preview'] or '[ocr vuoto]'}"
                )


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Analizza i campi ancora mancanti nel parser PDF e mostra "
            "pattern OCR candidati su cui lavorare."
        )
    )
    parser.add_argument(
        "directory",
        nargs="?",
        type=Path,
        default=IMAGES_DIR,
        help="Cartella con i PDF da analizzare. Default: ./images",
    )
    parser.add_argument(
        "--field",
        action="append",
        dest="fields",
        help="Campo specifico da analizzare. Ripetibile.",
    )
    parser.add_argument(
        "--sample-limit",
        type=int,
        default=3,
        help="Numero di esempi per sezione. Default: 3",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Stampa il report in JSON.",
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        help="Salva il report in JSON su file.",
    )
    args = parser.parse_args()

    pdf_dir = args.directory.resolve()
    pdfs = iter_pdfs(pdf_dir)
    focus_fields = args.fields or DEFAULT_FOCUS_FIELDS

    records = [
        build_structured_record(pdf_path, ocr_enabled=True)
        for pdf_path in pdfs
    ]

    report = build_missing_fields_report(
        records,
        focus_fields=focus_fields,
        sample_limit=args.sample_limit,
    )

    if args.output_json:
        args.output_json.resolve().write_text(
            json.dumps(report, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    if args.json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        print_missing_fields_report(report)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
