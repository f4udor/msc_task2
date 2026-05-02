from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from packaging_pdf_parser import IMAGES_DIR, build_automation_payload


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Entry point batch per automazione/n8n: analizza uno o più PDF "
            "e restituisce un JSON stabile."
        )
    )
    parser.add_argument(
        "input_path",
        nargs="?",
        type=Path,
        default=IMAGES_DIR,
        help="PDF singolo o cartella con PDF. Default: ./images",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Attiva OCR fallback con Tesseract.",
    )
    parser.add_argument(
        "--ocr-dpi",
        type=int,
        default=200,
        help="Risoluzione OCR. Default: 200",
    )
    parser.add_argument(
        "--include-ocr-text",
        action="store_true",
        help="Include il testo OCR nel payload JSON. Default: off",
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        help="Salva il payload JSON su file invece che stamparlo solo su stdout.",
    )
    args = parser.parse_args()

    try:
        payload = build_automation_payload(
            args.input_path,
            ocr_enabled=args.ocr,
            ocr_dpi=args.ocr_dpi,
            include_ocr_text=args.include_ocr_text,
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    rendered = json.dumps(payload, indent=2, ensure_ascii=False)
    if args.output_json:
        args.output_json.resolve().write_text(rendered, encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
