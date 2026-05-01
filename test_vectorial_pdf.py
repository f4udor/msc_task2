from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ModuleNotFoundError:
    print(
        "Errore: manca la libreria PyMuPDF.\n"
        "Installala con:\n"
        "  python3 -m pip install pymupdf\n",
        file=sys.stderr,
    )
    raise SystemExit(1)


IMAGES_DIR = Path(__file__).resolve().parent / "images"

STATIC_FIELDS = {
    "nome_fabbricante": "MySecretCase s.r.l.",
    "indirizzo_fabbricante": "Corso C. Colombo 7 - Milano 20144",
    "nome_importatore": "MySecretCase s.r.l.",
    "indirizzo_importatore": "Corso C. Colombo 7 - Milano 20144",
}

FILENAME_RE = re.compile(
    r"^(?P<ean>\d{13})_"
    r"(?P<pack_width_mm>\d+)x(?P<pack_height_mm>\d+)x(?P<pack_depth_mm>\d+)_"
    r"(?P<product_name>.+)$"
)

REQUIRED_FIELDS = [
    "nome_fabbricante",
    "indirizzo_fabbricante",
    "nome_importatore",
    "indirizzo_importatore",
    "ean",
    "product_name",
    "pack_width_mm",
    "pack_height_mm",
    "pack_depth_mm",
    "codice_smaltimento_scatola",
    "codice_smaltimento_sacchetto",
    "simboli_materiali_smaltimento",
    "contenuto_triman_corretto",
]


def value_entry(value: object, source: str, confidence: str = "high") -> dict[str, object]:
    return {
        "value": value,
        "source": source,
        "confidence": confidence,
    }


def analyze_pdf(pdf_path: Path) -> dict[str, object]:
    """Return vector/text stats and the extracted selectable text."""
    extracted_pages: list[str] = []
    vector_objects = 0
    raster_images = 0

    with fitz.open(pdf_path) as doc:
        for page_number, page in enumerate(doc, start=1):
            drawings = page.get_drawings()
            images = page.get_images(full=True)
            text = page.get_text("text", sort=True).strip()

            vector_objects += len(drawings)
            raster_images += len(images)

            if text:
                extracted_pages.append(
                    f"--- Pagina {page_number} ---\n{text}"
                )

        extracted_text = "\n\n".join(extracted_pages).strip()

        return {
            "pages": doc.page_count,
            "vector_objects": vector_objects,
            "raster_images": raster_images,
            "has_vector_content": vector_objects > 0,
            "has_selectable_text": bool(extracted_text),
            "text": extracted_text,
        }


def extract_words(pdf_path: Path) -> list[dict[str, object]]:
    words: list[dict[str, object]] = []

    with fitz.open(pdf_path) as doc:
        for page_number, page in enumerate(doc, start=1):
            page_width = page.rect.width
            page_height = page.rect.height

            for raw_word in page.get_text("words", sort=True):
                x0, y0, x1, y1, text, block_no, line_no, word_no = raw_word
                words.append(
                    {
                        "text": text,
                        "page": page_number,
                        "x0": x0,
                        "y0": y0,
                        "x1": x1,
                        "y1": y1,
                        "x0_norm": x0 / page_width,
                        "y0_norm": y0 / page_height,
                        "x1_norm": x1 / page_width,
                        "y1_norm": y1 / page_height,
                        "block_no": block_no,
                        "line_no": line_no,
                        "word_no": word_no,
                    }
                )

    return words


def parse_filename(pdf_path: Path) -> dict[str, dict[str, object]]:
    match = FILENAME_RE.match(pdf_path.stem)
    if not match:
        return {
            "filename_parse_error": value_entry(
                f"Nome file non conforme: {pdf_path.name}",
                "filename",
                "high",
            )
        }

    data = match.groupdict()
    return {
        "ean": value_entry(data["ean"], "filename"),
        "product_name": value_entry(data["product_name"].replace("_", " "), "filename"),
        "pack_width_mm": value_entry(int(data["pack_width_mm"]), "filename"),
        "pack_height_mm": value_entry(int(data["pack_height_mm"]), "filename"),
        "pack_depth_mm": value_entry(int(data["pack_depth_mm"]), "filename"),
    }


def parse_static_fields() -> dict[str, dict[str, object]]:
    return {
        field_name: value_entry(field_value, "static")
        for field_name, field_value in STATIC_FIELDS.items()
    }


def parse_disposal_codes(words: list[dict[str, object]]) -> dict[str, dict[str, object]]:
    tokens = [str(word["text"]).upper().strip() for word in words]
    token_set = set(tokens)
    result: dict[str, dict[str, object]] = {}

    has_pap21 = "PAP" in token_set and "21" in token_set
    has_cpe07 = "CPE" in token_set and ("7" in token_set or "07" in token_set)

    if has_pap21:
        result["codice_smaltimento_scatola"] = value_entry("PAP21", "pdf_text")

    if has_cpe07:
        result["codice_smaltimento_sacchetto"] = value_entry("CPE07", "pdf_text")

    if has_pap21 and has_cpe07:
        result["simboli_materiali_smaltimento"] = value_entry(
            "PAP21 / CPE07",
            "pdf_text",
        )
        result["contenuto_triman_corretto"] = value_entry(
            "scatola + sacchetto",
            "pdf_text",
            "medium",
        )
    elif has_pap21:
        result["simboli_materiali_smaltimento"] = value_entry("PAP21", "pdf_text")
        result["contenuto_triman_corretto"] = value_entry(
            "scatola",
            "pdf_text",
            "medium",
        )
    elif has_cpe07:
        result["simboli_materiali_smaltimento"] = value_entry("CPE07", "pdf_text")
        result["contenuto_triman_corretto"] = value_entry(
            "sacchetto",
            "pdf_text",
            "medium",
        )

    return result


def mark_missing_fields(fields: dict[str, dict[str, object]]) -> dict[str, dict[str, object]]:
    for field_name in REQUIRED_FIELDS:
        if field_name not in fields:
            fields[field_name] = value_entry(None, "missing", "low")

    return fields


def build_structured_record(pdf_path: Path) -> dict[str, object]:
    analysis = analyze_pdf(pdf_path)
    words = extract_words(pdf_path)

    fields: dict[str, dict[str, object]] = {}
    fields.update(parse_static_fields())
    fields.update(parse_filename(pdf_path))
    fields.update(parse_disposal_codes(words))
    fields = mark_missing_fields(fields)

    return {
        "file": pdf_path.name,
        "analysis": {
            "pages": analysis["pages"],
            "vector_objects": analysis["vector_objects"],
            "raster_images": analysis["raster_images"],
            "has_vector_content": analysis["has_vector_content"],
            "has_selectable_text": analysis["has_selectable_text"],
            "word_count": len(words),
        },
        "fields": fields,
    }


def iter_pdfs(directory: Path) -> list[Path]:
    return sorted(directory.glob("*.pdf"))


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Verifica se i PDF in una cartella hanno contenuto vettoriale "
            "e testo selezionabile, poi stampa il testo estratto."
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
        "--json",
        action="store_true",
        help="Stampa l'output strutturato JSON invece del report leggibile.",
    )
    args = parser.parse_args()

    pdf_dir = args.directory.resolve()
    if not pdf_dir.is_dir():
        print(f"Cartella non trovata: {pdf_dir}", file=sys.stderr)
        return 1

    pdfs = iter_pdfs(pdf_dir)
    if not pdfs:
        print(f"Nessun PDF trovato in: {pdf_dir}")
        return 0

    records = []

    for pdf_path in pdfs:
        record = build_structured_record(pdf_path)
        records.append(record)

        if args.json:
            continue

        print("=" * 80)
        print(f"PDF: {pdf_path.name}")

        analysis = record["analysis"]
        fields = record["fields"]

        print(f"Pagine: {analysis['pages']}")
        print(f"Oggetti vettoriali rilevati: {analysis['vector_objects']}")
        print(f"Immagini raster rilevate: {analysis['raster_images']}")
        print(f"Parole selezionabili rilevate: {analysis['word_count']}")
        print(
            "Vettoriale: "
            f"{'SI' if analysis['has_vector_content'] else 'NO'}"
        )
        print(
            "Testo selezionabile: "
            f"{'SI' if analysis['has_selectable_text'] else 'NO'}"
        )

        print("\nCAMPI ESTRATTI:")
        for field_name in REQUIRED_FIELDS:
            field = fields[field_name]
            print(
                f"- {field_name}: {field['value']} "
                f"(source={field['source']}, confidence={field['confidence']})"
            )

    if args.json:
        print(json.dumps(records, indent=2, ensure_ascii=False))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
