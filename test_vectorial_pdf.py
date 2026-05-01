from __future__ import annotations

import argparse
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
    args = parser.parse_args()

    pdf_dir = args.directory.resolve()
    if not pdf_dir.is_dir():
        print(f"Cartella non trovata: {pdf_dir}", file=sys.stderr)
        return 1

    pdfs = iter_pdfs(pdf_dir)
    if not pdfs:
        print(f"Nessun PDF trovato in: {pdf_dir}")
        return 0

    for pdf_path in pdfs:
        print("=" * 80)
        print(f"PDF: {pdf_path.name}")

        try:
            result = analyze_pdf(pdf_path)
        except Exception as exc:
            print(f"Errore durante l'analisi: {exc}")
            continue

        print(f"Pagine: {result['pages']}")
        print(f"Oggetti vettoriali rilevati: {result['vector_objects']}")
        print(f"Immagini raster rilevate: {result['raster_images']}")
        print(
            "Vettoriale: "
            f"{'SI' if result['has_vector_content'] else 'NO'}"
        )
        print(
            "Testo selezionabile: "
            f"{'SI' if result['has_selectable_text'] else 'NO'}"
        )

        if result["has_vector_content"] and result["has_selectable_text"]:
            print("\nTESTO ESTRATTO:")
            print(result["text"])
        else:
            print("\nTesto non stampato: il PDF non soddisfa entrambi i criteri.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
