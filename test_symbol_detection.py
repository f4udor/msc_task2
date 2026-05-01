from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from io import BytesIO
from pathlib import Path

import fitz

from test_vectorial_pdf import IMAGES_DIR, iter_pdfs, load_ocr_dependencies


ROI_PRESETS = {
    "footer_right": (0.58, 0.58, 1.00, 1.00),
    "footer_center": (0.28, 0.58, 0.82, 1.00),
}

SYMBOL_PATTERNS = {
    "simbolo_ce": [r"\bCE\b"],
    "simbolo_ukca": [r"\bUKCA\b"],
    "simbolo_triman": [r"\bTRIMAN\b"],
    "simbolo_raee": [r"\bRAEE\b", r"\bWEEE\b", r"wheeled"],
    "simbolo_libretto_informativo": [r"read", r"instruction", r"manuale", r"istruzioni"],
}


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def average_hash(image: object, size: int = 8) -> str:
    grayscale = image.convert("L").resize((size, size))
    pixels = list(grayscale.getdata())
    avg = sum(pixels) / len(pixels)
    return "".join("1" if value > avg else "0" for value in pixels)


def render_roi(page: fitz.Page, roi: tuple[float, float, float, float], dpi: int) -> object:
    image_cls, _, error = load_ocr_dependencies()
    if error:
        raise RuntimeError(error)

    width = page.rect.width
    height = page.rect.height
    x0, y0, x1, y1 = roi
    clip = fitz.Rect(x0 * width, y0 * height, x1 * width, y1 * height)
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=matrix, alpha=False, clip=clip)
    return image_cls.open(BytesIO(pix.tobytes("png")))


def ocr_roi_text(image: object) -> str:
    _, pytesseract, error = load_ocr_dependencies()
    if error:
        raise RuntimeError(error)
    return pytesseract.image_to_string(image, lang="eng").strip()


def detect_symbol_hits(text: str) -> dict[str, list[str]]:
    hits: dict[str, list[str]] = {}
    for symbol_name, patterns in SYMBOL_PATTERNS.items():
        symbol_hits = []
        for pattern in patterns:
            symbol_hits.extend(
                match.group(0)
                for match in re.finditer(pattern, text, flags=re.IGNORECASE)
            )
        if symbol_hits:
            hits[symbol_name] = symbol_hits
    return hits


def analyze_pdf_symbols(
    pdf_path: Path,
    *,
    dpi: int,
    dump_dir: Path | None,
) -> dict[str, object]:
    with fitz.open(pdf_path) as doc:
        page = doc[0]
        rois = {}
        symbol_counter = Counter()

        for roi_name, roi_box in ROI_PRESETS.items():
            image = render_roi(page, roi_box, dpi=dpi)
            text = ocr_roi_text(image)
            hits = detect_symbol_hits(text)
            symbol_counter.update(hits.keys())

            roi_record = {
                "ocr_text": normalize_text(text),
                "symbol_hits": hits,
                "average_hash": average_hash(image),
            }

            if dump_dir is not None:
                output_dir = dump_dir / pdf_path.stem
                output_dir.mkdir(parents=True, exist_ok=True)
                output_path = output_dir / f"{roi_name}.png"
                image.save(output_path)
                roi_record["image_path"] = str(output_path)

            rois[roi_name] = roi_record

    return {
        "file": pdf_path.name,
        "symbol_hit_count": sum(symbol_counter.values()),
        "unique_symbols": sorted(symbol_counter.keys()),
        "rois": rois,
    }


def build_summary(records: list[dict[str, object]]) -> dict[str, int]:
    counts = Counter()
    for record in records:
        for symbol_name in record["unique_symbols"]:
            counts[symbol_name] += 1
    return dict(counts)


def print_report(records: list[dict[str, object]]) -> None:
    summary = build_summary(records)
    print(f"PDF analizzati: {len(records)}")
    print("Copertura OCR sui simboli:")
    for symbol_name in sorted(SYMBOL_PATTERNS):
        print(f"- {symbol_name}: {summary.get(symbol_name, 0)}")

    print("\nEsempi:")
    for record in records[:5]:
        print("=" * 80)
        print(record["file"])
        print(f"unique_symbols={record['unique_symbols']}")
        for roi_name, roi_data in record["rois"].items():
            print(
                f"- {roi_name}: hits={list(roi_data['symbol_hits'].keys())} "
                f"hash={roi_data['average_hash'][:16]} "
                f"ocr={roi_data['ocr_text'][:180]}"
            )


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Script separato per esplorare symbol detection/template matching "
            "sulle zone footer dei PDF packaging."
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
        "--dpi",
        type=int,
        default=250,
        help="Risoluzione di rendering delle ROI. Default: 250",
    )
    parser.add_argument(
        "--dump-roi-dir",
        type=Path,
        help="Salva le ROI dei simboli come PNG per costruire template.",
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

    records = [
        analyze_pdf_symbols(
            pdf_path,
            dpi=args.dpi,
            dump_dir=args.dump_roi_dir.resolve() if args.dump_roi_dir else None,
        )
        for pdf_path in iter_pdfs(args.directory.resolve())
    ]

    report = {
        "total_pdfs": len(records),
        "summary": build_summary(records),
        "records": records,
    }

    if args.output_json:
        args.output_json.resolve().write_text(
            json.dumps(report, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    if args.json:
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        print_report(records)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
