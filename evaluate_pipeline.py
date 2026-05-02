from __future__ import annotations

import argparse
import json
import statistics
import time
from pathlib import Path

from test_vectorial_pdf import EXCEL_FIELDS, IMAGES_DIR, build_structured_record, iter_pdfs


NON_EVIDENCE_SOURCES = {"default_false"}
SOFT_EVIDENCE_SOURCES = {"inference"}


def is_filled(value: object) -> bool:
    return value not in (None, "")


def evaluate_records(records: list[dict[str, object]]) -> dict[str, object]:
    total_records = len(records)
    field_names = [field_name for _, field_name, _ in EXCEL_FIELDS]
    total_cells = total_records * len(field_names)

    asserted_filled = 0
    evidence_filled = 0
    strong_evidence_filled = 0
    source_counts: dict[str, int] = {}
    field_stats: list[dict[str, object]] = []

    for field_name in field_names:
        present = 0
        evidence_present = 0
        strong_evidence_present = 0
        per_source: dict[str, int] = {}

        for record in records:
            field = record["fields"][field_name]
            source = str(field["source"])
            value = field["value"]

            source_counts[source] = source_counts.get(source, 0) + 1
            per_source[source] = per_source.get(source, 0) + 1

            if is_filled(value):
                asserted_filled += 1
                present += 1

                if source not in NON_EVIDENCE_SOURCES:
                    evidence_filled += 1
                    evidence_present += 1

                if source not in NON_EVIDENCE_SOURCES | SOFT_EVIDENCE_SOURCES:
                    strong_evidence_filled += 1
                    strong_evidence_present += 1

        label = next(label for _, name, label in EXCEL_FIELDS if name == field_name)
        column = next(column for column, name, _ in EXCEL_FIELDS if name == field_name)
        field_stats.append(
            {
                "column": column,
                "field_name": field_name,
                "label": label,
                "asserted_coverage_pct": round(present / total_records * 100, 2),
                "evidence_coverage_pct": round(evidence_present / total_records * 100, 2),
                "strong_evidence_coverage_pct": round(strong_evidence_present / total_records * 100, 2),
                "source_counts": per_source,
            }
        )

    missing_counts = [record["missing_fields_count"] for record in records]

    return {
        "total_pdfs": total_records,
        "fields_per_pdf": len(field_names),
        "total_cells": total_cells,
        "asserted_fill_pct": round(asserted_filled / total_cells * 100, 2),
        "evidence_fill_pct": round(evidence_filled / total_cells * 100, 2),
        "strong_evidence_fill_pct": round(strong_evidence_filled / total_cells * 100, 2),
        "avg_missing_fields": round(statistics.mean(missing_counts), 2) if missing_counts else 0.0,
        "median_missing_fields": round(statistics.median(missing_counts), 2) if missing_counts else 0.0,
        "source_counts": source_counts,
        "field_stats": field_stats,
    }


def benchmark_run(pdfs: list[Path], *, ocr_enabled: bool, ocr_dpi: int) -> dict[str, object]:
    records: list[dict[str, object]] = []
    durations: list[float] = []

    started = time.perf_counter()
    for pdf_path in pdfs:
        per_pdf_started = time.perf_counter()
        record = build_structured_record(pdf_path, ocr_enabled=ocr_enabled, ocr_dpi=ocr_dpi)
        durations.append(time.perf_counter() - per_pdf_started)
        records.append(record)
    total_duration = time.perf_counter() - started

    evaluation = evaluate_records(records)

    return {
        "ocr_enabled": ocr_enabled,
        "ocr_dpi": ocr_dpi,
        "timing": {
            "total_seconds": round(total_duration, 3),
            "avg_seconds_per_pdf": round(statistics.mean(durations), 3) if durations else 0.0,
            "median_seconds_per_pdf": round(statistics.median(durations), 3) if durations else 0.0,
            "max_seconds_per_pdf": round(max(durations), 3) if durations else 0.0,
            "min_seconds_per_pdf": round(min(durations), 3) if durations else 0.0,
        },
        "evaluation": evaluation,
    }


def build_markdown_report(report: dict[str, object]) -> str:
    no_ocr = report["runs"]["no_ocr"]
    with_ocr = report["runs"]["with_ocr"]
    lines = []
    lines.append("# Pipeline Evaluation")
    lines.append("")
    lines.append(f"- PDF analizzati: {report['dataset']['total_pdfs']}")
    lines.append(f"- Cartella: `{report['dataset']['directory']}`")
    lines.append("")
    lines.append("## Efficienza")
    lines.append("")
    lines.append(f"- Senza OCR: {no_ocr['timing']['total_seconds']}s totali, {no_ocr['timing']['avg_seconds_per_pdf']}s/PDF")
    lines.append(f"- Con OCR: {with_ocr['timing']['total_seconds']}s totali, {with_ocr['timing']['avg_seconds_per_pdf']}s/PDF")
    lines.append("")
    lines.append("## Efficacia")
    lines.append("")
    lines.append(f"- Fill asserted senza OCR: {no_ocr['evaluation']['asserted_fill_pct']}%")
    lines.append(f"- Fill asserted con OCR: {with_ocr['evaluation']['asserted_fill_pct']}%")
    lines.append(f"- Fill con evidenza con OCR: {with_ocr['evaluation']['evidence_fill_pct']}%")
    lines.append(f"- Fill con evidenza forte con OCR: {with_ocr['evaluation']['strong_evidence_fill_pct']}%")
    lines.append(f"- Missing medi con OCR: {with_ocr['evaluation']['avg_missing_fields']}")
    lines.append("")
    lines.append("## Note")
    lines.append("")
    lines.append("- `asserted_fill_pct` conta anche i campi riempiti via `default_false`.")
    lines.append("- `evidence_fill_pct` esclude `default_false`.")
    lines.append("- `strong_evidence_fill_pct` esclude sia `default_false` sia `inference`.")
    lines.append("")
    lines.append("## Top Campi OCR/Inference")
    lines.append("")
    ranked = sorted(
        with_ocr["evaluation"]["field_stats"],
        key=lambda item: item["evidence_coverage_pct"],
        reverse=True,
    )
    for field in ranked[:15]:
        lines.append(
            f"- {field['column']} {field['field_name']}: "
            f"asserted={field['asserted_coverage_pct']}% "
            f"evidence={field['evidence_coverage_pct']}% "
            f"strong={field['strong_evidence_coverage_pct']}%"
        )
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Misura efficacia ed efficienza del parser PDF packaging, "
            "confrontando runs con e senza OCR."
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
        "--ocr-dpi",
        type=int,
        default=200,
        help="Risoluzione OCR. Default: 200",
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        default=Path("pipeline_evaluation.json"),
        help="Percorso del report JSON.",
    )
    parser.add_argument(
        "--output-md",
        type=Path,
        default=Path("pipeline_evaluation.md"),
        help="Percorso del report Markdown.",
    )
    args = parser.parse_args()

    pdfs = iter_pdfs(args.directory.resolve())

    no_ocr = benchmark_run(pdfs, ocr_enabled=False, ocr_dpi=args.ocr_dpi)
    with_ocr = benchmark_run(pdfs, ocr_enabled=True, ocr_dpi=args.ocr_dpi)

    report = {
        "dataset": {
            "directory": str(args.directory.resolve()),
            "total_pdfs": len(pdfs),
        },
        "runs": {
            "no_ocr": no_ocr,
            "with_ocr": with_ocr,
        },
    }

    json_path = args.output_json.resolve()
    md_path = args.output_md.resolve()
    json_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    md_path.write_text(build_markdown_report(report), encoding="utf-8")

    print(json_path)
    print(md_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
