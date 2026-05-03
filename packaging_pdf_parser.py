from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from io import BytesIO
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

EXCEL_FIELDS = [
    ("A", "nome_fabbricante", "Nome del fabbricante"),
    ("B", "indirizzo_fabbricante", "Indirizzo del fabbricante"),
    ("C", "nome_importatore", "Nome dell'importatore"),
    ("D", "indirizzo_importatore", "Indirizzo dell'importatore"),
    ("E", "tipo_o_modello", "Tipo o modello"),
    ("F", "numero_serie_lotto", "Numero di serie / lotto"),
    ("G", "lotto", "Lotto"),
    ("H", "simbolo_ce", "Simbolo CE"),
    ("I", "simbolo_raee", "Simbolo RAEE"),
    ("J", "simbolo_ukca", "Simbolo UKCA"),
    ("K", "simbolo_triman", "Simbolo TRIMAN"),
    ("L", "simbolo_smaltimento_spagnolo", "Simbolo smaltimento spagnolo"),
    ("M", "simboli_materiali_smaltimento", "Simboli materiali smaltimento"),
    ("N", "qr_code_junker", "QR code Junker"),
    ("O", "simbolo_garanzia_2_anni", "Simbolo garanzia 2 anni"),
    ("P", "simbolo_libretto_informativo", "Simbolo libretto informativo"),
    ("Q", "capacita_batteria_tensione", "Capacità batteria e tensione nominale"),
    ("R", "impermeabilita", "Impermeabilità"),
    ("S", "materiale", "Materiale"),
    ("T", "modalita_ricarica", "Modalità di ricarica"),
    ("U", "dimensioni", "Dimensioni"),
    ("V", "numero_vibrazioni", "N. vibrazioni"),
    ("W", "numero_velocita", "N. velocità"),
    ("X", "numero_modalita_suzione", "N. modalità suzione"),
    ("Y", "numero_modalita_tapping", "N. modalità tapping"),
    ("Z", "numero_modalita_rotazione", "N. modalità rotazione"),
    ("AA", "strap_on_compatibile", "Strap-on compatibile"),
    ("AB", "funzione_riscaldante", "Funzione riscaldante"),
    ("AC", "codice_asin", "Codice ASIN"),
    ("AD", "codice_smaltimento_scatola", "Codice smaltimento scatola"),
    ("AE", "codice_smaltimento_sacchetto", "Codice smaltimento sacchetto"),
    ("AF", "codice_smaltimento_doypack", "Codice smaltimento doypack"),
    ("AG", "contenuto_triman_corretto", "Contenuto TRIMAN corretto"),
    ("AH", "sexy_ideas", "Sexy Ideas"),
]

EXCEL_FIELD_NAMES = [field_name for _, field_name, _ in EXCEL_FIELDS]
PRESENCE_ONLY_FIELDS = {
    "simbolo_ce",
    "simbolo_raee",
    "simbolo_ukca",
    "simbolo_triman",
    "simbolo_smaltimento_spagnolo",
    "qr_code_junker",
    "simbolo_garanzia_2_anni",
    "simbolo_libretto_informativo",
    "strap_on_compatibile",
    "funzione_riscaldante",
    "sexy_ideas",
}
FIELD_MODES = {
    field_name: ("presence_only" if field_name in PRESENCE_ONLY_FIELDS else "value")
    for field_name in EXCEL_FIELD_NAMES
}

FILENAME_RE = re.compile(
    r"^(?P<ean>\d{13})_"
    r"(?P<pack_width_mm>\d+)x(?P<pack_height_mm>\d+)x(?P<pack_depth_mm>\d+)_"
    r"(?P<product_name>.+)$"
)
LOT_RE = re.compile(r"\bLOT[:\s-]*([A-Z0-9-]+)\b", re.IGNORECASE)
IP_RE = re.compile(r"\b[I1]PX?\d+\b", re.IGNORECASE)
BATTERY_RE = re.compile(
    r"Capacit[aà]\s+batteria[:\s]*([0-9]+(?:[.,][0-9]+)?)\s*mAh",
    re.IGNORECASE,
)
VOLTAGE_RE = re.compile(
    r"Tensione\s+nominale\s+batteria[:\s]*([0-9]+(?:[.,][0-9]+)?)\s*V",
    re.IGNORECASE,
)
DIMENSIONS_RE = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*c(?:m|em)?\s*[xX]\s*[@ØO09]?\s*(\d+(?:[.,]\d+)?)\s*c(?:m|em)?",
    re.IGNORECASE,
)
VIBRATIONS_RE = re.compile(r"(\d+)\s*vibrazioni|\b(\d+)\s*vibrations\b", re.IGNORECASE)
SPEED_RE = re.compile(r"(\d+)\s*(?:speed|speeds|velocit[aà]?)", re.IGNORECASE)
SUCTION_RE = re.compile(
    r"(\d+)\s*(?:modalit[aà]\s*(?:di)?\s*suz\w+|suction\s*mode|suction)",
    re.IGNORECASE,
)
TAPPING_RE = re.compile(
    r"(\d+)\s*(?:modalit[aà]\s*tapp\w*|tapping\s*mode|tapping)",
    re.IGNORECASE,
)
ROTATION_RE = re.compile(
    r"(\d+)\s*(?:modalit[aà]\s*rot\w*|rotation\s*mode|rotation)",
    re.IGNORECASE,
)
MATERIAL_RE = re.compile(
    r"\b(Sili\w*(?:[\s,\/-]+ABS)?|ABS(?:[\s,\/-]+Sili\w*)?)\b",
    re.IGNORECASE,
)
EXTRA_MATERIAL_PATTERNS = [
    (re.compile(r"\bTPE\b", re.IGNORECASE), "TPE"),
    (re.compile(r"\bPVC\b", re.IGNORECASE), "PVC"),
    (re.compile(r"\bmetal\b", re.IGNORECASE), "Metal"),
    (re.compile(r"\bpizzo\b", re.IGNORECASE), "Pizzo"),
    (re.compile(r"\bcotone\b", re.IGNORECASE), "Cotone"),
]
CE_BROAD_RE = re.compile(r"C€|\bCE\b|Ce RE|C[E€]\s+c", re.IGNORECASE)
FALSE_DEFAULT_FIELDS = {
    "simbolo_ce",
    "simbolo_raee",
    "simbolo_ukca",
    "simbolo_triman",
    "simbolo_smaltimento_spagnolo",
    "qr_code_junker",
    "simbolo_garanzia_2_anni",
    "simbolo_libretto_informativo",
    "numero_vibrazioni",
    "numero_velocita",
    "numero_modalita_suzione",
    "numero_modalita_tapping",
    "numero_modalita_rotazione",
    "strap_on_compatibile",
    "funzione_riscaldante",
    "codice_smaltimento_doypack",
    "sexy_ideas",
}
ANCHOR_PATTERNS = {
    "has_sexy_ideas": [r"SEXY IDEAS"],
    "has_qr_phrase": [r"TU CI PIACI", r"QR code", r"Scansiona il QR code"],
    "has_material": [r"Silicone", r"\bABS\b", r"\bTPE\b", r"\bPVC\b"],
    "has_ipx": [r"\b[I1]PX?\d+\b", r"Waterproof", r"Impermeabile"],
    "has_vibration": [r"vibraz", r"vibrations"],
    "has_speed": [r"speed", r"velocit"],
    "has_suction": [r"suction", r"suzione"],
    "has_tapping": [r"tapping", r"modalit[aà]\s*tapp"],
    "has_rotation": [r"rotation", r"modalit[aà]\s*rot"],
    "has_charge": [r"Ricarica", r"charge", r"USB-C", r"minijack", r"AAA Batteries"],
    "has_battery": [r"mAh", r"Tensione nominale", r"battery"],
    "has_lot": [r"\bLOT\b"],
    "has_pap": [r"\bPAP\b"],
    "has_cpe": [r"\bCPE\b"],
    "has_company_footer": [r"Prodotto e importato", r"MySecretCase", r"info@mysecretcase"],
    "has_environmental_label": [r"ETICHETTA AMBIENTALE"],
    "has_ce": [r"\bCE\b", r"C€", r"Ce RE"],
    "has_garanzia": [r"Garanzia 2 anni", r"2 ?years warranty"],
    "has_warming": [r"Riscaldante", r"Warming", r"Heating"],
    "has_strap_on": [r"\bStrap-?on\b"],
}


def value_entry(value: object, source: str, confidence: str = "high") -> dict[str, object]:
    return {
        "value": value,
        "source": source,
        "confidence": confidence,
    }


def excel_field_entry(
    field_name: str,
    value: object,
    source: str,
    confidence: str = "high",
) -> dict[str, object]:
    column, _, label = next(
        field for field in EXCEL_FIELDS if field[1] == field_name
    )
    entry = value_entry(value, source, confidence)
    entry["column"] = column
    entry["label"] = label
    entry["mode"] = FIELD_MODES[field_name]
    return entry


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


def load_ocr_dependencies() -> tuple[object | None, object | None, str | None]:
    try:
        from PIL import Image
    except ModuleNotFoundError:
        return None, None, "Manca Pillow. Installa con: .venv/bin/python -m pip install pillow"

    try:
        import pytesseract
    except ModuleNotFoundError:
        return None, None, "Manca pytesseract. Installa con: .venv/bin/python -m pip install pytesseract"

    try:
        pytesseract.get_tesseract_version()
    except Exception:
        return None, None, "Manca il binario tesseract. Installa con: brew install tesseract"

    return Image, pytesseract, None


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


def dump_page_images(pdf_path: Path, output_dir: Path, dpi: int = 200) -> list[str]:
    written_files: list[str] = []
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    pdf_output_dir = output_dir / pdf_path.stem
    pdf_output_dir.mkdir(parents=True, exist_ok=True)

    with fitz.open(pdf_path) as doc:
        for page_number, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            output_path = pdf_output_dir / f"page_{page_number:02d}.png"
            pix.save(output_path)
            written_files.append(str(output_path))

    return written_files


def extract_ocr_data(pdf_path: Path, dpi: int = 200) -> dict[str, object]:
    image_cls, pytesseract, error = load_ocr_dependencies()
    if error:
        return {
            "enabled": False,
            "error": error,
            "pages": [],
            "text": "",
        }

    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    pages: list[dict[str, object]] = []

    with fitz.open(pdf_path) as doc:
        for page_number, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = image_cls.open(BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(image, lang="eng").strip()
            pages.append(
                {
                    "page": page_number,
                    "text": text,
                }
            )

    full_text = "\n\n".join(
        f"--- Pagina {page['page']} ---\n{page['text']}"
        for page in pages
        if page["text"]
    ).strip()

    return {
        "enabled": True,
        "error": None,
        "pages": pages,
        "text": full_text,
    }


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
        "product_name": value_entry(
            data["product_name"].replace("_", " "),
            "filename",
        ),
        "pack_width_mm": value_entry(int(data["pack_width_mm"]), "filename"),
        "pack_height_mm": value_entry(int(data["pack_height_mm"]), "filename"),
        "pack_depth_mm": value_entry(int(data["pack_depth_mm"]), "filename"),
        "pack_dimensions_mm": value_entry(
            (
                f"{data['pack_width_mm']}x"
                f"{data['pack_height_mm']}x"
                f"{data['pack_depth_mm']}"
            ),
            "filename",
        ),
    }


def parse_static_fields() -> dict[str, dict[str, object]]:
    return {
        field_name: excel_field_entry(field_name, field_value, "static")
        for field_name, field_value in STATIC_FIELDS.items()
    }


def parse_disposal_codes(words: list[dict[str, object]]) -> dict[str, dict[str, object]]:
    tokens = [str(word["text"]).upper().strip() for word in words]
    token_set = set(tokens)
    result: dict[str, dict[str, object]] = {}

    has_pap21 = "PAP" in token_set and "21" in token_set
    has_cpe07 = "CPE" in token_set and ("7" in token_set or "07" in token_set)

    if has_pap21:
        result["codice_smaltimento_scatola"] = excel_field_entry(
            "codice_smaltimento_scatola",
            "PAP21",
            "pdf_text",
        )

    if has_cpe07:
        result["codice_smaltimento_sacchetto"] = excel_field_entry(
            "codice_smaltimento_sacchetto",
            "CPE07",
            "pdf_text",
        )

    if has_pap21 and has_cpe07:
        result["simboli_materiali_smaltimento"] = excel_field_entry(
            "simboli_materiali_smaltimento",
            "PAP21 / CPE07",
            "pdf_text",
        )
        result["contenuto_triman_corretto"] = excel_field_entry(
            "contenuto_triman_corretto",
            "scatola + sacchetto",
            "pdf_text",
            "medium",
        )
    elif has_pap21:
        result["simboli_materiali_smaltimento"] = excel_field_entry(
            "simboli_materiali_smaltimento",
            "PAP21",
            "pdf_text",
        )
        result["contenuto_triman_corretto"] = excel_field_entry(
            "contenuto_triman_corretto",
            "scatola",
            "pdf_text",
            "medium",
        )
    elif has_cpe07:
        result["simboli_materiali_smaltimento"] = excel_field_entry(
            "simboli_materiali_smaltimento",
            "CPE07",
            "pdf_text",
        )
        result["contenuto_triman_corretto"] = excel_field_entry(
            "contenuto_triman_corretto",
            "sacchetto",
            "pdf_text",
            "medium",
        )

    return result


def find_anchors(ocr_text: str) -> dict[str, bool]:
    anchors: dict[str, bool] = {}
    for anchor_name, patterns in ANCHOR_PATTERNS.items():
        anchors[anchor_name] = any(
            re.search(pattern, ocr_text, re.IGNORECASE)
            for pattern in patterns
        )
    return anchors


def infer_layout_zones(anchors: dict[str, bool]) -> dict[str, bool]:
    return {
        "marketing_panel": anchors.get("has_sexy_ideas", False),
        "technical_specs_panel": any(
            anchors.get(name, False)
            for name in (
                "has_material",
                "has_ipx",
                "has_vibration",
                "has_speed",
                "has_suction",
                "has_tapping",
                "has_rotation",
                "has_charge",
                "has_battery",
                "has_garanzia",
                "has_warming",
                "has_strap_on",
            )
        ),
        "qr_panel": anchors.get("has_qr_phrase", False),
        "compliance_panel": any(
            anchors.get(name, False)
            for name in (
                "has_lot",
                "has_pap",
                "has_cpe",
                "has_company_footer",
                "has_environmental_label",
                "has_ce",
            )
        ),
        "front_product_panel": anchors.get("has_company_footer", False),
    }


def parse_technical_specs_candidates(
    text: str,
    normalized_text: str,
    anchors: dict[str, bool],
) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}

    if anchors.get("has_battery"):
        battery_match = BATTERY_RE.search(text)
        voltage_match = VOLTAGE_RE.search(text)
        if battery_match:
            battery_value = battery_match.group(1).replace(",", ".")
            battery_text = f"{battery_value}mAh"
            if voltage_match:
                voltage_value = voltage_match.group(1).replace(",", ".")
                battery_text = f"{battery_text} / {voltage_value}V"
            result["capacita_batteria_tensione"] = excel_field_entry(
                "capacita_batteria_tensione",
                battery_text,
                "ocr",
                "medium",
            )

    if anchors.get("has_ipx"):
        ip_match = IP_RE.search(text)
        if ip_match:
            result["impermeabilita"] = excel_field_entry(
                "impermeabilita",
                ip_match.group(0).upper().replace("1P", "IP"),
                "ocr",
                "medium",
            )
        elif re.search(r"Non impermeabile|Not waterproof", text, re.IGNORECASE):
            result["impermeabilita"] = excel_field_entry(
                "impermeabilita",
                "❌",
                "inference",
                "medium",
            )

    if anchors.get("has_material"):
        material_match = MATERIAL_RE.search(normalized_text)
        if material_match:
            material_value = material_match.group(1)
            material_value = re.sub(r"sili\w*", "Silicone", material_value, flags=re.IGNORECASE)
            material_value = re.sub(r"\s+", "", material_value)
            material_value = material_value.replace(",", "/").replace("-", "/")
            result["materiale"] = excel_field_entry(
                "materiale",
                material_value,
                "ocr",
                "medium",
            )
        else:
            for pattern, material_value in EXTRA_MATERIAL_PATTERNS:
                if pattern.search(text):
                    result["materiale"] = excel_field_entry(
                        "materiale",
                        material_value,
                        "inference",
                        "medium",
                    )
                    break

    if any(anchors.get(name, False) for name in ("has_material", "has_charge", "has_vibration")):
        dimensions_match = DIMENSIONS_RE.search(normalized_text)
        if dimensions_match:
            first = dimensions_match.group(1).replace(",", ".")
            second = dimensions_match.group(2).replace(",", ".")
            second_value = float(second)
            first_value = float(first)
            if second_value > first_value and second.startswith("9") and len(second) > 2:
                second = second[1:]
            result["dimensioni"] = excel_field_entry(
                "dimensioni",
                f"{first}cm x Ø{second}cm",
                "ocr",
                "medium",
            )

    if anchors.get("has_charge"):
        if re.search(r"Ricarica magnetica|Magnetic charge", text, re.IGNORECASE):
            result["modalita_ricarica"] = excel_field_entry(
                "modalita_ricarica",
                "Ricarica magnetica",
                "ocr",
                "medium",
            )
        elif re.search(r"USB-C|USB C", text, re.IGNORECASE):
            result["modalita_ricarica"] = excel_field_entry(
                "modalita_ricarica",
                "Ricarica USB-C",
                "inference",
                "medium",
            )
        elif re.search(r"minijack", text, re.IGNORECASE):
            result["modalita_ricarica"] = excel_field_entry(
                "modalita_ricarica",
                "Ricarica minijack",
                "inference",
                "medium",
            )
        elif re.search(r"AAA Batteries|Batterie AAA", text, re.IGNORECASE):
            result["modalita_ricarica"] = excel_field_entry(
                "modalita_ricarica",
                "2 Batterie AAA",
                "inference",
                "medium",
            )

    if anchors.get("has_vibration"):
        vibrations_match = VIBRATIONS_RE.search(text)
        if vibrations_match:
            vibrations_value = vibrations_match.group(1) or vibrations_match.group(2)
            result["numero_vibrazioni"] = excel_field_entry(
                "numero_vibrazioni",
                vibrations_value,
                "ocr",
                "medium",
            )

    if anchors.get("has_speed"):
        speed_match = SPEED_RE.search(text)
        if speed_match:
            result["numero_velocita"] = excel_field_entry(
                "numero_velocita",
                speed_match.group(1),
                "ocr",
                "medium",
            )

    if anchors.get("has_suction"):
        suction_match = SUCTION_RE.search(text)
        if suction_match:
            result["numero_modalita_suzione"] = excel_field_entry(
                "numero_modalita_suzione",
                suction_match.group(1),
                "ocr",
                "medium",
            )

    if anchors.get("has_tapping"):
        tapping_match = TAPPING_RE.search(text)
        if tapping_match:
            result["numero_modalita_tapping"] = excel_field_entry(
                "numero_modalita_tapping",
                tapping_match.group(1),
                "ocr",
                "medium",
            )

    if anchors.get("has_rotation"):
        rotation_match = ROTATION_RE.search(text)
        if rotation_match:
            result["numero_modalita_rotazione"] = excel_field_entry(
                "numero_modalita_rotazione",
                rotation_match.group(1),
                "ocr",
                "medium",
            )

    if anchors.get("has_warming"):
        result["funzione_riscaldante"] = excel_field_entry(
            "funzione_riscaldante",
            "✅",
            "ocr",
            "low",
        )

    if anchors.get("has_strap_on"):
        result["strap_on_compatibile"] = excel_field_entry(
            "strap_on_compatibile",
            "✅",
            "ocr",
            "low",
        )

    if anchors.get("has_garanzia"):
        result["simbolo_garanzia_2_anni"] = excel_field_entry(
            "simbolo_garanzia_2_anni",
            "✅",
            "ocr",
            "low",
        )

    return result


def parse_marketing_candidates(
    text: str,
    anchors: dict[str, bool],
) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    if anchors.get("has_sexy_ideas") and re.search(r"SEXY IDEAS", text, re.IGNORECASE):
        result["sexy_ideas"] = excel_field_entry(
            "sexy_ideas",
            "✅",
            "ocr",
            "medium",
        )
    return result


def parse_qr_candidates(
    text: str,
    anchors: dict[str, bool],
) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    if anchors.get("has_qr_phrase") and re.search(r"Scansiona il QR code|QR code", text, re.IGNORECASE):
        result["qr_code_junker"] = excel_field_entry(
            "qr_code_junker",
            "✅",
            "ocr",
            "low",
        )
    return result


def parse_compliance_candidates(
    text: str,
    anchors: dict[str, bool],
) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}

    if anchors.get("has_lot"):
        lot_match = LOT_RE.search(text)
        if lot_match:
            lot_value = lot_match.group(1).strip()
            result["numero_serie_lotto"] = excel_field_entry(
                "numero_serie_lotto",
                f"LOT: {lot_value}",
                "ocr",
                "medium",
            )
            result["lotto"] = excel_field_entry(
                "lotto",
                lot_value,
                "ocr",
                "medium",
            )

    if anchors.get("has_ce") and CE_BROAD_RE.search(text):
        result["simbolo_ce"] = excel_field_entry(
            "simbolo_ce",
            "✅",
            "inference",
            "low",
        )

    return result


def parse_ocr_candidates(
    ocr_data: dict[str, object],
    anchors: dict[str, bool],
    zones: dict[str, bool],
) -> dict[str, dict[str, object]]:
    if not ocr_data.get("enabled"):
        return {}

    text = str(ocr_data.get("text", ""))
    normalized_text = text.replace("\n", " ")
    result: dict[str, dict[str, object]] = {}

    if zones.get("technical_specs_panel"):
        result.update(parse_technical_specs_candidates(text, normalized_text, anchors))

    if zones.get("marketing_panel"):
        result.update(parse_marketing_candidates(text, anchors))

    if zones.get("qr_panel"):
        result.update(parse_qr_candidates(text, anchors))

    if zones.get("compliance_panel"):
        result.update(parse_compliance_candidates(text, anchors))

    return result


def parse_inferred_candidates(
    fields: dict[str, dict[str, object]],
    ocr_data: dict[str, object] | None,
    anchors: dict[str, bool],
    zones: dict[str, bool],
) -> dict[str, dict[str, object]]:
    if ocr_data is None or not ocr_data.get("enabled"):
        return {}

    text = str(ocr_data.get("text", ""))
    result: dict[str, dict[str, object]] = {}

    if (
        fields.get("contenuto_triman_corretto", {}).get("value")
        and fields.get("simbolo_triman", {}).get("value") in (None, "")
        and zones.get("compliance_panel")
    ):
        result["simbolo_triman"] = excel_field_entry(
            "simbolo_triman",
            "✅",
            "inference",
            "medium",
        )

    if anchors.get("has_sexy_ideas") and re.search(r"SEXY IDEAS", text, re.IGNORECASE):
        result["sexy_ideas"] = excel_field_entry(
            "sexy_ideas",
            "✅",
            "ocr",
            "medium",
        )

    if anchors.get("has_ce") and CE_BROAD_RE.search(text):
        result["simbolo_ce"] = excel_field_entry(
            "simbolo_ce",
            "✅",
            "inference",
            "low",
        )

    if anchors.get("has_charge") and re.search(r"USB-C|USB C", text, re.IGNORECASE):
        result["modalita_ricarica"] = excel_field_entry(
            "modalita_ricarica",
            "Ricarica USB-C",
            "inference",
            "medium",
        )
    elif anchors.get("has_charge") and re.search(r"minijack", text, re.IGNORECASE):
        result["modalita_ricarica"] = excel_field_entry(
            "modalita_ricarica",
            "Ricarica minijack",
            "inference",
            "medium",
        )
    elif anchors.get("has_charge") and re.search(r"AAA Batteries|Batterie AAA", text, re.IGNORECASE):
        result["modalita_ricarica"] = excel_field_entry(
            "modalita_ricarica",
            "2 Batterie AAA",
            "inference",
            "medium",
        )

    if anchors.get("has_ipx") and re.search(r"Non impermeabile|Not waterproof", text, re.IGNORECASE):
        result["impermeabilita"] = excel_field_entry(
            "impermeabilita",
            "❌",
            "inference",
            "medium",
        )

    if anchors.get("has_material") and "materiale" not in result and fields.get("materiale", {}).get("value") in (None, ""):
        for pattern, material_value in EXTRA_MATERIAL_PATTERNS:
            if pattern.search(text):
                result["materiale"] = excel_field_entry(
                    "materiale",
                    material_value,
                    "inference",
                    "medium",
                )
                break

    return result


def parse_filename_candidates(
    filename_fields: dict[str, dict[str, object]],
) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    product_name = filename_fields.get("product_name")

    if product_name:
        result["tipo_o_modello"] = excel_field_entry(
            "tipo_o_modello",
            product_name["value"],
            "filename",
            "medium",
        )

    return result


def apply_false_defaults(fields: dict[str, dict[str, object]]) -> dict[str, dict[str, object]]:
    for field_name in FALSE_DEFAULT_FIELDS:
        if fields.get(field_name, {}).get("value") in (None, ""):
            fields[field_name] = excel_field_entry(
                field_name,
                "❌",
                "default_false",
                "low",
            )

    return fields


def mark_missing_fields(fields: dict[str, dict[str, object]]) -> dict[str, dict[str, object]]:
    for field_name in EXCEL_FIELD_NAMES:
        if field_name not in fields:
            fields[field_name] = excel_field_entry(field_name, None, "missing", "low")

    return fields


def build_sheet_row(fields: dict[str, dict[str, object]]) -> dict[str, object]:
    return {
        label: normalize_field_value(fields[field_name])
        for _, field_name, label in EXCEL_FIELDS
    }


def normalize_field_value(field: dict[str, object]) -> object:
    mode = field.get("mode")
    value = field.get("value")
    if mode == "presence_only":
        if value == "✅":
            return "✅"
        if value == "❌":
            return "❌"
        if value in (None, ""):
            return None
    return value


def count_missing_fields(fields: dict[str, dict[str, object]]) -> int:
    return sum(
        1
        for field_name in EXCEL_FIELD_NAMES
        if fields[field_name]["source"] == "missing"
    )


def is_filled_value(value: object) -> bool:
    return value not in (None, "")


def build_coverage_report(records: list[dict[str, object]]) -> dict[str, object]:
    total = len(records)
    sorted_by_missing = sorted(
        records,
        key=lambda record: (record["missing_fields_count"], record["file"]),
    )

    fields_report: list[dict[str, object]] = []
    for column, field_name, label in EXCEL_FIELDS:
        source_counts = {
            "static": 0,
            "filename": 0,
            "pdf_text": 0,
            "ocr": 0,
            "inference": 0,
            "default_false": 0,
            "missing": 0,
        }
        filled = 0
        examples: list[dict[str, object]] = []

        for record in records:
            field = record["fields"][field_name]
            source = field["source"]
            source_counts[source] = source_counts.get(source, 0) + 1

            if is_filled_value(field["value"]):
                filled += 1
                if len(examples) < 3:
                    examples.append(
                        {
                            "file": record["file"],
                            "value": field["value"],
                            "source": source,
                        }
                    )

        fields_report.append(
            {
                "column": column,
                "field_name": field_name,
                "label": label,
                "filled": filled,
                "total": total,
                "coverage_pct": round((filled / total) * 100, 2) if total else 0.0,
                "source_counts": source_counts,
                "examples": examples,
            }
        )

    return {
        "total_pdfs": total,
        "average_missing_fields": round(
            sum(record["missing_fields_count"] for record in records) / total,
            2,
        ) if total else 0.0,
        "best_pdfs": [
            {
                "file": record["file"],
                "missing_fields_count": record["missing_fields_count"],
            }
            for record in sorted_by_missing[:10]
        ],
        "worst_pdfs": [
            {
                "file": record["file"],
                "missing_fields_count": record["missing_fields_count"],
            }
            for record in sorted_by_missing[-10:]
        ],
        "fields": fields_report,
    }


def print_coverage_report(report: dict[str, object]) -> None:
    print("=" * 80)
    print("COVERAGE REPORT")
    print(f"PDF analizzati: {report['total_pdfs']}")
    print(f"Media campi mancanti: {report['average_missing_fields']}")

    print("\nMIGLIORI 10 PDF:")
    for item in report["best_pdfs"]:
        print(f"- {item['missing_fields_count']} missing: {item['file']}")

    print("\nPEGGIORI 10 PDF:")
    for item in report["worst_pdfs"]:
        print(f"- {item['missing_fields_count']} missing: {item['file']}")

    print("\nCOPERTURA PER CAMPO:")
    for field in sorted(report["fields"], key=lambda item: (item["filled"], item["column"])):
        sources = field["source_counts"]
        print(
            f"- {field['column']} {field['field_name']}: "
            f"{field['filled']}/{field['total']} "
            f"({field['coverage_pct']}%) "
            f"[static={sources['static']} "
            f"filename={sources['filename']} "
            f"pdf={sources['pdf_text']} "
            f"ocr={sources['ocr']} "
            f"inference={sources['inference']} "
            f"default_false={sources['default_false']} "
            f"missing={sources['missing']}]"
        )


def write_sheet_csv(records: list[dict[str, object]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[label for _, _, label in EXCEL_FIELDS],
        )
        writer.writeheader()
        for record in records:
            writer.writerow(record["sheet_row"])


def build_structured_record(
    pdf_path: Path,
    *,
    ocr_enabled: bool = False,
    ocr_dpi: int = 200,
    dump_images_dir: Path | None = None,
) -> dict[str, object]:
    analysis = analyze_pdf(pdf_path)
    words = extract_words(pdf_path)
    filename_fields = parse_filename(pdf_path)
    ocr_data = extract_ocr_data(pdf_path, dpi=ocr_dpi) if ocr_enabled else None
    anchors = find_anchors(str(ocr_data.get("text", ""))) if ocr_data is not None else {}
    zones = infer_layout_zones(anchors)
    image_files = (
        dump_page_images(pdf_path, dump_images_dir, dpi=ocr_dpi)
        if dump_images_dir is not None
        else []
    )

    fields: dict[str, dict[str, object]] = {}
    fields.update(parse_static_fields())
    fields.update(parse_filename_candidates(filename_fields))
    fields.update(parse_disposal_codes(words))
    if ocr_data is not None:
        fields.update(parse_ocr_candidates(ocr_data, anchors, zones))
        fields.update(parse_inferred_candidates(fields, ocr_data, anchors, zones))
    fields = apply_false_defaults(fields)
    fields = mark_missing_fields(fields)

    record = {
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
        "extra_fields": filename_fields,
        "anchors": anchors,
        "zones": zones,
        "review": build_review_summary(fields),
        "sheet_row": build_sheet_row(fields),
        "missing_fields_count": count_missing_fields(fields),
    }

    if ocr_data is not None:
        record["ocr"] = ocr_data

    if image_files:
        record["debug_images"] = image_files

    return record


def resolve_pdf_inputs(input_path: Path) -> list[Path]:
    resolved_path = input_path.resolve()
    if resolved_path.is_file():
        if resolved_path.suffix.lower() != ".pdf":
            raise ValueError(f"Il file non è un PDF: {resolved_path}")
        return [resolved_path]
    if resolved_path.is_dir():
        return iter_pdfs(resolved_path)
    raise ValueError(f"Percorso non trovato: {resolved_path}")


def iter_pdfs(directory: Path) -> list[Path]:
    return sorted(directory.glob("*.pdf"))


def simplify_field_entry(field: dict[str, object]) -> dict[str, object]:
    return {
        "column": field["column"],
        "label": field["label"],
        "mode": field["mode"],
        "value": normalize_field_value(field),
        "raw_value": field["value"],
        "source": field["source"],
        "confidence": field["confidence"],
    }


def build_review_summary(fields: dict[str, dict[str, object]]) -> dict[str, object]:
    source_counts: dict[str, int] = {}
    confidence_counts: dict[str, int] = {}
    review_fields: list[dict[str, object]] = []

    for field_name, field in fields.items():
        source = str(field["source"])
        confidence = str(field["confidence"])
        source_counts[source] = source_counts.get(source, 0) + 1
        confidence_counts[confidence] = confidence_counts.get(confidence, 0) + 1

        if source in {"inference", "missing"} or confidence == "low":
            review_fields.append(
                {
                    "field_name": field_name,
                    "column": field["column"],
                    "label": field["label"],
                    "value": normalize_field_value(field),
                    "raw_value": field["value"],
                    "source": source,
                    "confidence": confidence,
                }
            )

    review_fields.sort(key=lambda item: item["column"])
    return {
        "review_needed": bool(review_fields),
        "review_fields_count": len(review_fields),
        "source_counts": source_counts,
        "confidence_counts": confidence_counts,
        "review_fields": review_fields,
    }


def build_automation_record(
    record: dict[str, object],
    *,
    include_ocr_text: bool = False,
) -> dict[str, object]:
    simplified_fields = {
        field_name: simplify_field_entry(field)
        for field_name, field in record["fields"].items()
    }
    simplified_extra_fields = {
        field_name: field["value"]
        for field_name, field in record["extra_fields"].items()
    }
    automation_record = {
        "file": record["file"],
        "missing_fields_count": record["missing_fields_count"],
        "analysis": record["analysis"],
        "anchors": record["anchors"],
        "zones": record["zones"],
        "sheet_row": record["sheet_row"],
        "fields": simplified_fields,
        "extra_fields": simplified_extra_fields,
        "review": build_review_summary(record["fields"]),
    }

    if include_ocr_text and "ocr" in record:
        automation_record["ocr"] = {
            "enabled": record["ocr"]["enabled"],
            "error": record["ocr"]["error"],
            "text": record["ocr"]["text"],
        }

    return automation_record


def build_automation_payload(
    input_path: Path,
    *,
    ocr_enabled: bool = False,
    ocr_dpi: int = 200,
    include_ocr_text: bool = False,
) -> dict[str, object]:
    pdf_paths = resolve_pdf_inputs(input_path)
    records = [
        build_structured_record(
            pdf_path,
            ocr_enabled=ocr_enabled,
            ocr_dpi=ocr_dpi,
        )
        for pdf_path in pdf_paths
    ]
    return {
        "meta": {
            "input_path": str(input_path.resolve()),
            "total_pdfs": len(records),
            "ocr_enabled": ocr_enabled,
            "ocr_dpi": ocr_dpi,
        },
        "records": [
            build_automation_record(
                record,
                include_ocr_text=include_ocr_text,
            )
            for record in records
        ],
    }


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
    parser.add_argument(
        "--export-csv",
        type=Path,
        help="Esporta le righe allineate allo schema Excel in un CSV.",
    )
    parser.add_argument(
        "--dump-images",
        type=Path,
        help="Salva PNG delle pagine PDF per debug visuale.",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="Attiva OCR fallback con Tesseract sui PDF renderizzati.",
    )
    parser.add_argument(
        "--ocr-dpi",
        type=int,
        default=200,
        help="Risoluzione di rendering per OCR e dump immagini. Default: 200",
    )
    parser.add_argument(
        "--coverage-report",
        action="store_true",
        help="Stampa un riepilogo di copertura dei campi su tutti i PDF analizzati.",
    )
    parser.add_argument(
        "--coverage-json",
        type=Path,
        help="Salva il coverage report in formato JSON.",
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
        record = build_structured_record(
            pdf_path,
            ocr_enabled=args.ocr,
            ocr_dpi=args.ocr_dpi,
            dump_images_dir=args.dump_images,
        )
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
        print(f"Campi Excel mancanti: {record['missing_fields_count']}")
        print(
            "Vettoriale: "
            f"{'SI' if analysis['has_vector_content'] else 'NO'}"
        )
        print(
            "Testo selezionabile: "
            f"{'SI' if analysis['has_selectable_text'] else 'NO'}"
        )

        print("\nCAMPI ESTRATTI:")
        for field_name in EXCEL_FIELD_NAMES:
            field = fields[field_name]
            print(
                f"- {field['column']} {field_name}: {field['value']} "
                f"(source={field['source']}, confidence={field['confidence']})"
            )

        print("\nCAMPI EXTRA DAL FILENAME:")
        for field_name, field in record["extra_fields"].items():
            print(
                f"- {field_name}: {field['value']} "
                f"(source={field['source']}, confidence={field['confidence']})"
            )

        if args.ocr:
            ocr_data = record.get("ocr", {})
            print("\nOCR:")
            print(
                f"- enabled: {ocr_data.get('enabled')}"
            )
            if ocr_data.get("error"):
                print(f"- error: {ocr_data['error']}")
            else:
                ocr_text = str(ocr_data.get("text", ""))
                preview = ocr_text[:300].replace("\n", " | ")
                print(f"- preview: {preview or '[vuoto]'}")

        if record.get("debug_images"):
            print("\nDEBUG IMAGES:")
            for image_file in record["debug_images"]:
                print(f"- {image_file}")

    if args.export_csv:
        write_sheet_csv(records, args.export_csv.resolve())

    if args.coverage_report or args.coverage_json:
        coverage_report = build_coverage_report(records)
        if args.coverage_report and not args.json:
            print_coverage_report(coverage_report)
        if args.coverage_json:
            args.coverage_json.resolve().write_text(
                json.dumps(coverage_report, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

    if args.json:
        print(json.dumps(records, indent=2, ensure_ascii=False))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
