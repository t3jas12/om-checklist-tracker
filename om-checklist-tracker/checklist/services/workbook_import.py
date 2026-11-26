import csv
import io
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from django.db import transaction
from openpyxl import Workbook, load_workbook

from checklist.models import ChecklistItem, ChecklistTemplate, Site, SiteEquipment


_FREQ_MAP = {
    "d": "Daily",
    "daily": "Daily",
    "w": "Weekly",
    "weekly": "Weekly",
    "m": "Monthly",
    "monthly": "Monthly",
}

_FOOTER_PREFIXES = (
    "legends",
    "remarks",
    "signature",
    "if ok",
    "if not ok",
    "d - daily",
    "d -daily",
    "w - weekly",
    "w -weekly",
    "note :",
    "note:",
    "checked & done",
    "check & done",
    "verify by",
    "approved by",
)


def _cell_text(value):
    return "" if value is None else str(value).strip()


def _normalize_frequency(value):
    text = _cell_text(value)
    return _FREQ_MAP.get(text.lower(), text)


def _extract_sheet_metadata(sheet):
    """Extract document_number, company, and equipment_name from top rows of a sheet."""
    document_number = ""
    company = ""
    equipment_name = ""
    for row in sheet.iter_rows(min_row=1, max_row=6, values_only=True):
        for value in row:
            text = _cell_text(value)
            if not text:
                continue
            doc_match = re.match(r"^document\s*no\.?\s*:?\s*(.+)$", text, re.IGNORECASE)
            if doc_match and not document_number:
                document_number = doc_match.group(1).strip()
                continue
            equip_match = re.match(r"^equipment\s*name\s*:?\s*(.+)$", text, re.IGNORECASE)
            if equip_match and not equipment_name:
                equipment_name = equip_match.group(1).strip()
                continue
            low = text.lower()
            if not company and any(tok in low for tok in ("energy india", "pvt ltd", "private limited", "pvt. ltd")):
                company = text
    return document_number, company, equipment_name


def _answer_type(description, standard):
    desc_low = (description or "").lower()
    text = f"{desc_low} {(standard or '').lower()}"
    is_temp_reading = (
        "temperature" in text
        and "temperature sensor" not in text
        and not (
            "visual inspection" in desc_low
            and not any(k in text for k in ("notedown", "note down", "°c", "below", "≤", "<="))
        )
    )
    if is_temp_reading:
        return ChecklistItem.AnswerType.NUMBER
    if any(phrase in text for phrase in ("value notedown", "value note down", "note down", "notedown")):
        return ChecklistItem.AnswerType.TEXT
    if any(
        word in text
        for word in (
            "should be",
            "must be",
            "no leakage",
            "no sign",
            "no damage",
            "should open",
            "should be clean",
            "no dirt",
            "no any sign",
            "no spark",
            "no humming",
            "confirm with",
            "rustfree",
            "rust free",
            "working condition",
            "working condtion",
            "functional ok",
            "funtional ok",
        )
    ):
        return ChecklistItem.AnswerType.BOOLEAN
    return ChecklistItem.AnswerType.TEXT


def import_workbook(path, site_name=None):
    """Import checklist definitions from the sample workbook, ignoring daily values."""
    path = Path(path)
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["spv"] if "spv" in workbook.sheetnames else workbook.active
    document_number = ""
    company = ""
    equipment = ""
    for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
        for value in row:
            text = _cell_text(value)
            if text.startswith("Document No."):
                document_number = text.partition(":")[2].strip()
            elif "Energy India" in text or "Pvt Ltd" in text:
                company = text
            elif text.startswith("Equipment Name:"):
                equipment = text.partition(":")[2].strip()

    template, _ = ChecklistTemplate.objects.get_or_create(
        name="Trafo Daily Checklist", version=1,
        defaults={"document_number": document_number, "company": company, "equipment_name": equipment},
    )
    if template.items.exists():
        return template, 0

    current_section = ""
    created = 0
    for row in sheet.iter_rows(min_row=7, max_col=7, values_only=True):
        section, checkpoint_no, description, standard, method, frequency = row[1], row[2], row[3], row[4], row[5], row[6]
        if section:
            current_section = _cell_text(section)
        if not checkpoint_no or not description:
            continue
        description = _cell_text(description)
        standard = _cell_text(standard)
        kind = _answer_type(description, standard)
        unit = "°C" if "temperature" in description.lower() else ""
        ChecklistItem.objects.create(
            template=template,
            section=current_section,
            checkpoint_no=int(checkpoint_no),
            description=description,
            standard=standard,
            method=_cell_text(method),
            frequency=_cell_text(frequency),
            answer_type=kind,
            unit=unit,
            position=int(checkpoint_no),
        )
        created += 1

    if site_name:
        site, _ = Site.objects.get_or_create(
            name=site_name,
            defaults={"company": company, "location": "Jhatikara DL9"},
        )
        if equipment:
            SiteEquipment.objects.get_or_create(site=site, name=equipment)
    workbook.close()
    return template, created


def _norm(v):
    return re.sub(r'[\s\n_/]+', ' ', _cell_text(v)).lower().strip()


# Substring hint lists for header detection
_CP_HINTS = ["checkpoint", "check point", "point no", "s.no", "sl.no", "sr.no",
             "sl no", "sr no", "s/n", "item no", "no.", "no", "sl", "sr", "#"]
_DESC_HINTS = ["description", "job description", "observation", "activity",
               "parameter", "particulars", "check", "item", "title"]
_STD_HINTS = ["standard", "specification", "acceptance", "criteria", "limits", "norm", "requirement"]
_MTH_HINTS = ["method", "tool", "instrument", "procedure"]
_FRQ_HINTS = ["frequency", "freq", "period", "interval"]
_SEC_HINTS = ["section", "category", "system", "subsystem", "area", "group", "equipment"]
_STATUS_HINTS = ["operational", "status", "condition", "working", "inspection", "check", "state", "health", "result", "ok"]

# Strings that mark a header-only row (not a data row)
_HEADER_KEYWORDS = frozenset([
    "description", "job description", "observation", "check", "item",
    "checkpoint", "check point", "standard", "specification", "method", "tool",
    "frequency", "freq", "section", "category", "s.no", "sl.no", "sr.no", "sl", "sr", "no",
])


def _resolve_validation_options(formula1, wb_data, current_sheet):
    """
    Resolve an Excel data-validation list formula into a clean comma-separated
    string of options. Supports inline lists ("A,B,C"), cell ranges
    (=$Z$1:$Z$3 or =Lists!$A$1:$A$4), and workbook defined names (=Options).
    """
    f1 = (formula1 or "").strip()
    if not f1:
        return ""

    # 1. Inline list: "\"A,B,C\""
    if f1.startswith('"') and f1.endswith('"'):
        raw = f1[1:-1]
        sep = ";" if ";" in raw and "," not in raw else ","
        opts = [o.strip() for o in raw.split(sep) if o.strip()]
        return ",".join(opts)

    expr = f1[1:].strip() if f1.startswith("=") else f1

    # 2. Defined / Named range lookup
    if wb_data is not None and hasattr(wb_data, "defined_names") and expr in wb_data.defined_names:
        defn = wb_data.defined_names[expr]
        opts = []
        for s_name, coord in defn.destinations:
            if s_name in wb_data.sheetnames:
                opts.extend(_read_range_values(wb_data[s_name], coord))
        if opts:
            return ",".join(opts)

    # 3. Direct or cross-sheet cell range (e.g. "Lists!$A$1:$A$3" or "$Z$1:$Z$4")
    target_sheet = current_sheet
    coord = expr
    if "!" in expr:
        sheet_part, _, coord = expr.partition("!")
        sheet_part = sheet_part.strip().strip("'")
        if wb_data is not None and sheet_part in wb_data.sheetnames:
            target_sheet = wb_data[sheet_part]
        else:
            target_sheet = None

    if target_sheet is not None and coord:
        opts = _read_range_values(target_sheet, coord)
        if opts:
            return ",".join(opts)

    return ""


def _read_range_values(sheet, coord):
    """Read non-empty string values from an Excel cell or range coordinate."""
    clean_coord = coord.replace("$", "").strip()
    try:
        cells = sheet[clean_coord]
    except Exception:
        return []
    vals = []
    if hasattr(cells, "value"):
        t = _cell_text(cells.value)
        if t:
            vals.append(t)
    elif isinstance(cells, tuple):
        for item in cells:
            if isinstance(item, tuple):
                for cell in item:
                    t = _cell_text(cell.value)
                    if t:
                        vals.append(t)
            else:
                t = _cell_text(item.value)
                if t:
                    vals.append(t)
    return vals


def _extract_dropdown_info(ws_val, wb_data, current_sheet):
    """
    Extract Excel list data validations as:
      - col_options: {0-indexed col -> options_str}
      - dv_ranges: list of (min_row_1b, max_row_1b, min_col_0b, max_col_0b, options_str)
    """
    col_options = {}
    dv_ranges = []
    if ws_val and hasattr(ws_val, "data_validations"):
        for dv in ws_val.data_validations.dataValidation:
            if dv.type == "list":
                options_str = _resolve_validation_options(dv.formula1, wb_data, current_sheet)
                if not options_str:
                    continue
                for rng in dv.sqref.ranges:
                    dv_ranges.append((rng.min_row, rng.max_row, rng.min_col - 1, rng.max_col - 1, options_str))
                    for col_idx in range(rng.min_col, rng.max_col + 1):
                        col_options[col_idx - 1] = options_str
    return col_options, dv_ranges


def _get_dropdown_cols(ws):
    """Backward-compatible wrapper returning {0-indexed col -> options_str}."""
    col_options, _ = _extract_dropdown_info(ws, None, ws)
    return col_options


def _lookup_row_dropdown(dv_ranges, row_idx_1b, min_value_col_0b, fallback_cols=()):
    """
    Find the dropdown options string for a given 1-based Excel row index.
    Prioritizes data/day columns (>= min_value_col_0b) and prefers the most
    specific (smallest row-span) rule when multiple validations overlap.
    """
    matches = [
        (max_r - min_r, opts)
        for min_r, max_r, min_c, max_c, opts in dv_ranges
        if min_r <= row_idx_1b <= max_r and max_c >= min_value_col_0b
    ]
    if matches:
        matches.sort(key=lambda pair: pair[0])
        return matches[0][1]

    for col_idx in fallback_cols:
        if col_idx is None:
            continue
        col_matches = [
            (max_r - min_r, opts)
            for min_r, max_r, min_c, max_c, opts in dv_ranges
            if min_r <= row_idx_1b <= max_r and min_c <= col_idx <= max_c
        ]
        if col_matches:
            col_matches.sort(key=lambda pair: pair[0])
            return col_matches[0][1]
    return ""


def _classify_options(options_str):
    """
    Given a comma-separated option string, return (answer_type, choices_str):
    - If only Yes/No/OK/Pass/True/False variants → BOOLEAN, choices_str=""
    - Otherwise → CHOICES, choices_str=options_str
    """
    if not options_str:
        return ChecklistItem.AnswerType.BOOLEAN, ""
    opts = [o.strip() for o in options_str.split(",") if o.strip()]
    opts_lower = {o.lower() for o in opts}
    boolean_words = {"yes", "no", "ok", "pass", "fail", "true", "false", "y", "n"}
    if opts_lower and opts_lower.issubset(boolean_words):
        return ChecklistItem.AnswerType.BOOLEAN, ""
    return ChecklistItem.AnswerType.CHOICES, ",".join(opts)


def _is_header_row(row_texts):
    """Return True if the row looks like a header (no useful data)."""
    non_empty = [t for t in row_texts if t]
    if not non_empty:
        return False
    has_cp = any(any(h in t for h in _CP_HINTS) for t in non_empty)
    has_desc = any(any(h in t for h in _DESC_HINTS) for t in non_empty)
    keyword_hits = sum(1 for t in non_empty if t in _HEADER_KEYWORDS or any(h in t for h in _DESC_HINTS + _CP_HINTS))
    numeric_cells = sum(1 for t in non_empty if t.isdigit() or (t.replace('.', '', 1).isdigit() and t.count('.') <= 1))
    return has_cp and has_desc and keyword_hits >= 2 and numeric_cells == 0


def _is_footer_row(row_texts_raw):
    """Return True if a row marks the footer / legends / signature area."""
    for val in row_texts_raw:
        low = _cell_text(val).lower()
        if low and low.startswith(_FOOTER_PREFIXES):
            return True
    return False


def _detect_strategy1_columns(sheet):
    """
    Scan rows 1..15 of a sheet for a Strategy 1 table header containing distinct
    checkpoint and description columns (ignoring Document No. / Equipment metadata rows).
    Returns (header_row, col_map, header_texts) or (None, {}, []).
    """
    for ri, row in enumerate(sheet.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
        texts = [_norm(v) for v in row]
        # Ignore metadata banner rows (e.g. "Document No. WMS/CHECKLIST" contains "no." and "check")
        if any(t.startswith(("document no", "equipment name")) for t in texts if t):
            continue
        col_map = {}
        for ci, t in enumerate(texts):
            if not t:
                continue
            if "checkpoint" not in col_map and any(h in t for h in _CP_HINTS):
                col_map["checkpoint"] = ci
            elif "description" not in col_map and any(h in t for h in _DESC_HINTS):
                col_map["description"] = ci
            elif "standard" not in col_map and any(h in t for h in _STD_HINTS):
                col_map["standard"] = ci
            elif "method" not in col_map and any(h in t for h in _MTH_HINTS):
                col_map["method"] = ci
            elif "frequency" not in col_map and any(h in t for h in _FRQ_HINTS):
                col_map["frequency"] = ci
            elif "section" not in col_map and any(h in t for h in _SEC_HINTS):
                col_map["section"] = ci
        if "checkpoint" in col_map and "description" in col_map:
            return ri, col_map, texts
    return None, {}, []


def _import_workbook_pair(wb_data, wb_val, template_name):
    """Shared parser for an already-loaded openpyxl workbook pair."""
    sheets_to_process = []
    if "spv" in wb_data.sheetnames:
        sheets_to_process.append("spv")
    for name in wb_data.sheetnames:
        if name != "spv":
            sheets_to_process.append(name)

    strategy1_headers = {}
    for name in sheets_to_process:
        header_row, col_map, header_texts = _detect_strategy1_columns(wb_data[name])
        if header_row is not None:
            strategy1_headers[name] = (header_row, col_map, header_texts)

    parsed_items = []
    checkpoint_counter = 0
    seen_checkpoints = set()
    doc_number = ""
    company = ""
    equipment_name = ""

    for sheet_name in sheets_to_process:
        sheet = wb_data[sheet_name]
        ws_val = wb_val[sheet_name] if wb_val and sheet_name in wb_val.sheetnames else None
        dropdown_cols, dv_ranges = _extract_dropdown_info(ws_val, wb_data, sheet)

        s_doc, s_comp, s_equip = _extract_sheet_metadata(sheet)
        doc_number = doc_number or s_doc
        company = company or s_comp
        equipment_name = equipment_name or s_equip

        if sheet_name in strategy1_headers:
            # STRATEGY 1: Standard Checklist (Trafo / 11KV / WMS style)
            header_row, col_map, header_texts = strategy1_headers[sheet_name]
            col_map = dict(col_map)

            if (
                "section" in col_map
                and "checkpoint" in col_map
                and col_map["checkpoint"] - col_map["section"] == 2
                and not header_texts[col_map["section"] + 1]
            ):
                col_map["subsection"] = col_map["section"] + 1

            # Determine the first daily/value column (e.g. "Month-" or column after metadata)
            value_start_col = None
            for ci, t in enumerate(header_texts):
                if t.startswith("month") or t in {"value", "status", "result", "reading", "observation"}:
                    if ci not in col_map.values():
                        value_start_col = ci
                        break
            if value_start_col is None:
                value_start_col = (max(col_map.values()) + 1) if col_map else 0

            current_section = ""
            current_subsection = ""
            start_row_idx = (header_row or 1) + 1
            for ri, row in enumerate(sheet.iter_rows(min_row=start_row_idx, values_only=True), start=start_row_idx):
                row = list(row)
                if not any(v is not None and _cell_text(v) for v in row):
                    continue
                if _is_footer_row(row):
                    continue
                row_texts = [_norm(v) for v in row]
                if _is_header_row(row_texts):
                    continue

                def gcol(k):
                    idx = col_map.get(k)
                    return _cell_text(row[idx]) if idx is not None and idx < len(row) else ""

                sec = gcol("section")
                sub = gcol("subsection")
                if sec and not sec.isdigit():
                    current_section = sec
                    if not sub:
                        current_subsection = ""
                if sub and not sub.isdigit():
                    current_subsection = sub

                desc = gcol("description")
                if not desc or _norm(desc) in _HEADER_KEYWORDS:
                    continue

                cp_raw = gcol("checkpoint")
                try:
                    cp = int(float(cp_raw)) if cp_raw else None
                except (ValueError, TypeError):
                    cp = None

                if cp is None or cp <= 0 or cp in seen_checkpoints:
                    checkpoint_counter += 1
                    cp = checkpoint_counter
                else:
                    checkpoint_counter = max(checkpoint_counter, cp)
                seen_checkpoints.add(cp)

                std = gcol("standard")
                mth = gcol("method")
                frq = _normalize_frequency(gcol("frequency")) or "Daily"

                dropdown_opts = _lookup_row_dropdown(
                    dv_ranges,
                    ri,
                    value_start_col,
                    fallback_cols=(col_map.get("description"), col_map.get("standard")),
                )
                if dropdown_opts:
                    kind, choices_str = _classify_options(dropdown_opts)
                else:
                    kind = _answer_type(desc, std)
                    choices_str = ""

                unit = "°C" if kind == ChecklistItem.AnswerType.NUMBER and "temperature" in desc.lower() else ""

                if current_section and current_subsection:
                    combined_section = f"{current_section} — {current_subsection}"
                else:
                    combined_section = current_section or current_subsection or sheet_name

                parsed_items.append({
                    "checkpoint_no": cp,
                    "section": combined_section,
                    "description": desc,
                    "standard": std,
                    "method": mth,
                    "frequency": frq,
                    "answer_type": kind,
                    "choices": choices_str,
                    "unit": unit,
                    "position": len(parsed_items) + 1,
                })

        else:
            # STRATEGY 2: Equipment / Component Matrix Table (e.g. Dhansa Inverter1)
            matrix_header_row = None
            headers = []
            for ri, row in enumerate(sheet.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
                row_texts = [_cell_text(v) for v in row]
                non_empty = [t for t in row_texts if t]
                if len(non_empty) >= 2:
                    matrix_header_row = ri
                    headers = row_texts
                    break

            if matrix_header_row is None:
                continue

            check_col = None
            check_col_options = ""
            for ci in range(len(headers)):
                if ci in dropdown_cols:
                    check_col = ci
                    check_col_options = dropdown_cols[ci]
                    break
            if check_col is None:
                for ci, h in enumerate(headers):
                    if any(sh in _norm(h) for sh in _STATUS_HINTS):
                        check_col = ci
                        break
            if check_col is None and len(headers) > 1:
                for ci in range(len(headers) - 1, -1, -1):
                    if headers[ci]:
                        check_col = ci
                        break

            col_0 = 0
            while col_0 < len(headers) and not headers[col_0]:
                col_0 += 1

            col_0_header = headers[col_0] if col_0 < len(headers) else ""
            if col_0_header and _norm(col_0_header) not in ["no", "sl", "s.no", "item", "#"]:
                section_name = col_0_header
            else:
                section_name = sheet_name

            check_name = headers[check_col] if check_col is not None and check_col < len(headers) and headers[check_col] else "Operational"
            has_dropdown = (check_col in dropdown_cols) or any(sh in _norm(check_name) for sh in _STATUS_HINTS)

            if check_col_options:
                col_kind, col_choices = _classify_options(check_col_options)
            elif has_dropdown:
                col_kind, col_choices = ChecklistItem.AnswerType.BOOLEAN, ""
            else:
                col_kind, col_choices = None, ""

            for row in sheet.iter_rows(min_row=matrix_header_row + 1, values_only=True):
                row = list(row)
                if not any(v is not None and _cell_text(v) for v in row):
                    continue

                item_val = _cell_text(row[col_0]) if col_0 < len(row) else ""
                if not item_val or _norm(item_val) in [_norm(h) for h in headers if h]:
                    continue
                if _norm(item_val) in ["total", "grand total", "summary", "subtotal", "average", "avg"]:
                    continue

                checkpoint_counter += 1
                seen_checkpoints.add(checkpoint_counter)

                specs = []
                for ci in range(col_0 + 1, len(row)):
                    if ci == check_col:
                        continue
                    if ci < len(headers) and headers[ci] and _cell_text(row[ci]):
                        specs.append(f"{headers[ci]}: {_cell_text(row[ci])}")
                standard_str = " | ".join(specs)

                if check_name.lower() in item_val.lower():
                    desc = item_val
                else:
                    desc = f"{item_val} - {check_name} Check"

                if col_kind is not None:
                    kind = col_kind
                    choices_str = col_choices
                else:
                    kind = _answer_type(desc, standard_str)
                    choices_str = ""

                parsed_items.append({
                    "checkpoint_no": checkpoint_counter,
                    "section": section_name,
                    "description": desc,
                    "standard": standard_str,
                    "method": "Visual / SCADA",
                    "frequency": "Daily",
                    "answer_type": kind,
                    "choices": choices_str,
                    "unit": "",
                    "position": len(parsed_items) + 1,
                })

    with transaction.atomic():
        existing = ChecklistTemplate.objects.filter(name=template_name).order_by("-version").first()
        version = (existing.version + 1) if existing else 1
        template = ChecklistTemplate.objects.create(
            name=template_name,
            version=version,
            document_number=doc_number[:120],
            company=company[:200],
            equipment_name=equipment_name[:200],
        )

        items_to_create = [
            ChecklistItem(
                template=template,
                section=it["section"],
                checkpoint_no=it["checkpoint_no"],
                description=it["description"],
                standard=it["standard"],
                method=it["method"],
                frequency=it["frequency"],
                answer_type=it["answer_type"],
                choices=it.get("choices", ""),
                unit=it["unit"],
                position=it["position"],
            )
            for it in parsed_items
        ]
        ChecklistItem.objects.bulk_create(items_to_create)

    return template, len(items_to_create)


_OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP_MAGICS = (b"PK\x03\x04", b"PK\x05\x06")


def _convert_via_libreoffice(file_bytes, src_ext):
    """Convert .xls / XML / HTML spreadsheet bytes to .xlsx using LibreOffice."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        in_file = tmp_path / f"upload{src_ext}"
        out_dir = tmp_path / "out"
        profile_dir = tmp_path / "lo_profile"
        out_dir.mkdir()
        profile_dir.mkdir()
        in_file.write_bytes(file_bytes)
        cmd = [
            soffice,
            f"-env:UserInstallation=file://{profile_dir}",
            "--headless",
            "--norestore",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(out_dir),
            str(in_file),
        ]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=45)
        except Exception:
            return None
        out_file = out_dir / "upload.xlsx"
        if res.returncode == 0 and out_file.exists():
            return out_file.read_bytes()
    return None


def _looks_like_text_csv(file_bytes):
    """Return True if bytes look like plain-text CSV/TSV rather than binary or XML/HTML."""
    head = file_bytes[:4096]
    if b"\x00" in head:
        # Could be UTF-16 text; avoid misclassifying binary files
        if not (head.startswith(b"\xff\xfe") or head.startswith(b"\xfe\xff")):
            return False
    for enc in ("utf-8-sig", "utf-8", "utf-16", "latin-1"):
        try:
            text = head.decode(enc)
            stripped = text.lstrip()
            if stripped.startswith("<"):
                return False  # XML / HTML table
            if "\n" in text and ("," in text or "\t" in text):
                return True
        except UnicodeDecodeError:
            continue
    return False


def import_uploaded_excel(file_obj, template_name):
    """
    Parse an uploaded Excel file into a new ChecklistTemplate.
    Supports .xlsx/.xlsm, legacy .xls (converted via LibreOffice to preserve
    dropdown validations), XML/HTML spreadsheets, and CSV files saved with an
    Excel extension. Returns (template, items_created).
    """
    if hasattr(file_obj, "read"):
        file_bytes = file_obj.read()
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
    else:
        with open(file_obj, "rb") as fp:
            file_bytes = fp.read()

    if not file_bytes:
        raise ValueError("The uploaded file is empty.")

    if not file_bytes.startswith(_ZIP_MAGICS):
        # 1. Check for encrypted OLE2 Excel workbook
        if file_bytes.startswith(_OLE2_MAGIC) and b"E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00P\x00a\x00c\x00k\x00a\x00g\x00e" in file_bytes:
            raise ValueError("This Excel file is password-protected or encrypted. Please remove the password and re-upload.")

        # 2. Check if it's actually a CSV/TSV saved with .xls/.xlsx extension
        if not file_bytes.startswith(_OLE2_MAGIC) and _looks_like_text_csv(file_bytes):
            return import_uploaded_csv(io.BytesIO(file_bytes), template_name)

        # 3. Convert legacy .xls / XML / HTML / ODS to .xlsx via LibreOffice
        head_low = file_bytes[:512].lstrip().lower()
        if file_bytes.startswith(_OLE2_MAGIC):
            src_ext = ".xls"
        elif head_low.startswith(b"<?xml") or b"<workbook" in head_low:
            src_ext = ".xml"
        elif head_low.startswith((b"<html", b"<!doctype", b"<table")):
            src_ext = ".html"
        else:
            src_ext = ".xls"

        converted = _convert_via_libreoffice(file_bytes, src_ext)
        if converted:
            file_bytes = converted
        else:
            raise ValueError(
                "Could not open this Excel file (it is not a standard .xlsx ZIP workbook). "
                "Please re-save it as 'Excel Workbook (.xlsx)' or '.csv' and upload again."
            )

    try:
        wb_data = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except zipfile.BadZipFile as exc:
        raise ValueError(
            "Invalid or corrupted Excel file (not a valid .xlsx archive). "
            "Please open and re-save it as '.xlsx' or '.csv' and upload again."
        ) from exc

    try:
        wb_val = load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception:
        wb_val = None

    try:
        return _import_workbook_pair(wb_data, wb_val, template_name)
    finally:
        wb_data.close()
        if wb_val:
            wb_val.close()


def import_uploaded_csv(file_obj, template_name):
    """
    Parse an uploaded CSV file (such as an exported Excel sheet) into a new ChecklistTemplate
    by converting it to an in-memory Workbook and delegating to _import_workbook_pair.
    """
    if hasattr(file_obj, "read"):
        raw = file_obj.read()
        if hasattr(file_obj, "seek"):
            file_obj.seek(0)
    else:
        with open(file_obj, "rb") as fp:
            raw = fp.read()

    if isinstance(raw, bytes):
        if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
            text = raw.decode("utf-16")
        else:
            try:
                text = raw.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = raw.decode("latin-1")
    else:
        text = raw

    delimiter = "\t" if ("\t" in text and text.count("\t") > text.count(",")) else ","

    wb = Workbook()
    ws = wb.active
    ws.title = template_name[:31] or "Checklist"
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    for row in reader:
        ws.append(row)

    try:
        return _import_workbook_pair(wb, None, template_name)
    finally:
        wb.close()


def import_uploaded_pdf(file_obj, template_name):
    """
    Parse an uploaded PDF into a ChecklistTemplate using pdfplumber if available,
    falling back to a single placeholder item with instructions.
    Returns (template, items_created).
    """
    existing = ChecklistTemplate.objects.filter(name=template_name).order_by("-version").first()
    version = (existing.version + 1) if existing else 1
    template = ChecklistTemplate.objects.create(name=template_name, version=version)

    try:
        import pdfplumber  # optional dependency
    except ImportError:
        ChecklistItem.objects.create(
            template=template,
            checkpoint_no=1,
            description="[PDF import not available — install pdfplumber or re-upload as Excel]",
            answer_type=ChecklistItem.AnswerType.TEXT,
            position=1,
        )
        return template, 1

    created = 0
    checkpoint_counter = 0
    current_section = ""

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    row = [_cell_text(c) for c in row]
                    desc = row[1] if len(row) > 1 else row[0]
                    if not desc or desc.lower() in ("description", "observation", "job description"):
                        continue
                    checkpoint_counter += 1
                    standard = row[2] if len(row) > 2 else ""
                    kind = _answer_type(desc, standard)
                    ChecklistItem.objects.create(
                        template=template,
                        section=current_section,
                        checkpoint_no=checkpoint_counter,
                        description=desc,
                        standard=standard,
                        answer_type=kind,
                        position=checkpoint_counter,
                    )
                    created += 1

    if created == 0:
        ChecklistItem.objects.create(
            template=template,
            checkpoint_no=1,
            description="[No table rows found in PDF — please re-upload as Excel for best results]",
            answer_type=ChecklistItem.AnswerType.TEXT,
            position=1,
        )
        created = 1

    return template, created
