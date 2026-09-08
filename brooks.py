from __future__ import annotations

import csv
import io
import re
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
MASTER_TEMPLATE = BASE_DIR / "master_template.xlsm"

PRISMA_SHEET = "Prisma Export - Paste as values"
TRAFFIC_SHEET = "Traffic_Doc"
MULTI_SHEET = "Multi-Ad or Creative Rotation"

SEE_MULTI = "See Multi-Ad tab"
TRACKING_1X1 = "Tracking_1x1"


def _clean(v) -> str:
    return "" if v is None else str(v).strip()


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


def _read_uploaded_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


def _read_csv_rows(uploaded_file) -> list[list[str]]:
    text = _decode(_read_uploaded_bytes(uploaded_file))
    try:
        dialect = csv.Sniffer().sniff(text[:10000], delimiters=",;\t|")
        reader = csv.reader(io.StringIO(text), dialect)
    except csv.Error:
        reader = csv.reader(io.StringIO(text))
    return [list(r) for r in reader]


def _campaign_name(rows: list[list[str]]) -> str:
    for row in rows[:15]:
        if row and _clean(row[0]).lower().startswith("campaign name"):
            return _clean(row[1] if len(row) > 1 else "")
    return ""


def _find_header_row(rows: list[list[str]]) -> int:
    for i, row in enumerate(rows):
        vals = {_clean(x).lower() for x in row}
        if "ad server id" in vals and "placement name" in vals and "tag size" in vals:
            return i
    raise ValueError("Could not find the Prisma placement header row.")


def _colmap(header: list[str]) -> dict[str, int]:
    return {_clean(v).lower(): i for i, v in enumerate(header)}


def _parse_date(v):
    if isinstance(v, datetime):
        return v
    s = _clean(v)
    if not s:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y",
        "%d/%m/%Y", "%d-%m-%Y", "%m-%d-%Y",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return s


def _normalize_size(v: str) -> str:
    s = _clean(v).lower().replace(" ", "")
    s = s.replace("×", "x")
    if s in {"nanxnan", "0x0", "nan", ""}:
        return "1x1"
    m = re.search(r"(\d+)x(\d+)", s)
    return f"{m.group(1)}x{m.group(2)}" if m else s


def _extract_concept(campaign: str) -> str:
    # Confirmed Brooks sample: BRSP_DPS_F26 Cascadia H2T_Perf -> H2T
    m = re.search(r"\b([A-Za-z0-9]+)_Perf\b", campaign, flags=re.I)
    if m:
        return m.group(1)
    # Safe fallback: token immediately before a suffix beginning with Perf.
    parts = [p for p in re.split(r"[_\s]+", campaign) if p]
    for i, p in enumerate(parts):
        if p.lower().startswith("perf") and i:
            return parts[i - 1]
    return "BROOKS"


def _publisher_label(placement_name: str, supplier_name: str) -> str:
    parts = [p.strip() for p in placement_name.split("_")]
    if len(parts) >= 4 and parts[3]:
        return parts[3].title()
    supplier = _clean(supplier_name)
    if "outside" in supplier.lower():
        return "Outside"
    words = supplier.split()
    return (words[0].title() if words else "Publisher")


def _creative_names(creative_files) -> list[str]:
    names: list[str] = []
    for f in creative_files or []:
        name = getattr(f, "name", "") or ""
        data = _read_uploaded_bytes(f)
        if name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as z:
                    for n in z.namelist():
                        if not n.endswith("/"):
                            names.append(Path(n).name)
            except zipfile.BadZipFile:
                names.append(name)
        else:
            names.append(name)
    # Workbook examples omit extensions, so standardize to stem.
    result = []
    seen = set()
    for n in names:
        stem = Path(n).stem.strip()
        if stem and stem.lower() not in seen:
            seen.add(stem.lower())
            result.append(stem)
    return result


def _matches_dimension(name: str, size: str) -> bool:
    return re.search(rf"(?<!\d){re.escape(size)}(?!\d)", name, flags=re.I) is not None


def _copy_row_style(ws, source_row: int, target_row: int, max_col: int = 20):
    if target_row == source_row:
        return
    for c in range(1, max_col + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def _write_prisma(ws, rows: list[list[str]]):
    # Clear existing used values only; keep template formatting.
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(r_idx, c_idx).value = value


def _clear_output_rows(ws, start_row: int, end_row: int, max_col: int):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None


def preview_brooks_setup(prisma_file, creative_files):
    rows = _read_csv_rows(prisma_file)
    campaign = _campaign_name(rows)
    concept = _extract_concept(campaign)
    h = _find_header_row(rows)
    cols = _colmap(rows[h])
    creatives = _creative_names(creative_files)

    required = [
        "ad server id", "placement name", "tag size",
        "media outlet / supplier name (prisma)",
        "flight start date", "flight end date",
    ]
    missing = [x for x in required if x not in cols]
    if missing:
        raise ValueError("Missing Prisma columns: " + ", ".join(missing))

    placements = []
    warnings = []
    direct = multi = one_by_one = unmatched = 0

    for row in rows[h + 1:]:
        if not any(_clean(x) for x in row):
            continue
        def get(name):
            idx = cols.get(name)
            return row[idx] if idx is not None and idx < len(row) else ""

        pid = _clean(get("ad server id"))
        pname = _clean(get("placement name"))
        if not pid and not pname:
            continue

        size = _normalize_size(get("tag size") or get("ad size"))
        publisher = _publisher_label(pname, get("media outlet / supplier name (prisma)"))
        ad_name = f"{concept} - {publisher} - {size}"
        matches = [] if size == "1x1" else [n for n in creatives if _matches_dimension(n, size)]

        if size == "1x1":
            status = "Tracking 1x1"
            one_by_one += 1
        elif len(matches) == 0:
            status = "No creative match"
            unmatched += 1
            warnings.append(f"{pid or pname}: no creative matched {size}.")
        elif len(matches) == 1:
            status = "Direct"
            direct += 1
        else:
            status = f"Multi ({len(matches)})"
            multi += 1

        placements.append({
            "placement_id": pid,
            "placement_name": pname,
            "size": size,
            "publisher": publisher,
            "ad_name": ad_name,
            "start_date": _parse_date(get("flight start date")),
            "end_date": _parse_date(get("flight end date")),
            "matches": matches,
            "status": status,
        })

    return {
        "campaign": campaign,
        "concept": concept,
        "placements": placements,
        "creative_count": len(creatives),
        "direct_count": direct,
        "multi_count": multi,
        "tracking_1x1_count": one_by_one,
        "unmatched_count": unmatched,
        "warnings": warnings,
    }


def generate_brooks_tsheet(
    prisma_file,
    creative_files,
    apply_dynata_display: bool = False,
    template_path: str | Path | None = None,
):
    """
    Brooks automation confirmed from the two supplied Outside T-sheets.

    Automated:
      - Prisma import
      - concept extraction from Campaign Name (e.g. H2T from ...H2T_Perf)
      - publisher extraction (e.g. Outside)
      - Ad Name = <Concept> - <Publisher> - <Dimension>
      - 1x1 Creative File Name = Tracking_1x1
      - display creative matching by dimension
      - direct vs Multi-Ad routing
      - placement start/end dates
      - optional Dynata note/flag for Display

    Intentionally NOT automated:
      - click-through URL / UTM / tid construction
    """
    preview = preview_brooks_setup(prisma_file, creative_files)
    rows = _read_csv_rows(prisma_file)

    template = Path(template_path) if template_path else MASTER_TEMPLATE
    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")

    wb = load_workbook(template, keep_vba=True)
    if PRISMA_SHEET not in wb.sheetnames or TRAFFIC_SHEET not in wb.sheetnames or MULTI_SHEET not in wb.sheetnames:
        raise ValueError("Template is missing required Traffic Doc sheets.")

    prisma_ws = wb[PRISMA_SHEET]
    traffic_ws = wb[TRAFFIC_SHEET]
    multi_ws = wb[MULTI_SHEET]

    _write_prisma(prisma_ws, rows)

    # Header/default values.
    traffic_ws["B1"] = preview["campaign"]
    # Keep URL intentionally blank / manual. Do not create a UTM.
    traffic_ws["B5"] = "https://www.brooksrunning.com"

    # Clear prior output content while preserving template styles.
    _clear_output_rows(traffic_ws, 9, max(traffic_ws.max_row, 650), 20)
    _clear_output_rows(multi_ws, 2, max(multi_ws.max_row, 650), 16)

    warnings = list(preview["warnings"])
    multi_row = 2

    for i, p in enumerate(preview["placements"], start=9):
        _copy_row_style(traffic_ws, 9, i, 20)

        is_1x1 = p["size"] == "1x1"
        matches = p["matches"]

        # Traffic_Doc columns from supplied Brooks workbook.
        traffic_ws.cell(i, 1).value = "N" if is_1x1 else ("Y" if apply_dynata_display else "N")
        traffic_ws.cell(i, 2).value = p["publisher"].upper() if p["publisher"] == "Outside" else p["publisher"]
        # Preserve supplier display exactly where known.
        if p["publisher"] == "Outside":
            traffic_ws.cell(i, 2).value = "OUTSIDE INTEGRATED MEDIA"
        traffic_ws.cell(i, 3).value = int(p["placement_id"]) if p["placement_id"].isdigit() else p["placement_id"]
        traffic_ws.cell(i, 4).value = p["placement_name"]
        traffic_ws.cell(i, 5).value = p["size"]
        traffic_ws.cell(i, 8).value = p["ad_name"]
        traffic_ws.cell(i, 9).value = "Please add 2026 Dynata Pixel" if (apply_dynata_display and not is_1x1) else None
        traffic_ws.cell(i, 10).value = "New"
        traffic_ws.cell(i, 12).value = "N"
        traffic_ws.cell(i, 14).value = p["start_date"]
        traffic_ws.cell(i, 15).value = p["end_date"]

        if is_1x1:
            traffic_ws.cell(i, 11).value = TRACKING_1X1
            traffic_ws.cell(i, 13).value = "100%"
            traffic_ws.cell(i, 16).value = None
        elif len(matches) == 1:
            traffic_ws.cell(i, 11).value = matches[0]
            traffic_ws.cell(i, 13).value = "100%"
            traffic_ws.cell(i, 16).value = None
        elif len(matches) >= 2:
            traffic_ws.cell(i, 11).value = None
            traffic_ws.cell(i, 13).value = SEE_MULTI
            traffic_ws.cell(i, 16).value = SEE_MULTI
            for creative in matches:
                _copy_row_style(multi_ws, 2, multi_row, 16)
                multi_ws.cell(multi_row, 1).value = p["ad_name"]
                multi_ws.cell(multi_row, 2).value = "New"
                multi_ws.cell(multi_row, 4).value = creative
                multi_ws.cell(multi_row, 5).value = "N"
                multi_ws.cell(multi_row, 6).value = "Even"
                multi_ws.cell(multi_row, 7).value = p["start_date"]
                multi_ws.cell(multi_row, 8).value = p["end_date"]
                # URL/UTM intentionally blank for manual population.
                multi_ws.cell(multi_row, 9).value = None
                multi_row += 1
        else:
            traffic_ws.cell(i, 11).value = None
            traffic_ws.cell(i, 13).value = None
            traffic_ws.cell(i, 16).value = None

        for c in (14, 15):
            traffic_ws.cell(i, c).number_format = "mm/dd/yyyy"

    for r in range(2, multi_row):
        multi_ws.cell(r, 7).number_format = "mm/dd/yyyy"
        multi_ws.cell(r, 8).number_format = "mm/dd/yyyy"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), warnings, preview
