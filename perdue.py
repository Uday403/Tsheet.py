from __future__ import annotations

import csv
import io
import re
import zipfile
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Iterable
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
MASTER_TEMPLATE = BASE_DIR / "master_template.xlsm"
PRISMA_SHEET = "Prisma Export - Paste as values"
TRAFFIC_SHEET = "Traffic_Doc"
MULTI_SHEET = "Multi-Ad or Creative Rotation"

SUPPORTED_CREATIVE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".webp", ".html", ".htm",
    ".mp4", ".mov", ".m4v", ".zip",
}

# Confirmed from all 8 sheets of Perdue_Taxonomy_2027.
# Product alone is not always enough, so the most specific CreativeName/CTA
# combinations are checked first.
LANDING_PAGE_RULES = {
    ("CrispyStrips-Continuity", "Soggy", "LearnMore"): "https://www.perdue.com/perdue-way",
    ("CrispyStrips-Continuity", "CSCrave", "BuyNow"): "https://www.perdue.com/products/perdue-crispy-chicken-strips",
    ("CrispyStrips-Continuity", "CrispyVsSoggy", "BuyNow"): "https://www.perdue.com/products/perdue-crispy-chicken-strips",
    ("CrispyStrips-Continuity", "CrispyVsSoggy", "NA"): "https://www.perdue.com/products/perdue-crispy-chicken-strips",
    ("CrispyStrips-Continuity", "Fairy", "NA"): "https://www.perdue.com/products/perdue-crispy-chicken-strips",
    ("CrispyStrips-Continuity", "Fairy", "SaveNow"): "https://www.perdue.com/products/perdue-crispy-chicken-strips",
    ("PankoNuggets-Continuity", "PankoNuggets", "LearnMore"): "https://www.perdue.com/perdue-way",
    ("PankoNuggets-Continuity", "Taste", "BuyNow"): "https://www.perdue.com/products/perdue-panko-chicken-nuggets",
    ("PankoNuggets-Continuity", "Taste", "NA"): "https://www.perdue.com/perdue-way",
}

PRODUCT_LANDING_PAGES = {
    "GroundChicken-Continuity": "https://www.perdue.com/products/perdue-fresh-ground-chicken",
    "MasterBrand-Continuity": "https://www.perdue.com/perdue-way",
    "MasterBrand-Kelly&Mark": "https://www.perdue.com/perdue-way",
    "Powered-Continuity": "https://www.perdue.com/products/perdue-powered",
    "Yummy-PAWPatrol": "https://yummydinobuddies.com/pawpatrol",
}

TAXONOMY_FIELDS = [
    "Agency", "Channel", "SubChannel", "Partner/Vendor", "Tactic/Demo",
    "FreeHand", "Audience", "Data Source", "Bid", "Device", "Size",
    "Ad Server", "Region/Location", "Goal", "Creative Type", "Brand Type",
    "Family", "Product-Effort", "Creative-CTA",
]


def _clean(v) -> str:
    return "" if v is None else str(v).strip()


def _norm(v) -> str:
    return re.sub(r"[^a-z0-9]+", "", _clean(v).lower())


def _safe_name(v: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "-", _clean(v)).strip(" .")


def _decode_csv(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


def _read_upload_bytes(uploaded) -> bytes:
    uploaded.seek(0)
    return uploaded.read()


def read_prisma_export(prisma_file) -> tuple[list[list[str]], list[dict]]:
    text = _decode_csv(_read_upload_bytes(prisma_file))
    try:
        dialect = csv.Sniffer().sniff(text[:12000], delimiters=",;\t|")
        rows = list(csv.reader(io.StringIO(text), dialect))
    except csv.Error:
        rows = list(csv.reader(io.StringIO(text)))

    header_idx = None
    headers = None
    for i, row in enumerate(rows[:80]):
        normalized = [_norm(x) for x in row]
        if "placementname" in normalized:
            header_idx = i
            headers = [_clean(x) for x in row]
            break
    if header_idx is None:
        raise ValueError("Could not find Placement Name in the Prisma CSV.")

    records = []
    for source_row, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        padded = row + [""] * max(0, len(headers) - len(row))
        rec = dict(zip(headers, padded[:len(headers)]))
        rec["_source_excel_row"] = source_row
        placement = _get(rec, "Placement Name")
        if not placement:
            continue
        records.append(rec)
    return rows, records


def _get(rec: dict, *aliases: str) -> str:
    wanted = {_norm(x) for x in aliases}
    for k, v in rec.items():
        if _norm(k) in wanted and _clean(v):
            return _clean(v)
    return ""


def parse_perdue_placement(placement_name: str) -> dict:
    """Parse the 19-part Perdue Final Placement Name taxonomy."""
    parts = _clean(placement_name).split("_", 18)
    if len(parts) != 19:
        return {"valid": False, "placement": placement_name,
                "error": f"Expected 19 taxonomy sections, found {len(parts)}."}
    data = dict(zip(TAXONOMY_FIELDS, parts))
    creative_cta = data.pop("Creative-CTA")
    if "-" in creative_cta:
        creative, cta = creative_cta.rsplit("-", 1)
    else:
        creative, cta = creative_cta, ""
    data["CreativeName"] = creative
    data["CTA"] = cta
    data["placement"] = placement_name
    data["valid"] = True
    return data


def parse_landing_urls(text: str) -> dict[str, str]:
    """Parse user-provided BASE landing URLs.

    Supported formats:
      1) One URL only -> applies to every Perdue placement.
      2) Product-Effort<TAB>URL
      3) CreativeName-CTA<TAB>URL
      4) Product-Effort|CreativeName|CTA<TAB>URL

    The user supplies the landing URL; Perdue UTM parameters are appended by code.
    """
    lines = [line.strip() for line in _clean(text).splitlines() if line.strip()]
    result: dict[str, str] = {}
    bare_urls: list[str] = []

    for line in lines:
        parts = re.split(r"\t+", line, maxsplit=1)
        if len(parts) == 2 and parts[1].strip().startswith(("http://", "https://")):
            result[parts[0].strip()] = parts[1].strip()
        elif line.startswith(("http://", "https://")):
            bare_urls.append(line)

    if len(bare_urls) == 1:
        result["__default__"] = bare_urls[0]
    elif len(bare_urls) > 1:
        result["__multiple_unmapped__"] = "1"

    return result


def landing_page_for(tax: dict, user_urls: dict[str, str] | None = None) -> str:
    """Resolve ONLY from URLs supplied by the user; do not invent a landing page."""
    product = tax.get("Product-Effort", "")
    creative = tax.get("CreativeName", "")
    cta = tax.get("CTA", "")
    user_urls = user_urls or {}
    for key in (
        f"{product}|{creative}|{cta}",
        f"{creative}-{cta}",
        product,
        "__default__",
    ):
        if user_urls.get(key):
            return user_urls[key]
    return ""


def build_final_url(tax: dict, landing_page: str) -> str:
    if not landing_page:
        return ""
    params = {
        "utm_source": tax.get("Partner/Vendor", ""),
        "utm_medium": f'{tax.get("Channel", "")}_{tax.get("Size", "")}_{tax.get("SubChannel", "")}',
        "utm_campaign": f'{tax.get("Product-Effort", "")}_{tax.get("Family", "")}',
        "utm_term": f'{tax.get("Tactic/Demo", "")}_{tax.get("CreativeName", "")}-{tax.get("CTA", "")}',
    }
    # Keep any non-UTM query parameters already present in the supplied landing URL,
    # but replace/add the four Perdue UTM parameters deterministically.
    parts = urlsplit(landing_page.strip())
    existing = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True)
                if k.lower() not in {"utm_source", "utm_medium", "utm_campaign", "utm_term"}]
    query = urlencode(existing + list(params.items()), doseq=True, safe="-_|:")
    return urlunsplit((parts.scheme, parts.netloc, parts.path, query, parts.fragment))


def _platform_for_creative(tax: dict) -> str:
    partner = tax.get("Partner/Vendor", "")
    freehand = tax.get("FreeHand", "")
    # Historical Yahoo PMP placements have publisher in Partner/Vendor and Yahoo
    # in FreeHand. The creative itself is named as Yahoo.
    if _norm(freehand) == "yahoo":
        return "Yahoo"
    if _norm(freehand) == "nova" and _norm(partner) == "yahoo":
        return "Yahoo"
    return partner or freehand or "Perdue"


def build_creative_filename(tax: dict, extension: str = "") -> str:
    concept = f'{tax.get("CreativeName", "")}-{tax.get("CTA", "")}'.strip("-")
    platform = _platform_for_creative(tax)
    size = tax.get("Size", "")
    channel = tax.get("Channel", "")
    base = f"{concept}_PERD_{platform}_{size}_{channel}"
    ext = extension if extension.startswith(".") or not extension else "." + extension
    return _safe_name(base) + ext.lower()


def build_ad_name(tax: dict) -> str:
    """OLV/CTV are always the FIRST component, per confirmed Perdue rule."""
    channel = tax.get("Channel", "")
    concept = f'{tax.get("CreativeName", "")}-{tax.get("CTA", "")}'.strip("-")
    size = tax.get("Size", "")
    partner = tax.get("Partner/Vendor", "")
    tactic = tax.get("Tactic/Demo", "")
    freehand = tax.get("FreeHand", "")

    if _norm(partner) == "adtheorent":
        return _safe_name(f"{channel}_{concept}_PERD_AdTheorent_{size}")

    # Yahoo PMP historical structure: publisher is Partner/Vendor, Yahoo is FreeHand.
    if _norm(freehand) == "yahoo" and _norm(partner) != "nova":
        return _safe_name(f"{channel}_{tactic}_{concept}_{size}_{partner}")

    if _norm(partner) == "nova":
        return _safe_name(f"{channel}_{tactic}_{concept}_{size}_Nova")

    return _safe_name(f"{channel}_{tactic}_{concept}_{size}_{partner}")


def _extract_dimension(name: str) -> str:
    m = re.search(r"(?<!\d)(\d{2,4})\s*[xX×*]\s*(\d{2,4})(?!\d)", name)
    return f"{m.group(1)}x{m.group(2)}" if m else ""


def _extract_duration(name: str) -> str:
    m = re.search(r"(?<!\d)(6|10|15|20|30|60|90)\s*(?:s|sec|secs|second|seconds)(?![a-z0-9])", name, re.I)
    return f"{m.group(1)}s" if m else ""


def _target_size(tax: dict) -> str:
    return tax.get("Size", "")


def _creative_tokens(text: str) -> set[str]:
    stem = Path(text).stem
    stem = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", stem)
    stop = {"perd", "perdue", "final", "creative", "video", "display", "olv", "ctv", "jpg", "jpeg", "png", "gif", "mp4", "mov", "m4v"}
    return {x.lower() for x in re.findall(r"[A-Za-z0-9]+", stem) if len(x) >= 2 and x.lower() not in stop}


def _read_creatives(creative_files: Iterable) -> list[dict]:
    items = []
    for uploaded in creative_files or []:
        name = Path(uploaded.name).name
        data = _read_upload_bytes(uploaded)
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                for member in zf.namelist():
                    if member.endswith("/"):
                        continue
                    base = Path(member).name
                    if not base or base.startswith(".") or "__MACOSX" in member:
                        continue
                    ext = Path(base).suffix.lower()
                    if ext in SUPPORTED_CREATIVE_EXTENSIONS - {".zip"}:
                        items.append({"name": base, "bytes": zf.read(member), "extension": ext})
        else:
            ext = Path(name).suffix.lower()
            if ext in SUPPORTED_CREATIVE_EXTENSIONS - {".zip"}:
                items.append({"name": name, "bytes": data, "extension": ext})
    # duplicate-safe by original name
    out, seen = [], set()
    for item in items:
        key = item["name"].lower()
        if key not in seen:
            seen.add(key); out.append(item)
    return out


def parse_creative_mapping(text: str) -> dict[str, str]:
    """raw_filename<TAB>CreativeName-CTA (pipe is also accepted)."""
    result = {}
    for line in _clean(text).splitlines():
        if not line.strip():
            continue
        parts = re.split(r"\t+|\s*\|\s*", line.strip(), maxsplit=1)
        if len(parts) == 2:
            result[Path(parts[0].strip()).name.lower()] = parts[1].strip()
    return result


def _match_creative(tax: dict, creatives: list[dict], mapping: dict[str, str]) -> tuple[dict | None, str]:
    target = f'{tax.get("CreativeName", "")}-{tax.get("CTA", "")}'
    target_norm = _norm(target)
    required = _target_size(tax).lower()
    candidates = []

    for item in creatives:
        mapped = mapping.get(item["name"].lower(), "")
        if mapped:
            if _norm(mapped) == target_norm:
                candidates.append((10000, item))
            continue

        detected = _extract_dimension(item["name"]) or _extract_duration(item["name"])
        if required and detected and detected.lower() != required:
            continue

        tokens = _creative_tokens(item["name"])
        target_tokens = _creative_tokens(target)
        score = len(tokens & target_tokens) * 25
        if target_norm and target_norm in _norm(item["name"]):
            score += 100
        if detected and detected.lower() == required:
            score += 10
        if score > 0:
            candidates.append((score, item))

    if not candidates:
        return None, "Unmatched"
    best_score = max(x[0] for x in candidates)
    best = [x[1] for x in candidates if x[0] == best_score]
    if len(best) != 1:
        return None, "Ambiguous"
    return best[0], "Matched"


def preview_perdue_setup(prisma_file, creative_files, creative_mapping_text: str = "", landing_urls_text: str = "") -> dict:
    _, records = read_prisma_export(prisma_file)
    creatives = _read_creatives(creative_files)
    mapping = parse_creative_mapping(creative_mapping_text)
    user_urls = parse_landing_urls(landing_urls_text)
    rows, warnings = [], []
    if user_urls.get("__multiple_unmapped__"):
        warnings.append("Multiple bare landing URLs were pasted. Use Product-Effort<TAB>URL or CreativeName-CTA<TAB>URL so each URL can be mapped safely.")

    for rec in records:
        placement = _get(rec, "Placement Name")
        tax = parse_perdue_placement(placement)
        if not tax.get("valid"):
            warnings.append(f"Taxonomy parse failed: {placement} — {tax.get('error')}")
            rows.append({"placement_name": placement, "status": "Taxonomy error"})
            continue

        landing = landing_page_for(tax, user_urls)
        url = build_final_url(tax, landing) if landing else ""
        creative, status = _match_creative(tax, creatives, mapping)
        renamed = build_creative_filename(tax, creative["extension"]) if creative else ""
        ad_name = build_ad_name(tax)

        if not landing:
            warnings.append(f"User landing URL required for {tax['Product-Effort']} / {tax['CreativeName']}-{tax['CTA']}.")
        if status != "Matched":
            warnings.append(f"{status} creative for {placement}. Add Creative Mapping if the raw filename is generic.")

        rows.append({
            "placement_name": placement,
            "channel": tax.get("Channel", ""),
            "partner": tax.get("Partner/Vendor", ""),
            "size": tax.get("Size", ""),
            "concept": f'{tax.get("CreativeName", "")}-{tax.get("CTA", "")}',
            "ad_name": ad_name,
            "original_creative": creative["name"] if creative else "",
            "renamed_creative": renamed,
            "landing_page": landing,
            "final_url": url,
            "status": status if landing else f"{status}; Landing page missing",
            "_tax": tax,
            "_record": rec,
            "_creative": creative,
        })

    return {
        "rows": rows,
        "placements": rows,
        "creative_count": len(creatives),
        "matched_count": sum(1 for x in rows if x.get("original_creative")),
        "unmatched_count": sum(1 for x in rows if not x.get("original_creative")),
        "url_matched_count": sum(1 for x in rows if x.get("final_url")),
        "url_unmatched_count": sum(1 for x in rows if not x.get("final_url")),
        "warnings": warnings,
    }


def _find_header_row(ws) -> int:
    for r in range(1, min(ws.max_row, 40) + 1):
        vals = {_norm(ws.cell(r, c).value) for c in range(1, min(ws.max_column, 40) + 1)}
        if "placementname" in vals and "adname" in vals:
            return r
    raise ValueError("Traffic_Doc header row not found.")


def _header_map(ws, row: int) -> dict[str, int]:
    out = {}
    for c in range(1, ws.max_column + 1):
        v = _norm(ws.cell(row, c).value)
        if v == "additionalpixels": out["pixels"] = c
        elif v == "sitename": out["site"] = c
        elif v in {"dcmplacementid", "placementid"}: out["placement_id"] = c
        elif v == "placementname": out["placement"] = c
        elif v in {"dimensions", "dimension"}: out["dimension"] = c
        elif v == "videolength": out["video_length"] = c
        elif v in {"vastorvpaid", "vastvpaid"}: out["vast"] = c
        elif v == "adname": out["ad"] = c
        elif v == "action": out["action"] = c
        elif v == "creativefilename": out["creative"] = c
        elif v.startswith("studiocreative"): out["studio"] = c
        elif v.startswith("rotation"): out["rotation"] = c
        elif v == "startdate": out["start"] = c
        elif v == "enddate": out["end"] = c
        elif v.startswith("clickthroughurl") or v.startswith("clicktag1"): out["url"] = c
    return out


def _copy_style(src, dst):
    dst._style = copy(src._style)
    dst.number_format = src.number_format


def _parse_date(v):
    if isinstance(v, (datetime, date)): return v
    s = _clean(v)
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y"):
        try: return datetime.strptime(s, fmt)
        except ValueError: pass
    return s


def _write_prisma_sheet(wb, raw_rows: list[list[str]]) -> None:
    if PRISMA_SHEET not in wb.sheetnames:
        return
    ws = wb[PRISMA_SHEET]
    for row in ws.iter_rows():
        for cell in row: cell.value = None
    for r, values in enumerate(raw_rows, 1):
        for c, value in enumerate(values, 1):
            ws.cell(r, c).value = value


def build_renamed_creatives_zip(preview: dict) -> bytes:
    output = io.BytesIO()
    written = set()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for row in preview["rows"]:
            item = row.get("_creative")
            renamed = row.get("renamed_creative")
            if not item or not renamed:
                continue
            key = renamed.lower()
            if key in written:
                continue
            written.add(key)
            zf.writestr(renamed, item["bytes"])
    return output.getvalue()


def generate_perdue_tsheet(
    prisma_file,
    creative_files,
    creative_mapping_text: str = "",
    landing_urls_text: str = "",
    template_path: str | Path | None = None,
) -> tuple[bytes, bytes, list[str], dict]:
    """Return (xlsm_bytes, renamed_creatives_zip_bytes, warnings, stats)."""
    raw_rows, _ = read_prisma_export(prisma_file)
    preview = preview_perdue_setup(
        prisma_file=prisma_file,
        creative_files=creative_files,
        creative_mapping_text=creative_mapping_text,
        landing_urls_text=landing_urls_text,
    )

    template = Path(template_path) if template_path else MASTER_TEMPLATE
    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")
    wb = load_workbook(template, keep_vba=True)
    if TRAFFIC_SHEET not in wb.sheetnames:
        raise KeyError(f"Missing worksheet: {TRAFFIC_SHEET}")

    _write_prisma_sheet(wb, raw_rows)
    ws = wb[TRAFFIC_SHEET]
    header_row = _find_header_row(ws)
    first = header_row + 1
    cols = _header_map(ws, header_row)

    # Preserve first-row formatting, then clear old values.
    style_cells = {c: copy(ws.cell(first, c)._style) for c in range(1, ws.max_column + 1)}
    number_formats = {c: ws.cell(first, c).number_format for c in range(1, ws.max_column + 1)}
    for r in range(first, max(ws.max_row, first + len(preview["rows"]) + 20) + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None

    for i, item in enumerate(preview["rows"]):
        r = first + i
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c)._style = copy(style_cells[c]); ws.cell(r, c).number_format = number_formats[c]
        rec = item.get("_record", {})
        tax = item.get("_tax", {})
        channel = tax.get("Channel", "")
        size = tax.get("Size", "")
        values = {
            "site": _get(rec, "Media outlet / Supplier name (Prisma)", "Media outlet / Supplier name (ad server)") or tax.get("Partner/Vendor", ""),
            "placement_id": _get(rec, "Ad server ID", "Placement ID", "DCM Placement ID"),
            "placement": item.get("placement_name", ""),
            "dimension": "" if channel in {"OLV", "CTV"} else size,
            "video_length": size if channel in {"OLV", "CTV"} else "",
            "vast": "VAST" if channel in {"OLV", "CTV"} else "",
            "ad": item.get("ad_name", ""),
            "action": "New",
            "creative": item.get("renamed_creative", ""),
            "studio": "N",
            "rotation": 1,
            "start": _parse_date(_get(rec, "Flight start date", "Start Date")),
            "end": _parse_date(_get(rec, "Flight end date", "End Date")),
            "url": item.get("final_url", ""),
        }
        for key, value in values.items():
            if key in cols:
                ws.cell(r, cols[key]).value = value

    # Perdue direct creative workflow: clear old Multi rows to prevent stale rotations.
    if MULTI_SHEET in wb.sheetnames:
        multi = wb[MULTI_SHEET]
        for row in multi.iter_rows(min_row=2):
            for cell in row: cell.value = None

    out = io.BytesIO()
    wb.save(out)
    renamed_zip = build_renamed_creatives_zip(preview)
    stats = {
        "placement_count": len(preview["rows"]),
        "creative_count": preview["creative_count"],
        "matched_count": preview["matched_count"],
        "unmatched_count": preview["unmatched_count"],
        "url_matched_count": preview["url_matched_count"],
        "url_unmatched_count": preview["url_unmatched_count"],
    }
    return out.getvalue(), renamed_zip, preview["warnings"], stats
