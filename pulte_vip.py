from __future__ import annotations

import csv
import io
import re
from copy import copy
from pathlib import Path
from typing import BinaryIO, Iterable

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
MASTER_TEMPLATE = BASE_DIR / "master_template.xlsm"
TRACKING_CODES_FILE = BASE_DIR / "Pulte_Adobe_Tracking_Codes.xlsm"

PRISMA_SHEET = "Prisma Export - Paste as values"
TRAFFIC_SHEET = "Traffic_Doc"
ROTATION_SHEET = "Multi-Ad or Creative Rotation"

TRAFFIC_START_ROW = 8
TRAFFIC_LAST_COLUMN = 24


def _clean(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalize(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", _clean(value).lower())


def _split_placement(value: str) -> list[str]:
    return [_clean(part) for part in _clean(value).split("_")]


def _read_uploaded_bytes(uploaded_file) -> bytes:
    uploaded_file.seek(0)
    return uploaded_file.read()


def _decode_csv(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def read_prisma_csv(uploaded_file) -> tuple[list[list[str]], list[dict[str, str]]]:
    raw_text = _decode_csv(_read_uploaded_bytes(uploaded_file))
    raw_rows = list(csv.reader(io.StringIO(raw_text)))

    header_index = None
    for index, row in enumerate(raw_rows):
        normalized = {_clean(cell) for cell in row}
        if "Placement Name" in normalized and "Placement ID" in normalized:
            header_index = index
            break

    if header_index is None:
        raise ValueError(
            "The Prisma header row could not be found. "
            "The CSV must contain Placement Name and Placement ID columns."
        )

    headers = [_clean(value) for value in raw_rows[header_index]]
    records: list[dict[str, str]] = []

    for row in raw_rows[header_index + 1 :]:
        padded = row + [""] * (len(headers) - len(row))
        record = dict(zip(headers, padded))
        if _clean(record.get("Placement Name")):
            records.append(record)

    if not records:
        raise ValueError("No placement rows were found in the Prisma CSV.")

    return raw_rows, records


def clear_old_template_data(workbook) -> None:
    if PRISMA_SHEET in workbook.sheetnames:
        sheet = workbook[PRISMA_SHEET]
        max_row = max(sheet.max_row, 5000)
        max_col = max(sheet.max_column, 57)
        for row in sheet.iter_rows(
            min_row=1,
            max_row=max_row,
            min_col=1,
            max_col=max_col,
        ):
            for cell in row:
                cell.value = None

    if TRAFFIC_SHEET in workbook.sheetnames:
        sheet = workbook[TRAFFIC_SHEET]
        max_row = max(sheet.max_row, 5000)

        for row in sheet.iter_rows(
            min_row=TRAFFIC_START_ROW,
            max_row=max_row,
            min_col=1,
            max_col=TRAFFIC_LAST_COLUMN,
        ):
            for cell in row:
                cell.value = None

        for coordinate in ("B1", "B2", "B4", "B5"):
            sheet[coordinate] = None

    if ROTATION_SHEET in workbook.sheetnames:
        sheet = workbook[ROTATION_SHEET]
        max_row = max(sheet.max_row, 5000)
        max_col = max(sheet.max_column, 7)

        for row in sheet.iter_rows(
            min_row=2,
            max_row=max_row,
            min_col=1,
            max_col=max_col,
        ):
            for cell in row:
                cell.value = None

    for sheet_name in ("Native - DV360", "Native - TTD", "Native - Oath"):
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            max_row = max(sheet.max_row, 5000)
            max_col = max(sheet.max_column, 10)

            for row in sheet.iter_rows(
                min_row=2,
                max_row=max_row,
                min_col=1,
                max_col=max_col,
            ):
                for cell in row:
                    cell.value = None


def paste_prisma_export(workbook, raw_rows: list[list[str]]) -> None:
    if PRISMA_SHEET not in workbook.sheetnames:
        raise KeyError(f"Missing worksheet: {PRISMA_SHEET}")

    sheet = workbook[PRISMA_SHEET]

    for row_index, row_values in enumerate(raw_rows, start=1):
        for column_index, value in enumerate(row_values, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)


def _load_tracking_lookup() -> dict[str, dict[str, str]]:
    if not TRACKING_CODES_FILE.exists():
        raise FileNotFoundError(
            f"Tracking code workbook not found: {TRACKING_CODES_FILE.name}"
        )

    workbook = load_workbook(
        TRACKING_CODES_FILE,
        read_only=True,
        data_only=True,
    )

    sheet_name = "ChannelTrackingValues"
    if sheet_name not in workbook.sheetnames:
        raise KeyError(f"Missing worksheet in tracking workbook: {sheet_name}")

    sheet = workbook[sheet_name]

    sections = {
        "Medium": (1, 2),
        "Source": (3, 4),
        "Division": (5, 6),
        "Region": (7, 8),
        "Content": (9, 10),
        "Campaign": (11, 12),
        "Vendor": (13, 14),
        "Image": (15, 16),
    }

    lookup: dict[str, dict[str, str]] = {}

    for section, (category_col, abbreviation_col) in sections.items():
        lookup[section] = {}

        for row in range(4, sheet.max_row + 1):
            category = _clean(sheet.cell(row=row, column=category_col).value)
            abbreviation = _clean(
                sheet.cell(row=row, column=abbreviation_col).value
            )

            if category and abbreviation:
                lookup[section][_normalize(category)] = abbreviation

    return lookup


def _lookup_code(
    lookup: dict[str, dict[str, str]],
    section: str,
    value: str,
    default: str = "NA-_-",
) -> str:
    return lookup.get(section, {}).get(_normalize(value), default)


def _find_dimension(placement_name: str) -> str:
    match = re.search(r"(?<!\d)(\d{2,4})x(\d{2,4})(?!\d)", placement_name)
    if not match:
        return ""
    return f"{match.group(1)}x{match.group(2)}"


def _find_image_type(placement_name: str, creative_name: str = "") -> str:
    combined = f"{placement_name}_{creative_name}".upper()

    for image_type in (
        "EXTD",
        "EXTT",
        "EXT",
        "LIFE",
        "AMN",
        "OFPK",
        "OFP",
        "POOL",
        "FP",
        "VID",
    ):
        if re.search(rf"(^|_|\s){re.escape(image_type)}($|_|\s)", combined):
            return image_type

    return "NA"


def _source_from_placement(placement_name: str, supplier: str) -> str:
    combined = f"{placement_name} {supplier}".lower()

    if "zillow" in combined:
        return "zillow.com"
    if "realtor" in combined:
        return "realtor"
    if "newhomesource" in combined or "new home source" in combined:
        return "newhomesource.com"
    if "teads" in combined:
        return "teads"
    if "youtube" in combined:
        return "youtube.com"
    if "hulu" in combined:
        return "hulu"
    if "programmatic" in combined:
        return "programmatic"

    return supplier


def _site_name(source: str, supplier_name: str) -> str:
    normalized = _normalize(source)

    if normalized == "zillowcom":
        return "Zillow.com"
    if normalized == "realtor":
        return "Realtor"
    if normalized == "newhomesourcecom":
        return "NewHomeSource.com"

    return _clean(supplier_name) or source


def _brand_from_placement(placement_name: str) -> str:
    placement_lower = placement_name.lower()

    if "del webb" in placement_lower:
        return "Del Webb"
    if "centex" in placement_lower:
        return "Centex"
    if "divosta" in placement_lower:
        return "DiVosta"
    if "john wieland" in placement_lower or "wieland" in placement_lower:
        return "Wieland"

    return "Pulte"


def _community_id(placement_name: str) -> str:
    parts = _split_placement(placement_name)

    for part in reversed(parts):
        match = re.search(r"(?<!\d)(\d{5,7})(?!\d)", part)
        if match:
            return match.group(1)

    match = re.search(r"(?<!\d)(\d{5,7})(?!\d)", placement_name)
    return match.group(1) if match else ""


def _division_brand_campaign_community(
    placement_name: str,
) -> tuple[str, str, str, str]:
    parts = _split_placement(placement_name)

    brand_index = None
    for index, part in enumerate(parts):
        if _normalize(part) in {
            "pulte",
            "delwebb",
            "centex",
            "divosta",
            "wieland",
            "johnwieland",
        }:
            brand_index = index
            break

    if brand_index is None:
        return "", _brand_from_placement(placement_name), "", ""

    division = parts[brand_index - 1] if brand_index > 0 else ""
    brand = parts[brand_index]

    campaign = parts[brand_index + 1] if brand_index + 1 < len(parts) else ""
    community = parts[brand_index + 2] if brand_index + 2 < len(parts) else ""

    return division, brand, campaign, community


def _region_from_placement(placement_name: str, division: str) -> str:
    parts = _split_placement(placement_name)

    community_id = _community_id(placement_name)
    community_id_index = None

    for index, part in enumerate(parts):
        if community_id and community_id in part:
            community_id_index = index
            break

    if community_id_index is not None:
        trailing_parts = parts[community_id_index + 1 :]
        for part in trailing_parts:
            words = part.split()
            if words:
                first_word = words[0].strip()
                if first_word:
                    return first_word

    division_to_region = {
        "Southwest Florida": "fort myers-naples",
        "West Florida": "tampa",
        "Raleigh": "raleigh",
        "Indianapolis-Louisville": "indianapolis",
        "Central Florida": "orlando",
        "North Florida": "jacksonville",
        "Northeast Florida": "jacksonville",
        "Coastal Carolinas": "coastal carolinas",
        "East Carolina": "raleigh",
        "Charlotte": "charlotte",
    }

    return division_to_region.get(division, division)


def parse_pulte_placement(
    placement_name: str,
    supplier_name: str = "",
) -> dict[str, str]:
    division, brand, campaign, community = (
        _division_brand_campaign_community(placement_name)
    )

    source = _source_from_placement(placement_name, supplier_name)
    dimension = _find_dimension(placement_name)
    community_id = _community_id(placement_name)
    region = _region_from_placement(placement_name, division)

    return {
        "placement_name": placement_name,
        "source": source,
        "site_name": _site_name(source, supplier_name),
        "dimension": dimension,
        "division": division,
        "brand": brand,
        "campaign": campaign,
        "community": community,
        "community_id": community_id,
        "region": region,
    }


def _creative_names(creative_files: Iterable) -> list[str]:
    names = []
    for uploaded_file in creative_files:
        name = Path(uploaded_file.name).name
        if name:
            names.append(name)
    return names


def _creative_score(
    creative_name: str,
    parsed: dict[str, str],
    image_type: str,
) -> int:
    normalized_creative = _normalize(creative_name)
    score = 0

    checks = (
        (parsed["dimension"], 10),
        (parsed["community"], 8),
        (parsed["community_id"], 7),
        (parsed["brand"], 4),
        (parsed["source"], 3),
        (image_type, 5),
    )

    for value, points in checks:
        normalized_value = _normalize(value)
        if normalized_value and normalized_value in normalized_creative:
            score += points

    community_words = [
        word
        for word in re.split(r"\W+", parsed["community"].lower())
        if len(word) >= 4
    ]
    score += sum(
        2 for word in community_words if word in creative_name.lower()
    )

    return score


def match_creative(
    creative_names: list[str],
    parsed: dict[str, str],
) -> tuple[str, str]:
    image_type = _find_image_type(parsed["placement_name"])

    ranked = sorted(
        (
            (_creative_score(name, parsed, image_type), name)
            for name in creative_names
        ),
        reverse=True,
    )

    if not ranked or ranked[0][0] <= 0:
        return "", image_type

    matched_name = ranked[0][1]
    image_type = _find_image_type(
        parsed["placement_name"],
        matched_name,
    )
    return matched_name, image_type


def _parse_urls(landing_urls_text: str) -> list[str]:
    urls = []

    for line in landing_urls_text.splitlines():
        value = line.strip()
        if value:
            urls.append(value)

    return urls


def match_landing_url(
    urls: list[str],
    parsed: dict[str, str],
) -> str:
    community_id = parsed["community_id"]
    brand = parsed["brand"].lower()

    if community_id:
        for url in urls:
            if community_id in url:
                return url

    if brand == "del webb":
        for url in urls:
            if "delwebb.com" in url.lower():
                return url

    if brand == "pulte":
        for url in urls:
            if "pulte.com" in url.lower():
                return url

    if len(urls) == 1:
        return urls[0]

    return ""


def build_ad_name(parsed: dict[str, str]) -> str:
    parts = [
        parsed["division"],
        parsed["brand"],
        parsed["campaign"],
        parsed["community"],
        parsed["dimension"],
    ]
    return "_".join(part for part in parts if part)


def build_cmp_code(
    parsed: dict[str, str],
    image_type: str,
    lookup: dict[str, dict[str, str]],
) -> str:
    medium_code = _lookup_code(lookup, "Medium", "Endemic")
    source_code = _lookup_code(lookup, "Source", parsed["source"])
    division_code = _lookup_code(lookup, "Division", parsed["division"])
    region_code = _lookup_code(lookup, "Region", parsed["region"])

    content_value = parsed["brand"]
    content_code = _lookup_code(lookup, "Content", content_value, "NA")
    if parsed["community_id"]:
        content_code = f"{content_code}{parsed['community_id']}"

    campaign_code = _lookup_code(
        lookup,
        "Campaign",
        parsed["campaign"],
    )
    vendor_code = _lookup_code(lookup, "Vendor", "Assembly")
    image_code = _lookup_code(lookup, "Image", image_type)

    return "".join(
        [
            medium_code,
            source_code,
            division_code,
            region_code,
            f"{content_code}-_-",
            campaign_code,
            vendor_code,
            image_code,
        ]
    )


def _append_cmp(url: str, cmp_code: str) -> str:
    if not url:
        return ""

    separator = "&" if "?" in url else "?"
    return f"{url}{separator}cmp={cmp_code}"


def _campaign_name_from_raw_rows(raw_rows: list[list[str]]) -> str:
    for row in raw_rows:
        if row and _clean(row[0]) == "Campaign name:":
            return _clean(row[1]) if len(row) > 1 else ""
    return ""


def _copy_row_format(sheet, source_row: int, target_row: int) -> None:
    if source_row == target_row:
        return

    for column in range(1, TRAFFIC_LAST_COLUMN + 1):
        source = sheet.cell(row=source_row, column=column)
        target = sheet.cell(row=target_row, column=column)

        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.font:
            target.font = copy(source.font)
        if source.fill:
            target.fill = copy(source.fill)
        if source.border:
            target.border = copy(source.border)
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)

    sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height


def populate_traffic_sheet(
    workbook,
    raw_rows: list[list[str]],
    records: list[dict[str, str]],
    creative_files: Iterable,
    landing_urls_text: str,
) -> list[str]:
    if TRAFFIC_SHEET not in workbook.sheetnames:
        raise KeyError(f"Missing worksheet: {TRAFFIC_SHEET}")

    sheet = workbook[TRAFFIC_SHEET]
    lookup = _load_tracking_lookup()
    creative_names = _creative_names(creative_files)
    landing_urls = _parse_urls(landing_urls_text)
    warnings: list[str] = []

    campaign_name = _campaign_name_from_raw_rows(raw_rows)
    sheet["B1"] = campaign_name

    template_row = TRAFFIC_START_ROW

    for index, record in enumerate(records):
        output_row = TRAFFIC_START_ROW + index
        _copy_row_format(sheet, template_row, output_row)

        placement_name = _clean(record.get("Placement Name"))
        supplier_name = (
            _clean(record.get("Media outlet / Supplier name (ad server)"))
            or _clean(record.get("Media outlet / Supplier name (Prisma)"))
        )

        parsed = parse_pulte_placement(
            placement_name=placement_name,
            supplier_name=supplier_name,
        )

        creative_name, image_type = match_creative(
            creative_names,
            parsed,
        )

        landing_url = match_landing_url(
            landing_urls,
            parsed,
        )

        cmp_code = build_cmp_code(
            parsed,
            image_type,
            lookup,
        )

        final_url = _append_cmp(landing_url, cmp_code)
        ad_name = build_ad_name(parsed)

        if not creative_name:
            warnings.append(
                f"No creative matched placement ID "
                f"{_clean(record.get('Ad server ID'))}: {placement_name}"
            )

        if not landing_url:
            warnings.append(
                f"No landing URL matched placement ID "
                f"{_clean(record.get('Ad server ID'))}: {placement_name}"
            )

        values = [
            None,
            parsed["site_name"],
            _clean(record.get("Ad server ID")),
            placement_name,
            "1x1",
            None,
            None,
            ad_name,
            None,
            "New",
            creative_name,
            "Yes",
            None,
            _clean(record.get("Flight start date")),
            _clean(record.get("Flight end date")),
            final_url,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        ]

        for column, value in enumerate(values, start=1):
            sheet.cell(row=output_row, column=column, value=value)

    return warnings


def generate_pulte_tsheet(
    prisma_file,
    creative_files,
    landing_urls_text: str,
) -> tuple[bytes, list[str]]:
    if not MASTER_TEMPLATE.exists():
        raise FileNotFoundError(
            f"Master template not found: {MASTER_TEMPLATE.name}"
        )

    raw_rows, records = read_prisma_csv(prisma_file)

    workbook = load_workbook(
        MASTER_TEMPLATE,
        keep_vba=True,
    )

    clear_old_template_data(workbook)
    paste_prisma_export(workbook, raw_rows)

    warnings = populate_traffic_sheet(
        workbook=workbook,
        raw_rows=raw_rows,
        records=records,
        creative_files=creative_files,
        landing_urls_text=landing_urls_text,
    )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue(), warnings
