from __future__ import annotations

import csv
import io
import re
from copy import copy
from functools import lru_cache
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
MASTER_TEMPLATE = BASE_DIR / "master_template.xlsm"
TRACKING_CODES_FILE = BASE_DIR / "Pulte_Adobe_Tracking_Codes.xlsm"

PRISMA_SHEET = "Prisma Export - Paste as values"
TRAFFIC_SHEET = "Traffic_Doc"
ROTATION_SHEET = "Multi-Ad or Creative Rotation"

TRAFFIC_START_ROW = 8
TRAFFIC_LAST_COLUMN = 24


# ---------------------------------------------------------------------------
# BASIC HELPERS
# ---------------------------------------------------------------------------

def _clean(value) -> str:
    return "" if value is None else str(value).strip()


def _normalize(value) -> str:
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


def _find_traffic_header_row(sheet) -> int:
    for row_number in range(1, min(sheet.max_row, 30) + 1):
        values = {
            _normalize(sheet.cell(row=row_number, column=column).value)
            for column in range(1, min(sheet.max_column, 40) + 1)
        }
        if "adname" in values and "creativefilename" in values:
            return row_number
    return TRAFFIC_START_ROW - 1


def _traffic_data_start_row(sheet) -> int:
    return _find_traffic_header_row(sheet) + 1


# ---------------------------------------------------------------------------
# PRISMA
# ---------------------------------------------------------------------------

def read_prisma_csv(uploaded_file) -> tuple[list[list[str]], list[dict[str, str]]]:
    raw_text = _decode_csv(_read_uploaded_bytes(uploaded_file))

    try:
        dialect = csv.Sniffer().sniff(raw_text[:10000], delimiters=",;\t|")
        reader = csv.reader(io.StringIO(raw_text), dialect)
    except csv.Error:
        reader = csv.reader(io.StringIO(raw_text))

    raw_rows = list(reader)

    header_index = None
    headers = None

    for index, row in enumerate(raw_rows):
        cleaned = [_clean(cell).replace("\n", " ") for cell in row]
        normalized = {_normalize(cell): cell for cell in cleaned if cell}
        if "placementname" in normalized:
            header_index = index
            headers = cleaned
            break

    if header_index is None or headers is None:
        raise ValueError(
            "The Prisma header row could not be found. "
            "A Placement Name column is required."
        )

    records: list[dict[str, str]] = []

    for raw_index in range(header_index + 1, len(raw_rows)):
        row = raw_rows[raw_index]
        padded = row + [""] * max(0, len(headers) - len(row))
        record = dict(zip(headers, padded[:len(headers)]))

        placement_name = ""
        for key, value in record.items():
            if _normalize(key) == "placementname":
                placement_name = _clean(value)
                break

        if not placement_name:
            continue

        row_type = ""
        for candidate in ("Row Type", "Type", "Package / Placement"):
            for key, value in record.items():
                if _normalize(key) == _normalize(candidate):
                    row_type = _clean(value).lower()
                    break

        if row_type == "package" or placement_name.lower().startswith("package:"):
            continue

        record["Placement Name"] = placement_name
        record["_source_excel_row"] = str(raw_index + 1)
        records.append(record)

    if not records:
        raise ValueError(
            "The Placement Name header was found, but no placement rows "
            "were detected below it."
        )

    return raw_rows, records


def _record_value(record: dict[str, str], *names: str) -> str:
    normalized = {_normalize(k): v for k, v in record.items()}
    for name in names:
        value = normalized.get(_normalize(name))
        if value is not None:
            return _clean(value)
    return ""


def clear_old_template_data(workbook) -> None:
    if PRISMA_SHEET in workbook.sheetnames:
        sheet = workbook[PRISMA_SHEET]
        for row in sheet.iter_rows(
            min_row=1,
            max_row=max(sheet.max_row, 1),
            min_col=1,
            max_col=max(sheet.max_column, 57),
        ):
            for cell in row:
                cell.value = None

    if TRAFFIC_SHEET in workbook.sheetnames:
        sheet = workbook[TRAFFIC_SHEET]
        first_data_row = _traffic_data_start_row(sheet)

        for row in sheet.iter_rows(
            min_row=first_data_row,
            max_row=max(sheet.max_row, first_data_row),
            min_col=1,
            max_col=max(sheet.max_column, 40),
        ):
            for cell in row:
                cell.value = None

        for coordinate in ("B1", "B2", "B4", "B5"):
            sheet[coordinate] = None

    if ROTATION_SHEET in workbook.sheetnames:
        sheet = workbook[ROTATION_SHEET]
        for row in sheet.iter_rows(
            min_row=2,
            max_row=max(sheet.max_row, 2),
            min_col=1,
            max_col=max(sheet.max_column, 10),
        ):
            for cell in row:
                cell.value = None

    for sheet_name in ("Native - DV360", "Native - TTD", "Native - Oath"):
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows(
                min_row=2,
                max_row=max(sheet.max_row, 2),
                min_col=1,
                max_col=max(sheet.max_column, 1),
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


# ---------------------------------------------------------------------------
# PULTE TRACKING WORKBOOK
# ---------------------------------------------------------------------------

@lru_cache(maxsize=1)
def _load_tracking_data() -> dict:
    """
    Reads the official Pulte tracking workbook at runtime.

    Nothing important is hardcoded here:
      - Medium
      - Source
      - Division
      - Region
      - Content
      - Campaign
      - Vendor
      - Image
      - SEM Region Mapping Tab

    If Pulte updates the tracking workbook in GitHub, this script will use
    those updated mappings after the Streamlit app restarts.
    """
    if not TRACKING_CODES_FILE.exists():
        raise FileNotFoundError(
            f"Tracking code workbook not found: {TRACKING_CODES_FILE.name}"
        )

    workbook = load_workbook(
        TRACKING_CODES_FILE,
        read_only=True,
        data_only=True,
    )

    values_sheet_name = "ChannelTrackingValues"
    region_sheet_name = "SEM Region Mapping Tab"

    if values_sheet_name not in workbook.sheetnames:
        raise KeyError(
            f"Missing worksheet in tracking workbook: {values_sheet_name}"
        )

    values_sheet = workbook[values_sheet_name]

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
    reverse: dict[str, dict[str, str]] = {}

    for section, (category_col, abbreviation_col) in sections.items():
        lookup[section] = {}
        reverse[section] = {}

        for row in range(4, values_sheet.max_row + 1):
            category = _clean(values_sheet.cell(row=row, column=category_col).value)
            abbreviation = _clean(
                values_sheet.cell(row=row, column=abbreviation_col).value
            )

            if not category or not abbreviation:
                continue

            lookup[section][_normalize(category)] = abbreviation
            reverse[section][_normalize(abbreviation)] = category

    # Build region mapping from the official SEM mapping tab.
    region_rows: list[dict[str, str]] = []

    if region_sheet_name in workbook.sheetnames:
        region_sheet = workbook[region_sheet_name]

        # Row 2 contains:
        # Category | Abbreviation | SEM DMA | Division
        for row in range(3, region_sheet.max_row + 1):
            category = _clean(region_sheet.cell(row=row, column=1).value)
            abbreviation = _clean(region_sheet.cell(row=row, column=2).value)
            sem_dma = _clean(region_sheet.cell(row=row, column=3).value)
            division = _clean(region_sheet.cell(row=row, column=4).value)

            if not category:
                continue

            region_rows.append(
                {
                    "category": category,
                    "abbreviation": abbreviation,
                    "sem_dma": sem_dma,
                    "division": division,
                }
            )

    return {
        "lookup": lookup,
        "reverse": reverse,
        "region_rows": region_rows,
    }


def _lookup_code(
    tracking: dict,
    section: str,
    value: str,
    default: str = "",
) -> str:
    if not value:
        return default

    normalized = _normalize(value)
    section_lookup = tracking["lookup"].get(section, {})

    # Direct category lookup.
    if normalized in section_lookup:
        return section_lookup[normalized]

    # Also accept an abbreviation passed in by placement taxonomy.
    reverse_section = tracking["reverse"].get(section, {})
    if normalized in reverse_section:
        category = reverse_section[normalized]
        return section_lookup.get(_normalize(category), default)

    return default


def _match_tracking_category(
    text: str,
    tracking: dict,
    section: str,
) -> str:
    """
    Finds the best official category appearing in free-form placement text.
    Longest normalized match wins to avoid 'Florida' beating
    'Southwest Florida', etc.
    """
    normalized_text = _normalize(text)
    candidates = []

    for normalized_category in tracking["lookup"].get(section, {}):
        if not normalized_category or normalized_category == "choosevalue":
            continue
        if normalized_category in normalized_text:
            candidates.append(normalized_category)

    if not candidates:
        return ""

    winner = max(candidates, key=len)

    # Convert normalized category back to original category text.
    for code_norm, category in tracking["reverse"].get(section, {}).items():
        if _normalize(category) == winner:
            return category

    # Fallback: return normalized key; _lookup_code can still resolve it.
    return winner


# ---------------------------------------------------------------------------
# PLACEMENT PARSING
# ---------------------------------------------------------------------------

def _find_dimension(text: str) -> str:
    match = re.search(
        r"(?<!\d)(\d{1,4})\s*[xX]\s*(\d{1,4})(?!\d)",
        _clean(text),
    )
    return f"{match.group(1)}x{match.group(2)}" if match else ""


def _brand_from_placement(placement_name: str) -> str:
    normalized = _normalize(placement_name)

    if "delwebb" in normalized:
        return "Del Webb"
    if "centex" in normalized:
        return "Centex"
    if "divosta" in normalized:
        return "DiVosta"
    if "johnwieland" in normalized or "wieland" in normalized:
        return "Wieland"
    if "americanwest" in normalized:
        return "American West"

    return "Pulte"


def _community_id(placement_name: str) -> str:
    # Prefer IDs near the right side of the placement name.
    for part in reversed(_split_placement(placement_name)):
        match = re.search(r"(?<!\d)(\d{5,7})(?!\d)", part)
        if match:
            return match.group(1)

    return ""


def _source_from_placement(
    placement_name: str,
    supplier_name: str,
    tracking: dict,
) -> str:
    combined = f"{placement_name} {supplier_name}"

    # First use exact/known source aliases.
    aliases = [
        ("new home source", "newhomesource.com"),
        ("newhomesource", "newhomesource.com"),
        ("zillow", "zillow.com"),
        ("realtor", "realtor"),
        ("youtube", "youtube.com"),
        ("pinterest", "pinterest.com"),
        ("instagram", "instagram.com"),
        ("facebook", "facebook.com"),
        ("teads", "teads"),
        ("spotx", "SpotX"),
        ("hulu", "hulu"),
        ("programmatic", "programmatic"),
        ("google", "google.com"),
        ("bing", "bing.com"),
    ]

    combined_lower = combined.lower()
    for needle, category in aliases:
        if needle in combined_lower:
            return category

    # Then try any official Source value directly.
    detected = _match_tracking_category(combined, tracking, "Source")
    if detected:
        return detected

    # Supplier is only accepted if it is itself an official tracking source.
    supplier_code = _lookup_code(tracking, "Source", supplier_name)
    if supplier_code:
        return supplier_name

    return ""


def _medium_from_placement(
    placement_name: str,
    source: str,
    tracking: dict,
) -> str:
    detected = _match_tracking_category(placement_name, tracking, "Medium")
    if detected:
        return detected

    normalized = _normalize(placement_name)

    # Placement-taxonomy aliases.
    if "programmatic" in normalized:
        return "Programmatic"
    if "display" in normalized:
        return "Display"
    if "video" in normalized or "olv" in normalized:
        return "Video"
    if "socialpaid" in normalized or "paidsocial" in normalized:
        return "Social Paid"

    # Endemic sources such as Realtor/Zillow/NHS commonly use Endemic.
    if _normalize(source) in {
        "realtor",
        "zillowcom",
        "newhomesourcecom",
    }:
        return "Endemic"

    return ""


def _site_name(source: str, supplier_name: str) -> str:
    normalized = _normalize(source)

    names = {
        "zillowcom": "Zillow.com",
        "realtor": "Realtor",
        "newhomesourcecom": "NewHomeSource.com",
        "teads": "Teads",
        "youtubecom": "YouTube",
        "hulu": "Hulu",
        "programmatic": "Programmatic",
    }

    return names.get(normalized, _clean(supplier_name) or source)


def _division_from_placement(
    placement_name: str,
    tracking: dict,
) -> str:
    """
    Uses the official Division list. Longest match wins.

    Examples automatically supported from the workbook:
      Tennessee -> Tennessee
      Southwest Florida -> Southwest Florida
      Northeast Corridor -> Northeast Corridor
      Central Texas -> Central Texas (if present in current workbook)
      etc.
    """
    detected = _match_tracking_category(placement_name, tracking, "Division")
    if detected:
        return detected

    # Common taxonomy spelling aliases.
    aliases = {
        "tennessee": "Tennessee",
        "indianapoliskentucky": "Indianapolis-Louisville",
        "indianapolislouisville": "Indianapolis-Louisville",
        "midatlantic": "Mid-Atlantic",
        "northeastcorridor": "Northeast Corridor",
        "southwestflorida": "Southwest Florida",
        "southeastflorida": "Southeast Florida",
        "southerncalifornia": "Southern California",
        "northerncalifornia": "Northern California",
        "pacificnorthwest": "Pacific Northwest",
        "westflorida": "West Florida",
        "northflorida": "North Florida",
        "northeastflorida": "Northeast Florida",
        "eastcarolina": "East Carolina",
        "coastalcarolinas": "Coastal Carolinas",
        "newengland": "New England",
        "newmexico": "New Mexico",
        "sanantonio": "San Antonio",
    }

    normalized = _normalize(placement_name)
    matches = [
        (len(alias), official)
        for alias, official in aliases.items()
        if alias in normalized and _lookup_code(tracking, "Division", official)
    ]

    return max(matches)[1] if matches else ""


def _region_alias_matches(text: str, region_row: dict[str, str]) -> bool:
    """
    Region can appear as:
      - category: phoenix
      - SEM DMA: Phoenix AZ
      - abbreviation: PHX-_-
      - short DMA token in placement: PHX AZ / PHX
    """
    normalized_text = _normalize(text)

    category = _normalize(region_row["category"])
    sem_dma = _normalize(region_row["sem_dma"])
    abbreviation = _normalize(region_row["abbreviation"])

    abbreviation_short = re.sub(r"[^a-z0-9]", "", region_row["abbreviation"].split("-")[0].lower())

    candidates = {
        category,
        sem_dma,
        abbreviation,
        abbreviation_short,
    }

    # Also allow the first DMA word when sufficiently distinctive.
    dma_words = re.findall(r"[A-Za-z]+", region_row["sem_dma"])
    if dma_words and len(dma_words[0]) >= 4:
        candidates.add(_normalize(dma_words[0]))

    return any(
        candidate and candidate in normalized_text
        for candidate in candidates
    )


def _region_from_placement(
    placement_name: str,
    division: str,
    tracking: dict,
) -> tuple[str, str]:
    """
    Returns (region, warning).

    Priority:
      1. Exact region/DMA/abbreviation token found in placement name,
         scoped to the detected division.
      2. If that division has exactly one official SEM mapping, use it.
         This is how Tennessee automatically becomes Nashville.
      3. If division has multiple possible regions and no region token is
         present, DO NOT GUESS. Return warning.
    """
    rows = tracking["region_rows"]

    division_rows = [
        row for row in rows
        if _normalize(row["division"]) == _normalize(division)
    ]

    # Match placement against the division's permitted regions.
    scoped_matches = [
        row for row in division_rows
        if _region_alias_matches(placement_name, row)
    ]

    if scoped_matches:
        # Prefer longest category/SEM DMA match.
        winner = max(
            scoped_matches,
            key=lambda row: max(
                len(_normalize(row["category"])),
                len(_normalize(row["sem_dma"])),
            ),
        )
        return winner["category"], ""

    # Some rows in the official mapping have blank Division.
    # They can still be used if explicitly present in placement taxonomy.
    global_matches = [
        row for row in rows
        if _region_alias_matches(placement_name, row)
    ]

    if len(global_matches) == 1:
        return global_matches[0]["category"], ""

    # Unique-division fallback.
    unique_categories = []
    for row in division_rows:
        category = row["category"]
        if category and _normalize(category) != "choosevalue":
            if _normalize(category) not in {_normalize(x) for x in unique_categories}:
                unique_categories.append(category)

    if len(unique_categories) == 1:
        return unique_categories[0], ""

    # Explicit business-safe defaults where the workbook/business taxonomy
    # has a known umbrella market.
    explicit_defaults = {
        "Tennessee": "nashville",
        "Raleigh": "raleigh",
        "San Antonio": "san antonio",
        "West Florida": "tampa",
        "North Florida": "jacksonville",
        "Northeast Florida": "northeast florida",
        "Southwest Florida": "fort myers-naples",
        "Charlotte": "charlotte",
    }

    preferred = explicit_defaults.get(division)
    if preferred and _lookup_code(tracking, "Region", preferred):
        return preferred, ""

    if not division:
        return "", "Division could not be detected, so Region could not be resolved."

    return (
        "",
        f"Region could not be resolved safely for division '{division}'. "
        "This division has multiple possible DMA/region values and the "
        "placement name did not contain a recognizable region token.",
    )


def _campaign_from_placement(
    placement_name: str,
    tracking: dict,
) -> str:
    detected = _match_tracking_category(placement_name, tracking, "Campaign")
    if detected:
        return detected

    aliases = {
        "heavyup": "Heavy Up",
        "qmi": "QMI",
        "grandopening": "Grand Opening (inclusive of all openings)",
        "promotion": "Promotion (inclusive of Incentives)",
        "awareness": "Awareness",
        "community": "Community",
        "prospect": "Prospect",
        "lead": "Lead",
        "traffic": "Traffic",
        "comingsoon": "Coming Soon",
        "nurturing": "Nurturing",
    }

    normalized = _normalize(placement_name)
    for token, official in sorted(
        aliases.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    ):
        if token in normalized and _lookup_code(tracking, "Campaign", official):
            return official

    return ""


def _image_category(
    placement_name: str,
    creative_name: str,
    tracking: dict,
) -> str:
    combined = f"{placement_name} {creative_name}"

    # First detect full official Image category names.
    detected = _match_tracking_category(combined, tracking, "Image")
    if detected:
        return detected

    # Then common Pulte taxonomy abbreviations -> official Image categories.
    aliases = {
        "EXTD": "Exterior-Daylight",
        "EXTT": "Exterior-Twilight",
        "EXT": "Exterior",
        "LIFE": "Lifestyle",
        "AMN": "Amenity",
        "OFPK": "Open Floor Plan- Kitchen",
        "OFP": "Open Floor Plan",
        "POOL": "Pool",
        "FP": "Floor Plan",
        "STOR": "Instagram Story",
        "IFP": "Interactive Floor Plan",
        "PPC": "Pulte Planning Center",
        "SAM": "Site Availability Map",
        "SLIDE": "Slideshow",
        "SMHM": "Smart Home",
        "VID": "Video",
        "VIRT": "Virtual Reality",
        "BYD": "Backyard",
        "CAR": "Carousel",
        "FLXR": "Flex Room",
        "FYR": "Foyer",
        "LOFT": "Loft",
        "OBA": "Owners Bathroom",
        "OBR": "Owners Bedroom",
    }

    upper = combined.upper()

    for abbreviation, category in sorted(
        aliases.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    ):
        if re.search(
            rf"(^|[^A-Z0-9]){re.escape(abbreviation)}([^A-Z0-9]|$)",
            upper,
        ):
            if _lookup_code(tracking, "Image", category):
                return category

    return ""


def _content_value(
    placement_name: str,
    brand: str,
    community_id: str,
    tracking: dict,
) -> tuple[str, str]:
    """
    Returns (content_value, content_suffix).

    Market-wide examples:
      Pulte + Market Wide    -> Pulte Market Wide -> PULMKT
      Centex + Market Wide   -> Centex Market Wide -> CENMKT
      Del Webb + Market Wide -> Del Webb Market Wide -> DWMKT

    Community-specific example:
      Pulte + community 123456 -> PUL123456
    """
    normalized = _normalize(placement_name)
    market_wide = "marketwide" in normalized or "mktw" in normalized

    if market_wide:
        market_value = {
            "Pulte": "Pulte Market Wide",
            "Centex": "Centex Market Wide",
            "Del Webb": "Del Webb Market Wide",
            "DiVosta": "Divosta Market Wide",
            "Wieland": "Wieland Market Wide",
            "American West": "American West Market Wide",
        }.get(brand, f"{brand} Market Wide")

        if _lookup_code(tracking, "Content", market_value):
            return market_value, ""

    return brand, community_id


def parse_pulte_placement(
    placement_name: str,
    supplier_name: str = "",
) -> tuple[dict[str, str], list[str]]:
    tracking = _load_tracking_data()
    warnings: list[str] = []

    brand = _brand_from_placement(placement_name)
    community_id = _community_id(placement_name)

    division = _division_from_placement(placement_name, tracking)
    if not division:
        warnings.append("Division was not detected.")

    source = _source_from_placement(
        placement_name,
        supplier_name,
        tracking,
    )
    if not source:
        warnings.append("Source was not detected.")

    medium = _medium_from_placement(
        placement_name,
        source,
        tracking,
    )
    if not medium:
        warnings.append("Medium was not detected.")

    region, region_warning = _region_from_placement(
        placement_name,
        division,
        tracking,
    )
    if region_warning:
        warnings.append(region_warning)

    campaign = _campaign_from_placement(
        placement_name,
        tracking,
    )
    if not campaign:
        warnings.append("Campaign/purpose was not detected.")

    # Keep the older Ad Name convention by extracting the values surrounding
    # the Brand token where possible.
    parts = _split_placement(placement_name)
    brand_index = None
    brand_aliases = {
        "pulte", "delwebb", "centex", "divosta",
        "wieland", "johnwieland", "americanwest",
    }

    for index, part in enumerate(parts):
        if _normalize(part) in brand_aliases:
            brand_index = index
            break

    positional_campaign = ""
    community = ""

    if brand_index is not None:
        if brand_index + 1 < len(parts):
            positional_campaign = parts[brand_index + 1]
        if brand_index + 2 < len(parts):
            community = parts[brand_index + 2]

    return {
        "placement_name": placement_name,
        "source": source,
        "site_name": _site_name(source, supplier_name),
        "medium": medium,
        "dimension": _find_dimension(placement_name),
        "division": division,
        "brand": brand,
        "campaign": campaign,
        "ad_campaign": positional_campaign or campaign,
        "community": community,
        "community_id": community_id,
        "region": region,
    }, warnings


# ---------------------------------------------------------------------------
# CREATIVE MATCHING
# ---------------------------------------------------------------------------

def _creative_names(creative_files: Iterable) -> list[str]:
    return [
        Path(uploaded_file.name).name
        for uploaded_file in creative_files
        if getattr(uploaded_file, "name", None)
    ]


def _creative_brand_matches(creative_name: str, brand: str) -> bool:
    normalized = _normalize(creative_name)

    aliases = {
        "Centex": ("ctx", "centex"),
        "Del Webb": ("dwb", "delwebb"),
        "Pulte": ("pulte", "pul"),
        "DiVosta": ("div", "divosta"),
        "Wieland": ("jw", "wieland", "johnwieland"),
        "American West": ("aw", "americanwest"),
    }

    brand_aliases = aliases.get(brand, ())
    return any(alias in normalized for alias in brand_aliases)


def _creative_score(
    creative_name: str,
    parsed: dict[str, str],
) -> int:
    score = 0
    normalized_creative = _normalize(creative_name)

    # Dimension is mandatory whenever placement has a dimension.
    dimension = parsed["dimension"]
    if dimension:
        if _normalize(dimension) not in normalized_creative:
            return -1000
        score += 100

    # Brand is mandatory when creative filename carries recognizable brand
    # abbreviations. This prevents CTX from being selected for Del Webb.
    known_brand_marker = any(
        marker in normalized_creative
        for marker in (
            "ctx", "centex", "dwb", "delwebb",
            "divosta", "pulte", "johnwieland",
        )
    )

    if known_brand_marker:
        if not _creative_brand_matches(creative_name, parsed["brand"]):
            return -1000
        score += 80

    checks = (
        (parsed["community_id"], 40),
        (parsed["community"], 30),
        (parsed["region"], 15),
        (parsed["division"], 10),
    )

    for value, points in checks:
        normalized_value = _normalize(value)
        if normalized_value and normalized_value in normalized_creative:
            score += points

    return score


def match_creative(
    creative_names: list[str],
    parsed: dict[str, str],
) -> str:
    if not creative_names:
        return ""

    ranked = sorted(
        ((_creative_score(name, parsed), name) for name in creative_names),
        key=lambda item: (item[0], item[1]),
        reverse=True,
    )

    if not ranked or ranked[0][0] < 0:
        return ""

    # Do not silently choose between an exact tie.
    if (
        len(ranked) > 1
        and ranked[0][0] == ranked[1][0]
        and ranked[0][0] > 0
    ):
        return ""

    return ranked[0][1] if ranked[0][0] > 0 else ""


# ---------------------------------------------------------------------------
# LANDING URL MATCHING
# ---------------------------------------------------------------------------

def _parse_url_lines(landing_urls_text: str) -> list[dict[str, str]]:
    """
    Supports:
      URL
      placement<TAB>URL
      ad name<TAB>URL
      community ID<TAB>URL
      Centex<TAB>URL
      Del Webb<TAB>URL

    Never maps by row order.
    """
    parsed = []

    for raw_line in landing_urls_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        match = re.search(r"https?://\S+", line)
        if not match:
            continue

        url = match.group(0).rstrip(",;")
        label = line[:match.start()].strip(" \t:-|")
        parsed.append(
            {
                "line": line,
                "label": label,
                "url": url,
            }
        )

    return parsed


def match_landing_url(
    url_rows: list[dict[str, str]],
    parsed: dict[str, str],
    ad_name: str,
) -> str:
    if not url_rows:
        return ""

    placement_norm = _normalize(parsed["placement_name"])
    ad_norm = _normalize(ad_name)
    community_id = parsed["community_id"]
    brand_norm = _normalize(parsed["brand"])

    scored = []

    for item in url_rows:
        label_norm = _normalize(item["label"])
        line_norm = _normalize(item["line"])
        url_norm = _normalize(item["url"])

        score = 0

        if placement_norm and placement_norm in line_norm:
            score += 1000
        if ad_norm and ad_norm in line_norm:
            score += 900
        if community_id and community_id in item["line"]:
            score += 800
        if brand_norm and brand_norm in label_norm:
            score += 500

        # Domain-brand checks.
        if parsed["brand"] == "Del Webb" and "delwebbcom" in url_norm:
            score += 300
        elif parsed["brand"] in {"Pulte", "Centex"} and "pultecom" in url_norm:
            score += 200

        # Region/community hints.
        if parsed["region"] and _normalize(parsed["region"]) in line_norm:
            score += 100
        if parsed["community"] and _normalize(parsed["community"]) in line_norm:
            score += 100

        scored.append((score, item["url"]))

    scored.sort(reverse=True, key=lambda x: x[0])

    if scored and scored[0][0] > 0:
        if len(scored) > 1 and scored[0][0] == scored[1][0]:
            return ""
        return scored[0][1]

    # One supplied URL may safely apply to all.
    if len(url_rows) == 1:
        return url_rows[0]["url"]

    # Multiple URLs with no unique taxonomy match -> do not guess.
    return ""


# ---------------------------------------------------------------------------
# CMP CODE
# ---------------------------------------------------------------------------

def build_cmp_code(
    parsed: dict[str, str],
    image_category: str,
) -> tuple[str, list[str]]:
    tracking = _load_tracking_data()
    warnings: list[str] = []

    content_value, community_suffix = _content_value(
        parsed["placement_name"],
        parsed["brand"],
        parsed["community_id"],
        tracking,
    )

    components = [
        ("Medium", parsed["medium"]),
        ("Source", parsed["source"]),
        ("Division", parsed["division"]),
        ("Region", parsed["region"]),
        ("Content", content_value),
        ("Campaign", parsed["campaign"]),
        ("Vendor", "Assembly"),
        ("Image", image_category),
    ]

    codes: dict[str, str] = {}

    for section, value in components:
        code = _lookup_code(tracking, section, value)
        if not code:
            warnings.append(
                f"{section} mapping not found for '{value or 'blank'}'."
            )
        codes[section] = code

    if warnings:
        return "", warnings

    content_code = codes["Content"]

    # Content codes in the workbook (e.g. PUL/CEN/DW) do not all carry -_-.
    # Community IDs are appended immediately after the brand code.
    if community_suffix:
        content_code = f"{content_code}{community_suffix}"

    # Add separator after Content exactly once.
    if not content_code.endswith("-_-"):
        content_code = f"{content_code}-_-"

    cmp_code = "".join(
        [
            codes["Medium"],
            codes["Source"],
            codes["Division"],
            codes["Region"],
            content_code,
            codes["Campaign"],
            codes["Vendor"],
            codes["Image"],
        ]
    )

    return cmp_code, []


def _append_cmp(url: str, cmp_code: str) -> str:
    if not url or not cmp_code:
        return ""

    # Avoid double-appending CMP when a complete tracked URL is supplied.
    if re.search(r"[?&]cmp=", url, flags=re.IGNORECASE):
        return url

    separator = "&" if "?" in url else "?"
    return f"{url}{separator}cmp={cmp_code}"


# ---------------------------------------------------------------------------
# TRAFFIC DOC
# ---------------------------------------------------------------------------

def _campaign_name_from_raw_rows(raw_rows: list[list[str]]) -> str:
    for row in raw_rows:
        if row and _clean(row[0]) == "Campaign name:":
            return _clean(row[1]) if len(row) > 1 else ""
    return ""


def build_ad_name(parsed: dict[str, str]) -> str:
    parts = [
        parsed["division"],
        parsed["brand"],
        parsed["ad_campaign"],
        parsed["community"],
        parsed["dimension"],
    ]
    return "_".join(part for part in parts if part)


def _copy_row_format(sheet, source_row: int, target_row: int) -> None:
    if source_row == target_row:
        return

    for column in range(1, TRAFFIC_LAST_COLUMN + 1):
        source = sheet.cell(row=source_row, column=column)
        target = sheet.cell(row=target_row, column=column)

        target._style = copy(source._style)
        target.number_format = source.number_format
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)

    sheet.row_dimensions[target_row].height = (
        sheet.row_dimensions[source_row].height
    )


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
    creative_names = _creative_names(creative_files)
    url_rows = _parse_url_lines(landing_urls_text)
    warnings: list[str] = []

    campaign_name = _campaign_name_from_raw_rows(raw_rows)
    sheet["B1"] = campaign_name

    first_data_row = _traffic_data_start_row(sheet)
    template_row = first_data_row

    for index, record in enumerate(records):
        output_row = first_data_row + index
        _copy_row_format(sheet, template_row, output_row)

        placement_name = _record_value(record, "Placement Name")
        supplier_name = (
            _record_value(record, "Media outlet / Supplier name (ad server)")
            or _record_value(record, "Media outlet / Supplier name (Prisma)")
            or _record_value(record, "Media Outlet")
            or _record_value(record, "Supplier")
        )

        parsed, parse_warnings = parse_pulte_placement(
            placement_name,
            supplier_name,
        )

        ad_name = build_ad_name(parsed)
        creative_name = match_creative(creative_names, parsed)

        image_category = _image_category(
            placement_name,
            creative_name,
            _load_tracking_data(),
        )

        landing_url = match_landing_url(
            url_rows,
            parsed,
            ad_name,
        )

        cmp_code, cmp_warnings = build_cmp_code(
            parsed,
            image_category,
        )

        final_url = _append_cmp(landing_url, cmp_code)

        placement_id = (
            _record_value(record, "Ad server ID")
            or _record_value(record, "Placement ID")
            or _record_value(record, "DCM Placement ID")
        )

        for warning in parse_warnings:
            warnings.append(
                f"{placement_id or 'No ID'} — {warning} "
                f"Placement: {placement_name}"
            )

        for warning in cmp_warnings:
            warnings.append(
                f"{placement_id or 'No ID'} — {warning} "
                f"Placement: {placement_name}"
            )

        if not creative_name:
            warnings.append(
                f"{placement_id or 'No ID'} — No unique creative matched: "
                f"{placement_name}"
            )

        if not landing_url:
            warnings.append(
                f"{placement_id or 'No ID'} — No unique landing URL matched: "
                f"{placement_name}"
            )

        # Existing Pulte VIP behavior: 1x1 trafficking dimensions in Traffic_Doc.
        values = [
            None,                                   # A
            parsed["site_name"],                    # B Site Name
            placement_id,                           # C Placement ID
            placement_name,                         # D Placement Name
            "1x1",                                  # E Dimensions
            None,                                   # F Duration
            None,                                   # G
            ad_name,                                # H AD Name
            None,                                   # I
            "New",                                  # J Action
            creative_name,                          # K Creative File Name
            "Yes",                                  # L Studio Creative?
            None,                                   # M Rotation
            _record_value(record, "Flight start date"),  # N
            _record_value(record, "Flight end date"),    # O
            final_url,                              # P Click Through URL
            None, None, None, None, None, None, None, None,
        ]

        for column, value in enumerate(values, start=1):
            sheet.cell(row=output_row, column=column, value=value)

    return warnings


# ---------------------------------------------------------------------------
# PUBLIC ENTRY POINT
# ---------------------------------------------------------------------------

def generate_pulte_tsheet(
    prisma_file,
    creative_files,
    landing_urls_text: str,
) -> tuple[bytes, list[str]]:
    """
    Drop-in replacement for the existing Pulte VIP module.

    Key safeguards:
      - Uses official Pulte tracking workbook mappings.
      - Tennessee -> Nashville automatically through SEM Region Mapping.
      - Other divisions use exact DMA/region tokens when multiple regions exist.
      - Never maps URLs by row order.
      - Never guesses when multiple URLs/regions are ambiguous.
      - Brand + dimension aware creative matching.
      - Market Wide content codes handled correctly.
      - Existing function signature is unchanged, so Tsheet.py does not need
        to be changed.
    """
    if not MASTER_TEMPLATE.exists():
        raise FileNotFoundError(
            f"Master template not found: {MASTER_TEMPLATE.name}"
        )

    # Fail early if tracking workbook is missing/broken.
    _load_tracking_data()

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

    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue(), warnings
