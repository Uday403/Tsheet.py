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
) -> tuple[str, str]:
    """
    Returns:
        (tracking_division, business_division)

    Some Pulte business divisions map to a different official CMP Division.

    Official hierarchy examples:
        Tennessee -> Nashville -> NAS-_-
        St. Louis -> Illinois-St. Louis -> ILS-_-
        Southern Nevada -> Las Vegas -> LSV-_-
        Indianapolis-Kentucky -> Indianapolis-Louisville -> INK-_-
        Mid Atlantic -> Mid-Atlantic -> MAT-_-
        Central Texas -> Austin -> AUS-_-
    """
    normalized = _normalize(placement_name)

    hierarchy_aliases = {
        "indianapoliskentucky": ("Indianapolis-Louisville", "Indianapolis-Kentucky"),
        "indianapolislouisville": ("Indianapolis-Louisville", "Indianapolis-Kentucky"),
        "illinoisstlouis": ("Illinois-St. Louis", "St. Louis"),
        "centraltexas": ("Austin", "Central Texas"),
        "southernnevada": ("Las Vegas", "Southern Nevada"),
        "tennessee": ("Nashville", "Tennessee"),
        "stlouis": ("Illinois-St. Louis", "St. Louis"),
        "midatlantic": ("Mid-Atlantic", "Mid Atlantic"),
        "southwestflorida": ("Southwest Florida", "Southwest Florida"),
        "southeastflorida": ("Southeast Florida", "Southeast Florida"),
        "southerncalifornia": ("Southern California", "Southern California"),
        "northerncalifornia": ("Northern California", "Northern California"),
        "pacificnorthwest": ("Pacific Northwest", "Pacific Northwest"),
        "northeastcorridor": ("Northeast Corridor", "Northeast Corridor"),
        "northeastflorida": ("Northeast Florida", "Northeast Florida"),
        "coastalcarolinas": ("Coastal Carolinas", "Coastal Carolinas"),
        "eastcarolina": ("East Carolina", "East Carolina"),
        "westflorida": ("West Florida", "West Florida"),
        "northflorida": ("North Florida", "North Florida"),
        "newengland": ("New England", "New England"),
        "newmexico": ("New Mexico", "New Mexico"),
        "sanantonio": ("San Antonio", "San Antonio"),
        "arizona": ("Arizona", "Arizona"),
        "charlotte": ("Charlotte", "Charlotte"),
        "cleveland": ("Cleveland", "Cleveland"),
        "columbus": ("Columbus", "Columbus"),
        "dallas": ("Dallas", "Dallas"),
        "georgia": ("Georgia", "Georgia"),
        "houston": ("Houston", "Houston"),
        "michigan": ("Michigan", "Michigan"),
        "minnesota": ("Minnesota", "Minnesota"),
        "raleigh": ("Raleigh", "Raleigh"),
        "utah": ("Utah", "Utah"),
        "national": ("National", "National"),
    }

    for alias, (tracking_division, business_division) in sorted(
        hierarchy_aliases.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    ):
        if alias in normalized and _lookup_code(
            tracking,
            "Division",
            tracking_division,
        ):
            return tracking_division, business_division

    detected = _match_tracking_category(placement_name, tracking, "Division")
    if detected:
        return detected, detected

    return "", ""


def _sem_division_candidates(
    tracking_division: str,
    business_division: str,
) -> set[str]:
    candidates = {
        _normalize(tracking_division),
        _normalize(business_division),
    }

    reverse_hierarchy = {
        "nashville": "Tennessee",
        "illinoisstlouis": "St. Louis",
        "lasvegas": "Southern Nevada",
        "indianapolislouisville": "Indianapolis-Kentucky",
        "midatlantic": "Mid Atlantic",
        "austin": "Central Texas",
    }

    mapped = reverse_hierarchy.get(_normalize(tracking_division))
    if mapped:
        candidates.add(_normalize(mapped))

    return {value for value in candidates if value}

def _region_alias_matches(
    text: str,
    region_row: dict[str, str],
    business_division: str = "",
) -> bool:
    """
    Match only REGION-SPECIFIC evidence.

    Broad division labels must not be allowed to select a region.
    Example: SEM DMA='Southwest Florida' on the Sarasota row must not
    cause every Southwest Florida placement to map to SAR.
    """
    normalized_text = _normalize(text)

    category = _normalize(region_row["category"])
    sem_dma_raw = _clean(region_row["sem_dma"])
    sem_dma = _normalize(sem_dma_raw)
    abbreviation = _normalize(region_row["abbreviation"])
    division_norm = _normalize(region_row["division"])
    business_division_norm = _normalize(business_division)

    abbreviation_short = _normalize(
        region_row["abbreviation"].split("-")[0]
    )

    candidates: set[str] = set()

    if category and category != "choosevalue":
        candidates.add(category)

    if abbreviation:
        candidates.add(abbreviation)

    if abbreviation_short and len(abbreviation_short) >= 2:
        candidates.add(abbreviation_short)

    # Exclude broad DMA/division labels from acting as region evidence.
    broad_labels = {
        division_norm,
        business_division_norm,
        "southwestflorida",
        "southeastflorida",
        "southerncalifornia",
        "northerncalifornia",
        "coastalcarolinas",
        "eastcarolina",
        "newengland",
        "midatlantic",
        "centraltexas",
        "tennessee",
        "stlouis",
    }

    if sem_dma and sem_dma not in broad_labels:
        candidates.add(sem_dma)

    stop_words = {
        "fl", "az", "tx", "ca", "nc", "sc", "ga", "mo", "md", "va",
        "ma", "nh", "ri", "ct", "nm", "wa", "mn", "oh", "mi", "ky",
        "in", "ny", "pa",
        "southwest", "southeast", "southern", "northern", "north",
        "south", "east", "west", "central", "florida", "carolina",
        "california", "texas", "tennessee",
    }

    for word in re.findall(r"[A-Za-z]+", sem_dma_raw.lower()):
        if len(word) >= 4 and word not in stop_words:
            candidates.add(_normalize(word))

    return any(
        candidate and candidate in normalized_text
        for candidate in candidates
    )

def _region_from_prisma_dma_token(
    placement_name: str,
    business_division: str,
    tracking: dict,
) -> str:
    """
    Resolve compact DMA tokens used in Prisma placement names BEFORE
    broader region matching.

    Examples:
      FM FL  -> fort myers-naples -> FMNA-_-
      PHX AZ -> phoenix           -> PHX-_-
      TUC AZ -> tucson            -> TUC-_-
      WLM NC -> wilmington        -> WIL-_-

    Matching uses placement segments/tokens rather than loose substring
    matching, preventing short abbreviations from matching unrelated words.
    """
    parts = _split_placement(placement_name)

    # Normalize each underscore-delimited placement segment while preserving
    # the ability to recognize values such as "FM FL".
    segment_norms = {_normalize(part) for part in parts if _clean(part)}

    # Also collect individual word tokens for placements where DMA/state are
    # separated differently.
    word_tokens = {
        _normalize(token)
        for token in re.split(r"[_\s\-]+", placement_name)
        if _normalize(token)
    }

    # (business division, accepted placement tokens) -> official Region value
    rules = [
        ("Southwest Florida", {"fmfl", "fmna", "fortmyers", "ftmyers", "alva"}, "fort myers-naples"),
        ("Southwest Florida", {"napfl", "nap", "naples", "avemaria"}, "naples"),
        ("Southwest Florida", {"sarfl", "sar", "sarasota"}, "sarasota"),

        ("Arizona", {"phxaz", "phx", "phoenix"}, "phoenix"),
        ("Arizona", {"tucaz", "tuc", "tucson"}, "tucson"),

        ("East Carolina", {"wlmnc", "wlm", "wilmington"}, "wilmington"),
        ("East Carolina", {"myrsc", "myr", "myrtlebeach"}, "myrtle beach"),

        ("Coastal Carolinas", {"chssc", "chs", "charleston"}, "charleston"),
        ("Coastal Carolinas", {"savga", "sav", "savannah"}, "savannah"),

        ("Mid Atlantic", {"bwi", "baltimore"}, "baltimore"),
        ("Mid Atlantic", {"dca", "dc", "dcmetro", "washingtondc"}, "dc metro"),
        ("Mid Atlantic", {"ric", "richmond"}, "richmond"),

        ("Northeast Corridor", {"nyc", "newyork"}, "New York"),
        ("Northeast Corridor", {"phl", "philadelphia"}, "philadelphia"),

        ("Southeast Florida", {"mia", "miami", "fortlauderdale", "ftlauderdale"}, "Miami-Ft. Lauderdale"),
        ("Southeast Florida", {"pbi", "palmbeach", "westpalmbeach"}, "palm beach"),

        ("Northern California", {"sfo", "sjc", "oak", "sanfrancisco", "sanjose", "oakland"}, "bay area"),
        ("Northern California", {"sac", "sacramento"}, "sacramento"),
        ("Northern California", {"fat", "fresno"}, "central-valley"),

        ("Southern California", {"lax", "losangeles"}, "los angeles"),
        ("Southern California", {"san", "sandiego"}, "southern california"),
    ]

    division_norm = _normalize(business_division)

    for rule_division, aliases, region_name in rules:
        if _normalize(rule_division) != division_norm:
            continue

        matched = False
        for alias in aliases:
            alias_norm = _normalize(alias)

            # Exact placement segment is strongest, e.g. "FM FL".
            if alias_norm in segment_norms:
                matched = True
                break

            # Allow exact word token only for 3+ character abbreviations or
            # full locality words. Two-letter codes like FM are intentionally
            # not used alone.
            if len(alias_norm) >= 3 and alias_norm in word_tokens:
                matched = True
                break

            # Full locality names can appear inside a longer segment.
            if len(alias_norm) >= 5 and alias_norm in _normalize(placement_name):
                matched = True
                break

        if matched and _lookup_code(tracking, "Region", region_name):
            return region_name

    return ""


def _region_from_placement(
    placement_name: str,
    tracking_division: str,
    business_division: str,
    tracking: dict,
) -> tuple[str, str]:
    """
    Resolve Region safely.

    Examples:
      Tennessee -> Nashville -> NASH-_-
      Arizona + Phoenix/PHX -> PHX-_-
      Arizona + Tucson/TUC -> TUC-_-
      Southwest Florida + Fort Myers/Alva/Naples -> FMNA-_-
      Southwest Florida + Sarasota -> SAR-_-

    If a division has multiple valid regions and the placement has no
    region-specific clue, do not guess.
    """
    rows = tracking["region_rows"]
    placement_norm = _normalize(placement_name)

    # Highest priority: explicit compact DMA token from Prisma taxonomy.
    # This fixes placements such as ..._209822_FM FL_PT where the business
    # division is Southwest Florida but the correct Region is Fort Myers-Naples.
    dma_region = _region_from_prisma_dma_token(
        placement_name,
        business_division,
        tracking,
    )
    if dma_region:
        return dma_region, ""

    sem_divisions = _sem_division_candidates(
        tracking_division,
        business_division,
    )

    division_rows = [
        row for row in rows
        if _normalize(row["division"]) in sem_divisions
    ]

    # 1. Specific region evidence inside the allowed division.
    scoped_matches = [
        row for row in division_rows
        if _region_alias_matches(
            placement_name,
            row,
            business_division=business_division,
        )
    ]

    if scoped_matches:
        winner = max(
            scoped_matches,
            key=lambda row: max(
                len(_normalize(row["category"])),
                len(_normalize(row["sem_dma"])),
                len(_normalize(row["abbreviation"])),
            ),
        )
        return winner["category"], ""

    # 2. Common locality aliases from placement/community naming.
    locality_aliases = {
        # Southwest Florida
        ("Southwest Florida", "fortmyers"): "fort myers-naples",
        ("Southwest Florida", "ftmyers"): "fort myers-naples",
        ("Southwest Florida", "alva"): "fort myers-naples",
        ("Southwest Florida", "naples"): "fort myers-naples",
        ("Southwest Florida", "sarasota"): "sarasota",

        # Arizona
        ("Arizona", "phoenix"): "phoenix",
        ("Arizona", "phx"): "phoenix",
        ("Arizona", "tucson"): "tucson",
        ("Arizona", "tuc"): "tucson",

        # Coastal Carolinas
        ("Coastal Carolinas", "charleston"): "charleston",
        ("Coastal Carolinas", "savannah"): "savannah",

        # East Carolina
        ("East Carolina", "myrtlebeach"): "myrtle beach",
        ("East Carolina", "wilmington"): "wilmington",
        ("East Carolina", "wlm"): "wilmington",

        # Mid Atlantic
        ("Mid Atlantic", "baltimore"): "baltimore",
        ("Mid Atlantic", "washingtondc"): "dc metro",
        ("Mid Atlantic", "dcmetro"): "dc metro",
        ("Mid Atlantic", "marylandbeaches"): "maryland beaches",
        ("Mid Atlantic", "norfolk"): "northfolk-portsmout-newsport",
        ("Mid Atlantic", "richmond"): "richmond",

        # New England
        ("New England", "boston"): "greater boston area",
        ("New England", "providence"): "greater boston area",
        ("New England", "hartford"): "hartford-new haven",
        ("New England", "newhaven"): "hartford-new haven",

        # Northeast Corridor
        ("Northeast Corridor", "newyork"): "New York",
        ("Northeast Corridor", "philadelphia"): "philadelphia",

        # Northern California
        ("Northern California", "sanfrancisco"): "bay area",
        ("Northern California", "oakland"): "bay area",
        ("Northern California", "sanjose"): "bay area",
        ("Northern California", "sacramento"): "sacramento",
        ("Northern California", "fresno"): "central-valley",

        # Southern California
        ("Southern California", "losangeles"): "los angeles",
        ("Southern California", "sandiego"): "southern california",

        # Southeast Florida
        ("Southeast Florida", "miami"): "Miami-Ft. Lauderdale",
        ("Southeast Florida", "fortlauderdale"): "Miami-Ft. Lauderdale",
        ("Southeast Florida", "palmbeach"): "palm beach",
        ("Southeast Florida", "westpalmbeach"): "palm beach",
    }

    for (division_name, alias), region_name in sorted(
        locality_aliases.items(),
        key=lambda item: len(item[0][1]),
        reverse=True,
    ):
        if (
            _normalize(division_name) in sem_divisions
            and alias in placement_norm
        ):
            if _lookup_code(tracking, "Region", region_name):
                return region_name, ""

    # 3. Exactly one official region for the business division -> safe default.
    unique_categories: list[str] = []
    seen = set()

    for row in division_rows:
        category = _clean(row["category"])
        norm = _normalize(category)

        if not category or norm == "choosevalue":
            continue

        if norm not in seen:
            seen.add(norm)
            unique_categories.append(category)

    if len(unique_categories) == 1:
        return unique_categories[0], ""

    # 4. Known one-market hierarchy defaults.
    explicit_defaults = {
        "Tennessee": "nashville",
        "St. Louis": "st.louis",
        "Southern Nevada": "las-vegas",
        "Raleigh": "raleigh",
        "San Antonio": "san antonio",
        "West Florida": "tampa",
        "Charlotte": "charlotte",
        "Dallas": "dallas",
        "Georgia": "atlanta",
        "Houston": "houston",
        "Michigan": "detroit",
        "Minnesota": "the twin cities",
        "New Mexico": "albuquerque",
        "Pacific Northwest": "seattle",
    }

    preferred = explicit_defaults.get(business_division)
    if preferred and _lookup_code(tracking, "Region", preferred):
        return preferred, ""

    if not tracking_division:
        return "", "Division could not be detected, so Region could not be resolved."

    return (
        "",
        f"Region could not be resolved safely for division "
        f"'{business_division or tracking_division}'. "
        "This division has multiple possible regions and the placement "
        "did not contain a specific DMA/city/region signal."
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

    division, business_division = _division_from_placement(
        placement_name,
        tracking,
    )
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
        business_division,
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
        "business_division": business_division,
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
    """
    Match Pulte brand codes in creative filenames safely.

    Supported examples:
      SWFL_DW_...   -> Del Webb
      SWFL_DWB_...  -> Del Webb
      SWFL_PU_...   -> Pulte
      SWFL_PUL_...  -> Pulte
      ..._CTX_...   -> Centex
      ..._CEN_...   -> Centex

    Tokens are compared as filename tokens so short codes such as DW/PU
    do not accidentally match unrelated text.
    """
    stem = Path(creative_name).stem
    tokens = {
        _normalize(token)
        for token in re.split(r"[_\-\s]+", stem)
        if _normalize(token)
    }
    normalized = _normalize(stem)

    token_aliases = {
        "Centex": {"ctx", "cen", "centex"},
        "Del Webb": {"dw", "dwb", "delwebb"},
        "Pulte": {"pu", "pul", "pulte"},
        "DiVosta": {"div", "divosta"},
        "Wieland": {"jw", "wieland", "johnwieland"},
        "American West": {"aw", "americanwest"},
    }

    aliases = token_aliases.get(brand, set())

    if tokens.intersection(aliases):
        return True

    # Full brand names can appear without separators.
    full_aliases = {
        "Centex": ("centex",),
        "Del Webb": ("delwebb",),
        "Pulte": ("pulte",),
        "DiVosta": ("divosta",),
        "Wieland": ("johnwieland", "wieland"),
        "American West": ("americanwest",),
    }

    return any(alias in normalized for alias in full_aliases.get(brand, ()))


def _creative_has_known_brand_marker(creative_name: str) -> bool:
    stem = Path(creative_name).stem
    tokens = {
        _normalize(token)
        for token in re.split(r"[_\-\s]+", stem)
        if _normalize(token)
    }
    normalized = _normalize(stem)

    short_markers = {
        "ctx", "cen", "dw", "dwb", "pu", "pul",
        "div", "jw", "aw",
    }
    full_markers = {
        "centex", "delwebb", "pulte", "divosta",
        "johnwieland", "wieland", "americanwest",
    }

    return bool(tokens.intersection(short_markers)) or any(
        marker in normalized for marker in full_markers
    )

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
    known_brand_marker = _creative_has_known_brand_marker(creative_name)

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


def _region_from_landing_url(
    landing_url: str,
    business_division: str,
    tracking: dict,
) -> str:
    """
    Use the matched landing-page URL as the strongest Region signal.

    Southwest Florida examples:
      /naples/ or /ave-maria/ -> naples -> NAP-_-
      /fort-myers/ or /alva/  -> fort myers-naples -> FMNA-_-
      /sarasota/              -> sarasota -> SAR-_-
    """
    url = _clean(landing_url)
    if not url:
        return ""

    # Normalize URL path/text so hyphenated locations such as ave-maria and
    # fort-myers are recognized safely.
    url_norm = _normalize(re.sub(r"[%+_\-]+", " ", url.lower()))
    division_norm = _normalize(business_division)

    rules = {
        "southwestflorida": [
            (("naples", "avemaria"), "naples"),
            (("fortmyers", "ftmyers", "alva"), "fort myers-naples"),
            (("sarasota",), "sarasota"),
        ],
        "arizona": [
            (("phoenix",), "phoenix"),
            (("tucson",), "tucson"),
        ],
    }

    for aliases, region_name in rules.get(division_norm, []):
        if any(_normalize(alias) in url_norm for alias in aliases):
            # Only accept a Region that actually exists in the official
            # tracking workbook.
            if _lookup_code(tracking, "Region", region_name):
                return region_name

    return ""


# ---------------------------------------------------------------------------
# CMP CODE
# ---------------------------------------------------------------------------

def build_cmp_code(
    parsed: dict[str, str],
    image_category: str,
    landing_url: str = "",
) -> tuple[str, list[str]]:
    tracking = _load_tracking_data()
    warnings: list[str] = []

    # URL is the strongest Region signal when it clearly identifies a market.
    url_region = _region_from_landing_url(
        landing_url,
        parsed.get("business_division", ""),
        tracking,
    )
    if url_region:
        parsed = dict(parsed)
        parsed["region"] = url_region

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
        # 1x1/tracking placements may not carry an image taxonomy.
        # In that case use the official workbook's Choose Value = NA-_-.
        ("Image", image_category or "Choose Value"),
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
            landing_url,
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
