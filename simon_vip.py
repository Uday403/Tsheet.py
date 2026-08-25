
from __future__ import annotations

import io
import re
import zipfile
from copy import copy
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
MASTER_TEMPLATE = BASE_DIR / "master_template.xlsm"

TRAFFIC_SHEET = "Traffic_Doc"
MULTI_SHEET = "Multi-Ad or Creative Rotation"

TRACKING_1X1 = "Tracking_1x1"


def _clean(value) -> str:
    return "" if value is None else str(value).strip()


def _normalize(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", _clean(value).lower())


def _words(value) -> list[str]:
    return re.findall(r"[A-Za-z0-9]+", _clean(value).lower())


# ============================================================
# PLACEMENT TAXONOMY INPUT
# ============================================================

def parse_placement_taxonomy(placement_text: str) -> list[str]:
    """
    Accepts one placement taxonomy per line.
    Also works when users paste one Excel column directly.
    """
    placements = []

    for raw_line in placement_text.splitlines():
        value = raw_line.strip()

        if not value:
            continue

        # Ignore a pasted header.
        if _normalize(value) in {
            "placement",
            "placementname",
            "placementtaxonomy",
        }:
            continue

        placements.append(value)

    if not placements:
        raise ValueError(
            "No placement taxonomy was detected. "
            "Paste one Placement Name per line."
        )

    return placements


# ============================================================
# OUTLET + UTM MAPPING
# ============================================================

def parse_outlet_utm_mapping(
    outlet_utm_text: str,
) -> list[dict[str, str]]:
    """
    Paste two Excel columns:
        Outlet Name<TAB>UTM
    """
    mappings = []

    for raw_line in outlet_utm_text.splitlines():
        line = raw_line.strip()

        if not line:
            continue

        cells = []

        if "\t" in line:
            cells = line.split("\t", 1)
        elif "|" in line:
            cells = line.split("|", 1)
        else:
            match = re.search(r"https?://\S+", line)

            if match:
                cells = [
                    line[:match.start()].strip(),
                    match.group(0).strip(),
                ]

        if len(cells) < 2:
            continue

        outlet = _clean(cells[0])
        url = _clean(cells[1])

        if not outlet or not url:
            continue

        if (
            _normalize(outlet)
            in {
                "outlet",
                "outletname",
                "property",
                "propertyname",
                "mall",
                "mallname",
            }
            and "url" in _normalize(url)
        ):
            continue

        mappings.append(
            {
                "outlet": outlet,
                "outlet_normalized": _normalize(outlet),
                "url": url,
            }
        )

    if not mappings:
        raise ValueError(
            "No Outlet / UTM mapping was detected. "
            "Paste Outlet Name and UTM side by side from Excel."
        )

    # Prefer the longest matching outlet name.
    mappings.sort(
        key=lambda item: len(item["outlet_normalized"]),
        reverse=True,
    )

    return mappings


def parse_outlet_date_mapping(
    outlet_date_text: str,
) -> list[dict[str, str]]:
    """
    Paste three Excel columns:
        Outlet Name<TAB>Start Date<TAB>End Date

    Example:
        Arundel Mills    08/01/2026    08/31/2026
    """
    mappings = []

    for raw_line in outlet_date_text.splitlines():
        line = raw_line.strip()

        if not line:
            continue

        if "\t" in line:
            cells = line.split("\t")
        elif "|" in line:
            cells = line.split("|")
        else:
            # Space/comma separated text is intentionally not guessed because
            # outlet names themselves contain spaces.
            continue

        if len(cells) < 3:
            continue

        outlet = _clean(cells[0])
        start_date = _clean(cells[1])
        end_date = _clean(cells[2])

        if not outlet:
            continue

        if _normalize(outlet) in {
            "outlet",
            "outletname",
            "property",
            "propertyname",
            "mall",
            "mallname",
        }:
            continue

        mappings.append(
            {
                "outlet": outlet,
                "outlet_normalized": _normalize(outlet),
                "start_date": start_date,
                "end_date": end_date,
            }
        )

    if not mappings:
        raise ValueError(
            "No Outlet / Start Date / End Date mapping was detected. "
            "Paste three columns from Excel."
        )

    mappings.sort(
        key=lambda item: len(item["outlet_normalized"]),
        reverse=True,
    )

    return mappings


def match_outlet_dates(
    placement_name: str,
    mappings: list[dict[str, str]],
) -> tuple[str, str, str]:
    normalized_placement = _normalize(placement_name)

    matches = [
        mapping
        for mapping in mappings
        if (
            mapping["outlet_normalized"]
            and mapping["outlet_normalized"] in normalized_placement
        )
    ]

    if not matches:
        return "", "", ""

    best = matches[0]

    return (
        best["outlet"],
        best["start_date"],
        best["end_date"],
    )


def match_outlet_utm(
    placement_name: str,
    mappings: list[dict[str, str]],
) -> tuple[str, str]:
    normalized_placement = _normalize(placement_name)

    matches = [
        mapping
        for mapping in mappings
        if (
            mapping["outlet_normalized"]
            and mapping["outlet_normalized"] in normalized_placement
        )
    ]

    if not matches:
        return "", ""

    best = matches[0]
    return best["outlet"], best["url"]


# ============================================================
# DIMENSION + CREATIVE MATCHING
# ============================================================

def _extract_dimension(value: str) -> str:
    match = re.search(
        r"(?<!\d)(\d{1,4})\s*[xX×*]\s*(\d{1,4})(?!\d)",
        _clean(value),
    )

    if not match:
        return ""

    return f"{match.group(1)}x{match.group(2)}"


def _creative_file_names(
    creative_files: Iterable | None,
) -> list[str]:
    names = []

    for uploaded_file in creative_files or []:
        file_name = Path(uploaded_file.name).name

        if file_name.lower().endswith(".zip"):
            uploaded_file.seek(0)

            with zipfile.ZipFile(uploaded_file, "r") as archive:
                for member in archive.namelist():
                    if member.endswith("/"):
                        continue

                    name = Path(member).name

                    if name:
                        names.append(name)
        else:
            names.append(file_name)

    return list(dict.fromkeys(name for name in names if name))


def _creative_content_tokens(
    creative_name: str,
) -> set[str]:
    stem = Path(creative_name).stem

    stem = re.sub(
        r"(?<!\d)\d{1,4}\s*[xX×*]\s*\d{1,4}(?!\d)",
        " ",
        stem,
    )

    generic = {
        "simon",
        "premium",
        "outlet",
        "outlets",
        "creative",
        "display",
        "banner",
        "static",
        "image",
        "jpg",
        "jpeg",
        "png",
        "gif",
        "webp",
        "ad",
        "ads",
        "version",
        "ver",
        "new",
    }

    return {
        word
        for word in _words(stem)
        if len(word) >= 3 and word not in generic
    }


def _creative_score(
    creative_name: str,
    placement_name: str,
    dimension: str,
) -> float:
    creative_dimension = _extract_dimension(creative_name)

    # For normal display sizes, dimension must match.
    if dimension and dimension.lower() != "1x1":
        if not creative_dimension:
            return -1000

        if creative_dimension.lower() != dimension.lower():
            return -1000

    score = 0.0

    if (
        dimension
        and creative_dimension
        and creative_dimension.lower() == dimension.lower()
    ):
        score += 50

    placement_normalized = _normalize(placement_name)

    for token in _creative_content_tokens(creative_name):
        if _normalize(token) in placement_normalized:
            score += 10

    stem = Path(creative_name).stem

    for chunk in re.split(r"[_\-\s]+", stem):
        chunk_normalized = _normalize(chunk)

        if (
            len(chunk_normalized) >= 6
            and chunk_normalized in placement_normalized
        ):
            score += 15

    return score


def match_creative(
    creative_names: list[str],
    placement_name: str,
    dimension: str,
) -> tuple[str, list[str]]:
    """
    Rules:
    - If no creatives were uploaded -> Tracking_1x1
    - If creatives exist -> match by dimension first, then content/name
    """
    if not creative_names:
        return TRACKING_1X1, []

    ranked = sorted(
        (
            (
                _creative_score(
                    creative_name=name,
                    placement_name=placement_name,
                    dimension=dimension,
                ),
                name,
            )
            for name in creative_names
        ),
        reverse=True,
    )

    valid = [item for item in ranked if item[0] > 0]

    if not valid:
        if dimension.lower() == "1x1":
            return TRACKING_1X1, []

        return "", []

    best_score = valid[0][0]
    tied = [
        name
        for score, name in valid
        if score == best_score
    ]

    if len(tied) > 1:
        return "", tied

    return valid[0][1], []


# ============================================================
# EXCEL HELPERS
# ============================================================

def _find_traffic_layout(sheet) -> tuple[int, int]:
    for row in range(1, min(sheet.max_row, 30) + 1):
        placement_header = _clean(
            sheet.cell(row=row, column=4).value
        ).lower()

        ad_header = _clean(
            sheet.cell(row=row, column=8).value
        ).lower()

        if (
            "placement name" in placement_header
            and "ad name" in ad_header
        ):
            return row, row + 1

    raise ValueError(
        "Traffic_Doc headers could not be located."
    )


def _snapshot_row_format(
    sheet,
    row_number: int,
    max_column: int,
) -> dict:
    snapshot = {}

    for column in range(1, max_column + 1):
        cell = sheet.cell(row=row_number, column=column)

        snapshot[column] = {
            "style": copy(cell._style),
            "number_format": cell.number_format,
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "protection": copy(cell.protection),
        }

    return {
        "cells": snapshot,
        "height": sheet.row_dimensions[row_number].height,
    }


def _apply_row_format(
    sheet,
    row_number: int,
    snapshot: dict,
) -> None:
    for column, style in snapshot["cells"].items():
        cell = sheet.cell(row=row_number, column=column)

        cell._style = copy(style["style"])
        cell.number_format = style["number_format"]
        cell.font = copy(style["font"])
        cell.fill = copy(style["fill"])
        cell.border = copy(style["border"])
        cell.alignment = copy(style["alignment"])
        cell.protection = copy(style["protection"])

    sheet.row_dimensions[row_number].height = snapshot["height"]


def _clear_values(
    sheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> None:
    for row in sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    ):
        for cell in row:
            cell.value = None


def _clear_multi_sheet(workbook) -> None:
    if MULTI_SHEET not in workbook.sheetnames:
        return

    sheet = workbook[MULTI_SHEET]

    for merged_range in list(sheet.merged_cells.ranges):
        if merged_range.min_row >= 2:
            sheet.unmerge_cells(str(merged_range))

    _clear_values(
        sheet,
        min_row=2,
        max_row=max(sheet.max_row, 100),
        min_col=1,
        max_col=max(sheet.max_column, 16),
    )


# ============================================================
# PREVIEW
# ============================================================

def preview_simon_vip_setup(
    placement_text: str,
    creative_files,
    outlet_utm_text: str,
    outlet_date_text: str,
) -> dict:
    placements = parse_placement_taxonomy(placement_text)
    mappings = parse_outlet_utm_mapping(outlet_utm_text)
    date_mappings = parse_outlet_date_mapping(outlet_date_text)
    creative_names = _creative_file_names(creative_files)

    rows = []
    warnings = []

    for placement_name in placements:
        dimension = _extract_dimension(placement_name)

        outlet, url = match_outlet_utm(
            placement_name,
            mappings,
        )

        date_outlet, start_date, end_date = match_outlet_dates(
            placement_name,
            date_mappings,
        )

        creative, tied = match_creative(
            creative_names=creative_names,
            placement_name=placement_name,
            dimension=dimension,
        )

        if not outlet:
            warnings.append(
                f"No outlet/UTM match: {placement_name}"
            )

        if not date_outlet:
            warnings.append(
                f"No outlet/date match: {placement_name}"
            )

        if tied:
            warnings.append(
                f"Ambiguous creative match: "
                f"{placement_name} -> "
                + ", ".join(tied)
            )

        if not creative:
            warnings.append(
                f"No creative matched: {placement_name}"
            )

        rows.append(
            {
                "placement_name": placement_name,
                "ad_name": placement_name,
                "dimension": dimension,
                "outlet": outlet,
                "url": url,
                "start_date": start_date,
                "end_date": end_date,
                "creative": creative,
                "creative_candidates": tied,
            }
        )

    matched_count = sum(1 for row in rows if row["url"])

    return {
        "rows": rows,
        "outlet_mapping_count": len(mappings),
        "placement_count": len(placements),
        "utm_matched_count": matched_count,
        "utm_unmatched_count": len(placements) - matched_count,
        "creative_count": len(creative_names),
        "using_tracking_1x1": len(creative_names) == 0,
        "warnings": warnings,
    }


# ============================================================
# GENERATION
# ============================================================

def generate_simon_vip_tsheet(
    placement_text: str,
    creative_files,
    outlet_utm_text: str,
    outlet_date_text: str,
) -> tuple[bytes, list[str]]:
    """
    Simon VIP:
    - NO Prisma input
    - Placement taxonomy pasted directly in dashboard
    - Placement Name = Ad Name
    - Outlet name inside Placement Name determines provided UTM
    - Outlet name also determines Start Date and End Date
    - If no creatives -> Tracking_1x1
    - If creatives -> dimension + content/name matching
    """
    if not MASTER_TEMPLATE.exists():
        raise FileNotFoundError(
            f"Master template not found: {MASTER_TEMPLATE.name}"
        )

    placements = parse_placement_taxonomy(placement_text)
    mappings = parse_outlet_utm_mapping(outlet_utm_text)
    date_mappings = parse_outlet_date_mapping(outlet_date_text)
    creative_names = _creative_file_names(creative_files)

    workbook = load_workbook(
        MASTER_TEMPLATE,
        keep_vba=True,
    )

    if TRAFFIC_SHEET not in workbook.sheetnames:
        raise KeyError(
            f"Missing worksheet in master template: {TRAFFIC_SHEET}"
        )

    sheet = workbook[TRAFFIC_SHEET]
    _, first_data_row = _find_traffic_layout(sheet)

    max_column = max(sheet.max_column, 24)

    style_snapshot = _snapshot_row_format(
        sheet,
        first_data_row,
        max_column,
    )

    # Remove every old Traffic_Doc row value.
    _clear_values(
        sheet,
        min_row=first_data_row,
        max_row=max(
            sheet.max_row,
            first_data_row + len(placements) + 10,
        ),
        min_col=1,
        max_col=max_column,
    )

    _clear_multi_sheet(workbook)

    warnings = []

    for index, placement_name in enumerate(placements):
        row = first_data_row + index

        _apply_row_format(
            sheet,
            row,
            style_snapshot,
        )

        dimension = _extract_dimension(placement_name)

        outlet_name, matched_url = match_outlet_utm(
            placement_name,
            mappings,
        )

        date_outlet, start_date, end_date = match_outlet_dates(
            placement_name,
            date_mappings,
        )

        creative_name, tied = match_creative(
            creative_names=creative_names,
            placement_name=placement_name,
            dimension=dimension,
        )

        # Placement Name = Ad Name
        sheet.cell(row=row, column=4).value = placement_name
        sheet.cell(row=row, column=5).value = dimension
        sheet.cell(row=row, column=8).value = placement_name
        sheet.cell(row=row, column=10).value = "New"
        sheet.cell(row=row, column=11).value = creative_name
        sheet.cell(row=row, column=13).value = "100%"
        sheet.cell(row=row, column=14).value = start_date
        sheet.cell(row=row, column=15).value = end_date
        sheet.cell(row=row, column=16).value = matched_url

        if not matched_url:
            warnings.append(
                f"No UTM matched outlet in placement: {placement_name}"
            )

        if not date_outlet:
            warnings.append(
                f"No Start/End Date matched outlet in placement: {placement_name}"
            )

        if tied:
            warnings.append(
                f"Ambiguous creative match — manual review required: "
                f"{placement_name} -> "
                + ", ".join(tied)
            )

        if not creative_name:
            warnings.append(
                f"No creative matched: {placement_name}"
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
