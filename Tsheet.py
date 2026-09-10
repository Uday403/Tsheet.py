import csv
import itertools
import os
import io
from datetime import datetime

import streamlit as st

from aaa import (
    creative_version_key,
    generate_aaa_tsheet,
    preview_aaa_setup,
    validate_multi_rotation,
)
from anthem import (
    generate_anthem_tsheet,
    preview_anthem_setup,
)
from brooks import (
    generate_brooks_tsheet,
    preview_brooks_setup,
)
from pulte_normal import generate_normal_pulte_tsheet
from pulte_vip import generate_pulte_tsheet
from simon_vip import (
    generate_simon_vip_tsheet,
    preview_simon_vip_setup,
)


ACCOUNT_NAMES = [
    "Naming Convention Generator",
    "Pulte",
    "Pulte VIP",
    "AAA",
    "Simon VIP",
    "Anthem / Elevance",
    "Brooks",
    "UPS Store",
    "Hyatt",
    "USTA",
    "HMH",
    "ConEd",
    "Ascensus",
    "Simon",
    "Tillamook",
    "Fossil",
    "Lenovo",
    "Thrivent",
    "Vivid Seats",
    "Ace Hardware",
    "Arby's",
    "Bank OZK",
    "Best Friends",
    "Famous Footwear",
    "Tradex",
    "ASI",
    "IMC",
]


# ============================================================
# TRACKING CONFIGURATION
# ============================================================

TRACKING_FILE = "dashboard_tracking.csv"

TRACKING_FIELDS = [
    "timestamp",
    "account",
    "action",
    "output_file",
    "ads_processed",
    "direct_count",
    "multi_count",
    "unmatched_count",
    "creative_count",
    "warning_count",
    "estimated_minutes_saved",
]


def log_dashboard_usage(
    account,
    action,
    output_file="",
    ads_processed=0,
    direct_count=0,
    multi_count=0,
    unmatched_count=0,
    creative_count=0,
    warning_count=0,
    estimated_minutes_saved=0,
):
    """
    Append one successful generation event to dashboard_tracking.csv.

    IMPORTANT:
    Streamlit Community Cloud local files may be reset after app
    restarts/redeployments. This is good for testing the tracking UI.
    For permanent organization-wide tracking, later connect this
    function to SharePoint, Google Sheets, or a database.
    """
    row = {
        "timestamp": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        "account": account,
        "action": action,
        "output_file": output_file,
        "ads_processed": int(ads_processed or 0),
        "direct_count": int(direct_count or 0),
        "multi_count": int(multi_count or 0),
        "unmatched_count": int(unmatched_count or 0),
        "creative_count": int(creative_count or 0),
        "warning_count": int(warning_count or 0),
        "estimated_minutes_saved": int(
            estimated_minutes_saved or 0
        ),
    }

    file_exists = os.path.exists(TRACKING_FILE)

    with open(
        TRACKING_FILE,
        "a",
        newline="",
        encoding="utf-8",
    ) as file:
        writer = csv.DictWriter(
            file,
            fieldnames=TRACKING_FIELDS,
        )

        if not file_exists:
            writer.writeheader()

        writer.writerow(row)


def load_tracking_rows():
    if not os.path.exists(TRACKING_FILE):
        return []

    try:
        with open(
            TRACKING_FILE,
            "r",
            newline="",
            encoding="utf-8",
        ) as file:
            return list(csv.DictReader(file))

    except Exception:
        return []


def tracking_summary():
    rows = load_tracking_rows()

    successful_rows = [
        row
        for row in rows
        if row.get("action")
        in {
            "T-Sheet Generated",
            "Naming Generated",
        }
    ]

    total_ads = sum(
        int(float(row.get("ads_processed") or 0))
        for row in successful_rows
    )

    total_minutes_saved = sum(
        int(
            float(
                row.get(
                    "estimated_minutes_saved"
                )
                or 0
            )
        )
        for row in successful_rows
    )

    total_warnings = sum(
        int(float(row.get("warning_count") or 0))
        for row in successful_rows
    )

    return {
        "rows": rows,
        "generated_count": len(successful_rows),
        "total_ads": total_ads,
        "hours_saved": round(
            total_minutes_saved / 60,
            1,
        ),
        "total_warnings": total_warnings,
    }


def count_prisma_placements(prisma_file):
    """
    Count only actual Prisma placement rows.

    This follows the same rules used by pulte_vip.py:
    - locate the real 'Placement Name' header
    - ignore metadata rows above the header
    - ignore blank rows
    - ignore Package rows
    - count only rows containing a Placement Name
    """
    if prisma_file is None:
        return 0

    try:
        raw = prisma_file.getvalue()
        decoded = raw.decode("utf-8-sig", errors="replace")

        try:
            dialect = csv.Sniffer().sniff(
                decoded[:10000],
                delimiters=",;\t|",
            )
            rows = list(
                csv.reader(io.StringIO(decoded), dialect)
            )
        except csv.Error:
            rows = list(
                csv.reader(io.StringIO(decoded))
            )

        header_index = None
        headers = None

        def normalize(value):
            return "".join(
                ch.lower()
                for ch in str(value or "")
                if ch.isalnum()
            )

        for index, row in enumerate(rows):
            cleaned = [
                str(cell or "").strip().replace("\n", " ")
                for cell in row
            ]
            if any(
                normalize(cell) == "placementname"
                for cell in cleaned
            ):
                header_index = index
                headers = cleaned
                break

        if header_index is None or headers is None:
            return 0

        placement_count = 0

        for row in rows[header_index + 1:]:
            padded = row + [""] * max(
                0,
                len(headers) - len(row),
            )
            record = dict(
                zip(headers, padded[:len(headers)])
            )

            placement_name = ""
            row_type = ""

            for key, value in record.items():
                key_norm = normalize(key)

                if key_norm == "placementname":
                    placement_name = str(
                        value or ""
                    ).strip()

                if key_norm in {
                    "rowtype",
                    "type",
                    "packageplacement",
                }:
                    if not row_type:
                        row_type = str(
                            value or ""
                        ).strip().lower()

            if not placement_name:
                continue

            if (
                row_type == "package"
                or placement_name.lower().startswith(
                    "package:"
                )
            ):
                continue

            placement_count += 1

        return placement_count

    except Exception:
        return 0


# ============================================================
# STREAMLIT PAGE
# ============================================================

st.set_page_config(
    page_title="Traffic Sheet Generator",
    page_icon="📄",
    layout="wide",
)

# Header with Assembly logo
header_left, header_right = st.columns([4, 1])

with header_left:
    st.title("Traffic Sheet Generator")
    st.caption(
        "Select an account and generate the required trafficking sheet."
    )

with header_right:
    st.image(
        "assembly_Logo.png",
        width=220,
    )


# ============================================================
# TRACKING DASHBOARD
# ============================================================

summary = tracking_summary()

metric1, metric2, metric3, metric4 = st.columns(4)

with metric1:
    st.metric(
        "Sheets Generated",
        f"{summary['generated_count']:,}",
    )

with metric2:
    st.metric(
        "Ads / Placements Processed",
        f"{summary['total_ads']:,}",
    )

with metric3:
    st.metric(
        "Estimated Hours Saved",
        f"{summary['hours_saved']:,.1f}",
    )

with metric4:
    st.metric(
        "Warnings Logged",
        f"{summary['total_warnings']:,}",
    )

with st.expander(
    "Usage Tracking",
    expanded=False,
):
    tracking_rows = summary["rows"]

    if tracking_rows:
        st.dataframe(
            list(reversed(tracking_rows)),
            use_container_width=True,
            hide_index=True,
        )

        try:
            with open(
                TRACKING_FILE,
                "rb",
            ) as tracking_file:
                tracking_bytes = (
                    tracking_file.read()
                )

            st.download_button(
                "Download Tracking CSV",
                data=tracking_bytes,
                file_name=(
                    "dashboard_tracking.csv"
                ),
                mime="text/csv",
                use_container_width=True,
            )

        except Exception:
            pass

    else:
        st.info(
            "No successful generations have "
            "been tracked yet."
        )


selected_account = st.selectbox(
    "Select Account",
    ACCOUNT_NAMES,
    index=0,
)


def common_upload_fields(
    key_prefix: str,
    allow_zip: bool = False,
):
    prisma = st.file_uploader(
        "Upload Prisma CSV",
        type=["csv", "txt"],
        key=f"{key_prefix}_prisma",
    )

    creative_types = [
        "jpg",
        "jpeg",
        "png",
        "gif",
        "webp",
        "html",
        "htm",
        "mp4",
    ]

    if allow_zip:
        creative_types.append("zip")

    creatives = st.file_uploader(
        "Upload Creative Files",
        type=creative_types,
        accept_multiple_files=True,
        key=f"{key_prefix}_creatives",
    )

    return prisma, creatives


# ============================================================
# PULTE
# ============================================================

if selected_account == "Pulte":
    st.success(
        "Normal Pulte automation is ready."
    )

    st.info(
        "Paste the complete URLs/UTMs exactly as "
        "provided by the team. The dashboard will "
        "not create or modify the UTM."
    )

    prisma_file, creative_files = (
        common_upload_fields("pulte")
    )

    complete_urls_text = st.text_area(
        "Paste Complete URLs / UTMs",
        placeholder=(
            "Paste one complete URL per line"
        ),
        height=160,
        key="pulte_urls",
    )

    output_name = st.text_input(
        "Output File Name",
        value="Pulte_Tsheet.xlsm",
        key="pulte_output",
    )

    if not output_name.lower().endswith(
        ".xlsm"
    ):
        output_name += ".xlsm"

    if st.button(
        "Generate Pulte T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error(
                "Please upload the Prisma CSV."
            )

        elif not creative_files:
            st.error(
                "Please upload at least one "
                "creative file."
            )

        elif not complete_urls_text.strip():
            st.error(
                "Please paste the complete "
                "URLs/UTMs."
            )

        else:
            try:
                with st.spinner(
                    "Generating the normal "
                    "Pulte T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_normal_pulte_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            complete_urls_text=(
                                complete_urls_text
                            ),
                        )
                    )

                ads_processed = count_prisma_placements(prisma_file)

                log_dashboard_usage(
                    account="Pulte",
                    action=(
                        "T-Sheet Generated"
                    ),
                    output_file=output_name,
                    ads_processed=(
                        ads_processed
                    ),
                    creative_count=len(
                        creative_files
                    ),
                    warning_count=len(
                        warnings
                    ),
                    estimated_minutes_saved=45,
                )

                st.success(
                    "Pulte T-Sheet generated "
                    "successfully."
                )

                st.caption(
                    "Tracking recorded: "
                    f"{ads_processed:,} Ads "
                    "processed."
                )

                if warnings:
                    with st.expander(
                        "Review matching and "
                        "dimension warnings"
                    ):
                        for warning in warnings:
                            st.warning(
                                warning
                            )

                st.download_button(
                    "Download Pulte T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.ms-excel."
                        "sheet.macroEnabled.12"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)


# ============================================================
# PULTE VIP
# ============================================================

elif selected_account == "Pulte VIP":
    st.success(
        "Pulte VIP automation is ready."
    )

    prisma_file, creative_files = (
        common_upload_fields(
            "pulte_vip"
        )
    )

    landing_urls_text = st.text_area(
        "Paste Landing URLs",
        placeholder=(
            "Paste one landing URL per line"
        ),
        height=160,
        key="pulte_vip_urls",
    )

    output_name = st.text_input(
        "Output File Name",
        value="Pulte_VIP_Tsheet.xlsm",
        key="pulte_vip_output",
    )

    if not output_name.lower().endswith(
        ".xlsm"
    ):
        output_name += ".xlsm"

    if st.button(
        "Generate Pulte VIP T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error(
                "Please upload the Prisma CSV."
            )

        elif not creative_files:
            st.error(
                "Please upload at least one "
                "creative file."
            )

        elif not landing_urls_text.strip():
            st.error(
                "Please paste at least one "
                "landing URL."
            )

        else:
            try:
                with st.spinner(
                    "Generating the Pulte VIP "
                    "T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_pulte_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            landing_urls_text=(
                                landing_urls_text
                            ),
                        )
                    )

                ads_processed = count_prisma_placements(prisma_file)

                log_dashboard_usage(
                    account="Pulte VIP",
                    action=(
                        "T-Sheet Generated"
                    ),
                    output_file=output_name,
                    ads_processed=(
                        ads_processed
                    ),
                    creative_count=len(
                        creative_files
                    ),
                    warning_count=len(
                        warnings
                    ),
                    estimated_minutes_saved=45,
                )

                st.success(
                    "Pulte VIP T-Sheet generated "
                    "successfully."
                )

                st.caption(
                    "Tracking recorded: "
                    f"{ads_processed:,} Ads "
                    "processed."
                )

                if warnings:
                    with st.expander(
                        "Review warnings"
                    ):
                        for warning in warnings:
                            st.warning(
                                warning
                            )

                st.download_button(
                    "Download Pulte VIP T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.ms-excel."
                        "sheet.macroEnabled.12"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)


# ============================================================
# AAA
# ============================================================

elif selected_account == "AAA":
    st.success(
        "AAA automation is ready."
    )

    st.info(
        "AAA rules: Placement Name = Ad Name. "
        "Enter the BASE landing URL only; the "
        "dashboard creates the AAA pmed "
        "automatically."
    )

    prisma_file, creative_files = (
        common_upload_fields(
            "aaa",
            allow_zip=True,
        )
    )

    creative_setup = st.radio(
        "Creative setup type",
        [
            "Single creative per ad",
            "Multiple creatives per ad",
        ],
        key="aaa_creative_setup",
    )

    default_base_url = st.text_input(
        "Base Landing URL",
        placeholder=(
            "https://www.ace.aaa.com/"
            "travel/category/cruises.html"
        ),
        key="aaa_default_url",
    )

    override_dates = st.checkbox(
        "Override Prisma flight dates",
        value=False,
        key="aaa_override_dates",
    )

    override_start_date = None
    override_end_date = None

    if override_dates:
        col1, col2 = st.columns(2)

        with col1:
            override_start_date = (
                st.date_input(
                    "Start Date",
                    key="aaa_start_date",
                )
            )

        with col2:
            override_end_date = (
                st.date_input(
                    "End Date",
                    key="aaa_end_date",
                )
            )

    rotation_by_version = {}
    separate_url_by_version = {}
    preview = None

    if (
        prisma_file is not None
        and creative_files
    ):
        try:
            preview = preview_aaa_setup(
                prisma_file=prisma_file,
                creative_files=creative_files,
                creative_setup=creative_setup,
            )

            with st.expander(
                "AAA Creative Matching Preview",
                expanded=True,
            ):
                for placement in (
                    preview["placements"]
                ):
                    matches = (
                        placement["matches"]
                    )

                    st.write(
                        f"**{placement['dimension'] or 'No dimension'}** "
                        f"— {placement['placement_name']}"
                    )

                    if matches:
                        for creative in matches:
                            st.caption(
                                f"↳ {creative}"
                            )

                    else:
                        st.warning(
                            "No creative matched "
                            "this placement."
                        )

            for warning in (
                preview["warnings"]
            ):
                st.warning(warning)

        except Exception as exc:
            st.error(
                "Unable to preview AAA "
                f"matching: {exc}"
            )

    if (
        creative_setup
        == "Multiple creatives per ad"
        and preview is not None
    ):
        st.subheader(
            "Multi Creative Rotation"
        )

        st.caption(
            "Rotation is entered once per "
            "creative VERSION and is reused "
            "across all matching dimensions. "
            "For example V1 can be 19% for "
            "160x600, 300x250, etc."
        )

        version_groups = (
            preview["version_groups"]
        )

        for index, (
            version,
            files,
        ) in enumerate(
            version_groups.items()
        ):
            st.markdown(
                f"**{version}**"
            )

            st.caption(
                " / ".join(files)
            )

            rotation_by_version[
                version
            ] = st.number_input(
                f"Rotation % — {version}",
                min_value=0.0,
                max_value=100.0,
                value=0.0,
                step=1.0,
                key=(
                    f"aaa_rotation_{index}"
                ),
            )

            use_separate_url = (
                st.checkbox(
                    "Use a separate landing "
                    f"URL for {version}",
                    value=False,
                    key=(
                        "aaa_separate_url_"
                        f"check_{index}"
                    ),
                )
            )

            if use_separate_url:
                separate_url_by_version[
                    version
                ] = st.text_input(
                    "Separate Base URL — "
                    f"{version}",
                    placeholder=(
                        default_base_url
                    ),
                    key=(
                        "aaa_separate_url_"
                        f"{index}"
                    ),
                )

        rotation_errors = (
            validate_multi_rotation(
                preview=preview,
                rotation_by_version=(
                    rotation_by_version
                ),
            )
        )

        if rotation_errors:
            st.warning(
                "Rotation must total 100% "
                "for every Multi placement."
            )

            for error in rotation_errors:
                st.caption(error)

        else:
            st.success(
                "Rotation validation passed: "
                "each matched Multi placement "
                "totals 100%."
            )

    output_name = st.text_input(
        "Output File Name",
        value="AAA_Tsheet.xlsm",
        key="aaa_output",
    )

    if not output_name.lower().endswith(
        ".xlsm"
    ):
        output_name += ".xlsm"

    if st.button(
        "Generate AAA T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error(
                "Please upload the Prisma CSV."
            )

        elif not creative_files:
            st.error(
                "Please upload creative files."
            )

        elif not default_base_url.strip():
            st.error(
                "Please enter the base "
                "landing URL."
            )

        elif (
            creative_setup
            == "Multiple creatives per ad"
            and preview is None
        ):
            st.error(
                "AAA creative preview could "
                "not be created."
            )

        else:
            try:
                if (
                    creative_setup
                    == "Multiple creatives per ad"
                ):
                    rotation_errors = (
                        validate_multi_rotation(
                            preview=preview,
                            rotation_by_version=(
                                rotation_by_version
                            ),
                        )
                    )

                    if rotation_errors:
                        st.error(
                            "Fix the rotation "
                            "percentages before "
                            "generating the "
                            "T-Sheet."
                        )
                        st.stop()

                with st.spinner(
                    "Generating the AAA "
                    "T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_aaa_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            creative_setup=creative_setup,
                            default_base_url=default_base_url,
                            rotation_by_version=(
                                rotation_by_version
                            ),
                            separate_base_url_by_version=(
                                separate_url_by_version
                            ),
                            override_start_date=(
                                override_start_date
                            ),
                            override_end_date=(
                                override_end_date
                            ),
                        )
                    )

                ads_processed = (
                    len(preview.get("placements", []))
                    if preview is not None
                    else 0
                )

                log_dashboard_usage(
                    account="AAA",
                    action=(
                        "T-Sheet Generated"
                    ),
                    output_file=output_name,
                    ads_processed=(
                        ads_processed
                    ),
                    creative_count=len(
                        creative_files
                    ),
                    warning_count=len(
                        warnings
                    ),
                    estimated_minutes_saved=60,
                )

                st.success(
                    "AAA T-Sheet generated "
                    "successfully."
                )

                st.caption(
                    "Tracking recorded: "
                    f"{ads_processed:,} Ads "
                    "processed."
                )

                if warnings:
                    with st.expander(
                        "AAA Review Warnings"
                    ):
                        for warning in warnings:
                            st.warning(
                                warning
                            )

                st.download_button(
                    "Download AAA T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.ms-excel."
                        "sheet.macroEnabled.12"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)


# ============================================================
# ANTHEM / ELEVANCE
# ============================================================

elif selected_account == "Anthem / Elevance":
    st.success(
        "Anthem / Elevance automation is ready."
    )

    st.info(
        "Upload the Prisma CSV and Anthem "
        "creative files/ZIPs. One creative "
        "goes directly to Traffic_Doc; two "
        "or more matching creatives go to "
        "the Multi-Ad or Creative Rotation "
        "tab."
    )

    prisma_file, creative_files = (
        common_upload_fields(
            "anthem",
            allow_zip=True,
        )
    )

    url_mapping_text = st.text_area(
        "Paste Anthem URL Mapping",
        placeholder=(
            "Paste the Anthem URLs, one per line.\n"
            "The dashboard detects the required "
            "mapping from the URL."
        ),
        height=220,
        key="anthem_url_mapping",
    )

    override_dates = st.checkbox(
        "Override Prisma flight dates",
        value=False,
        key="anthem_override_dates",
    )

    override_start_date = None
    override_end_date = None

    if override_dates:
        col1, col2 = st.columns(2)

        with col1:
            override_start_date = (
                st.date_input(
                    "Start Date",
                    key="anthem_start_date",
                )
            )

        with col2:
            override_end_date = (
                st.date_input(
                    "End Date",
                    key="anthem_end_date",
                )
            )

    preview = None

    if (
        prisma_file is not None
        and creative_files
    ):
        try:
            preview = (
                preview_anthem_setup(
                    prisma_file=prisma_file,
                    creative_files=creative_files,
                    url_mapping_text=(
                        url_mapping_text
                    ),
                )
            )

            placements = (
                preview["placements"]
            )

            direct_count = sum(
                1
                for row in placements
                if row[
                    "creative_destination"
                ]
                == "Traffic_Doc"
            )

            multi_count = sum(
                1
                for row in placements
                if row[
                    "creative_destination"
                ]
                == "Multi"
            )

            unmatched_count = sum(
                1
                for row in placements
                if row[
                    "creative_destination"
                ]
                == "Unmatched"
            )

            metric1, metric2, metric3 = (
                st.columns(3)
            )

            with metric1:
                st.metric(
                    "Direct to Traffic_Doc",
                    direct_count,
                )

            with metric2:
                st.metric(
                    "Multi Creative Ads",
                    multi_count,
                )

            with metric3:
                st.metric(
                    "Unmatched Placements",
                    unmatched_count,
                )

            with st.expander(
                "Anthem Creative Matching "
                "Preview",
                expanded=True,
            ):
                for row in placements:
                    st.write(
                        f"**{row['ad_name'] or 'Ad Name not detected'}**"
                    )

                    st.caption(
                        "Placement: "
                        f"{row['placement_name']}"
                    )

                    st.caption(
                        "Language / Channel: "
                        f"{row['language'] or 'Not detected'} / "
                        f"{row['channel'] or 'Not detected'}"
                    )

                    st.caption(
                        "Destination: "
                        f"{row['creative_destination']}"
                    )

                    if row["matches"]:
                        for creative in (
                            row["matches"]
                        ):
                            st.caption(
                                f"↳ {creative}"
                            )

                    else:
                        st.warning(
                            "No creative matched "
                            "this placement."
                        )

                    if not row["url"]:
                        st.warning(
                            "No URL mapping found "
                            "for this placement."
                        )

                    st.divider()

            if preview["warnings"]:
                with st.expander(
                    "Anthem Preview Warnings"
                ):
                    for warning in (
                        preview["warnings"]
                    ):
                        st.warning(
                            warning
                        )

        except Exception as exc:
            st.error(
                "Unable to preview Anthem "
                f"matching: {exc}"
            )

    output_name = st.text_input(
        "Output File Name",
        value="Anthem_Tsheet.xlsm",
        key="anthem_output",
    )

    if not output_name.lower().endswith(
        ".xlsm"
    ):
        output_name += ".xlsm"

    if st.button(
        "Generate Anthem T-Sheet",
        type="primary",
        use_container_width=True,
        key="generate_anthem_tsheet",
    ):
        if prisma_file is None:
            st.error(
                "Please upload the Prisma CSV."
            )

        elif not creative_files:
            st.error(
                "Please upload Anthem creative "
                "files or ZIPs."
            )

        elif not url_mapping_text.strip():
            st.error(
                "Please paste the Anthem "
                "URL mapping."
            )

        else:
            try:
                with st.spinner(
                    "Generating the Anthem "
                    "T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_anthem_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            url_mapping_text=(
                                url_mapping_text
                            ),
                            override_start_date=(
                                override_start_date
                            ),
                            override_end_date=(
                                override_end_date
                            ),
                        )
                    )

                placements = (
                    preview.get("placements", [])
                    if preview is not None
                    else []
                )

                ads_processed = len(placements)

                direct_count = 0
                multi_count = 0
                unmatched_count = 0

                if preview is not None:

                    direct_count = sum(
                        1
                        for row in placements
                        if row.get(
                            "creative_destination"
                        )
                        == "Traffic_Doc"
                    )

                    multi_count = sum(
                        1
                        for row in placements
                        if row.get(
                            "creative_destination"
                        )
                        == "Multi"
                    )

                    unmatched_count = sum(
                        1
                        for row in placements
                        if row.get(
                            "creative_destination"
                        )
                        == "Unmatched"
                    )

                log_dashboard_usage(
                    account=(
                        "Anthem / Elevance"
                    ),
                    action=(
                        "T-Sheet Generated"
                    ),
                    output_file=output_name,
                    ads_processed=(
                        ads_processed
                    ),
                    direct_count=(
                        direct_count
                    ),
                    multi_count=(
                        multi_count
                    ),
                    unmatched_count=(
                        unmatched_count
                    ),
                    creative_count=len(
                        creative_files
                    ),
                    warning_count=len(
                        warnings
                    ),
                    estimated_minutes_saved=60,
                )

                st.success(
                    "Anthem T-Sheet generated "
                    "successfully."
                )

                st.caption(
                    "Tracking recorded: "
                    f"{ads_processed:,} Ads "
                    "processed."
                )

                if warnings:
                    with st.expander(
                        "Review Anthem warnings"
                    ):
                        for warning in warnings:
                            st.warning(
                                warning
                            )

                st.download_button(
                    "Download Anthem T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.ms-excel."
                        "sheet.macroEnabled.12"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)


# ============================================================
# SIMON VIP
# ============================================================

elif selected_account == "Simon VIP":
    st.success(
        "Simon VIP automation is ready."
    )

    st.info(
        "Simon VIP does not use Prisma. Paste "
        "the Placement taxonomy directly below. "
        "Placement Name = Ad Name. If no "
        "creatives are uploaded, Tracking_1x1 "
        "will be used."
    )

    placement_text = st.text_area(
        "Paste Placement Names / Taxonomy",
        placeholder=(
            "Paste one Placement Name per line\n"
            "Example: Simon_..._Arundel Mills_"
            "..._300x250"
        ),
        height=220,
        key="simon_vip_placements",
    )

    creative_files = st.file_uploader(
        "Upload Creative Files (optional)",
        type=[
            "jpg",
            "jpeg",
            "png",
            "gif",
            "webp",
            "html",
            "htm",
            "mp4",
            "zip",
        ],
        accept_multiple_files=True,
        key="simon_vip_creatives",
    )

    outlet_utm_text = st.text_area(
        "Paste Outlet Name + UTM",
        placeholder=(
            "Arundel Mills\thttps://...\n"
            "Desert Hills Premium Outlets\t"
            "https://..."
        ),
        height=220,
        key="simon_vip_utm",
    )

    outlet_date_text = st.text_area(
        "Paste Outlet Name + Start Date + End Date",
        placeholder=(
            "Arundel Mills\t08/01/2026\t"
            "08/31/2026\n"
            "Desert Hills Premium Outlets\t"
            "08/01/2026\t08/31/2026"
        ),
        height=180,
        key="simon_vip_dates",
    )

    preview = None

    if (
        placement_text.strip()
        and outlet_utm_text.strip()
        and outlet_date_text.strip()
    ):
        try:
            preview = (
                preview_simon_vip_setup(
                    placement_text=placement_text,
                    creative_files=creative_files,
                    outlet_utm_text=outlet_utm_text,
                    outlet_date_text=(
                        outlet_date_text
                    ),
                )
            )

            (
                metric_col1,
                metric_col2,
                metric_col3,
            ) = st.columns(3)

            with metric_col1:
                st.metric(
                    "Outlet mappings loaded",
                    preview[
                        "outlet_mapping_count"
                    ],
                )

            with metric_col2:
                st.metric(
                    "Placements matched",
                    preview[
                        "utm_matched_count"
                    ],
                )

            with metric_col3:
                st.metric(
                    "Unmatched placements",
                    preview[
                        "utm_unmatched_count"
                    ],
                )

            if preview[
                "using_tracking_1x1"
            ]:
                st.info(
                    "No creatives uploaded — "
                    "Tracking_1x1 will be used "
                    "for all placements."
                )

            with st.expander(
                "Simon VIP Matching Preview",
                expanded=True,
            ):
                for row in (
                    preview["rows"]
                ):
                    st.write(
                        f"**{row['placement_name']}**"
                    )

                    st.caption(
                        "Outlet: "
                        f"{row['outlet'] or 'Not matched'}"
                    )

                    st.caption(
                        "Creative: "
                        f"{row['creative'] or 'Not matched'}"
                    )

                    st.caption(
                        "Dates: "
                        f"{row['start_date'] or 'Not matched'} "
                        "to "
                        f"{row['end_date'] or 'Not matched'}"
                    )

            if preview["warnings"]:
                with st.expander(
                    "Simon VIP Preview Warnings"
                ):
                    for warning in (
                        preview["warnings"]
                    ):
                        st.warning(
                            warning
                        )

        except Exception as exc:
            st.error(
                "Unable to preview Simon VIP "
                f"matching: {exc}"
            )

    output_name = st.text_input(
        "Output File Name",
        value="Simon_VIP_Tsheet.xlsm",
        key="simon_vip_output",
    )

    if not output_name.lower().endswith(
        ".xlsm"
    ):
        output_name += ".xlsm"

    if st.button(
        "Generate Simon VIP T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if not placement_text.strip():
            st.error(
                "Please paste the Placement "
                "taxonomy."
            )

        elif not outlet_utm_text.strip():
            st.error(
                "Please paste Outlet Name "
                "and UTM mapping."
            )

        elif not outlet_date_text.strip():
            st.error(
                "Please paste Outlet Name, "
                "Start Date and End Date "
                "mapping."
            )

        else:
            try:
                with st.spinner(
                    "Generating the Simon VIP "
                    "T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_simon_vip_tsheet(
                            placement_text=(
                                placement_text
                            ),
                            creative_files=(
                                creative_files
                            ),
                            outlet_utm_text=(
                                outlet_utm_text
                            ),
                            outlet_date_text=(
                                outlet_date_text
                            ),
                        )
                    )

                ads_processed = (
                    len(preview.get("rows", []))
                    if preview is not None
                    else 0
                )

                unmatched_count = 0

                if preview is not None:
                    unmatched_count = (
                        preview.get(
                            "utm_unmatched_count",
                            0,
                        )
                    )

                log_dashboard_usage(
                    account="Simon VIP",
                    action=(
                        "T-Sheet Generated"
                    ),
                    output_file=output_name,
                    ads_processed=(
                        ads_processed
                    ),
                    unmatched_count=(
                        unmatched_count
                    ),
                    creative_count=len(
                        creative_files or []
                    ),
                    warning_count=len(
                        warnings
                    ),
                    estimated_minutes_saved=45,
                )

                st.success(
                    "Simon VIP T-Sheet generated "
                    "successfully."
                )

                st.caption(
                    "Tracking recorded: "
                    f"{ads_processed:,} Ads "
                    "processed."
                )

                if warnings:
                    with st.expander(
                        "Review Simon VIP warnings"
                    ):
                        for warning in warnings:
                            st.warning(
                                warning
                            )

                st.download_button(
                    "Download Simon VIP T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime=(
                        "application/vnd.ms-excel."
                        "sheet.macroEnabled.12"
                    ),
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)


# ============================================================
# NAMING CONVENTION GENERATOR
# ============================================================

elif selected_account == (
    "Naming Convention Generator"
):
    st.success(
        "Naming Convention Generator is ready."
    )

    st.info(
        "Copy the complete taxonomy table from "
        "Excel, including the header row, and "
        "paste it below. Any number of columns "
        "and values can be used."
    )

    taxonomy_text = st.text_area(
        "Paste Taxonomy Table from Excel",
        placeholder=(
            "Header 1\tLOB\tGeo\tCreativeSize\n"
            "ASM\tTRV\tCA\t300x250\n"
            "\tINS\tIN\t320x480\n"
            "\tBrand\tOH\t160x600\n"
            "\t\tVI\t300x600\n"
            "\t\tBA\t728x90\n"
            "\t\t\t970x250"
        ),
        height=300,
        key="naming_taxonomy_table",
    )

    separator = st.selectbox(
        "Naming Separator",
        ["_", "-", "|"],
        index=0,
        key="naming_separator",
    )

    column_values = []
    usable_columns = []
    total_combinations = 0
    table_valid = False

    if taxonomy_text.strip():
        try:
            rows = [
                row.split("\t")
                for row in (
                    taxonomy_text.splitlines()
                )
                if row.strip()
            ]

            if len(rows) < 2:
                st.warning(
                    "Paste the header row and "
                    "at least one row of values."
                )

            else:
                headers = [
                    header.strip()
                    for header in rows[0]
                ]

                for column_index in range(
                    len(headers)
                ):
                    values = []

                    for row in rows[1:]:
                        if (
                            column_index
                            < len(row)
                        ):
                            value = (
                                row[
                                    column_index
                                ].strip()
                            )

                            if (
                                value
                                and value
                                not in values
                            ):
                                values.append(
                                    value
                                )

                    column_values.append(
                        values
                    )

                usable_columns = [
                    (h, v)
                    for h, v in zip(
                        headers,
                        column_values,
                    )
                    if v
                ]

                if usable_columns:
                    table_valid = True

                    st.subheader(
                        "Detected Taxonomy"
                    )

                    preview_count = min(
                        len(usable_columns),
                        4,
                    )

                    preview_columns = (
                        st.columns(
                            preview_count
                        )
                    )

                    for index, (
                        header,
                        values,
                    ) in enumerate(
                        usable_columns
                    ):
                        with preview_columns[
                            index
                            % preview_count
                        ]:
                            st.metric(
                                header
                                or (
                                    f"Column "
                                    f"{index + 1}"
                                ),
                                len(values),
                            )

                            preview_text = (
                                " | ".join(
                                    values[:8]
                                )
                            )

                            if len(values) > 8:
                                preview_text += (
                                    " | ..."
                                )

                            st.caption(
                                preview_text
                            )

                    total_combinations = 1

                    for _, values in (
                        usable_columns
                    ):
                        total_combinations *= (
                            len(values)
                        )

                    st.info(
                        "Total naming conventions "
                        "that will be generated: "
                        f"{total_combinations:,}"
                    )

                    if (
                        total_combinations
                        > 100000
                    ):
                        st.warning(
                            "This taxonomy will "
                            "generate more than "
                            "100,000 combinations. "
                            "Consider reducing the "
                            "number of values before "
                            "generating."
                        )

        except Exception as exc:
            st.error(
                "Unable to read the pasted "
                f"taxonomy: {exc}"
            )

    if st.button(
        "Generate Naming Conventions",
        type="primary",
        use_container_width=True,
        key="generate_naming_conventions",
    ):
        if not taxonomy_text.strip():
            st.error(
                "Please paste the taxonomy "
                "table from Excel."
            )

        elif not table_valid:
            st.error(
                "No usable taxonomy values "
                "were detected."
            )

        elif total_combinations > 500000:
            st.error(
                "More than 500,000 combinations "
                "were detected. Please reduce "
                "the taxonomy before generating."
            )

        else:
            try:
                usable_values = [
                    values
                    for _, values
                    in usable_columns
                ]

                generated_names = [
                    separator.join(
                        combination
                    )
                    for combination
                    in itertools.product(
                        *usable_values
                    )
                ]

                output_text = "\n".join(
                    generated_names
                )

                log_dashboard_usage(
                    account=(
                        "Naming Convention "
                        "Generator"
                    ),
                    action=(
                        "Naming Generated"
                    ),
                    output_file=(
                        "Naming_Conventions.txt"
                    ),
                    ads_processed=len(
                        generated_names
                    ),
                    estimated_minutes_saved=30,
                )

                st.success(
                    f"{len(generated_names):,} "
                    "naming conventions "
                    "generated successfully."
                )

                st.text_area(
                    "Generated Naming "
                    "Conventions",
                    value=output_text,
                    height=400,
                    key=(
                        "naming_generated_output"
                    ),
                )

                st.download_button(
                    "Download Naming Conventions",
                    data=output_text,
                    file_name=(
                        "Naming_Conventions.txt"
                    ),
                    mime="text/plain",
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)


# ============================================================
# ACCOUNTS NOT YET AUTOMATED
# ============================================================


# ============================================================
# BROOKS
# ============================================================
elif selected_account == "Brooks":
    st.success("Brooks automation is ready.")
    st.info(
        "Automates Prisma mapping, Ad Names, creative matching, 1x1 tracking, "
        "Multi-Ad rotation and dates. Paste all complete URLs/UTMs in one box; "
        "the dashboard maps them to the matching Brooks creative concept."
    )

    prisma_file, creative_files = common_upload_fields("brooks", allow_zip=True)

    brooks_urls_text = st.text_area(
        "Paste Complete URLs / UTMs",
        placeholder=(
            "Paste all complete Brooks URLs/UTMs here, one per line.\n"
            "Example: https://www.brooksrunning.com/en_us/.../?tid=..."
        ),
        height=180,
        key="brooks_urls",
    )

    apply_dynata = st.checkbox(
        "Apply 2026 Dynata pixel note to Display placements",
        value=False,
        help="Enable only when the trafficking request requires the Dynata pixel.",
        key="brooks_dynata",
    )

    output_name = st.text_input(
        "Output File Name",
        value="Brooks_Tsheet.xlsm",
        key="brooks_output",
    )
    if not output_name.lower().endswith(".xlsm"):
        output_name += ".xlsm"

    if st.button("Preview Brooks Matching", use_container_width=True):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        else:
            try:
                preview = preview_brooks_setup(prisma_file, creative_files or [], brooks_urls_text)
                st.session_state["brooks_preview"] = preview
            except Exception as exc:
                st.exception(exc)

    preview = st.session_state.get("brooks_preview")
    if preview:
        st.write(f"**Campaign:** {preview['campaign']}")
        st.write(f"**Concept extracted:** {preview['concept']}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Placements", len(preview["placements"]))
        c2.metric("Direct", preview["direct_count"])
        c3.metric("Multi", preview["multi_count"])
        c4.metric("1x1", preview["tracking_1x1_count"])

        preview_rows = [
            {
                "Placement ID": p["placement_id"],
                "Dimension": p["size"],
                "Ad Name": p["ad_name"],
                "Status": p["status"],
                "Matched Creatives": ", ".join(p["matches"]),
            }
            for p in preview["placements"]
        ]
        st.dataframe(preview_rows, use_container_width=True, hide_index=True)
        if preview["warnings"]:
            with st.expander("Preview warnings"):
                for warning in preview["warnings"]:
                    st.warning(warning)

    if st.button(
        "Generate Brooks T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        else:
            try:
                with st.spinner("Generating Brooks T-Sheet..."):
                    output_bytes, warnings, preview = generate_brooks_tsheet(
                        prisma_file,
                        creative_files or [],
                        brooks_urls_text,
                        apply_dynata_display=apply_dynata,
                    )

                log_dashboard_usage(
                    account="Brooks",
                    action="T-Sheet Generated",
                    output_file=output_name,
                    ads_processed=len(preview["placements"]),
                    direct_count=preview["direct_count"],
                    multi_count=preview["multi_count"],
                    unmatched_count=preview["unmatched_count"],
                    creative_count=preview["creative_count"],
                    warning_count=len(warnings),
                    estimated_minutes_saved=45,
                )

                st.success("Brooks T-Sheet generated successfully.")
                st.caption(
                    f"{len(preview['placements'])} placements processed | "
                    f"{preview['tracking_1x1_count']} Tracking_1x1 | "
                    f"{preview['multi_count']} Multi-Ad placements"
                )

                if warnings:
                    with st.expander("Review Brooks warnings"):
                        for warning in warnings:
                            st.warning(warning)

                st.download_button(
                    "Download Brooks T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    use_container_width=True,
                )
            except Exception as exc:
                st.exception(exc)

else:
    st.info(
        f"{selected_account} is visible in the "
        "dashboard. Its account-specific "
        "automation will be added later."
    )
