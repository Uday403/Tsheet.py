import csv
import itertools
import os
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
from bfas import (
    generate_bfas_tsheet,
    preview_bfas_setup,
)
from coned import (
    generate_coned_tsheet,
    preview_coned_setup,
)

from perdue import (
    generate_perdue_tsheet,
    preview_perdue_setup,
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
    "Perdue",
    "BFAS",
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

                ads_processed = 0

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

                ads_processed = 0

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
    st.success("Anthem / Elevance automation is ready.")
    st.info(
        "Upload the Prisma CSV, then upload English and Spanish creatives separately. "
        "Paste English and Spanish tagged URLs/UTMs in their respective sections. "
        "One matching creative goes to Traffic_Doc; two or more matching creatives "
        "go to the Multi-Ad or Creative Rotation tab."
    )

    prisma_file = st.file_uploader(
        "Upload Prisma CSV", type=["csv", "txt"], key="anthem_prisma"
    )

    col_en, col_sp = st.columns(2)
    with col_en:
        st.subheader("English")
        english_creative_files = st.file_uploader(
            "Upload English Creatives",
            type=["jpg", "jpeg", "png", "gif", "webp", "html", "htm", "mp4", "zip"],
            accept_multiple_files=True, key="anthem_en_creatives"
        )
        english_url_mapping_text = st.text_area(
            "Paste English UTMs / URLs",
            placeholder="Paste English tagged URLs, one per line.",
            height=220, key="anthem_en_urls"
        )

    with col_sp:
        st.subheader("Spanish")
        spanish_creative_files = st.file_uploader(
            "Upload Spanish Creatives",
            type=["jpg", "jpeg", "png", "gif", "webp", "html", "htm", "mp4", "zip"],
            accept_multiple_files=True, key="anthem_sp_creatives"
        )
        spanish_url_mapping_text = st.text_area(
            "Paste Spanish UTMs / URLs",
            placeholder="Paste Spanish tagged URLs, one per line.",
            height=220, key="anthem_sp_urls"
        )

    override_dates = st.checkbox("Override Prisma flight dates", value=False, key="anthem_override_dates")
    override_start_date = override_end_date = None
    if override_dates:
        c1, c2 = st.columns(2)
        with c1:
            override_start_date = st.date_input("Start Date", key="anthem_start_date")
        with c2:
            override_end_date = st.date_input("End Date", key="anthem_end_date")

    preview = None
    all_creatives = list(english_creative_files or []) + list(spanish_creative_files or [])
    if prisma_file is not None and all_creatives:
        try:
            preview = preview_anthem_setup(
                prisma_file=prisma_file,
                english_creative_files=english_creative_files,
                spanish_creative_files=spanish_creative_files,
                english_url_mapping_text=english_url_mapping_text,
                spanish_url_mapping_text=spanish_url_mapping_text,
            )
            placements = preview.get("placements", [])
            direct_count = sum(1 for r in placements if r.get("creative_destination") == "Traffic_Doc")
            multi_count = sum(1 for r in placements if r.get("creative_destination") == "Multi")
            unmatched_count = sum(1 for r in placements if r.get("creative_destination") == "Unmatched")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Placements", len(placements))
            m2.metric("Direct", direct_count)
            m3.metric("Multi", multi_count)
            m4.metric("Unmatched", unmatched_count)
            with st.expander("Anthem Matching Preview", expanded=False):
                st.dataframe(placements, use_container_width=True)
            for warning in preview.get("warnings", []):
                st.warning(warning)
        except Exception as exc:
            st.exception(exc)

    if st.button("Generate Anthem T-Sheet", type="primary", use_container_width=True, key="generate_anthem"):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        elif not all_creatives:
            st.error("Please upload at least one English or Spanish creative.")
        else:
            try:
                output_bytes, warnings = generate_anthem_tsheet(
                    prisma_file=prisma_file,
                    english_creative_files=english_creative_files,
                    spanish_creative_files=spanish_creative_files,
                    english_url_mapping_text=english_url_mapping_text,
                    spanish_url_mapping_text=spanish_url_mapping_text,
                    override_start_date=override_start_date,
                    override_end_date=override_end_date,
                )
                output_name = "Anthem_Tsheet.xlsm"
                placements = (preview or {}).get("placements", [])
                direct_count = sum(1 for r in placements if r.get("creative_destination") == "Traffic_Doc")
                multi_count = sum(1 for r in placements if r.get("creative_destination") == "Multi")
                unmatched_count = sum(1 for r in placements if r.get("creative_destination") == "Unmatched")
                log_dashboard_usage("Anthem / Elevance", "Generate", output_name, len(placements), direct_count, multi_count, unmatched_count, len((preview or {}).get("creative_names", [])), len(warnings))
                st.success("Anthem T-Sheet generated successfully.")
                for warning in warnings:
                    st.warning(warning)
                st.download_button("Download Anthem T-Sheet", data=output_bytes, file_name=output_name, mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True)
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
# BFAS
# ============================================================

elif selected_account == "BFAS":
    st.success("BFAS automation is ready.")
    st.info("BFAS does not require a Prisma upload. Paste Placement Names, upload creatives, and provide the URL mapping.")
    placement_text = st.text_area("Paste BFAS Placement Names", placeholder="Paste one Placement Name per line.", height=220, key="bfas_placements")
    creative_files = st.file_uploader("Upload BFAS Creative Files", type=["jpg", "jpeg", "png", "gif", "webp", "html", "htm", "mp4", "zip"], accept_multiple_files=True, key="bfas_creatives")
    url_mapping_text = st.text_area("Paste BFAS URL Mapping", placeholder="One URL for all creatives, or: Creative Set<TAB>URL", height=200, key="bfas_urls")
    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("Start Date", key="bfas_start")
    with c2:
        end_date = st.date_input("End Date", key="bfas_end")
    campaign_name = st.text_input("Campaign Name (optional)", key="bfas_campaign")
    site_name = st.text_input("Site Name", value="Nexxen", key="bfas_site")
    preview = None
    if placement_text.strip() and creative_files:
        try:
            preview = preview_bfas_setup(placement_text=placement_text, creative_files=creative_files, url_mapping_text=url_mapping_text)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Placements", len(preview.get("rows", [])))
            m2.metric("Creative Matches", preview.get("matched_count", 0))
            m3.metric("Unmatched", preview.get("unmatched_count", 0))
            m4.metric("URL Matches", preview.get("url_matched_count", 0))
            with st.expander("BFAS Matching Preview", expanded=False):
                st.dataframe(preview.get("rows", []), use_container_width=True)
            for warning in preview.get("warnings", []): st.warning(warning)
        except Exception as exc:
            st.exception(exc)
    if st.button("Generate BFAS T-Sheet", type="primary", use_container_width=True, key="generate_bfas"):
        if not placement_text.strip(): st.error("Please paste BFAS Placement Names.")
        elif not creative_files: st.error("Please upload BFAS creatives.")
        else:
            try:
                output_bytes, warnings, generation_info = generate_bfas_tsheet(placement_text=placement_text, creative_files=creative_files, url_mapping_text=url_mapping_text, start_date=start_date, end_date=end_date, campaign_name=campaign_name, site_name=site_name)
                output_name = "BFAS_Tsheet.xlsm"
                p = preview or {}
                log_dashboard_usage("BFAS", "Generate", output_name, len(p.get("rows", [])), p.get("matched_count", 0), 0, p.get("unmatched_count", 0), p.get("creative_count", 0), len(warnings))
                st.success("BFAS T-Sheet generated successfully.")
                for warning in warnings: st.warning(warning)
                st.download_button("Download BFAS T-Sheet", data=output_bytes, file_name=output_name, mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True)
            except Exception as exc:
                st.exception(exc)


# ============================================================
# CONED
# ============================================================

elif selected_account == "ConEd":
    st.success("ConEd automation is ready.")
    st.info("Upload the Prisma CSV and ConEd creatives, then paste the tagged URLs/UTMs. Matching is dynamic and uses placement/creative naming plus technical attributes.")
    prisma_file, creative_files = common_upload_fields("coned", allow_zip=True)
    urls_text = st.text_area("Paste ConEd URLs / UTMs", placeholder="Paste one tagged URL per line.", height=220, key="coned_urls")
    override_dates = st.checkbox("Override Prisma flight dates", value=False, key="coned_override_dates")
    override_start_date = override_end_date = None
    if override_dates:
        c1, c2 = st.columns(2)
        with c1: override_start_date = st.date_input("Start Date", key="coned_start")
        with c2: override_end_date = st.date_input("End Date", key="coned_end")
    preview = None
    if prisma_file is not None and creative_files:
        try:
            preview = preview_coned_setup(prisma_file=prisma_file, creative_files=creative_files, urls_text=urls_text)
            rows = preview.get("rows", [])
            direct_count = sum(1 for r in rows if r.get("Destination") == "Direct")
            multi_count = sum(1 for r in rows if r.get("Destination") == "Multi")
            unmatched_count = sum(1 for r in rows if r.get("Destination") in {"Unmatched", "Ambiguous"})
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Placements", preview.get("placement_count", len(rows)))
            m2.metric("Direct", direct_count)
            m3.metric("Multi", multi_count)
            m4.metric("Unmatched / Ambiguous", unmatched_count)
            with st.expander("ConEd Matching Preview", expanded=False): st.dataframe(rows, use_container_width=True)
            for warning in preview.get("warnings", []): st.warning(warning)
        except Exception as exc:
            st.exception(exc)
    if st.button("Generate ConEd T-Sheet", type="primary", use_container_width=True, key="generate_coned"):
        if prisma_file is None: st.error("Please upload the Prisma CSV.")
        elif not creative_files: st.error("Please upload ConEd creatives.")
        else:
            try:
                output_bytes, warnings = generate_coned_tsheet(prisma_file=prisma_file, creative_files=creative_files, urls_text=urls_text, override_start_date=override_start_date, override_end_date=override_end_date)
                output_name = "ConEd_Tsheet.xlsm"
                rows = (preview or {}).get("rows", [])
                direct_count = sum(1 for r in rows if r.get("Destination") == "Direct")
                multi_count = sum(1 for r in rows if r.get("Destination") == "Multi")
                unmatched_count = sum(1 for r in rows if r.get("Destination") in {"Unmatched", "Ambiguous"})
                log_dashboard_usage("ConEd", "Generate", output_name, len(rows), direct_count, multi_count, unmatched_count, (preview or {}).get("creative_count", 0), len(warnings))
                st.success("ConEd T-Sheet generated successfully.")
                for warning in warnings: st.warning(warning)
                st.download_button("Download ConEd T-Sheet", data=output_bytes, file_name=output_name, mime="application/vnd.ms-excel.sheet.macroEnabled.12", use_container_width=True)
            except Exception as exc:
                st.exception(exc)



# ============================================================
# PERDUE
# ============================================================
elif selected_account == "Perdue":
    st.success("Perdue automation is ready.")
    st.info(
        "Upload the Prisma CSV and raw creative files. Perdue placement taxonomy is "
        "parsed automatically to create the Ad Name. You provide the base landing URL and the dashboard adds the Final URL + UTM. Raw creative "
        "files are matched and renamed; a ZIP of the renamed creatives is generated "
        "with the T-Sheet. No taxonomy workbook is required."
    )

    prisma_file, creative_files = common_upload_fields("perdue", allow_zip=True)

    st.subheader("Creative Mapping")
    st.caption(
        "Only use this when incoming creative filenames are generic or ambiguous. "
        "Enter one mapping per line as: Original File<TAB>CreativeName-CTA. "
        "Example: video1.mp4<TAB>BouncyBalls-SaveNow"
    )
    creative_mapping_text = st.text_area(
        "Optional Creative Mapping",
        placeholder=(
            "video1.mp4\tBouncyBalls-SaveNow\n"
            "video2.mp4\tFairy-SaveNow\n"
            "300x250.png\tCSCrave-BuyNow"
        ),
        height=150,
        key="perdue_creative_mapping",
    )

    st.subheader("Landing URL")
    st.caption(
        "Paste the BASE landing URL supplied by the team. The dashboard will create and append "
        "the Perdue UTM automatically from the Placement Name. If one URL applies to every "
        "placement, paste only that URL. If different URLs are required, map them as "
        "Product-Effort<TAB>URL or CreativeName-CTA<TAB>URL."
    )
    landing_urls_text = st.text_area(
        "Paste Base Landing URL(s)",
        placeholder=(
            "https://www.perdue.com/products/perdue-crispy-chicken-strips\n\n"
            "OR for multiple URLs:\n"
            "CrispyStrips-Continuity\thttps://www.perdue.com/products/perdue-crispy-chicken-strips\n"
            "GroundChicken-Continuity\thttps://www.perdue.com/products/perdue-fresh-ground-chicken"
        ),
        height=180,
        key="perdue_landing_urls",
    )

    output_name = st.text_input(
        "Output File Name",
        value="Perdue_Tsheet.xlsm",
        key="perdue_output",
    )
    if not output_name.lower().endswith(".xlsm"):
        output_name += ".xlsm"

    preview = None
    if prisma_file is not None and creative_files:
        try:
            preview = preview_perdue_setup(
                prisma_file=prisma_file,
                creative_files=creative_files,
                creative_mapping_text=creative_mapping_text,
                landing_urls_text=landing_urls_text,
            )

            rows = preview.get("rows", [])
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Placements", len(rows))
            m2.metric("Creatives Matched", preview.get("matched_count", 0))
            m3.metric("Unmatched", preview.get("unmatched_count", 0))
            m4.metric("Final URLs Created", preview.get("url_matched_count", 0))

            preview_rows = []
            for row in rows:
                preview_rows.append({
                    "Placement Name": row.get("placement_name", ""),
                    "Channel": row.get("channel", ""),
                    "Partner": row.get("partner", ""),
                    "Size": row.get("size", ""),
                    "Concept": row.get("concept", ""),
                    "Ad Name": row.get("ad_name", ""),
                    "Original Creative": row.get("original_creative", ""),
                    "Renamed Creative": row.get("renamed_creative", ""),
                    "Final URL + UTM": row.get("final_url", ""),
                    "Status": row.get("status", ""),
                })

            with st.expander("Perdue Matching Preview", expanded=True):
                st.dataframe(preview_rows, use_container_width=True, hide_index=True)

            if preview.get("unmatched_count", 0):
                st.warning(
                    "Some creatives could not be matched safely. Add them to Optional "
                    "Creative Mapping above before generating the final files."
                )

            if preview.get("url_unmatched_count", 0):
                st.warning(
                    "Some placements do not have a supplied landing URL. Paste one base URL for all placements, "
                    "or map different URLs using Product-Effort<TAB>URL or CreativeName-CTA<TAB>URL."
                )

            if preview.get("warnings"):
                with st.expander("Perdue Preview Warnings", expanded=False):
                    for warning in preview["warnings"]:
                        st.warning(warning)

        except Exception as exc:
            st.exception(exc)

    if st.button(
        "Generate Perdue T-Sheet + Renamed Creatives",
        type="primary",
        use_container_width=True,
        key="generate_perdue",
    ):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        elif not landing_urls_text.strip():
            st.error("Please paste the Perdue base landing URL. The dashboard will add the UTM automatically.")
        elif not creative_files:
            st.error("Please upload the Perdue creative files.")
        else:
            try:
                with st.spinner("Generating Perdue T-Sheet and renamed creatives..."):
                    output_bytes, renamed_zip, warnings, stats = generate_perdue_tsheet(
                        prisma_file=prisma_file,
                        creative_files=creative_files,
                        creative_mapping_text=creative_mapping_text,
                        landing_urls_text=landing_urls_text,
                    )

                log_dashboard_usage(
                    account="Perdue",
                    action="T-Sheet Generated",
                    output_file=output_name,
                    ads_processed=stats.get("placement_count", 0),
                    direct_count=stats.get("matched_count", 0),
                    multi_count=0,
                    unmatched_count=stats.get("unmatched_count", 0),
                    creative_count=stats.get("creative_count", 0),
                    warning_count=len(warnings),
                    estimated_minutes_saved=60,
                )

                st.success("Perdue T-Sheet and renamed creatives generated successfully.")
                st.caption(
                    f"{stats.get('placement_count', 0)} placements processed | "
                    f"{stats.get('matched_count', 0)} creative matches | "
                    f"{stats.get('url_matched_count', 0)} Final URLs created"
                )

                if warnings:
                    with st.expander("Review Perdue warnings", expanded=False):
                        for warning in warnings:
                            st.warning(warning)

                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        "Download Perdue T-Sheet",
                        data=output_bytes,
                        file_name=output_name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        use_container_width=True,
                    )
                with c2:
                    st.download_button(
                        "Download Renamed Creatives ZIP",
                        data=renamed_zip,
                        file_name="Perdue_Renamed_Creatives.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

            except Exception as exc:
                st.exception(exc)


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
