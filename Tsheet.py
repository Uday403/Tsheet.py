import itertools

import streamlit as st

from aaa import (
    creative_version_key,
    generate_aaa_tsheet,
    preview_aaa_setup,
    validate_multi_rotation,
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


if selected_account == "Pulte":
    st.success("Normal Pulte automation is ready.")

    st.info(
        "Paste the complete URLs/UTMs exactly as provided by the team. "
        "The dashboard will not create or modify the UTM."
    )

    prisma_file, creative_files = common_upload_fields("pulte")

    complete_urls_text = st.text_area(
        "Paste Complete URLs / UTMs",
        placeholder="Paste one complete URL per line",
        height=160,
        key="pulte_urls",
    )

    output_name = st.text_input(
        "Output File Name",
        value="Pulte_Tsheet.xlsm",
        key="pulte_output",
    )

    if not output_name.lower().endswith(".xlsm"):
        output_name += ".xlsm"

    if st.button(
        "Generate Pulte T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        elif not creative_files:
            st.error("Please upload at least one creative file.")
        elif not complete_urls_text.strip():
            st.error("Please paste the complete URLs/UTMs.")
        else:
            try:
                with st.spinner(
                    "Generating the normal Pulte T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_normal_pulte_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            complete_urls_text=complete_urls_text,
                        )
                    )

                st.success(
                    "Pulte T-Sheet generated successfully."
                )

                if warnings:
                    with st.expander(
                        "Review matching and dimension warnings"
                    ):
                        for warning in warnings:
                            st.warning(warning)

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


elif selected_account == "Pulte VIP":
    st.success("Pulte VIP automation is ready.")

    prisma_file, creative_files = common_upload_fields(
        "pulte_vip"
    )

    landing_urls_text = st.text_area(
        "Paste Landing URLs",
        placeholder="Paste one landing URL per line",
        height=160,
        key="pulte_vip_urls",
    )

    output_name = st.text_input(
        "Output File Name",
        value="Pulte_VIP_Tsheet.xlsm",
        key="pulte_vip_output",
    )

    if not output_name.lower().endswith(".xlsm"):
        output_name += ".xlsm"

    if st.button(
        "Generate Pulte VIP T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        elif not creative_files:
            st.error("Please upload at least one creative file.")
        elif not landing_urls_text.strip():
            st.error("Please paste at least one landing URL.")
        else:
            try:
                with st.spinner(
                    "Generating the Pulte VIP T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_pulte_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            landing_urls_text=landing_urls_text,
                        )
                    )

                st.success(
                    "Pulte VIP T-Sheet generated successfully."
                )

                if warnings:
                    with st.expander("Review warnings"):
                        for warning in warnings:
                            st.warning(warning)

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


elif selected_account == "AAA":
    st.success("AAA automation is ready.")

    st.info(
        "AAA rules: Placement Name = Ad Name. "
        "Enter the BASE landing URL only; the dashboard creates "
        "the AAA pmed automatically."
    )

    prisma_file, creative_files = common_upload_fields(
        "aaa",
        allow_zip=True,
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
            "https://www.ace.aaa.com/travel/category/cruises.html"
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
            override_start_date = st.date_input(
                "Start Date",
                key="aaa_start_date",
            )

        with col2:
            override_end_date = st.date_input(
                "End Date",
                key="aaa_end_date",
            )

    rotation_by_version = {}
    separate_url_by_version = {}
    preview = None

    if prisma_file is not None and creative_files:
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
                for placement in preview["placements"]:
                    matches = placement["matches"]

                    st.write(
                        f"**{placement['dimension'] or 'No dimension'}** "
                        f"— {placement['placement_name']}"
                    )

                    if matches:
                        for creative in matches:
                            st.caption(f"↳ {creative}")
                    else:
                        st.warning("No creative matched this placement.")

            for warning in preview["warnings"]:
                st.warning(warning)

        except Exception as exc:
            st.error(f"Unable to preview AAA matching: {exc}")

    if (
        creative_setup == "Multiple creatives per ad"
        and preview is not None
    ):
        st.subheader("Multi Creative Rotation")

        st.caption(
            "Rotation is entered once per creative VERSION and is "
            "reused across all matching dimensions. "
            "For example V1 can be 19% for 160x600, 300x250, etc."
        )

        version_groups = preview["version_groups"]

        for index, (version, files) in enumerate(
            version_groups.items()
        ):
            st.markdown(f"**{version}**")
            st.caption(" / ".join(files))

            rotation_by_version[version] = st.number_input(
                f"Rotation % — {version}",
                min_value=0.0,
                max_value=100.0,
                value=0.0,
                step=1.0,
                key=f"aaa_rotation_{index}",
            )

            use_separate_url = st.checkbox(
                f"Use a separate landing URL for {version}",
                value=False,
                key=f"aaa_separate_url_check_{index}",
            )

            if use_separate_url:
                separate_url_by_version[version] = (
                    st.text_input(
                        f"Separate Base URL — {version}",
                        placeholder=default_base_url,
                        key=f"aaa_separate_url_{index}",
                    )
                )

        rotation_errors = validate_multi_rotation(
            preview=preview,
            rotation_by_version=rotation_by_version,
        )

        if rotation_errors:
            st.warning(
                "Rotation must total 100% for every Multi placement."
            )

            for error in rotation_errors:
                st.caption(error)
        else:
            st.success(
                "Rotation validation passed: each matched Multi "
                "placement totals 100%."
            )

    output_name = st.text_input(
        "Output File Name",
        value="AAA_Tsheet.xlsm",
        key="aaa_output",
    )

    if not output_name.lower().endswith(".xlsm"):
        output_name += ".xlsm"

    if st.button(
        "Generate AAA T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        elif not creative_files:
            st.error("Please upload creative files.")
        elif not default_base_url.strip():
            st.error("Please enter the base landing URL.")
        elif (
            creative_setup == "Multiple creatives per ad"
            and preview is None
        ):
            st.error(
                "AAA creative preview could not be created."
            )
        else:
            try:
                if creative_setup == "Multiple creatives per ad":
                    rotation_errors = validate_multi_rotation(
                        preview=preview,
                        rotation_by_version=rotation_by_version,
                    )

                    if rotation_errors:
                        st.error(
                            "Fix the rotation percentages before "
                            "generating the T-Sheet."
                        )
                        st.stop()

                with st.spinner(
                    "Generating the AAA T-Sheet..."
                ):
                    output_bytes, warnings = generate_aaa_tsheet(
                        prisma_file=prisma_file,
                        creative_files=creative_files,
                        creative_setup=creative_setup,
                        default_base_url=default_base_url,
                        rotation_by_version=rotation_by_version,
                        separate_base_url_by_version=(
                            separate_url_by_version
                        ),
                        override_start_date=override_start_date,
                        override_end_date=override_end_date,
                    )

                st.success(
                    "AAA T-Sheet generated successfully."
                )

                if warnings:
                    with st.expander("AAA Review Warnings"):
                        for warning in warnings:
                            st.warning(warning)

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



elif selected_account == "Simon VIP":
    st.success("Simon VIP automation is ready.")

    st.info(
        "Simon VIP does not use Prisma. Paste the Placement taxonomy "
        "directly below. Placement Name = Ad Name. "
        "If no creatives are uploaded, Tracking_1x1 will be used."
    )

    placement_text = st.text_area(
        "Paste Placement Names / Taxonomy",
        placeholder=(
            "Paste one Placement Name per line\n"
            "Example: Simon_..._Arundel Mills_..._300x250"
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
            "Desert Hills Premium Outlets\thttps://..."
        ),
        height=220,
        key="simon_vip_utm",
    )

    outlet_date_text = st.text_area(
        "Paste Outlet Name + Start Date + End Date",
        placeholder=(
            "Arundel Mills\t08/01/2026\t08/31/2026\n"
            "Desert Hills Premium Outlets\t08/01/2026\t08/31/2026"
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
            preview = preview_simon_vip_setup(
                placement_text=placement_text,
                creative_files=creative_files,
                outlet_utm_text=outlet_utm_text,
                outlet_date_text=outlet_date_text,
            )

            metric_col1, metric_col2, metric_col3 = st.columns(3)

            with metric_col1:
                st.metric(
                    "Outlet mappings loaded",
                    preview["outlet_mapping_count"],
                )

            with metric_col2:
                st.metric(
                    "Placements matched",
                    preview["utm_matched_count"],
                )

            with metric_col3:
                st.metric(
                    "Unmatched placements",
                    preview["utm_unmatched_count"],
                )

            if preview["using_tracking_1x1"]:
                st.info(
                    "No creatives uploaded — Tracking_1x1 will be used "
                    "for all placements."
                )

            with st.expander(
                "Simon VIP Matching Preview",
                expanded=True,
            ):
                for row in preview["rows"]:
                    st.write(f"**{row['placement_name']}**")
                    st.caption(
                        f"Outlet: {row['outlet'] or 'Not matched'}"
                    )
                    st.caption(
                        f"Creative: {row['creative'] or 'Not matched'}"
                    )
                    st.caption(
                        f"Dates: {row['start_date'] or 'Not matched'} "
                        f"to {row['end_date'] or 'Not matched'}"
                    )

            if preview["warnings"]:
                with st.expander("Simon VIP Preview Warnings"):
                    for warning in preview["warnings"]:
                        st.warning(warning)

        except Exception as exc:
            st.error(
                f"Unable to preview Simon VIP matching: {exc}"
            )

    output_name = st.text_input(
        "Output File Name",
        value="Simon_VIP_Tsheet.xlsm",
        key="simon_vip_output",
    )

    if not output_name.lower().endswith(".xlsm"):
        output_name += ".xlsm"

    if st.button(
        "Generate Simon VIP T-Sheet",
        type="primary",
        use_container_width=True,
    ):
        if not placement_text.strip():
            st.error(
                "Please paste the Placement taxonomy."
            )

        elif not outlet_utm_text.strip():
            st.error(
                "Please paste Outlet Name and UTM mapping."
            )

        elif not outlet_date_text.strip():
            st.error(
                "Please paste Outlet Name, Start Date and End Date mapping."
            )

        else:
            try:
                with st.spinner(
                    "Generating the Simon VIP T-Sheet..."
                ):
                    output_bytes, warnings = (
                        generate_simon_vip_tsheet(
                            placement_text=placement_text,
                            creative_files=creative_files,
                            outlet_utm_text=outlet_utm_text,
                            outlet_date_text=outlet_date_text,
                        )
                    )

                st.success(
                    "Simon VIP T-Sheet generated successfully."
                )

                if warnings:
                    with st.expander(
                        "Review Simon VIP warnings"
                    ):
                        for warning in warnings:
                            st.warning(warning)

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



elif selected_account == "Naming Convention Generator":
    st.success("Naming Convention Generator is ready.")
    st.info("Copy the complete taxonomy table from Excel, including the header row, and paste it below. Any number of columns and values can be used.")

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
        "Naming Separator", ["_", "-", "|"], index=0, key="naming_separator"
    )

    column_values = []
    usable_columns = []
    total_combinations = 0
    table_valid = False

    if taxonomy_text.strip():
        try:
            rows = [row.split("\t") for row in taxonomy_text.splitlines() if row.strip()]
            if len(rows) < 2:
                st.warning("Paste the header row and at least one row of values.")
            else:
                headers = [header.strip() for header in rows[0]]
                for column_index in range(len(headers)):
                    values = []
                    for row in rows[1:]:
                        if column_index < len(row):
                            value = row[column_index].strip()
                            if value and value not in values:
                                values.append(value)
                    column_values.append(values)

                usable_columns = [(h, v) for h, v in zip(headers, column_values) if v]
                if usable_columns:
                    table_valid = True
                    st.subheader("Detected Taxonomy")
                    preview_count = min(len(usable_columns), 4)
                    preview_columns = st.columns(preview_count)
                    for index, (header, values) in enumerate(usable_columns):
                        with preview_columns[index % preview_count]:
                            st.metric(header or f"Column {index + 1}", len(values))
                            preview_text = " | ".join(values[:8])
                            if len(values) > 8:
                                preview_text += " | ..."
                            st.caption(preview_text)

                    total_combinations = 1
                    for _, values in usable_columns:
                        total_combinations *= len(values)
                    st.info(f"Total naming conventions that will be generated: {total_combinations:,}")
                    if total_combinations > 100000:
                        st.warning("This taxonomy will generate more than 100,000 combinations. Consider reducing the number of values before generating.")
        except Exception as exc:
            st.error(f"Unable to read the pasted taxonomy: {exc}")

    if st.button("Generate Naming Conventions", type="primary", use_container_width=True, key="generate_naming_conventions"):
        if not taxonomy_text.strip():
            st.error("Please paste the taxonomy table from Excel.")
        elif not table_valid:
            st.error("No usable taxonomy values were detected.")
        elif total_combinations > 500000:
            st.error("More than 500,000 combinations were detected. Please reduce the taxonomy before generating.")
        else:
            try:
                usable_values = [values for _, values in usable_columns]
                generated_names = [separator.join(combination) for combination in itertools.product(*usable_values)]
                output_text = "\n".join(generated_names)
                st.success(f"{len(generated_names):,} naming conventions generated successfully.")
                st.text_area("Generated Naming Conventions", value=output_text, height=400, key="naming_generated_output")
                st.download_button("Download Naming Conventions", data=output_text, file_name="Naming_Conventions.txt", mime="text/plain", use_container_width=True)
            except Exception as exc:
                st.exception(exc)


else:
    st.info(
        f"{selected_account} is visible in the dashboard. "
        "Its account-specific automation will be added later."
    )
