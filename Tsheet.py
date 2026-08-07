import streamlit as st

from pulte_normal import generate_normal_pulte_tsheet
from pulte_vip import generate_pulte_tsheet
from aaa import (
    generate_aaa_tsheet,
    preview_aaa_setup,
    validate_multi_rotation,
)


ACCOUNT_NAMES = [
    "Pulte",
    "Pulte VIP",
    "Anthem / Elevance",
    "UPS Store",
    "Hyatt",
    "AAA",
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

st.title("Traffic Sheet Generator")
st.caption("Select an account and generate the required trafficking sheet.")

selected_account = st.selectbox(
    "Select Account",
    ACCOUNT_NAMES,
    index=0,
)


def common_upload_fields(key_prefix: str):
    prisma = st.file_uploader(
        "Upload Prisma CSV",
        type=["csv", "txt"],
        key=f"{key_prefix}_prisma",
    )

    creatives = st.file_uploader(
        "Upload Creative Files",
        type=["jpg", "jpeg", "png", "gif"],
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
                with st.spinner("Generating the normal Pulte T-Sheet..."):
                    output_bytes, warnings = (
                        generate_normal_pulte_tsheet(
                            prisma_file=prisma_file,
                            creative_files=creative_files,
                            complete_urls_text=complete_urls_text,
                        )
                    )

                st.success("Pulte T-Sheet generated successfully.")

                if warnings:
                    with st.expander("Review matching and dimension warnings"):
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

    prisma_file, creative_files = common_upload_fields("pulte_vip")

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
                with st.spinner("Generating the Pulte VIP T-Sheet..."):
                    output_bytes, warnings = generate_pulte_tsheet(
                        prisma_file=prisma_file,
                        creative_files=creative_files,
                        landing_urls_text=landing_urls_text,
                    )

                st.success("Pulte VIP T-Sheet generated successfully.")

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


else:
    st.info(
        f"{selected_account} is visible in the dashboard. "
        "Its account-specific automation will be added later."
    )
