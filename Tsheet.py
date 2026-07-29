import streamlit as st

from pulte_vip import generate_pulte_tsheet


ACCOUNT_NAMES = [
    "Pulte VIP",
    "Pulte",
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

if selected_account == "Pulte VIP":
    st.success("Pulte VIP automation is ready.")

    prisma_file = st.file_uploader(
        "Upload Prisma CSV",
        type=["csv"],
        key="pulte_prisma",
    )

    creative_files = st.file_uploader(
        "Upload Creative Files",
        type=["jpg", "jpeg", "png", "gif"],
        accept_multiple_files=True,
        key="pulte_creatives",
    )

    landing_urls_text = st.text_area(
        "Paste Landing URLs",
        placeholder="Paste one landing URL per line",
        height=140,
    )

    output_name = st.text_input(
        "Output File Name",
        value="Pulte_Tsheet.xlsm",
    )

    if not output_name.lower().endswith(".xlsm"):
        output_name = f"{output_name}.xlsm"

    generate_clicked = st.button(
        "Generate Pulte T-Sheet",
        type="primary",
        use_container_width=True,
    )

    if generate_clicked:
        if prisma_file is None:
            st.error("Please upload the Prisma CSV.")
        elif not creative_files:
            st.error("Please upload at least one creative file.")
        elif not landing_urls_text.strip():
            st.error("Please paste at least one landing URL.")
        else:
            try:
                with st.spinner("Generating the Pulte trafficking sheet..."):
                    output_bytes, warnings = generate_pulte_tsheet(
                        prisma_file=prisma_file,
                        creative_files=creative_files,
                        landing_urls_text=landing_urls_text,
                    )

                st.success("Pulte trafficking sheet generated successfully.")

                if warnings:
                    with st.expander("Review warnings"):
                        for warning in warnings:
                            st.warning(warning)

                st.download_button(
                    label="Download Pulte T-Sheet",
                    data=output_bytes,
                    file_name=output_name,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    use_container_width=True,
                )

            except Exception as exc:
                st.exception(exc)

else:
    st.info(
        f"{selected_account} is visible in the dashboard. "
        "Its account-specific automation will be added later."
    )
