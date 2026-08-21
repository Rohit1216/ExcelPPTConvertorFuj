import io
from datetime import date

import streamlit as st

from utils.ppt_generator import load_teams_from_workbook, generate_presentation

st.set_page_config(page_title="LinkedIn Team Deck Generator", page_icon="📊", layout="centered")

st.title("📊 LinkedIn Team → PowerPoint Generator")
st.write(
    "Upload an Excel file (one sheet per team/section) with columns "
    "**Name, Designation, LinkedinUrl, PhotoUrl** (Phone/Email/Location optional) "
    "and get a branded .pptx deck back — one card grid per person, name "
    "hyperlinked to their LinkedIn profile, photo pulled from PhotoUrl."
)

with st.sidebar:
    st.header("Report settings")
    report_title = st.text_input("Report title", value="Account Intelligence Report")
    report_subtitle = st.text_input("Subtitle", value=date.today().strftime("%b %d, %Y"))
    st.markdown("---")
    st.subheader("Featured first row")
    st.caption(
        "For these sheets, the first person (e.g. a CEO / Secretary General) "
        "is rendered as a large highlighted card above the grid."
    )
    featured_input = st.text_input(
        "Sheet names (comma-separated)", value="Executive Management"
    )
    st.markdown("---")
    template_file = st.file_uploader(
        "Optional: base template .pptx (uses its theme/master)", type=["pptx"]
    )

uploaded = st.file_uploader("Upload Excel workbook (.xlsx)", type=["xlsx"])

if uploaded is not None:
    featured_sheets = [s.strip() for s in featured_input.split(",") if s.strip()]

    with st.spinner("Reading workbook..."):
        try:
            teams = load_teams_from_workbook(uploaded, featured_sheets=featured_sheets)
        except Exception as e:
            st.error(f"Could not read the Excel file: {e}")
            st.stop()

    if not teams:
        st.warning("No sheets with a 'Name' column were found.")
        st.stop()

    st.success(f"Found {len(teams)} sheet(s):")
    for t in teams:
        st.write(f"- **{t.sheet_name}** — {len(t.people)} people"
                  + (" (featured first row)" if t.featured_first else ""))

    if st.button("Generate PPTX", type="primary"):
        progress = st.progress(0, text="Downloading photos and building slides...")
        template_path = None
        if template_file is not None:
            template_path = io.BytesIO(template_file.read())

        try:
            pptx_bytes = generate_presentation(
                teams,
                title=report_title,
                subtitle=report_subtitle,
                template_path=template_path,
            )
        except Exception as e:
            st.error(f"Failed to generate the deck: {e}")
            st.stop()
        progress.progress(100, text="Done!")

        st.download_button(
            "⬇️ Download PowerPoint",
            data=pptx_bytes,
            file_name="team_report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
else:
    st.info("👆 Upload an .xlsx file to get started.")