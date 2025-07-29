import streamlit as st
import pandas as pd
import zipfile
import tempfile
import re
import os

st.set_page_config(page_title="NYSAND Region Splitter", layout="centered")
st.title("📍 NYSAND Region-Based Member Splitter")
st.markdown("""
Upload your **Member Export CSV** and the **NYSAND Region Zipcodes Excel file**.
The app will:
- Clean and match ZIP codes
- Add Region and County
- Split the data by Region
- Provide a file for unmatched/out-of-state members
- Let you download everything in a single ZIP
""")

member_file = st.file_uploader("📄 Upload Member Export CSV", type="csv")
region_file = st.file_uploader("📄 Upload NYSAND Region Zipcodes Excel", type=["xls", "xlsx"])

if member_file and region_file:
    with st.spinner("Processing files..."):
        # Load member file
        members = pd.read_csv(member_file)

        # Clean ZIP
        def clean_zip(zipcode):
            if pd.isna(zipcode): return None
            match = re.search(r"\b\d{5}\b", str(zipcode))
            return match.group(0) if match else None

        members['Zip_clean'] = members['Zip'].apply(clean_zip)

        # Load zip region mapping
        zip_map_all = pd.DataFrame()
        region_xls = pd.read_excel(region_file, sheet_name=None)

        for _, df in region_xls.items():
            df = df.iloc[:, :3]  # First 3 columns
            df.columns = ['County', 'Zip', 'Region']
            df['Zip'] = df['Zip'].astype(str).str.zfill(5)
            zip_map_all = pd.concat([zip_map_all, df], ignore_index=True)

        # Merge
        merged = pd.merge(members, zip_map_all, left_on='Zip_clean', right_on='Zip', how='left')

        # Group by Region
        grouped = merged[merged['Region'].notna()].groupby('Region')

        # Write to temp dir
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "NYSAND_Member_Files.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for region, df in grouped:
                    safe_region = region.replace("/", "-").replace(" ", "_")
                    fname = f"{safe_region}_Members.xlsx"
                    fpath = os.path.join(tmpdir, fname)
                    df.to_excel(fpath, index=False)
                    zipf.write(fpath, arcname=fname)

                # Unmatched
                unmatched = merged[merged['Region'].isna()]
                unmatched_path = os.path.join(tmpdir, "Unmatched_OutOfState_Members.xlsx")
                unmatched.to_excel(unmatched_path, index=False)
                zipf.write(unmatched_path, arcname="Unmatched_OutOfState_Members.xlsx")

            with open(zip_path, "rb") as f:
                st.success("✅ Done! Download your ZIP file below.")
                st.download_button("📥 Download All Files (ZIP)", f.read(), file_name="NYSAND_Member_Files.zip")
