# NYSAND_Regional_List_creator_API.py
import streamlit as st
import pandas as pd
import zipfile
import tempfile
import re
import os
import requests, xmltodict
from datetime import datetime

# =========================
# Branding header (unchanged)
# =========================
header_left, header_right = st.columns([3, 8])

def _find_logo():
    for p in ("assets/logo.png", "logo.png"):
        if os.path.exists(p):
            return p
    return None

with header_left:
    _logo = _find_logo()
    if _logo:
        st.image(_logo, width=220)
    else:
        st.caption("(logo not found: assets/logo.png or logo.png)")
with header_right:
    st.markdown("## NYSAND Region-Based Member Splitter")

st.markdown("""
Upload your **Member Export CSV** and the **NYSAND Region Zipcodes Excel file**, **or** fetch the member list via the EatRight SOAP API, then:
- Clean and match ZIP codes  
- Add Region and County  
- Split the data by Region  
- Provide an unmatched/out-of-state file  
- Download everything in a single ZIP
""")

# =========================
# SOAP API helpers
# =========================
SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
API_NS  = "http://eatright/membership"
ENDPOINT = "https://ws.eatright.org/service/service.svc"

def _soap_envelope(body_xml: str, *, access_key: str | None) -> str:
    header = (
        f"""
        <s:Header>
          <AccessKey xmlns="{API_NS}" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
            <Value>{access_key}</Value>
          </AccessKey>
        </s:Header>
        """
        if access_key else "<s:Header/>"
    )
    return f"""<s:Envelope xmlns:s="{SOAP_NS}">
{header}
  <s:Body>
{body_xml}
  </s:Body>
</s:Envelope>""".strip()

def _post_soap(action: str, envelope_xml: str) -> dict:
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": f"{API_NS}/{action}",
    }
    r = requests.post(ENDPOINT, data=envelope_xml.encode("utf-8"), headers=headers, timeout=60)
    r.raise_for_status()
    return xmltodict.parse(r.text)

def _validate_access_key(access_key: str) -> bool:
    body = f"""
    <ValidateAccessKey xmlns="{API_NS}">
      <key>{access_key}</key>
    </ValidateAccessKey>
    """.strip()
    env = _soap_envelope(body, access_key=None)
    data = _post_soap("ValidateAccessKey", env)
    try:
        result = data["s:Envelope"]["s:Body"]["ValidateAccessKeyResponse"]["ValidateAccessKeyResult"]
        return str(result.get("a:Success", "false")).lower() == "true"
    except Exception:
        return False

def fetch_members_via_api(access_key: str, group_key: str) -> pd.DataFrame:
    """RetrieveGroupMembersWithCustomProperties → pandas DataFrame"""
    body = f"""
    <RetrieveGroupMembersWithCustomProperties xmlns="{API_NS}">
      <groupKey>{group_key}</groupKey>
    </RetrieveGroupMembersWithCustomProperties>
    """.strip()
    env = _soap_envelope(body, access_key=access_key)
    data = _post_soap("RetrieveGroupMembersWithCustomProperties", env)

    node = data["s:Envelope"]["s:Body"]["RetrieveGroupMembersWithCustomPropertiesResponse"]["RetrieveGroupMembersWithCustomPropertiesResult"]

    # Find list of members defensively
    def _to_rows(obj):
        # common shapes: obj["a:Members"]["a:Member"] or already a list
        if isinstance(obj, dict):
            for k in ("a:Members", "Members"):
                if k in obj:
                    obj = obj[k]
                    break
            for k in ("a:Member", "Member"):
                if isinstance(obj, dict) and k in obj:
                    obj = obj[k]
                    break
        if isinstance(obj, list):
            return obj
        if isinstance(obj, dict):
            return [obj]
        return []

    rows = _to_rows(node)

    def strip_ns(d):
        return {k.split(":", 1)[-1]: v for k, v in d.items()} if isinstance(d, dict) else {}

    cleaned = [strip_ns(r) for r in rows]

    # Expand CustomProperties → columns
    def props_to_dict(v):
        if not isinstance(v, dict):
            return {}
        items = v.get("a:CustomProperty") or v.get("CustomProperty") or []
        if isinstance(items, dict):
            items = [items]
        out = {}
        for it in items:
            if not isinstance(it, dict):
                continue
            name = it.get("a:Name") or it.get("Name")
            val  = it.get("a:Value") or it.get("Value")
            if name:
                out[str(name)] = val
        return out

    df = pd.DataFrame(cleaned)
    if "CustomProperties" in df.columns:
        props = df["CustomProperties"].apply(props_to_dict).apply(pd.Series)
        df = pd.concat([df.drop(columns=["CustomProperties"]), props], axis=1)

    # Normalize: create a 'Zip' column that matches your current pipeline
    candidate_cols = [c for c in df.columns if c.lower() in ("zip", "postalcode", "postal_code", "zipcode")]
    if candidate_cols:
        df["Zip"] = df[candidate_cols[0]]
    else:
        # If no postal field exists, create empty (keeps pipeline from breaking)
        df["Zip"] = None

    return df

# =========================
# Shared processing
# =========================
def _clean_zip(zipcode):
    if pd.isna(zipcode):
        return None
    m = re.search(r"\b\d{5}\b", str(zipcode))
    return m.group(0) if m else None

def process_and_package(members: pd.DataFrame, region_xlsx_bytes) -> bytes:
    # Clean ZIP and build Zip_clean used for merge
    members = members.copy()
    members["Zip_clean"] = members["Zip"].apply(_clean_zip)

    # Load zip region mapping (all sheets, first 3 cols → County, Zip, Region)
    zip_map_all = pd.DataFrame()
    region_xls = pd.read_excel(region_xlsx_bytes, sheet_name=None)
    for _, df in region_xls.items():
        df = df.iloc[:, :3]
        df.columns = ["County", "Zip", "Region"]
        df["Zip"] = df["Zip"].astype(str).str.zfill(5)
        zip_map_all = pd.concat([zip_map_all, df], ignore_index=True)

    # Merge + group
    merged = pd.merge(members, zip_map_all, left_on="Zip_clean", right_on="Zip", how="left")
    grouped = merged[merged["Region"].notna()].groupby("Region")

    # Write ZIP archive of region files + unmatched
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "NYSAND_Member_Files.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for region, df in grouped:
                safe_region = str(region).replace("/", "-").replace(" ", "_")
                fname = f"{safe_region}_Members.xlsx"
                fpath = os.path.join(tmpdir, fname)
                df.to_excel(fpath, index=False)
                zipf.write(fpath, arcname=fname)

            unmatched = merged[merged["Region"].isna()]
            unmatched_path = os.path.join(tmpdir, "Unmatched_OutOfState_Members.xlsx")
            unmatched.to_excel(unmatched_path, index=False)
            zipf.write(unmatched_path, arcname="Unmatched_OutOfState_Members.xlsx")

        return open(zip_path, "rb").read()

# =========================
# UI: Source selection
# =========================
with st.sidebar:
    source = st.radio("Data source", ["Manual Upload", "EatRight SOAP API"], index=0)
    st.caption("Run API fetch first, then process into region files.")

region_file = st.file_uploader("📄 Upload NYSAND Region Zipcodes Excel", type=["xls", "xlsx"], key="regionfile")

if source == "Manual Upload":
    member_file = st.file_uploader("📄 Upload Member Export CSV", type="csv")
    if member_file and region_file:
        with st.spinner("Processing files..."):
            members_df = pd.read_csv(member_file)
            if "Zip" not in members_df.columns:
                st.error("Uploaded CSV must contain a 'Zip' column.")
            else:
                blob = process_and_package(members_df, region_file)
                st.success("✅ Done! Download your ZIP below.")
                st.download_button("📥 Download All Files (ZIP)", blob, file_name="NYSAND_Member_Files.zip")

elif source == "EatRight SOAP API":
    st.info("Uses secrets: EATR_ACCESS_KEY and EATR_GROUP_KEY")
    if st.button("🔄 Fetch members via API") and region_file:
        try:
            with st.spinner("Validating AccessKey…"):
                ak = st.secrets["EATR_ACCESS_KEY"]
                gk = st.secrets["EATR_GROUP_KEY"]
                if not _validate_access_key(ak):
                    st.error("AccessKey invalid. Check Vendor Access in the portal.")
                else:
                    with st.spinner("Fetching members from EatRight API…"):
                        api_df = fetch_members_via_api(ak, gk)
                        st.success(f"Fetched {len(api_df):,} records.")
                        st.dataframe(api_df.head(25))

                    with st.spinner("Creating region files…"):
                        blob = process_and_package(api_df, region_file)
                        today = datetime.now().strftime("%Y-%m-%d")
                        st.success("✅ Done! Download your ZIP below.")
                        st.download_button(
                            "📥 Download All Files (ZIP)",
                            blob,
                            file_name=f"NYSAND_Member_Files_{today}.zip"
                        )
        except KeyError as e:
            st.error(f"Missing secret: {e}. Please set EATR_ACCESS_KEY and EATR_GROUP_KEY.")
        except Exception as e:
            st.error(f"API fetch failed: {e}")
