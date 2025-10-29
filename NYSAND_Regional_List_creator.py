# NYSAND_Regional_List_creator.py
import streamlit as st
import pandas as pd
import zipfile
import tempfile
import re
import os
import requests, xmltodict
from datetime import datetime

# --- Region mapping loader (bundled first, uploader as fallback) ---
DEFAULT_REGION_PATH = "assets/nysand_region_zips.xlsx"

def load_region_mapping(region_xlsx_file=None):
    """Return a dict of DataFrames keyed by sheet name, or None if not found."""
    import pandas as pd, os
    if region_xlsx_file is not None:
        # User uploaded a file this session
        return pd.read_excel(region_xlsx_file, sheet_name=None)
    if os.path.exists(DEFAULT_REGION_PATH):
        # Use the bundled repo file
        return pd.read_excel(DEFAULT_REGION_PATH, sheet_name=None)
    return None


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

# Per ADA docs, use HTTP only (their HTTPS certificate has expired)
ENDPOINT = "http://ws.eatright.org/service/service.svc"

# --- WSDL action map helpers ---
WSDL_URL = "http://ws.eatright.org/service/service.svc?wsdl"

def _wsdl_actions_map() -> dict:
    """Return {operationName: [soapAction, ...]} by parsing the WSDL (best-effort)."""
    import requests, xmltodict
    try:
        r = requests.get(WSDL_URL, timeout=60)
        r.raise_for_status()
        doc = xmltodict.parse(r.text)
    except Exception:
        return {}

    actions = {}
    defs = doc.get("wsdl:definitions") or doc.get("definitions") or {}
    bindings = defs.get("wsdl:binding") or defs.get("binding") or []
    if isinstance(bindings, dict):
        bindings = [bindings]
    for b in bindings:
        ops = b.get("wsdl:operation") or b.get("operation") or []
        if isinstance(ops, dict):
            ops = [ops]
        for op in ops:
            name = op.get("@name")
            if not name:
                continue
            # SOAP 1.1
            soap_op = op.get("soap:operation")
            if isinstance(soap_op, dict) and "@soapAction" in soap_op:
                actions.setdefault(name, []).append(soap_op["@soapAction"])
            # SOAP 1.2
            soap12_op = op.get("soap12:operation")
            if isinstance(soap12_op, dict) and "@soapAction" in soap12_op:
                actions.setdefault(name, []).append(soap12_op["@soapAction"])
    return actions

_WSDL_ACTIONS = _wsdl_actions_map()

def _wsdl_action_for(method: str) -> list[str]:
    return _WSDL_ACTIONS.get(method, [])


# --- Keep this header builder exactly as is ---
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
    """
    SOAP 1.1 over HTTP only.
    1) Try WSDL-declared soapAction(s), quoted.
    2) Fall back to known WCF patterns, quoted (no SOAP 1.2).
    Surfaces SOAP Faults or raw body on error.
    """
    import requests, xmltodict

    base = API_NS  # "http://eatright/membership"

    # Prioritize WSDL-declared actions (if any)
    candidates = list(dict.fromkeys(_wsdl_action_for(action)))  # dedupe, preserve order

    # Fallback patterns commonly seen in WCF
    candidates += [
        f"{base}/{action}",
        f"{base}/IService/{action}",
        f"{base}/IWcfAdaMembership/{action}",
        "",  # empty SOAPAction (some WCF allow this)
    ]

    errors = []
    for sa in candidates:
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "Accept": "text/xml",
        }
        # SOAP 1.1 uses SOAPAction header; WCF prefers it quoted
        if sa is not None:
            headers["SOAPAction"] = f"\"{sa}\"" if sa else ""

        try:
            r = requests.post(ENDPOINT, data=envelope_xml.encode("utf-8"), headers=headers, timeout=60)
            if r.status_code >= 400:
                # Try to extract a SOAP Fault
                try:
                    doc = xmltodict.parse(r.text)
                    fault = (doc.get("s:Envelope", {}).get("s:Body", {}).get("s:Fault")
                             or doc.get("Envelope", {}).get("Body", {}).get("Fault"))
                    if fault:
                        fs = fault.get("faultstring") or fault.get("faultcode") or "SOAP Fault"
                        errors.append(f"1.1 SOAPAction={headers.get('SOAPAction','(none)')} → Fault: {fs}")
                        continue
                except Exception:
                    pass
                snippet = (r.text or "").strip()
                if len(snippet) > 1200:
                    snippet = snippet[:1200] + " …(truncated)…"
                errors.append(f"1.1 SOAPAction={headers.get('SOAPAction','(none)')} → HTTP {r.status_code}: {snippet}")
                continue

            # Success — parse and return
            return xmltodict.parse(r.text)

        except Exception as e:
            errors.append(f"1.1 SOAPAction={headers.get('SOAPAction','(none)')} → {e}")

    raise RuntimeError(" / ".join(errors) or "All SOAP 1.1 variants failed.")


def _validate_access_key(access_key: str) -> bool:
    body = f"""
    <ValidateAccessKey xmlns="{API_NS}">
      <key>{access_key}</key>
    </ValidateAccessKey>
    """.strip()
    # IMPORTANT: per docs, do NOT send the AccessKey header for ValidateAccessKey
    env = _soap_envelope(body, access_key=None)
    data = _post_soap("ValidateAccessKey", env)
    try:
        result = data["s:Envelope"]["s:Body"]["ValidateAccessKeyResponse"]["ValidateAccessKeyResult"]
        # Some stacks use 'a:Success'; keep a fallback for 'Success'
        return str(result.get("a:Success", result.get("Success", "false"))).lower() == "true"
    except Exception:
        return False

def fetch_members_via_api(access_key: str, group_key: str) -> pd.DataFrame:
    """RetrieveGroupMembersWithCustomProperties → pandas DataFrame"""
    body = f"""
    <RetrieveGroupMembersWithCustomProperties xmlns="{API_NS}">
      <groupKey>{group_key}</groupKey>
    </RetrieveGroupMembersWithCustomProperties>
    """.strip()
    # For all other methods, include the AccessKey header
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

def process_and_package(members: pd.DataFrame, region_sheets: dict) -> bytes:
    import zipfile, tempfile, os
    # Clean ZIP and build Zip_clean used for merge
    members = members.copy()
    members["Zip_clean"] = members["Zip"].apply(_clean_zip)

    # Build one tall map from all sheets (first 3 columns: County, Zip, Region)
    zip_map_all = pd.DataFrame()
    for _, df in region_sheets.items():
        df = df.iloc[:, :3].copy()
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

if source == "Manual Upload":
    member_file = st.file_uploader("📄 Upload Member Export CSV", type="csv")
    # Optional override of bundled mapping:
    region_file = st.file_uploader("📄 (Optional) Upload NYSAND Region Zipcodes Excel (overrides bundled)", type=["xls", "xlsx"], key="regionfile")

    if member_file:
        region_sheets = load_region_mapping(region_file)
        if region_sheets is None:
            st.error("Region mapping not found. Please add assets/nysand_region_zips.xlsx to the repo or upload it above.")
        else:
            with st.spinner("Processing files..."):
                members_df = pd.read_csv(member_file)
                if "Zip" not in members_df.columns:
                    st.error("Uploaded CSV must contain a 'Zip' column.")
                else:
                    blob = process_and_package(members_df, region_sheets)
                    st.success("✅ Done! Download your ZIP below.")
                    st.download_button("📥 Download All Files (ZIP)", blob, file_name="NYSAND_Member_Files.zip")

elif source == "EatRight SOAP API":
    st.info("Uses secrets: EATR_ACCESS_KEY and EATR_GROUP_KEY")

    # Optional override of bundled mapping:
    region_file = st.file_uploader("📄 (Optional) Upload NYSAND Region Zipcodes Excel (overrides bundled)", type=["xls", "xlsx"], key="regionfile_api")

    fetch_clicked = st.button("🔄 Fetch members via API")
    if fetch_clicked:
        try:
            ak = st.secrets["EATR_ACCESS_KEY"].strip()
            gk = st.secrets["EATR_GROUP_KEY"].strip()

            with st.spinner("Validating AccessKey…"):
                if not _validate_access_key(ak):
                    st.error("AccessKey invalid. Check Vendor Access in the portal.")
                    st.stop()

            with st.spinner("Fetching members from EatRight API…"):
                api_df = fetch_members_via_api(ak, gk)
            st.success(f"Fetched {len(api_df):,} records.")
            st.dataframe(api_df.head(25))

            region_sheets = load_region_mapping(region_file)
            if region_sheets is None:
                st.error("Region mapping not found. Please add assets/nysand_region_zips.xlsx to the repo or upload it above.")
                st.stop()

            with st.spinner("Creating region files…"):
                blob = process_and_package(api_df, region_sheets)
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
            st.exception(e)
