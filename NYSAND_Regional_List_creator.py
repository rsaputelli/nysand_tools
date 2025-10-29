# NYSAND_Regional_List_creator.py  (drop-in repaired)
import os
import re
import zipfile
import tempfile
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
import xmltodict

# =========================
# Region mapping loader (bundled first, uploader as fallback)
# =========================
DEFAULT_REGION_PATH = "assets/nysand_region_zips.xlsx"

def load_region_mapping(region_xlsx_file=None):
    """Return a dict of DataFrames keyed by sheet name, or None if not found."""
    if region_xlsx_file is not None:
        return pd.read_excel(region_xlsx_file, sheet_name=None)
    if os.path.exists(DEFAULT_REGION_PATH):
        return pd.read_excel(DEFAULT_REGION_PATH, sheet_name=None)
    return None


# =========================
# Branding header
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

# Per ADA docs / observed cert problem — use HTTP only
ENDPOINT = "http://ws.eatright.org/service/service.svc"
WSDL_URL = "http://ws.eatright.org/service/service.svc?wsdl"

def _wsdl_actions_map() -> dict:
    """Return {operationName: [soapAction, ...]} by parsing the WSDL (best-effort)."""
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
            # SOAP 1.2 (we won't post 1.2, but collect actions just in case)
            soap12_op = op.get("soap12:operation")
            if isinstance(soap12_op, dict) and "@soapAction" in soap12_op:
                actions.setdefault(name, []).append(soap12_op["@soapAction"])
    return actions

_WSDL_ACTIONS = _wsdl_actions_map()

def _wsdl_action_for(method: str) -> list[str]:
    return _WSDL_ACTIONS.get(method, [])


# =========================
# Shared processing helpers
# =========================
def _clean_zip(z):
    """Extract first 5 digits, return as 5-char string (drops +4 and non-digits)."""
    if z is None or (isinstance(z, float) and pd.isna(z)):
        return None
    s = str(z).strip()
    if not s:
        return None
    m = re.search(r"\b(\d{5})\b", s)
    return m.group(1) if m else None

# Accept any plausible ZIP field name that could come from API or CSV/XLSX
ZIP_CANDIDATE_COLS = [
    "Zip", "ZIP", "zip",
    "PostalCode", "Postal Code", "postal_code", "postal", "postal code",
    "ZipCode", "Zip Code", "zipcode", "ZIPCODE",
]

def _ensure_member_zip_column(members_df: pd.DataFrame) -> str | None:
    """
    Ensure members_df has a 'Zip_clean' column, derived from any plausible ZIP field.
    Returns the source column name if found, else None.
    """
    for c in ZIP_CANDIDATE_COLS:
        if c in members_df.columns:
            z = members_df[c].map(_clean_zip)
            if z.notna().any():
                members_df["Zip_clean"] = z.astype(str).str.zfill(5)
                return c
    members_df["Zip_clean"] = None  # still create column so downstream code runs
    return None


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
    base = API_NS  # "http://eatright/membership"
    candidates = list(dict.fromkeys(_wsdl_action_for(action)))  # dedupe & preserve order
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
        return str(result.get("a:Success", result.get("Success", "false"))).lower() == "true"
    except Exception:
        return False


def _find_members_anywhere(obj):
    """
    Recursively search a parsed xmltodict structure for a list/dict of members.
    We look for arrays whose items look like person records (have RecordNumber/LoginName/PostalCode/etc.).
    Returns a list[dict] (possibly empty).
    """
    CANDIDATE_KEYS = {"RecordNumber", "LoginName", "PostalCode", "Zip", "FirstName", "LastName", "Email"}
    out = []

    def is_member_dict(d):
        if not isinstance(d, dict):
            return False
        # strip namespace prefixes for key comparison
        keys = {k.split(":", 1)[-1] for k in d.keys()}
        return len(CANDIDATE_KEYS.intersection(keys)) >= 2  # at least two familiar fields

    def walk(x):
        nonlocal out
        if isinstance(x, list):
            # if this list already looks like members, keep items that are dict-like
            if x and all(isinstance(i, dict) for i in x) and any(is_member_dict(i) for i in x):
                out.extend([i for i in x if isinstance(i, dict)])
                return
            for i in x:
                walk(i)
        elif isinstance(x, dict):
            # if this dict itself looks like a member, capture it
            if is_member_dict(x):
                out.append(x)
                return
            for v in x.values():
                walk(v)

    walk(obj)
    return out


def fetch_members_via_api(access_key: str, group_key: str) -> pd.DataFrame:
    """RetrieveGroupMembersWithCustomProperties → pandas DataFrame (robust)."""
    body = f"""
    <RetrieveGroupMembersWithCustomProperties xmlns="{API_NS}">
      <groupKey>{group_key}</groupKey>
    </RetrieveGroupMembersWithCustomProperties>
    """.strip()
    env = _soap_envelope(body, access_key=access_key)

    # Get raw SOAP and save for debugging
    r = requests.post(
        ENDPOINT,
        data=env.encode("utf-8"),
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "Accept": "text/xml",
            "SOAPAction": f"\"{API_NS}/RetrieveGroupMembersWithCustomProperties\"",
        },
        timeout=60,
    )
    r.raise_for_status()
    raw_xml = r.text
    try:
        with open("/tmp/soap_response.xml", "w", encoding="utf-8") as f:
            f.write(raw_xml)
    except Exception:
        pass  # best-effort

    # Parse xml → dict
    data = xmltodict.parse(raw_xml)

    # Navigate to the ...Result node if present, then search recursively
    node = (data.get("s:Envelope") or data.get("Envelope") or {}).get("s:Body") or data.get("Body") or {}
    node = (node.get("RetrieveGroupMembersWithCustomPropertiesResponse")
            or node.get("RetrieveGroupMembersWithCustomPropertiesResult")
            or node)

    # In some shapes, the response puts Result under the Response object:
    if isinstance(node, dict) and "RetrieveGroupMembersWithCustomPropertiesResult" in node:
        node = node["RetrieveGroupMembersWithCustomPropertiesResult"]

    # Try our robust recursive finder
    rows = _find_members_anywhere(node)

    # As a fallback, also look for explicit Members/Member nesting
    if not rows:
        container = node
        for k in ("a:Members", "Members"):
            if isinstance(container, dict) and k in container:
                container = container[k]
        members = container.get("a:Member") if isinstance(container, dict) else None
        if not members and isinstance(container, dict):
            members = container.get("Member")
        if isinstance(members, dict):
            rows = [members]
        elif isinstance(members, list):
            rows = members

    # If still nothing, return empty DataFrame (the UI will show it clearly)
    if not rows:
        return pd.DataFrame()

    # Strip namespace prefixes in keys
    def strip_ns(d):
        return {k.split(":", 1)[-1]: v for k, v in d.items()} if isinstance(d, dict) else {}

    cleaned = [strip_ns(r) for r in rows]

    # Expand CustomProperties into columns, if present
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

    # Create a Zip column from any plausible field (we’ll normalize later)
    candidate_cols = [c for c in df.columns if c.lower() in ("zip", "postalcode", "postal_code", "zipcode")]
    df["Zip"] = df[candidate_cols[0]] if candidate_cols else None

    return df



def _debug_merge_preview(members_df: pd.DataFrame, region_sheets: dict):
    """
    Build a standardized region map, normalize member ZIPs, and merge — for debugging.
    Returns: merged, zip_map_all, source_zip_col
    """
    # 1) Normalize members ZIP
    members = members_df.copy()
    source_col = _ensure_member_zip_column(members)  # creates Zip_clean

    # 2) Standardize region map: expect County, Zip, Region in first 3 cols
    zip_map_all = pd.DataFrame()
    for _, df in region_sheets.items():
        d = df.iloc[:, :3].copy()
        d.columns = ["County", "Zip", "Region"]
        zip_map_all = pd.concat([zip_map_all, d], ignore_index=True)

    # 3) Clean region zips
    zip_map_all["Zip"] = zip_map_all["Zip"].map(_clean_zip)
    zip_map_all = zip_map_all[zip_map_all["Zip"].notna()].copy()
    zip_map_all["Zip"] = zip_map_all["Zip"].astype(str).str.zfill(5)
    zip_map_all = zip_map_all.drop_duplicates(subset=["Zip"], keep="first")

    # 4) Merge
    merged = pd.merge(members, zip_map_all, left_on="Zip_clean", right_on="Zip", how="left")
    return merged, zip_map_all, source_col


# =========================
# Core processing
# =========================
def process_and_package(members: pd.DataFrame, region_sheets: dict) -> bytes:
    """Return a bytes ZIP containing per-region workbooks + unmatched workbook."""
    # --- Normalize member ZIPs from any known column name ---
    members = members.copy()
    source_col = _ensure_member_zip_column(members)  # may be None if no ZIP-like column is present

    # --- Build one tall map from all sheets (expecting County, Zip, Region) ---
    zip_map_all = pd.DataFrame()
    for _, df in region_sheets.items():
        d = df.iloc[:, :3].copy()
        d.columns = ["County", "Zip", "Region"]
        zip_map_all = pd.concat([zip_map_all, d], ignore_index=True)

    # Robust ZIP normalization on the region map (handles numbers-as-text, +4, etc.)
    zip_map_all["Zip"] = zip_map_all["Zip"].map(_clean_zip)
    zip_map_all = zip_map_all[zip_map_all["Zip"].notna()].copy()
    zip_map_all["Zip"] = zip_map_all["Zip"].astype(str).str.zfill(5)
    zip_map_all = zip_map_all.drop_duplicates(subset=["Zip"], keep="first")

    # --- Merge + group ---
    merged = pd.merge(members, zip_map_all, left_on="Zip_clean", right_on="Zip", how="left")
    grouped = merged[merged["Region"].notna()].groupby("Region")

    # UI: quick match stats
    try:
        matched = merged["Region"].notna().sum()
        total = len(merged)
        st.info(
            f"Matched {matched} of {total} members to regions"
            + (f" using '{source_col}'" if source_col else " (no ZIP column detected)")
        )
    except Exception:
        pass

    # --- Write ZIP archive of region files + unmatched ---
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
    region_file = st.file_uploader(
        "📄 (Optional) Upload NYSAND Region Zipcodes Excel (overrides bundled)",
        type=["xls", "xlsx"],
        key="regionfile",
    )

    if member_file:
        region_sheets = load_region_mapping(region_file)
        if region_sheets is None:
            st.error("Region mapping not found. Please add assets/nysand_region_zips.xlsx to the repo or upload it above.")
        else:
            members_df = pd.read_csv(member_file)

            # --- Debug: Inspect uploaded data and ZIP matching ---
            with st.expander("🔎 Debug — Inspect uploaded data and ZIP matching", expanded=False):
                try:
                    st.write("**Uploaded columns:**", list(members_df.columns))
                    st.dataframe(members_df.head(10))

                    merged_dbg, zip_map_dbg, src_zip = _debug_merge_preview(members_df, region_sheets)
                    matched_ct = merged_dbg["Region"].notna().sum()
                    total_ct = len(merged_dbg)
                    st.info(
                        f"Matched **{matched_ct:,} / {total_ct:,}** members"
                        + (f" using ZIP column **{src_zip}**" if src_zip else " (no ZIP column detected)")
                    )

                    # Downloads
                    st.download_button(
                        "⬇️ Download Uploaded RAW (csv)",
                        members_df.to_csv(index=False).encode("utf-8"),
                        file_name="uploaded_raw.csv",
                    )
                    st.download_button(
                        "⬇️ Download Members + Zip_clean (csv)",
                        merged_dbg.drop(
                            columns=[c for c in ["County", "Region", "Zip_y"] if c in merged_dbg.columns],
                            errors="ignore",
                        ).to_csv(index=False).encode("utf-8"),
                        file_name="members_with_zip_clean.csv",
                    )
                    st.download_button(
                        "⬇️ Download Region Map (standardized) (csv)",
                        zip_map_dbg.to_csv(index=False).encode("utf-8"),
                        file_name="region_map_standardized.csv",
                    )
                    st.download_button(
                        "⬇️ Download MERGED (full) (csv)",
                        merged_dbg.to_csv(index=False).encode("utf-8"),
                        file_name="merged_full.csv",
                    )
                    st.download_button(
                        "⬇️ Download MATCHED only (csv)",
                        merged_dbg[merged_dbg["Region"].notna()].to_csv(index=False).encode("utf-8"),
                        file_name="merged_matched_only.csv",
                    )
                    st.download_button(
                        "⬇️ Download UNMATCHED only (csv)",
                        merged_dbg[merged_dbg["Region"].isna()].to_csv(index=False).encode("utf-8"),
                        file_name="merged_unmatched_only.csv",
                    )
                except Exception as e:
                    st.warning(f"Debug panel could not render: {e}")

            # --- Create outputs ---
            with st.spinner("Processing files..."):
                blob = process_and_package(members_df, region_sheets)
            st.success("✅ Done! Download your ZIP below.")
            st.download_button("📥 Download All Files (ZIP)", blob, file_name="NYSAND_Member_Files.zip")

elif source == "EatRight SOAP API":
    st.info("Uses secrets: EATR_ACCESS_KEY and EATR_GROUP_KEY")
    region_file = st.file_uploader(
        "📄 (Optional) Upload NYSAND Region Zipcodes Excel (overrides bundled)",
        type=["xls", "xlsx"],
        key="regionfile_api",
    )

    if st.button("🔄 Fetch members via API"):
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

            # --- DEBUG PANEL: inspect API data and merge behavior ---
            with st.expander("🔎 Debug — Inspect API data and ZIP matching", expanded=True):
                st.write("**API columns:**", list(api_df.columns))
                st.dataframe(api_df.head(10))

                merged_dbg, zip_map_dbg, src_zip = _debug_merge_preview(api_df, region_sheets)
                matched_ct = merged_dbg["Region"].notna().sum()
                total_ct = len(merged_dbg)
                st.info(
                    f"Matched **{matched_ct:,} / {total_ct:,}** members"
                    + (f" using ZIP column **{src_zip}**" if src_zip else " (no ZIP column detected)")
                )

                # Downloads
                st.download_button(
                    "⬇️ Download RAW API (csv)",
                    api_df.to_csv(index=False).encode("utf-8"),
                    file_name="api_raw.csv",
                )
                members_with_zip = merged_dbg.drop(
                    columns=[c for c in ["County", "Region", "Zip_y"] if c in merged_dbg.columns],
                    errors="ignore",
                )
                st.download_button(
                    "⬇️ Download Members + Zip_clean (csv)",
                    members_with_zip.to_csv(index=False).encode("utf-8"),
                    file_name="members_with_zip_clean.csv",
                )
                st.download_button(
                    "⬇️ Download Region Map (standardized) (csv)",
                    zip_map_dbg.to_csv(index=False).encode("utf-8"),
                    file_name="region_map_standardized.csv",
                )
                st.download_button(
                    "⬇️ Download MERGED (full) (csv)",
                    merged_dbg.to_csv(index=False).encode("utf-8"),
                    file_name="merged_full.csv",
                )
                st.download_button(
                    "⬇️ Download MATCHED only (csv)",
                    merged_dbg[merged_dbg["Region"].notna()].to_csv(index=False).encode("utf-8"),
                    file_name="merged_matched_only.csv",
                )
                st.download_button(
                    "⬇️ Download UNMATCHED only (csv)",
                    merged_dbg[merged_dbg["Region"].isna()].to_csv(index=False).encode("utf-8"),
                    file_name="merged_unmatched_only.csv",
                )

            # --- Create outputs ---
            with st.spinner("Creating region files…"):
                blob = process_and_package(api_df, region_sheets)
            today = datetime.now().strftime("%Y-%m-%d")
            st.success("✅ Done! Download your ZIP below.")
            st.download_button(
                "📥 Download All Files (ZIP)",
                blob,
                file_name=f"NYSAND_Member_Files_{today}.zip",
            )
        except KeyError as e:
            st.error(f"Missing secret: {e}. Please set EATR_ACCESS_KEY and EATR_GROUP_KEY.")
        except Exception as e:
            st.error(f"API fetch failed: {e}")
            st.exception(e)
