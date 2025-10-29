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

def load_region_mapping(uploaded_file):
    """
    Load the region zip mapping workbook. Try the user upload first (if provided),
    else fall back to the bundled DEFAULT_REGION_PATH. Return a dict of sheet->DataFrame
    with normalized columns: Zip (str), County (str), Region (str).
    """
    def _normalize_df(df: pd.DataFrame) -> pd.DataFrame:
        # Try to find zip-ish column
        cols_lower = {c.lower(): c for c in df.columns}
        zip_col = None
        for cand in ["zip", "zipcode", "zip_code", "postal", "postalcode"]:
            if cand in cols_lower:
                zip_col = cols_lower[cand]
                break
        if zip_col is None:
            # If first column looks numeric/zip-like, use it
            zip_col = df.columns[0]

        # Standardized output
        out = pd.DataFrame()
        out["Zip"] = df[zip_col].astype(str).str.replace(r"\D+", "", regex=True).str.zfill(5)

        # County/Region best-effort
        county_col = None
        region_col = None
        for cand in ["county", "region"]:
            if cand in cols_lower:
                if cand == "county":
                    county_col = cols_lower[cand]
                else:
                    region_col = cols_lower[cand]

        out["County"] = df[county_col].astype(str).str.strip() if county_col else ""
        out["Region"] = df[region_col].astype(str).str.strip() if region_col else ""

        # Remove empties
        out = out[out["Zip"].str.len() == 5]
        out = out.drop_duplicates(subset=["Zip"]).reset_index(drop=True)
        return out

    try:
        if uploaded_file is not None:
            xls = pd.ExcelFile(uploaded_file)
        else:
            # bundled file next to app (allow working dir or current path)
            path_try = DEFAULT_REGION_PATH
            if not os.path.exists(path_try):
                # try app-relative path
                here = os.path.dirname(__file__)
                path_try = os.path.join(here, DEFAULT_REGION_PATH)
            xls = pd.ExcelFile(path_try)
    except Exception:
        return None

    sheets = {}
    for name in xls.sheet_names:
        df = xls.parse(name)
        sheets[name] = _normalize_df(df)
    return sheets

# =========================
# Assets / logo helper
# =========================
def _find_logo():
    candidates = [
        "assets/nysand_logo.png",
        "assets/logo.png",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
        try:
            here = os.path.dirname(__file__)
            maybe = os.path.join(here, p)
            if os.path.exists(maybe):
                return maybe
        except Exception:
            pass
    return None

# =========================
# SOAP helpers
# =========================
def _wsdl_actions_map():
    # Minimal map for reference; you can extend as needed
    return {
        "ValidateAccessKey": "http://eatright.org/ValidateAccessKey",
        "GetMembers": "http://eatright.org/GetMembers",
    }

def _wsdl_action_for(op):
    return _wsdl_actions_map().get(op, "")

def _clean_zip(s):
    return re.sub(r"\D+", "", str(s or ""))[:5]

def _ensure_member_zip_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure there's a member ZIP column named Zip_clean for matching.
    Tries a variety of common column names (Zip, ZIP Code, Postal, etc.).
    """
    zcols = [c for c in df.columns if c.lower() in ("zip", "zipcode", "zip code", "postal", "postalcode", "home zip", "primary zip")]
    if not zcols:
        # keep as-is; downstream will show 0 matches
        df["Zip_clean"] = ""
        return df
    # pick best
    z = zcols[0]
    df["Zip_clean"] = df[z].map(_clean_zip).fillna("")
    return df

def _soap_envelope(access_key: str, group_key: str, include_custom: bool = True):
    # Basic SOAP envelope; customize as needed if API requires paging/filters
    custom_tag = f"<IncludeCustomProperties>{'true' if include_custom else 'false'}</IncludeCustomProperties>"
    return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <GetMembers xmlns="http://eatright.org/">
      <AccessKey>{access_key}</AccessKey>
      <GroupKey>{group_key}</GroupKey>
      {custom_tag}
    </GetMembers>
  </soap:Body>
</soap:Envelope>
"""

def _post_soap(endpoint_url: str, soap_action: str, body_xml: str) -> str:
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": soap_action,
    }
    resp = requests.post(endpoint_url, data=body_xml.encode("utf-8"), headers=headers, timeout=60)
    resp.raise_for_status()
    # Save last response for debug
    try:
        with open("/tmp/soap_response.xml", "w", encoding="utf-8") as f:
            f.write(resp.text)
    except Exception:
        pass
    return resp.text

def _validate_access_key(endpoint_url: str, access_key: str) -> bool:
    # Optionally call ValidateAccessKey if required by service
    try:
        envelope = f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <ValidateAccessKey xmlns="http://eatright.org/">
      <AccessKey>{access_key}</AccessKey>
    </ValidateAccessKey>
  </soap:Body>
</soap:Envelope>
"""
        xml = _post_soap(endpoint_url, _wsdl_action_for("ValidateAccessKey"), envelope)
        d = xmltodict.parse(xml)
        # Interpret success; if service returns boolean/status, adapt here
        return True if d else False
    except Exception:
        return False

def _find_members_anywhere(d: dict):
    """
    Walk a parsed SOAP dict to find the members array, tolerating nesting.
    Returns a list[dict] of member objects (best-effort).
    """
    out = []

    def is_member_dict(candidate):
        # Heuristic: member-ish dict has name/zip/email fields (very loose)
        if not isinstance(candidate, dict):
            return False
        keys = {k.lower() for k in candidate.keys()}
        signals = ["name", "firstname", "lastname", "zip", "zipcode", "email"]
        return any(s in keys for s in signals)

    def walk(node):
        if isinstance(node, dict):
            # If looks like a member, capture it
            if is_member_dict(node):
                out.append(node)
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)

    walk(d)
    return out

def is_member_dict(candidate):
    return isinstance(candidate, dict)

def fetch_members_via_api(access_key: str, group_key: str, include_custom_props: bool = True) -> pd.DataFrame:
    """
    Call the EatRight SOAP API, parse xml to dict, find a members array, and
    return a DataFrame. Also add Zip_clean for downstream matching.
    """
    endpoint = st.secrets.get("EATR_ENDPOINT_URL", "https://reports.eatright.org/MemberServices/MemberService.asmx")
    # Optionally validate key (no-op if service doesn't require)
    _ = _validate_access_key(endpoint, access_key)

    envelope = _soap_envelope(access_key, group_key, include_custom_props)
    xml = _post_soap(endpoint, _wsdl_action_for("GetMembers"), envelope)

    try:
        parsed = xmltodict.parse(xml)
    except Exception as e:
        raise RuntimeError(f"SOAP XML parse error: {e}") from e

    members = _find_members_anywhere(parsed)
    if not members:
        # Try to find any table-ish structure
        # If service returns a single dict, wrap it
        if isinstance(parsed, dict):
            members = [parsed]
    # Normalize to DataFrame
    df = pd.json_normalize(members, max_level=3)
    df = df.drop_duplicates().reset_index(drop=True)

    # Ensure zip-clean column for matching
    df = _ensure_member_zip_column(df)
    return df

def strip_ns(colname: str) -> str:
    # Strip XML namespaces from column names for nicer display
    return re.sub(r"(^.*:)", "", colname or "")

def props_to_dict(df: pd.DataFrame) -> pd.DataFrame:
    """
    If any columns are nested custom properties dicts, flatten them.
    Best-effort; safe no-op if already flat.
    """
    # Discover dict-like columns
    dict_cols = [c for c in df.columns if df[c].apply(lambda v: isinstance(v, dict)).any()]
    out = df.copy()
    for c in dict_cols:
        expanded = pd.json_normalize(out[c]).add_prefix(strip_ns(c) + ".")
        out = pd.concat([out.drop(columns=[c]), expanded], axis=1)
    return out

# =========================
# Debug merge & packaging
# =========================
def _debug_merge_preview(members_df: pd.DataFrame, region_sheets: dict):
    """
    Merge members_df with region_sheets by ZIP for a quick preview.
    Returns (merged, zip_map_concat, src_zip_colname).
    """
    zip_map = []
    for _, zdf in region_sheets.items():
        # keep only Zip, County, Region
        zdf2 = zdf[["Zip", "County", "Region"]].copy()
        zip_map.append(zdf2)
    region_map = pd.concat(zip_map, ignore_index=True).drop_duplicates(subset=["Zip"])

    # detect source zip
    src_zip = "Zip_clean" if "Zip_clean" in members_df.columns else None
    merged = members_df.merge(region_map, left_on=src_zip, right_on="Zip", how="left") if src_zip else members_df.copy()
    return merged, region_map, src_zip

def process_and_package(members_df: pd.DataFrame, region_sheets: dict) -> bytes:
    """
    Create one CSV per Region, plus summary sheets, and return a ZIP blob.
    """
    merged, region_map, src_zip = _debug_merge_preview(members_df, region_sheets)

    # Prepare per-Region member CSVs
    if "Region" in merged.columns:
        regions = sorted(merged["Region"].dropna().unique().tolist())
    else:
        regions = []

    # Build ZIP in memory
    with tempfile.TemporaryDirectory() as tmpdir:
        zippath = os.path.join(tmpdir, "NYSAND_Member_Files.zip")
        with zipfile.ZipFile(zippath, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            # Per-region files
            for r in regions:
                r_df = merged[merged["Region"] == r].copy()
                safe_r = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(r))
                csv_path = os.path.join(tmpdir, f"members_{safe_r}.csv")
                r_df.to_csv(csv_path, index=False)
                zf.write(csv_path, arcname=f"regions/members_{safe_r}.csv")

            # Summary: merged, region_map, unmatched
            merged_path = os.path.join(tmpdir, "merged_full.csv")
            merged.to_csv(merged_path, index=False)
            zf.write(merged_path, arcname="debug/merged_full.csv")

            map_path = os.path.join(tmpdir, "region_map_standardized.csv")
            region_map.to_csv(map_path, index=False)
            zf.write(map_path, arcname="debug/region_map_standardized.csv")

            if "Region" in merged.columns:
                unmatched = merged[merged["Region"].isna()].copy()
            else:
                unmatched = merged.copy()
            um_path = os.path.join(tmpdir, "unmatched.csv")
            unmatched.to_csv(um_path, index=False)
            zf.write(um_path, arcname="debug/unmatched.csv")

            # Optional: include a logo if present
            logo_path = _find_logo()
            if logo_path:
                zf.write(logo_path, arcname="assets/logo.png")

        with open(zippath, "rb") as f:
            blob = f.read()
    return blob

# =========================
# UI
# =========================
st.title("NYSAND Member Region Files")
st.caption(
    "Upload your **Member Export CSV** and the **NYSAND Region Zip...*, **or** fetch the member list via the EatRight SOAP API, then:"
)
st.markdown(
    "- Inspect the merged data in **Debug**\n"
    "- Download per-region CSVs in a single ZIP\n"
    "- Download helpful intermediate CSVs"
)

with st.sidebar:
    source = st.radio("Data source", ["Manual Upload", "EatRight SOAP API"], index=0)

if source == "Manual Upload":
    st.subheader("Manual Upload")
    members_csv = st.file_uploader("📥 Upload Member Export (CSV)", type=["csv"])
    region_file = st.file_uploader(
        "📄 (Optional) Upload NYSAND Region Zipcodes Excel (overrides bundled)",
        type=["xls", "xlsx"],
    )
    if st.button("Process uploaded files", type="primary") and members_csv:
        try:
            members_df = pd.read_csv(members_csv)
            members_df = _ensure_member_zip_column(members_df)
            region_sheets = load_region_mapping(region_file)
            if region_sheets is None:
                st.error("Region mapping not found. Upload it above or add assets/nysand_region_zips.xlsx to the repo.")
                st.stop()

            merged_dbg, zip_map_dbg, src_zip = _debug_merge_preview(members_df, region_sheets)

            with st.expander("🔎 Debug — Inspect upload & ZIP matching", expanded=True):
                try:
                    st.write("**Member columns:**", list(members_df.columns))
                    st.dataframe(members_df.head(10))
                    matched_ct = merged_dbg["Region"].notna().sum() if "Region" in merged_dbg.columns else 0
                    total_ct = len(merged_dbg)
                    st.info(
                        f"Matched **{matched_ct:,} / {total_ct:,}** members"
                        + (f" using ZIP column **{src_zip}**" if src_zip else "")
                    )

                    # Diagnostics / raw downloads
                    st.download_button(
                        "⬇️ Download Members RAW (csv)",
                        members_df.to_csv(index=False).encode("utf-8"),
                        file_name="members_raw.csv",
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
                    if "Region" in merged_dbg.columns:
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
            
        except Exception as e:
            st.error(f"Manual upload failed: {e}")
            

elif source == "EatRight SOAP API":
    st.info("Uses secrets: EATR_ACCESS_KEY and EATR_GROUP_KEY")

    # ---------- Session persistence keys ----------
    _STATE_KEYS = [
        "api_df",          # pd.DataFrame of raw members
        "merged_dbg",      # pd.DataFrame merged with regions (for debug)
        "zip_map_dbg",     # pd.DataFrame standardized region map
        "src_zip",         # str of the detected ZIP column
        "region_blob",     # bytes of the generated ZIP file
        "fetched_on",      # timestamp string
    ]
    for k in _STATE_KEYS:
        st.session_state.setdefault(k, None)

    # Optional region map override
    region_file = st.file_uploader(
        "📄 (Optional) Upload NYSAND Region Zipcodes Excel (overrides bundled)",
        type=["xls", "xlsx"],
        key="regionfile_api",
    )

    cols = st.columns([1,1,1])
    with cols[0]:
        do_fetch = st.button("🔄 Fetch members via API", type="primary")
    with cols[1]:
        do_reset = st.button("🧹 Reset session")

    if do_reset:
        for k in _STATE_KEYS:
            st.session_state[k] = None
        st.success("Session cleared. You can fetch again.")
        st.stop()

    # ---------- On-demand fetch (write into session_state) ----------
    if do_fetch:
        try:
            ak = st.secrets["EATR_ACCESS_KEY"].strip()
            gk = st.secrets["EATR_GROUP_KEY"].strip()

            use_custom = st.checkbox(
                "Include custom properties (slower, richer)",
                value=True,
                key="api_custom_props"
            )

            with st.spinner("Fetching members from EatRight API…"):
                api_df = fetch_members_via_api(ak, gk, include_custom_props=use_custom)

            region_sheets = load_region_mapping(region_file)
            if region_sheets is None:
                st.error("Region mapping not found. Upload it above or add assets/nysand_region_zips.xlsx to the repo.")
                st.stop()

            merged_dbg, zip_map_dbg, src_zip = _debug_merge_preview(api_df, region_sheets)

            with st.spinner("Creating region files…"):
                region_blob = process_and_package(api_df, region_sheets)

            # Persist to session so downloads don't disappear on rerun
            st.session_state["api_df"] = api_df
            st.session_state["merged_dbg"] = merged_dbg
            st.session_state["zip_map_dbg"] = zip_map_dbg
            st.session_state["src_zip"] = src_zip
            st.session_state["region_blob"] = region_blob
            st.session_state["fetched_on"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            st.success(f"Fetched {len(api_df):,} records and built region files.")
        except KeyError as e:
            st.error(f"Missing secret: {e}. Please set EATR_ACCESS_KEY and EATR_GROUP_KEY.")
        except Exception as e:
            st.error(f"API fetch failed: {e}")
            st.exception(e)

    # ---------- Render from session_state on every rerun ----------
    if st.session_state["api_df"] is not None:
        api_df = st.session_state["api_df"]
        merged_dbg = st.session_state["merged_dbg"]
        zip_map_dbg = st.session_state["zip_map_dbg"]
        src_zip = st.session_state["src_zip"]
        region_blob = st.session_state["region_blob"]
        fetched_on = st.session_state["fetched_on"]

        st.info(f"Session has {len(api_df):,} members. Last fetched: {fetched_on}")

        st.dataframe(api_df.head(25))

        # Optional: Last SOAP XML (if your fetcher wrote it)
        try:
            with open("/tmp/soap_response.xml", "r", encoding="utf-8") as f:
                raw_xml = f.read()
            st.download_button(
                "⬇️ Download last SOAP response (xml)",
                raw_xml.encode("utf-8"),
                file_name="soap_response.xml",
                key="dl_last_soap_xml_ok",
            )
        except Exception:
            pass

        # ---- Debug expander
        with st.expander("🔎 Debug — Inspect API data and ZIP matching", expanded=True):
            st.write("**API columns:**", list(api_df.columns))
            st.dataframe(api_df.head(10))

            matched_ct = merged_dbg["Region"].notna().sum() if "Region" in merged_dbg.columns else 0
            total_ct = len(merged_dbg)
            st.info(
                f"Matched **{matched_ct:,} / {total_ct:,}** members"
                + (f" using ZIP column **{src_zip}**" if src_zip else "")
            )

            # Diagnostics / raw downloads (persist across reruns)
            st.download_button(
                "⬇️ Download RAW API (csv)",
                api_df.to_csv(index=False).encode("utf-8"),
                file_name="api_raw.csv",
                key="dl_api_raw",
            )

            members_with_zip = merged_dbg.drop(
                columns=[c for c in ["County", "Region", "Zip_y"] if c in merged_dbg.columns],
                errors="ignore",
            )
            st.download_button(
                "⬇️ Download Members + Zip_clean (csv)",
                members_with_zip.to_csv(index=False).encode("utf-8"),
                file_name="members_with_zip_clean.csv",
                key="dl_zip_clean",
            )
            st.download_button(
                "⬇️ Download Region Map (standardized) (csv)",
                zip_map_dbg.to_csv(index=False).encode("utf-8"),
                file_name="region_map_standardized.csv",
                key="dl_region_map",
            )
            st.download_button(
                "⬇️ Download MERGED (full) (csv)",
                merged_dbg.to_csv(index=False).encode("utf-8"),
                file_name="merged_full.csv",
                key="dl_merged_full",
            )
            st.download_button(
                "⬇️ Download MATCHED only (csv)",
                merged_dbg[merged_dbg["Region"].notna()].to_csv(index=False).encode("utf-8"),
                file_name="merged_matched_only.csv",
                key="dl_matched",
            )
            st.download_button(
                "⬇️ Download UNMATCHED only (csv)",
                merged_dbg[merged_dbg["Region"].isna()].to_csv(index=False).encode("utf-8"),
                file_name="merged_unmatched_only.csv",
                key="dl_unmatched",
            )

        # ---- Region ZIP download (persists)
        today = datetime.now().strftime("%Y-%m-%d")
        st.download_button(
            "📥 Download All Files (ZIP)",
            region_blob,
            file_name=f"NYSAND_Member_Files_{today}.zip",
            key="dl_region_zip",
        )
