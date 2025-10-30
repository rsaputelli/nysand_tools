# tools/generate_region_zip.py
# Headless NYSAND region ZIP generator for CI / schedulers

from urllib.parse import urlparse, urlunparse
from datetime import datetime
import argparse
import requests
import xmltodict
import pandas as pd
import re, io, zipfile, tempfile, sys, os

# ---------- Config & helpers ----------
DEFAULT_REGION_ASSET = "assets/nysand_region_zips.xlsx"

def force_http(url: str | None) -> str:
    """
    Coerce any HTTPS endpoint to HTTP and normalize known ADA endpoints to HTTP.
    """
    if not url:
        return "http://reports.eatright.org/MemberServices/MemberService.asmx"
    u = url.strip()
    p = urlparse(u)
    if p.scheme.lower() == "https":
        p = p._replace(scheme="http")
        u = urlunparse(p)
    if "reports.eatright.org/MemberServices/MemberService.asmx" in u:
        return "http://reports.eatright.org/MemberServices/MemberService.asmx"
    if "ws.eatright.org/service/service.svc" in u:
        return "http://ws.eatright.org/service/service.svc"
    return u

def _post_http_no_https_redirect(url: str, data: bytes, headers: dict, timeout: int = 90) -> requests.Response:
    """
    POST without following redirects. If server tries to redirect to HTTPS,
    rewrite back to HTTP and retry once.
    """
    r = requests.post(url, data=data, headers=headers, timeout=timeout, allow_redirects=False)
    if r.is_redirect or r.status_code in (301, 302, 303, 307, 308):
        loc = r.headers.get("Location", "")
        if loc:
            p = urlparse(loc)
            if p.scheme.lower() == "https":
                p = p._replace(scheme="http")
                http_loc = urlunparse(p)
                r2 = requests.post(http_loc, data=data, headers=headers, timeout=timeout, allow_redirects=False)
                return r2
    return r

def _clean_zip(s):
    return re.sub(r"\D+", "", str(s or ""))[:5]

def ensure_zip_col(df: pd.DataFrame) -> pd.DataFrame:
    zcols = [c for c in df.columns if c.lower() in
             ("zip","zipcode","zip code","postal","postalcode","home zip","primary zip")]
    if not zcols:
        df["Zip_clean"] = ""
        return df
    z = zcols[0]
    df["Zip_clean"] = df[z].map(_clean_zip).fillna("")
    return df

def load_region_mapping(region_xlsx_path: str | None) -> dict | None:
    """Return dict of DataFrames keyed by sheet name, or None."""
    cand = None
    if region_xlsx_path and os.path.exists(region_xlsx_path):
        cand = region_xlsx_path
    elif os.path.exists(DEFAULT_REGION_ASSET):
        cand = DEFAULT_REGION_ASSET
    elif os.path.exists("nysand_region_zips.xlsx"):
        cand = "nysand_region_zips.xlsx"
    if not cand:
        return None
    return pd.read_excel(cand, sheet_name=None)

def region_concat_map(sheets: dict) -> pd.DataFrame:
    acc = []
    for _, df in sheets.items():
        cols = {c.lower(): c for c in df.columns}
        zip_col = cols.get("zip") or cols.get("zipcode") or list(df.columns)[0]
        out = pd.DataFrame()
        out["Zip"] = df[zip_col].astype(str).str.replace(r"\D+","", regex=True).str.zfill(5)
        out["County"] = df[cols.get("county")].astype(str).str.strip() if "county" in cols else ""
        out["Region"] = df[cols.get("region")].astype(str).str.strip() if "region" in cols else ""
        out = out[out["Zip"].str.len()==5]
        acc.append(out[["Zip","County","Region"]])
    m = pd.concat(acc, ignore_index=True)
    return m.drop_duplicates(subset=["Zip"]).reset_index(drop=True)

def soap_envelope_xmlns(access_key: str, group_key: str, include_custom=True, xmlns="http://eatright/membership") -> str:
    """
    Build SOAP 1.1 envelope with a selectable XML namespace for the GetMembers op.
    Some hosts expect http://eatright/membership, others http://eatright.org/.
    """
    return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Body>
    <GetMembers xmlns="{xmlns}">
      <AccessKey>{access_key}</AccessKey>
      <GroupKey>{group_key}</GroupKey>
      <IncludeCustomProperties>{'true' if include_custom else 'false'}</IncludeCustomProperties>
    </GetMembers>
  </soap:Body>
</soap:Envelope>
"""

def _try_post(endpoint: str, access_key: str, group_key: str, include_custom: bool) -> requests.Response:
    """
    Try both SOAPAction/namespace combos against the given endpoint:
      1) SOAPAction: http://eatright/membership/GetMembers  (xmlns=http://eatright/membership)
      2) SOAPAction: http://eatright.org/GetMembers        (xmlns=http://eatright.org/)
    Return the first 2xx response; otherwise return the last response for error handling.
    """
    combos = [
        ("http://eatright/membership/GetMembers", "http://eatright/membership"),
        ("http://eatright.org/GetMembers",        "http://eatright.org/"),
    ]
    last = None
    for action, ns in combos:
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": action,
        }
        body = soap_envelope_xmlns(access_key, group_key, include_custom, xmlns=ns).encode("utf-8")
        resp = _post_http_no_https_redirect(endpoint, data=body, headers=headers, timeout=90)
        print(f"[{datetime.utcnow().isoformat()}Z] Tried SOAPAction={action} ns={ns} → status={resp.status_code}")
        last = resp
        if 200 <= resp.status_code < 300:
            return resp
    return last

def fetch_members(endpoint: str, access_key: str, group_key: str, include_custom=True) -> pd.DataFrame:
    endpoint = force_http(endpoint)
    print(f"[{datetime.utcnow().isoformat()}Z] Using endpoint scheme={urlparse(endpoint).scheme}")

    # 1) Try given endpoint with both SOAPAction variants
    resp = _try_post(endpoint, access_key, group_key, include_custom)

    # 2) If not OK, try alternate host with both variants (HTTP only)
    if not (200 <= resp.status_code < 300):
        alt = "http://ws.eatright.org/service/service.svc" \
              if "reports.eatright.org" in endpoint else \
              "http://reports.eatright.org/MemberServices/MemberService.asmx"
        print(f"[{datetime.utcnow().isoformat()}Z] Falling back to alternate host: {alt}")
        resp = _try_post(alt, access_key, group_key, include_custom)

    resp.raise_for_status()
    parsed = xmltodict.parse(resp.text)

    # best-effort crawl to find member dicts
    out = []
    def is_member(d):
        if not isinstance(d, dict): return False
        k = {k.lower() for k in d.keys()}
        return any(t in k for t in ("name","firstname","lastname","zip","zipcode","email"))
    def walk(n):
        if isinstance(n, dict):
            if is_member(n): out.append(n)
            for v in n.values(): walk(v)
        elif isinstance(n, list):
            for v in n: walk(v)
    walk(parsed)
    if not out:
        out = [parsed]
    df = pd.json_normalize(out, max_level=3).drop_duplicates().reset_index(drop=True)
    return ensure_zip_col(df)

def build_zip_blob(members_df: pd.DataFrame, region_sheets: dict) -> bytes:
    region_map = region_concat_map(region_sheets)
    merged = members_df.merge(region_map, left_on="Zip_clean", right_on="Zip", how="left")
    regions = sorted(merged["Region"].dropna().unique().tolist()) if "Region" in merged.columns else []

    with tempfile.TemporaryDirectory() as td:
        zip_path = os.path.join(td, "NYSAND_Member_Files.zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            # per-region
            for r in regions:
                r_df = merged[merged["Region"] == r].copy()
                safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(r))
                p = os.path.join(td, f"members_{safe}.csv")
                r_df.to_csv(p, index=False)
                zf.write(p, arcname=f"regions/members_{safe}.csv")
            # debug
            p = os.path.join(td, "debug_merged_full.csv"); merged.to_csv(p, index=False); zf.write(p, arcname="debug/merged_full.csv")
            p = os.path.join(td, "debug_region_map.csv"); region_map.to_csv(p, index=False); zf.write(p, arcname="debug/region_map_standardized.csv")
            p = os.path.join(td, "debug_unmatched.csv"); merged[merged["Region"].isna()].to_csv(p, index=False); zf.write(p, arcname="debug/unmatched.csv")
        with open(zip_path, "rb") as f:
            return f.read()

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Generate NYSAND per-region ZIP (headless)")
    ap.add_argument("--region-xlsx", help="Path to region zipcodes Excel (optional; defaults to assets)")
    ap.add_argument("--out", default="out/NYSAND_Member_Files.zip", help="Output ZIP path")
    ap.add_argument("--include-custom", action="store_true", help="Include custom properties (slower, richer)")
    args = ap.parse_args()

    endpoint = os.environ.get("EATR_ENDPOINT_URL") or "http://reports.eatright.org/MemberServices/MemberService.asmx"
    access  = os.environ.get("EATR_ACCESS_KEY")
    group   = os.environ.get("EATR_GROUP_KEY")
    if not access or not group:
        print("ERROR: EATR_ACCESS_KEY and EATR_GROUP_KEY must be set in env.", file=sys.stderr)
        return 2

    try:
        print(f"[{datetime.utcnow().isoformat()}Z] Fetching members… endpoint={endpoint}")
        members = fetch_members(endpoint, access, group, include_custom=args.include_custom)
        print(f"Fetched {len(members):,} members")

        region_sheets = load_region_mapping(args.region_xlsx)
        if region_sheets is None:
            print("ERROR: Region mapping Excel not found (assets/nysand_region_zips.xlsx or provided path).", file=sys.stderr)
            return 3

        blob = build_zip_blob(members, region_sheets)
        os.makedirs(os.path.dirname(args.out), exist_ok=True)
        with open(args.out, "wb") as f:
            f.write(blob)
        print(f"Wrote ZIP → {args.out}")
        return 0
    except requests.HTTPError as e:
        print(f"HTTPError: {e}", file=sys.stderr)
        return 10
    except Exception as e:
        print(f"Unhandled error: {e}", file=sys.stderr)
        return 99

if __name__ == "__main__":
    sys.exit(main())
