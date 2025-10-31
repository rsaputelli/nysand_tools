import os, sys, re, zipfile, tempfile, base64, json
from datetime import datetime
import pandas as pd
import requests
import xmltodict
import argparse  # ok to import; just don't PARSE at import time

# ---------- Config / Inputs ----------

def _endpoint() -> str:
    """Resolve endpoint with environment or fallback to default WCF service."""
    return os.getenv("EATR_ENDPOINT_URL", "http://ws.eatright.org/service/service.svc").strip()

ACCESS_KEY   = os.getenv("EATR_ACCESS_KEY", "").strip()
GROUP_KEY    = os.getenv("EATR_GROUP_KEY", "").strip()
ENDPOINT     = _endpoint()
WSDL_URL     = ENDPOINT + "?wsdl" if "?" not in ENDPOINT else ENDPOINT
REGION_XLSX  = os.getenv("REGION_ZIPS_PATH", "assets/nysand_region_zips.xlsx")
CSV_FALLBACK = os.getenv("API_CSV_FALLBACK", "assets/api_seed.csv")
CSV_B64_ENV  = os.getenv("API_CSV_BASE64", "")

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
API_NS  = "http://eatright/membership"

# --- diagnostics ---
print(f"[{datetime.utcnow().isoformat()}Z] Using endpoint → {ENDPOINT}")

def _wsdl_actions_map() -> dict:
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

def _soap_envelope(body_xml: str, access_key: str | None) -> str:
    header = (
        f"""
        <s:Header>
          <AccessKey xmlns="{API_NS}" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">
            <Value>{access_key}</Value>
          </AccessKey>
        </s:Header>
        """ if access_key else "<s:Header/>"
    )
    return f"""<s:Envelope xmlns:s="{SOAP_NS}">
{header}
  <s:Body>
{body_xml}
  </s:Body>
</s:Envelope>""".strip()

def _post_soap(action: str, envelope_xml: str) -> dict:
    """Send SOAP request trying membership and org namespaces (WCF tolerant)."""
    candidates = [
        f"http://eatright/membership/{action}",
        f"http://eatright/membership/IService/{action}",
        f"http://eatright/membership/IWcfAdaMembership/{action}",
        f"http://eatright.org/membership/{action}",             # added
        f"http://eatright.org/membership/IService/{action}",    # added
        f"http://eatright.org/membership/IWcfAdaMembership/{action}",  # added
        ""  # sometimes empty works
    ]

    last_body, errors = "", []
    for sa in candidates:
        headers = {
            "Content-Type": "text/xml; charset=utf-8",
            "Accept": "text/xml",
            "SOAPAction": f"\"{sa}\""
        }
        try:
            r = requests.post(ENDPOINT, data=envelope_xml.encode("utf-8"), headers=headers, timeout=60)
            last_body = r.text or ""
            with open("/tmp/soap_response.xml", "w", encoding="utf-8") as f:
                f.write(last_body)
            if r.status_code >= 400:
                errors.append(f"{sa} → HTTP {r.status_code}")
                continue
            return xmltodict.parse(last_body)
        except Exception as e:
            errors.append(f"{sa} → {e}")

    if last_body:
        with open("/tmp/soap_response.xml", "w", encoding="utf-8") as f:
            f.write(last_body)
    raise RuntimeError("All SOAP attempts failed: " + "; ".join(errors))


def _find_members_anywhere(obj):
    CANDIDATE_KEYS = {"RecordNumber", "LoginName", "PostalCode", "Zip", "FirstName", "LastName", "Email"}
    out = []
    def is_member_dict(d):
        if not isinstance(d, dict): return False
        keys = {k.split(":", 1)[-1] for k in d.keys()}
        return len(CANDIDATE_KEYS.intersection(keys)) >= 2
    def walk(x):
        nonlocal out
        if isinstance(x, list):
            if x and all(isinstance(i, dict) for i in x) and any(is_member_dict(i) for i in x):
                out.extend([i for i in x if isinstance(i, dict)])
                return
            for i in x: walk(i)
        elif isinstance(x, dict):
            if is_member_dict(x):
                out.append(x); return
            for v in x.values(): walk(v)
    walk(obj)
    return out

def _clean_zip(z):
    if z is None or (isinstance(z, float) and pd.isna(z)): return None
    s = str(z).strip()
    if not s: return None
    m = re.search(r"\b(\d{5})\b", s)
    return m.group(1) if m else None

def _ensure_member_zip_column(df: pd.DataFrame) -> None:
    cand = [c for c in df.columns if c.lower() in ("zip", "zipcode", "postalcode", "postal_code")]
    df["Zip"] = df[cand[0]] if cand else None

def fetch_members_via_api(include_custom_props: bool = True) -> pd.DataFrame:
    method = "RetrieveGroupMembersWithCustomProperties" if include_custom_props else "RetrieveGroupMembers"
    body = f"""
    <{method} xmlns="{API_NS}">
      <groupKey>{GROUP_KEY}</groupKey>
    </{method}>
    """.strip()
    env = _soap_envelope(body, access_key=ACCESS_KEY)
    data = _post_soap(method, env)

    body_node = (data.get("s:Envelope") or data.get("Envelope") or {}).get("s:Body") or data.get("Body") or {}
    resp_key, res_key = f"{method}Response", f"{method}Result"
    node = body_node.get(resp_key) or body_node.get(res_key) or body_node
    if isinstance(node, dict) and res_key in node:
        node = node[res_key]

    rows = _find_members_anywhere(node)
    if not rows and isinstance(node, dict):
        container = node
        for k in ("a:Members", "Members"):
            if k in container:
                container = container[k]
        members = None
        if isinstance(container, dict):
            members = container.get("a:Member") or container.get("Member")
        rows = [members] if isinstance(members, dict) else (members or [])
    if not rows:
        return pd.DataFrame()

    def strip_ns(d):
        return {k.split(":", 1)[-1]: v for k, v in d.items()} if isinstance(d, dict) else {}
    df = pd.DataFrame([strip_ns(r) for r in rows])

    if "CustomProperties" in df.columns:
        def props_to_dict(v):
            if not isinstance(v, dict): return {}
            items = v.get("a:CustomProperty") or v.get("CustomProperty") or []
            if isinstance(items, dict): items = [items]
            out = {}
            for it in items:
                if not isinstance(it, dict): continue
                name = it.get("a:Name") or it.get("Name")
                val  = it.get("a:Value") or it.get("Value")
                if name: out[str(name)] = val
            return out
        props = df["CustomProperties"].apply(props_to_dict).apply(pd.Series)
        df = pd.concat([df.drop(columns=["CustomProperties"]), props], axis=1)

    _ensure_member_zip_column(df)
    return df

def _load_region_map(xlsx_path: str) -> pd.DataFrame:
    sheets = pd.read_excel(xlsx_path, sheet_name=None)
    z = pd.DataFrame()
    for _, d in sheets.items():
        t = d.iloc[:, :3].copy()
        t.columns = ["County", "Zip", "Region"]
        z = pd.concat([z, t], ignore_index=True)
    z["Zip"] = z["Zip"].map(_clean_zip)
    z = z[z["Zip"].notna()].copy()
    z["Zip"] = z["Zip"].astype(str).str.zfill(5)
    z = z.drop_duplicates(subset=["Zip"], keep="first")
    return z

def _group_and_zip(members_df: pd.DataFrame, region_map: pd.DataFrame) -> bytes:
    members = members_df.copy()
    members["Zip_clean"] = members["Zip"].map(_clean_zip).astype("string")
    members.loc[members["Zip_clean"].notna(), "Zip_clean"] = members.loc[members["Zip_clean"].notna(), "Zip_clean"].str.zfill(5)
    merged = pd.merge(members, region_map, left_on="Zip_clean", right_on="Zip", how="left")

    with tempfile.TemporaryDirectory() as td:
        out_zip = os.path.join(td, "NYSAND_Member_Files.zip")
        with zipfile.ZipFile(out_zip, "w") as zf:
            # per-region
            grp = merged[merged["Region"].notna()].groupby("Region")
            for region, df in grp:
                safe = str(region).replace("/", "-").replace(" ", "_")
                f = os.path.join(td, f"{safe}_Members.xlsx")
                df.to_excel(f, index=False)
                zf.write(f, arcname=os.path.basename(f))
            # unmatched
            unmatched = merged[merged["Region"].isna()]
            f = os.path.join(td, "Unmatched_OutOfState_Members.xlsx")
            unmatched.to_excel(f, index=False)
            zf.write(f, arcname="Unmatched_OutOfState_Members.xlsx")
        return open(out_zip, "rb").read()

def _load_fallback_csv() -> pd.DataFrame:
    if CSV_B64_ENV:
        try:
            raw = base64.b64decode(CSV_B64_ENV)
            return pd.read_csv(pd.io.common.BytesIO(raw))
        except Exception:
            pass
    if CSV_FALLBACK and os.path.exists(CSV_FALLBACK):
        return pd.read_csv(CSV_FALLBACK)
    raise RuntimeError("No API_CSV_BASE64 or assets/api_seed.csv found for fallback")

def run_build(out_path: str):
    """Main build: fetch members, load region map, zip outputs, write to out_path."""
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    print(f"[{datetime.utcnow().isoformat()}Z] Fetching members…")
    if ENDPOINT.lower().startswith("http://"):
        print(f"[{datetime.utcnow().isoformat()}Z] Using endpoint scheme=http")

    try:
        members = fetch_members_via_api(include_custom_props=True)
        if members.empty:
            raise RuntimeError("API returned 0 rows (empty).")
        print(f"[{datetime.utcnow().isoformat()}Z] API returned {len(members)} rows")
    except Exception as e:
        print(f"[{datetime.utcnow().isoformat()}Z] API fetch failed: {e}", file=sys.stderr)
        print(f"[{datetime.utcnow().isoformat()}Z] Falling back to CSV…", file=sys.stderr)
        members = _load_fallback_csv()
        print(f"[{datetime.utcnow().isoformat()}Z] Fallback CSV rows: {len(members)}")

    region = _load_region_map(REGION_XLSX)
    blob = _group_and_zip(members, region)
    with open(out_path, "wb") as f:
        f.write(blob)
    print(f"[{datetime.utcnow().isoformat()}Z] Wrote ZIP → {out_path}")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, help="output ZIP path")
    args = ap.parse_args()
    run_build(args.out)
