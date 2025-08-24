# app.py — TPD Draft Generator (industry-aware + roll-forward vs rewrite + DOCX formatting + .DOC conversion)
from __future__ import annotations
import io, re, json, os, subprocess, tempfile
from typing import Dict, Any, List, Optional, Tuple

import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup

# Optional deps
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from docx import Document as DocxDocument
    from docx.text.paragraph import Paragraph
except Exception:
    DocxDocument = None  # guarded below

st.set_page_config(page_title="TPD Draft Generator", layout="wide")
st.sidebar.title("TPA (Transfer Pricing Associate)")
st.sidebar.caption("Roll-forward TPD • Industry-aware • Formatting preserved")

page = st.sidebar.radio(
    "Choose function",
    ["TPD Draft", "TNMM Review", "CUT/CUP Review", "Information Request List", "Advisory / Opportunity Spotting"],
)

# ==========================
# Helpers: PDF / DOCX text
# ==========================
def read_pdf(file_like) -> str:
    if pdfplumber is None:
        return ""
    try:
        with pdfplumber.open(file_like) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""

def read_docx_text_bytes(docx_bytes: bytes) -> str:
    """Lightweight text extraction from DOCX for industry detection (doesn't alter formatting)."""
    if DocxDocument is None:
        return ""
    try:
        bio = io.BytesIO(docx_bytes)
        doc = DocxDocument(bio)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""

# ==========================
# Helpers: .DOC → .DOCX conversion (best-effort)
# ==========================
def _try_libreoffice_convert(doc_bytes: bytes) -> Optional[bytes]:
    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, "in.doc")
        out_path = os.path.join(tmpdir, "in.docx")
        with open(in_path, "wb") as f:
            f.write(doc_bytes)
        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "docx", "--outdir", tmpdir, in_path],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            if os.path.exists(out_path):
                with open(out_path, "rb") as f:
                    return f.read()
        except Exception:
            return None
    return None

def _try_pandoc_convert(doc_bytes: bytes) -> Optional[bytes]:
    try:
        import pypandoc  # type: ignore
    except Exception:
        return None
    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, "in.doc")
        out_path = os.path.join(tmpdir, "out.docx")
        with open(in_path, "wb") as f:
            f.write(doc_bytes)
        try:
            pypandoc.convert_file(in_path, "docx", outputfile=out_path)
            if os.path.exists(out_path):
                with open(out_path, "rb") as f:
                    return f.read()
        except Exception:
            return None
    return None

def convert_doc_to_docx_bytes(doc_bytes: bytes) -> bytes:
    converted = _try_libreoffice_convert(doc_bytes)
    if converted:
        return converted
    converted = _try_pandoc_convert(doc_bytes)
    if converted:
        return converted
    raise RuntimeError(
        "Could not convert .doc to .docx automatically. "
        "Please save the file as .docx in Microsoft Word and re-upload."
    )

# ==========================
# Helpers: DOCX formatting-preserving replacements
# ==========================
def _iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                yield p
            for tbl in section.header.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p
        if section.footer:
            for p in section.footer.paragraphs:
                yield p
            for tbl in section.footer.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p

def _replace_preserving_style(paragraph: "Paragraph", old: str, new: str) -> int:
    if old not in paragraph.text:
        return 0
    runs = paragraph.runs
    full = "".join(r.text for r in runs)
    n = full.count(old)
    full_new = full.replace(old, new)
    if runs:
        style = runs[0].style
        for r in runs: r.text = ""
        runs[0].text = full_new
        runs[0].style = style
    else:
        paragraph.add_run(full_new)
    return n

def docx_replace_text_everywhere(doc: "DocxDocument", replacements: Dict[str, str]) -> int:
    total = 0
    for p in _iter_all_paragraphs(doc):
        for old, new in replacements.items():
            total += _replace_preserving_style(p, old, new)
    return total

def detect_years(text: str) -> set:
    years = set(re.findall(r"(?:FY\s*-?_?\s*)?(20\d{2})", text, flags=re.I))
    ranges = re.findall(r"(?:FY\s*)?(20\d{2})\s*/\s*(\d{2})", text, flags=re.I)
    for y, yy in ranges:
        try:
            y1 = int(y); y2 = (y1 // 100) * 100 + int(yy)
            years.add(str(y1)); years.add(str(y2))
        except Exception:
            pass
    return years

def bump_range_token(token: str, new_start_year: int) -> str:
    m = re.search(r"(20\d{2})\s*/\s*(\d{2})", token)
    if not m:
        return token
    y1 = new_start_year; y2 = (y1 % 100) + 1
    return re.sub(r"(20\d{2})\s*/\s*(\d{2})", f"{y1}/{y2:02d}", token)

def build_rollforward_replacements(doc: "DocxDocument", new_fy: int, report_date: str) -> Dict[str, str]:
    text_all = "\n".join(p.text for p in _iter_all_paragraphs(doc))
    years = detect_years(text_all) or {str(new_fy - 1)}
    repl: Dict[str, str] = {}
    for y in years:
        repl[f"FY{y}"] = f"FY{new_fy}"
        repl[f"FY {y}"] = f"FY {new_fy}"
        repl[f"FYE {y}"] = f"FYE {new_fy}"
        repl[f"Financial Year {y}"] = f"Financial Year {new_fy}"
        repl[f"Fiscal Year {y}"] = f"Fiscal Year {new_fy}"
    tokens = re.findall(r"(?:FY\s*)?(20\d{2}\s*/\s*\d{2})", text_all)
    for t in set(tokens):
        repl[t] = bump_range_token(t, new_fy)
    if report_date:
        repl["Report Date"] = f"Report Date: {report_date}"
    return repl

# ==========================
# Industry detection (from prior TPD text)
# ==========================
INDUSTRY_KEYWORDS = {
    "ICT / Technology": [
        "software", "it services", "saas", "cloud", "telecom", "telecommunications",
        "data center", "ai", "machine learning", "cybersecurity", "ict", "internet", "platform"
    ],
    "Manufacturing": [
        "manufactur", "production facility", "plant", "factory", "assembly", "oem", "industrial"
    ],
    "Agriculture / Food": [
        "farming", "agricultur", "crop", "livestock", "food processing", "beverage", "dairy", "meat"
    ],
    "Energy / Utilities": [
        "electricity", "power generation", "renewable", "oil", "gas", "pipeline", "utility", "solar", "wind"
    ],
    "Financial Services": [
        "bank", "insurance", "fintech", "payment", "lending", "asset management", "securities", "brokerage"
    ],
    "Retail / Wholesale": [
        "retail", "wholesale", "store", "e-commerce", "omnichannel", "merchandising", "distribution network"
    ],
    "Healthcare / Pharma": [
        "pharma", "pharmaceutical", "biotech", "clinical", "medical device", "hospital", "healthcare", "diagnostic"
    ],
    "Transport / Logistics": [
        "logistics", "freight", "shipping", "airline", "rail", "warehouse", "3pl", "last-mile", "fleet"
    ],
    "Professional Services": [
        "consulting", "legal services", "accounting", "advisory", "engineering services", "staff augmentation"
    ],
}
def detect_industry_label(text: str) -> str:
    low = (text or "").lower()
    scores = {label: 0 for label in INDUSTRY_KEYWORDS}
    for label, kws in INDUSTRY_KEYWORDS.items():
        for k in kws:
            if k in low:
                scores[label] += 1
    best = max(scores.items(), key=lambda x: x[1])
    return best[0] if best[1] > 0 else "General / Macro"

# ==========================
# Auto Industry Research (World Bank) — sector packs (open-ended default)
# ==========================
WB_BASE = "https://api.worldbank.org/v2"

def wb_get_countries() -> List[Dict[str, Any]]:
    try:
        r = requests.get(f"{WB_BASE}/country?format=json&per_page=400", timeout=15)
        data = r.json()
        return data[1] if isinstance(data, list) and len(data) > 1 else []
    except Exception:
        return []

def wb_resolve_country(user_input: str) -> Optional[Tuple[str, str]]:
    ui = (user_input or "").strip().lower()
    countries = wb_get_countries()
    for c in countries:
        name = (c.get("name") or "").lower()
        iso2 = (c.get("id") or "").lower()
        if ui == name or ui == iso2 or ui in name:
            return (c.get("id"), c.get("name"))
    return None

# Indicator catalog (default pack + sector add-ons)
WB_INDICATORS_PACKS: Dict[str, Dict[str, Tuple[str, str]]] = {
    "General / Macro": {
        "gdp_growth": ("NY.GDP.MKTP.KD.ZG", "GDP growth (annual %)"),
        "inflation": ("FP.CPI.TOTL.ZG", "Inflation, consumer prices (annual %)"),
        "services_share": ("NV.SRV.TOTL.ZS", "Services value added (% of GDP)"),
        "manufacturing_share": ("NV.IND.MANF.ZS", "Manufacturing value added (% of GDP)"),
        "industry_share": ("NV.IND.TOTL.ZS", "Industry (incl. construction) value added (% of GDP)"),
    },
    "ICT / Technology": {
        "internet_users": ("IT.NET.USER.ZS", "Individuals using the Internet (% of population)"),
        "ict_goods_exports": ("TX.VAL.ICTG.ZS.UN", "ICT goods exports (% of total goods exports)"),
        "hightech_exports": ("TX.VAL.TECH.MF.ZS", "High-technology exports (% of manufactured exports)"),
    },
    "Manufacturing": {
        "manufacturing_share": ("NV.IND.MANF.ZS", "Manufacturing value added (% of GDP)"),
        "hightech_exports": ("TX.VAL.TECH.MF.ZS", "High-technology exports (% of manufactured exports)"),
    },
    "Agriculture / Food": {
        "agri_share": ("NV.AGR.TOTL.ZS", "Agriculture, forestry, and fishing value added (% of GDP)"),
    },
    "Energy / Utilities": {
        "access_electricity": ("EG.ELC.ACCS.ZS", "Access to electricity (% of population)"),
        "renewable_output": ("EG.ELC.RNEW.ZS", "Renewable electricity output (% of total electricity output)"),
    },
    "Financial Services": {
        "domestic_credit_banks": ("FS.AST.PRVT.GD.ZS", "Domestic credit to private sector by banks (% of GDP)"),
    },
    "Retail / Wholesale": {
        "internet_users": ("IT.NET.USER.ZS", "Individuals using the Internet (% of population)"),
    },
    "Healthcare / Pharma": {
        "internet_users": ("IT.NET.USER.ZS", "Individuals using the Internet (% of population)"),
    },
    "Transport / Logistics": {
        "internet_users": ("IT.NET.USER.ZS", "Individuals using the Internet (% of population)"),
    },
    "Professional Services": {
        "services_share": ("NV.SRV.TOTL.ZS", "Services value added (% of GDP)"),
    },
}

def wb_fetch_indicator_series(iso2: str, indicator: str) -> Dict[str, Any]:
    url = f"{WB_BASE}/country/{iso2}/indicator/{indicator}?format=json&per_page=70"
    series = []
    latest_year = None
    latest_value = None
    try:
        r = requests.get(url, timeout=20)
        data = r.json()
        if isinstance(data, list) and len(data) > 1 and isinstance(data[1], list):
            for row in data[1]:
                year = row.get("date"); val = row.get("value")
                if year is None or val is None: continue
                try:
                    y = int(year); v = float(val)
                except Exception:
                    continue
                series.append((y, v))
            series.sort()
            for y, v in reversed(series):
                latest_year, latest_value = y, v
                break
    except Exception:
        pass
    return {"latest_year": latest_year, "latest_value": latest_value, "series": series, "source_url": url}

def auto_sector_research(country_input: str, industry_label: str) -> Dict[str, Any]:
    resolved = wb_resolve_country(country_input)
    if not resolved:
        return {"note": "Could not resolve country; please use a standard name (e.g., Singapore).", "items": {}}
    iso2, country_name = resolved
    pack = {**WB_INDICATORS_PACKS["General / Macro"], **WB_INDICATORS_PACKS.get(industry_label, {})}
    out = {"country": country_name, "iso2": iso2, "industry": industry_label, "items": {}, "notes": []}
    for key, (code, label) in pack.items():
        data = wb_fetch_indicator_series(iso2, code)
        out["items"][key] = {**data, "code": code, "label": f"World Bank — {label}"}
        if data["latest_year"] is None:
            out["notes"].append(f"Missing recent data: {label}")
    return out

def format_sector_update_text(sector_pack: Dict[str, Any]) -> Tuple[List[str], List[Tuple[int, str]]]:
    lines: List[str] = []
    foots: List[Tuple[int, str]] = []
    if not sector_pack or "items" not in sector_pack: return lines, foots
    items = sector_pack["items"]

    def add_line(key: str, fmt: str):
        if key in items and items[key]["latest_year"] is not None:
            y = items[key]["latest_year"]; v = items[key]["latest_value"]; url = items[key]["source_url"]
            try:
                lines.append(fmt.format(v=float(v), y=int(y)))
            except Exception:
                lines.append(fmt.format(v=v, y=y))
            foots.append((len(foots) + 1, url))

    # General
    add_line("gdp_growth", "GDP growth: {v:.1f}% in {y}.")
    add_line("inflation", "Inflation (CPI): {v:.1f}% in {y}.")
    add_line("services_share", "Services share of GDP: {v:.1f}% (in {y}).")
    add_line("manufacturing_share", "Manufacturing share of GDP: {v:.1f}% (in {y}).")
    add_line("industry_share", "Industry (incl. construction) share of GDP: {v:.1f}% (in {y}).")

    # ICT/Tech
    add_line("internet_users", "Internet usage: {v:.1f}% of population (in {y}).")
    add_line("ict_goods_exports", "ICT goods exports: {v:.1f}% of total goods exports (in {y}).")
    add_line("hightech_exports", "High-tech exports: {v:.1f}% of manufactured exports (in {y}).")

    # Agriculture
    add_line("agri_share", "Agriculture, forestry & fishing: {v:.1f}% of GDP (in {y}).")

    # Energy
    add_line("access_electricity", "Access to electricity: {v:.1f}% of population (in {y}).")
    add_line("renewable_output", "Renewable electricity output: {v:.1f}% of total electricity output (in {y}).")

    # Finance
    add_line("domestic_credit_banks", "Domestic credit to private sector by banks: {v:.1f}% of GDP (in {y}).")

    return lines, foots

def fetch_title(url: str) -> str:
    try:
        r = requests.get(url, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        title = soup.title.string.strip() if soup.title and soup.title.string else url
        return title[:120]
    except Exception:
        return url

# ==========================
# PAGE: TPD Draft (industry-aware + roll-forward vs rewrite)
# ==========================
if page == "TPD Draft":
    st.title("TPD Draft Generator")
    st.write("Upload prior-year TPD as **Microsoft Word (.docx or .doc)** to preserve fonts/colours/sizes. PDFs are supported but styles cannot be preserved.")

    # 1) Upload prior TPD
    prior = st.file_uploader(
        "Upload Prior TPD (DOCX/DOC preferred; PDF supported as JSON fallback)",
        type=["docx", "doc", "pdf"],
        accept_multiple_files=False
    )

    # 2) New FY
    new_fy = st.number_input("New FY (e.g., 2024)", min_value=1990, max_value=2100, value=2024)

    # 3) Financial year end (new field)
    fye_date = st.text_input("Financial Year End (e.g., 31 December 2024)", value="")

    # 4) Country of report
    override_country = st.text_input("Country of report (for auto research)", value="Singapore")

    # 5) Information available
    mode = st.radio("Additional information available?", [
        "No information",
        "Client information request",
        "Benchmark study",
        "Both (IRL + Benchmark)",
    ])

    bench_df: Optional[pd.DataFrame] = None
    irl_text: Optional[str] = None

    if mode in ("Benchmark study", "Both (IRL + Benchmark)"):
        bench_file = st.file_uploader("Attach benchmark export (CSV/XLSX)", type=["csv", "xlsx"], key="bench")
        if bench_file is not None:
            try:
                bench_df = pd.read_csv(bench_file) if bench_file.name.endswith(".csv") else pd.read_excel(bench_file)
                st.caption("Loaded benchmark for inclusion in draft.")
            except Exception as e:
                st.error(f"Could not read benchmark: {e}")

    if mode in ("Client information request", "Both (IRL + Benchmark)"):
        irl_up = st.file_uploader("Attach client info (TXT/CSV) to insert", type=["txt", "csv"], key="irl")
        if irl_up is not None:
            try:
                if irl_up.name.endswith(".csv"):
                    _df = pd.read_csv(irl_up)
                    irl_text = "\n".join("- " + " | ".join(map(str, row)) for _, row in _df.iterrows())
                else:
                    irl_text = irl_up.read().decode("utf-8", errors="ignore")
                st.caption("Loaded client information for inclusion.")
            except Exception as e:
                st.error(f"Could not read client info: {e}")

    # 6) Industry analysis mode
    st.subheader("Industry Analysis Mode")
    industry_mode = st.radio(
        "How should we handle Industry Analysis?",
        ["Roll-forward (update facts & stats)", "Full Rewrite"],
        help="Roll-forward: update outdated numbers and citations only, keeping prior narrative. Full Rewrite: rebuild the section from scratch."
    )

    # Detect industry from prior TPD text
    detected_industry = "General / Macro"
    prior_text_for_detection = ""
    if prior is not None:
        name = prior.name.lower()
        if name.endswith(".docx"):
            prior_text_for_detection = read_docx_text_bytes(prior.getvalue())
        elif name.endswith(".doc"):
            try:
                converted = convert_doc_to_docx_bytes(prior.getvalue())
                prior_text_for_detection = read_docx_text_bytes(converted)
            except Exception:
                prior_text_for_detection = ""
        elif name.endswith(".pdf"):
            prior_text_for_detection = read_pdf(io.BytesIO(prior.getvalue()))
        detected_industry = detect_industry_label(prior_text_for_detection)

    st.write("**Detected industry (from prior TPD, editable):**")
    industry_choice = st.selectbox(
        "Industry",
        options=list(WB_INDICATORS_PACKS.keys()),
        index=list(WB_INDICATORS_PACKS.keys()).index(detected_industry) if detected_industry in WB_INDICATORS_PACKS else 0,
        help="Auto-detected from prior TPD text. You can override."
    )

    # 7) Industry sources
    st.subheader("Industry sources (optional)")
    urls = st.text_area("Extra source URLs (one per line, optional)", value="")
    user_url_list = [u.strip() for u in urls.splitlines() if u.strip()]
    user_reports = st.file_uploader("Upload market/industry reports (PDF/DOCX/TXT — optional)", type=["pdf","docx","txt"], accept_multiple_files=True)

    # Advanced text replacements
    adv = st.expander("Advanced: custom replacements (JSON)", expanded=False)
    with adv:
        st.write('Example: {"{{ENTITY}}": "ABC Pte Ltd", "{{COUNTRY}}": "Singapore"}')
        repl_json = st.text_area("Key-value JSON (optional)", value="")

    # --- Generate draft (same as before) ---
    if st.button("Generate TPD draft now", type="primary"):
        # [generation logic stays unchanged below...]
        ...


# ==========================
# Other pages (kept simple)
# ==========================
elif page == "TNMM Review":
    st.title("TNMM Benchmark Review (demo)")
    up = st.file_uploader("Upload Benchmark (CSV/XLSX)", type=["csv", "xlsx"])
    if up:
        try:
            df = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
            st.dataframe(df)
            if "Decision" in df.columns and "Reason" in df.columns:
                flags = df[(df["Decision"].astype(str).str.lower() == "reject") & (df["Reason"].astype(str).str.strip() == "")]
                st.subheader("⚠️ Rejects with empty reason")
                st.dataframe(flags)
                if not flags.empty:
                    st.download_button(
                        "Download flags (CSV)",
                        data=flags.to_csv(index=False).encode("utf-8"),
                        file_name="tnmm_flags.csv",
                        mime="text/csv",
                    )
        except Exception as e:
            st.error(f"Could not read file: {e}")

elif page == "CUT/CUP Review":
    st.title("CUT / CUP Agreements Review (demo)")
    st.info("Clause extraction + scoring can be wired next; current focus is the TPD generator.")

elif page == "Information Request List":
    st.title("Information Request List (IRL)")
    industry = st.text_input("Industry", value="Technology / Services")
    transactions = st.text_area("Transactions in-scope (comma-separated)", value="intra-group services, distribution")
    if st.button("Generate IRL"):
        required = {
            "financials": ["Trial balance FY", "Segmented P&L by service line", "Intercompany charges by counterparty"],
            "legal": ["Latest org chart", "All intercompany agreements", "Board minutes re: restructuring"],
            "operational": ["Headcount by function", "KPIs / cost drivers", "Descriptions of services performed"],
        }
        email = (
            f"Dear Client,\n\nTo complete the FY TPD update for {industry}, please provide the following:\n- "
            + "\n- ".join(sum(required.values(), []))
            + "\n\nTransactions in scope: "
            + transactions
            + "\n\nKind regards,\nTP Team"
        )
        out = {"requests": required, "email": email}
        st.json(out)
        st.download_button(
            "Download IRL (JSON)",
            data=json.dumps(out, indent=2).encode("utf-8"),
            file_name="info_requests.json",
            mime="application/json",
        )

else:
    st.title("Advisory / Opportunity Spotting (demo)")
    st.info("Upload a benchmark on the TNMM page to explore opportunities; simplified here.")



