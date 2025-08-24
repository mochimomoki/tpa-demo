# app.py — TPD Draft Generator (jurisdiction-aware + roll-forward vs rewrite + DOCX formatting + .DOC conversion)
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
    if DocxDocument is None:
        return ""
    try:
        bio = io.BytesIO(docx_bytes)
        doc = DocxDocument(bio)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""

# ==========================
# Helpers: .DOC → .DOCX conversion
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
    raise RuntimeError("Could not convert .doc to .docx automatically. Please save the file as .docx in Microsoft Word and re-upload.")

# ==========================
# Compliance Rules
# ==========================
JURISDICTION_RULES = {
    "OECD": {
        "required_sections": ["Group Overview", "Local Entity FAR", "Controlled Transactions", "Economic Analysis", "Conclusion"],
        "thresholds": {},
        "citation": "OECD Transfer Pricing Guidelines (2022)"
    },
    "Singapore": {
        "required_sections": ["Group Overview", "Entity Overview", "FAR", "Related Party Transactions", "Economic Analysis", "Supporting Documents"],
        "thresholds": {"mandatory_local_file_revenue": 10000000},
        "citation": "IRAS Transfer Pricing Guidelines (6th edition, 2024)"
    },
    "Australia": {
        "required_sections": ["Master File", "Local File"],
        "thresholds": {"CbCR_consolidated_revenue": 1000000000},
        "citation": "ATO Transfer Pricing Guidelines"
    },
}

def check_jurisdiction_compliance(country: str, detected_sections: List[str], revenue: Optional[float] = None) -> List[str]:
    issues = []
    rules = JURISDICTION_RULES.get(country, JURISDICTION_RULES["OECD"])
    for required in rules["required_sections"]:
        if required not in detected_sections:
            issues.append(f"Missing section: {required}")
    if revenue is not None:
        for k, v in rules.get("thresholds", {}).items():
            if "revenue" in k and revenue < v:
                issues.append(f"Entity revenue {revenue} below threshold ({v}); section may not be mandatory")
    return issues

# ==========================
# Industry Detection (simplified)
# ==========================
INDUSTRY_KEYWORDS = {
    "ICT / Technology": ["software","it services","cloud","telecom","platform"],
    "Manufacturing": ["manufactur","plant","factory"],
    "Healthcare / Pharma": ["pharma","biotech","medical","hospital"],
    "General / Macro": []
}
def detect_industry_label(text: str) -> str:
    low=(text or "").lower()
    scores={label:0 for label in INDUSTRY_KEYWORDS}
    for label,kws in INDUSTRY_KEYWORDS.items():
        for k in kws:
            if k in low: scores[label]+=1
    best=max(scores.items(),key=lambda x:x[1])
    return best[0] if best[1]>0 else "General / Macro"

# ==========================
# World Bank Sector Stats (simplified demo)
# ==========================
WB_BASE="https://api.worldbank.org/v2"
def wb_fetch_indicator_series(iso2: str, indicator: str) -> Dict[str,Any]:
    url=f"{WB_BASE}/country/{iso2}/indicator/{indicator}?format=json&per_page=70"
    try:
        r=requests.get(url,timeout=15); data=r.json()
        if isinstance(data,list) and len(data)>1 and isinstance(data[1],list):
            for row in data[1]:
                if row.get("value"): return {"latest_year":row["date"],"latest_value":row["value"],"source_url":url}
    except Exception: pass
    return {"latest_year":None,"latest_value":None,"source_url":url}

def format_sector_update_text(country: str) -> List[str]:
    lines=[]
    for code,label in [("NY.GDP.MKTP.KD.ZG","GDP growth"),("FP.CPI.TOTL.ZG","Inflation")]:
        data=wb_fetch_indicator_series(country,code)
        if data["latest_year"]:
            lines.append(f"{label}: {data['latest_value']} in {data['latest_year']}. Source: {data['source_url']}")
    return lines

# ==========================
# PAGE: TPD Draft
# ==========================
if page=="TPD Draft":
    st.title("TPD Draft Generator")

    # 1) Upload prior TPD
    prior=st.file_uploader("Upload Prior TPD (.docx/.doc/.pdf)",type=["docx","doc","pdf"])

    # 2) New FY
    new_fy=st.number_input("New FY (e.g. 2024)",1990,2100,2024)

    # 3) Financial Year End
    fye_date=st.text_input("Financial Year End (e.g. 31 Dec 2024)")

    # 4) Country of report
    country_choice=st.selectbox("Select jurisdiction",list(JURISDICTION_RULES.keys()),index=1)

    # 5) Information available
    mode=st.radio("Additional information available?",["No information","Client information request","Benchmark study","Both"])
    bench_df=None; irl_text=None

    if mode in("Benchmark study","Both"):
        bench_file=st.file_uploader("Upload benchmark (CSV/XLSX)",type=["csv","xlsx"])
        if bench_file is not None:
            bench_df=pd.read_csv(bench_file) if bench_file.name.endswith(".csv") else pd.read_excel(bench_file)

    if mode in("Client information request","Both"):
        irl_up=st.file_uploader("Upload client info (TXT/CSV)",type=["txt","csv"])
        if irl_up is not None:
            irl_text=irl_up.read().decode("utf-8",errors="ignore")

    # 6) Industry Analysis Mode
    industry_mode=st.radio("Industry Analysis Mode",["Roll-forward","Full Rewrite"])

    # Detect industry
    detected="General / Macro"
    if prior is not None:
        if prior.name.endswith(".docx"): text=read_docx_text_bytes(prior.getvalue())
        elif prior.name.endswith(".pdf"): text=read_pdf(io.BytesIO(prior.getvalue()))
        else: text=""
        detected=detect_industry_label(text)
    industry_choice=st.selectbox("Industry",list(INDUSTRY_KEYWORDS.keys()),index=list(INDUSTRY_KEYWORDS.keys()).index(detected))

    # 7) Industry sources
    urls=st.text_area("Extra source URLs (one per line)")
    user_url_list=[u.strip() for u in urls.splitlines() if u.strip()]

    # Generate
    if st.button("Generate TPD Draft",type="primary"):
        if prior is None:
            st.error("Upload a TPD first")
        else:
            detected_sections=["FAR","Economic Analysis","Related Party Transactions"]  # placeholder detection
            issues=check_jurisdiction_compliance(country_choice,detected_sections,revenue=20000000)

            if issues:
                st.warning("⚠️ Compliance Gaps Found:")
                for i in issues: st.write(f"- {i}")
            else:
                st.success(f"Draft complies with {country_choice} guidelines.")

            lines=format_sector_update_text("SG")
            draft=f"--- Draft TPD ---\nJurisdiction: {country_choice}\nNew FY: {new_fy}\nFYE: {fye_date}\n\nIndustry Analysis ({industry_mode}): {industry_choice}\n"
            for ln in lines: draft+=f"\n- {ln}"
            if user_url_list: draft+="\nExtra sources:\n"+"\n".join(user_url_list)

            st.text_area("Preview Draft",draft,height=300)

# ==========================
# Other pages (unchanged)
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
        email = f"Dear Client,\n\nTo complete the FY TPD update for {industry}, please provide the following:\n- " + "\n- ".join(sum(required.values(), []))
        st.json({"requests": required, "email": email})

else:
    st.title("Advisory / Opportunity Spotting (demo)")
    st.info("Upload a benchmark on the TNMM page to explore opportunities; simplified here.")




