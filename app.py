# app.py — TPD Draft Generator (jurisdiction-aware + industry-aware + roll-forward)
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
st.sidebar.caption("Roll-forward TPD • Industry-aware • Jurisdiction-aware • Formatting preserved")

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
# Compliance Rules (sample jurisdictions)
# ==========================
JURISDICTION_RULES = {
    "OECD": {
        "required_sections": ["Group Overview","Local Entity FAR","Controlled Transactions","Economic Analysis","Conclusion"],
        "thresholds": {},
        "citation": "OECD Transfer Pricing Guidelines (2022)"
    },
    "Singapore": {
        "required_sections": ["Group Overview","Entity Overview","FAR","Related Party Transactions","Economic Analysis","Supporting Documents"],
        "thresholds": {"mandatory_local_file_revenue": 10000000},
        "citation": "IRAS Transfer Pricing Guidelines (6th edition, 2024)"
    },
    "Australia": {
        "required_sections": ["Master File","Local File"],
        "thresholds": {"CbCR_consolidated_revenue": 1000000000},
        "citation": "ATO Transfer Pricing Guidelines"
    },
    "United States": {
        "required_sections": ["Controlled Transactions","Intercompany Agreements","Economic Analysis","Supporting Documentation"],
        "thresholds": {"mandatory_doc_revenue": 50000000},
        "citation": "US IRC §482 & Treas. Regs."
    },
    "United Kingdom": {
        "required_sections": ["Master File","Local File","CbC Report"],
        "thresholds": {"CbCR_consolidated_revenue": 750000000},
        "citation": "UK HMRC TP Guidelines"
    },
    "India": {
        "required_sections": ["Entity Overview","International Transactions","Economic Analysis","Accountant’s Report (Form 3CEB)"],
        "thresholds": {"mandatory_local_file_revenue": 100000000},
        "citation": "Indian Income Tax Rules, Sec. 92E"
    },
    "European Union": {
        "required_sections": ["Master File","Local File"],
        "thresholds": {"CbCR_consolidated_revenue": 750000000},
        "citation": "EU Transfer Pricing Directive"
    },
}

def check_jurisdiction_compliance(country: str, detected_sections: List[str], revenue: Optional[float] = None) -> List[str]:
    issues=[]
    rules=JURISDICTION_RULES.get(country,JURISDICTION_RULES["OECD"])
    for required in rules["required_sections"]:
        if required not in detected_sections:
            issues.append(f"Missing section: {required}")
    if revenue is not None:
        for k,v in rules.get("thresholds",{}).items():
            if "revenue" in k and revenue < v:
                issues.append(f"Entity revenue {revenue} below threshold ({v}); section may not be mandatory")
    return issues

# ==========================
# Industry Detection
# ==========================
INDUSTRY_KEYWORDS = {
    "ICT / Technology":["software","it services","saas","cloud","telecom","platform","cybersecurity","ai"],
    "Manufacturing":["manufactur","plant","factory","assembly","production"],
    "Agriculture / Food":["farming","crop","livestock","food","dairy","meat","beverage"],
    "Energy / Utilities":["oil","gas","electricity","renewable","pipeline","solar","wind"],
    "Financial Services":["bank","insurance","fintech","asset management","brokerage"],
    "Retail / Wholesale":["retail","wholesale","store","e-commerce","distribution"],
    "Healthcare / Pharma":["pharma","medical","biotech","hospital","diagnostic","clinical"],
    "Transport / Logistics":["logistics","freight","shipping","rail","warehouse","fleet"],
    "Professional Services":["consulting","legal","accounting","advisory","engineering"],
    "General / Macro":[]
}
def detect_industry_label(text:str)->str:
    low=(text or "").lower()
    scores={label:0 for label in INDUSTRY_KEYWORDS}
    for label,kws in INDUSTRY_KEYWORDS.items():
        for k in kws:
            if k in low: scores[label]+=1
    best=max(scores.items(),key=lambda x:x[1])
    return best[0] if best[1]>0 else "General / Macro"

# ==========================
# World Bank Data (credible sources)
# ==========================
WB_BASE="https://api.worldbank.org/v2"

def wb_get_countries()->List[Dict[str,Any]]:
    try:
        r=requests.get(f"{WB_BASE}/country?format=json&per_page=400",timeout=15)
        data=r.json()
        return data[1] if isinstance(data,list) and len(data)>1 else []
    except Exception: return []

def wb_resolve_country(user_input:str)->Optional[Tuple[str,str]]:
    ui=(user_input or "").strip().lower()
    for c in wb_get_countries():
        if ui in (c.get("name") or "").lower(): return (c.get("id"),c.get("name"))
    return None

def wb_fetch_indicator_series(iso2:str,indicator:str)->Dict[str,Any]:
    url=f"{WB_BASE}/country/{iso2}/indicator/{indicator}?format=json&per_page=70"
    try:
        r=requests.get(url,timeout=15); data=r.json()
        if isinstance(data,list) and len(data)>1 and isinstance(data[1],list):
            for row in data[1]:
                if row.get("value"): return {"latest_year":row["date"],"latest_value":row["value"],"source_url":url}
    except Exception: pass
    return {"latest_year":None,"latest_value":None,"source_url":url}

def format_sector_update_text(country:str)->List[str]:
    lines=[]
    for code,label in [("NY.GDP.MKTP.KD.ZG","GDP growth"),("FP.CPI.TOTL.ZG","Inflation")]:
        data=wb_fetch_indicator_series(country,code)
        if data["latest_year"]:
            lines.append(f"{label}: {data['latest_value']} in {data['latest_year']} (World Bank)")
    return lines

# ==========================
# PAGE: TPD Draft
# ==========================
if page=="TPD Draft":
    from docx import Document

if st.button("Generate TPD Draft", type="primary", key="tpd_generate"):
    if prior is None:
        st.error("Please upload a prior TPD (.docx/.doc/.pdf)")
    else:
        # Compliance check
        detected_sections = ["FAR","Economic Analysis","Related Party Transactions"]  # TODO: parse actual doc headings
        issues = check_jurisdiction_compliance(country_choice, detected_sections, revenue=20000000)

        if issues:
            st.warning("⚠️ Compliance Gaps Found:")
            for i in issues:
                st.write(f"- {i}")
        else:
            st.success(f"Draft complies with {country_choice} guidelines.")

        # --- Build the draft TPD ---
        doc = Document()
        doc.add_heading("Transfer Pricing Documentation — Draft", level=1)
        doc.add_paragraph(f"Jurisdiction: {country_choice}")
        doc.add_paragraph(f"New FY: {new_fy}")
        doc.add_paragraph(f"Financial Year End: {fye_date}")
        doc.add_paragraph(f"Industry: {industry_choice}")
        doc.add_paragraph(f"Industry Analysis Mode: {industry_mode}")

        doc.add_heading("Industry Update", level=2)
        resolved = wb_resolve_country(country_choice)
        lines = format_sector_update_text(resolved[0]) if resolved else []
        for ln in lines:
            doc.add_paragraph(f"- {ln}")

        if user_url_list:
            doc.add_heading("Additional Sources", level=2)
            for u in user_url_list:
                doc.add_paragraph(u)

        if bench_df is not None:
            doc.add_heading("Benchmark Study", level=2)
            acc = (bench_df.get("Decision", "").astype(str).str.lower() == "accept").sum()
            rej = (bench_df.get("Decision", "").astype(str).str.lower() == "reject").sum()
            doc.add_paragraph(f"Vendor study: {acc} accepted, {rej} rejected, {len(bench_df)} total.")

        if irl_text:
            doc.add_heading("Client Information", level=2)
            for line in irl_text.splitlines():
                if line.strip():
                    doc.add_paragraph(f"• {line.strip()}")

        # Save to buffer
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)

        st.download_button(
            "⬇️ Download Draft TPD (DOCX)",
            data=out,
            file_name="TPD_Draft.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


# ==========================
# Other pages (TNMM, CUT/CUP, IRL, Advisory)
# ==========================
elif page=="TNMM Review":
    st.title("TNMM Benchmark Review (demo)")
    up=st.file_uploader("Upload Benchmark",type=["csv","xlsx"])
    if up:
        try:
            df=pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
            st.dataframe(df)
        except Exception as e: st.error(f"Could not read file: {e}")

elif page=="CUT/CUP Review":
    st.title("CUT / CUP Agreements Review (demo)")
    st.info("Clause extraction + scoring can be wired next; current focus is the TPD generator.")

elif page=="Information Request List":
    st.title("Information Request List (IRL)")
    industry=st.text_input("Industry",value="Technology / Services")
    transactions=st.text_area("Transactions in-scope",value="intra-group services, distribution")
    if st.button("Generate IRL"):
        required={"financials":["Trial balance FY","Segmented P&L","Intercompany charges"],"legal":["Org chart","Intercompany agreements"],"ops":["Headcount","KPIs"]}
        email=f"Dear Client,\n\nTo complete the FY TPD update for {industry}, please provide:\n- "+"\n- ".join(sum(required.values(),[]))
        st.json({"requests":required,"email":email})

else:
    st.title("Advisory / Opportunity Spotting (demo)")
    st.info("Upload a benchmark on the TNMM page to explore opportunities.")





