# app.py — TPD Draft Generator (jurisdiction-aware + industry-aware + roll-forward + DOCX output)
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






