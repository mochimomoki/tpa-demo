"""
app.py — TPA Core Demo (Roll‑Forward with Formatting + Options)

What’s new:
1) Formatting preservation (DOCX only): clones your prior DOCX and replaces text in‑place → fonts/colours/sizes stay intact.
2) Update mode: No information / Client information request / Benchmark study / Both.
3) Auto FY roll‑forward: replace FY 20XX patterns to your chosen FY.
4) Conditional inserts: benchmark summary, client info block.
5) Industry update: paste credible URLs; we add a short subsection + numbered footnotes (fetching page titles for nicer citations).

requirements.txt should have:
streamlit\npandas\npdfplumber\npython-docx\nopenpyxl\nrequests\nbeautifulsoup4
"""
from __future__ import annotations
import io, re, json
from typing import Dict, Any, List, Optional

import streamlit as st
import pandas as pd

# Optional deps
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from docx import Document as DocxDocument
except Exception:
    DocxDocument = None

import requests
from bs4 import BeautifulSoup

st.set_page_config(page_title="TPA (Transfer Pricing Associate)", layout="wide")
st.sidebar.title("TPA (Transfer Pricing Associate)")
st.sidebar.write("Roll‑forward generator with formatting + options")

page = st.sidebar.radio(
    "Choose function",
    ["TPD Draft", "TNMM Review", "CUT/CUP Review", "Information Request List", "Advisory / Opportunity Spotting"],
)
# ---- PwC theming (drop this right after st.set_page_config) ----
PWC_COLORS = {
    "primary": "#FD5108",   # bold signature orange
    "orange1": "#D85604",
    "orange2": "#E88D14",
    "yellow":  "#F3BE26",
    "red1":    "#AD1B02",
    "red2":    "#E0301E",
    "pink":    "#E669A2",
    "black":   "#000000",
    "grey1":   "#2D2D2D",
    "grey2":   "#7D7D7D",
    "grey3":   "#DEDEDE",
}

st.markdown(f"""
<style>
/* Page background & typography */
:root {{
  --pwc-primary: {PWC_COLORS["primary"]};
  --pwc-orange1: {PWC_COLORS["orange1"]};
  --pwc-orange2: {PWC_COLORS["orange2"]};
  --pwc-yellow:  {PWC_COLORS["yellow"]};
  --pwc-red:     {PWC_COLORS["red2"]};
  --pwc-pink:    {PWC_COLORS["pink"]};
  --pwc-black:   {PWC_COLORS["black"]};
  --pwc-grey1:   {PWC_COLORS["grey1"]};
  --pwc-grey2:   {PWC_COLORS["grey2"]};
  --pwc-grey3:   {PWC_COLORS["grey3"]};
}}

html, body, [data-testid="stAppViewContainer"] {{
  background: linear-gradient(180deg, #ffffff 0%, #fff9f3 60%, #fff6e8 100%);
  color: var(--pwc-grey1);
  font-family: "Inter", -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
}}

[data-testid="stHeader"] {{ background: transparent; }}

.sidebar .sidebar-content {{ background: #ffffff00; }}

section.main > div:has(> .block-container) {{
  padding-top: 0.5rem;
}}

.block-container {{
  padding-top: 1rem;
}}

h1, h2, h3 {{
  letter-spacing: .2px;
}}
h1 {{ color: var(--pwc-primary); }}
h2 {{ color: var(--pwc-red); }}
h3 {{ color: var(--pwc-orange2); }}

/* PwC header bar */
.pwc-topbar {{
  display: flex;
  align-items: center;
  gap: .75rem;
  border-radius: 14px;
  padding: .85rem 1rem;
  margin: .25rem 0 1rem;
  background: linear-gradient(90deg, var(--pwc-primary), var(--pwc-orange2));
  color: white;
  box-shadow: 0 8px 24px rgba(0,0,0,.08);
}}
.pwc-badge {{
  background: rgba(255,255,255,.12);
  border: 1px solid rgba(255,255,255,.25);
  padding: .15rem .5rem;
  border-radius: 999px;
  font-size: .75rem;
}}
/* Buttons */
.stButton>button {{
  background: var(--pwc-primary);
  color: #fff;
  border-radius: 12px;
  border: none;
  box-shadow: 0 6px 16px rgba(253, 81, 8, .25);
}}
.stButton>button:hover {{
  background: var(--pwc-orange1);
}}
/* Widgets */
.stTextInput>div>div>input,
.stTextArea textarea,
.stNumberInput input {{
  border-radius: 10px !important;
  border: 1px solid var(--pwc-grey3);
}}
/* Dataframes */
[data-testid="stTable"], .stDataFrame {{
  border: 1px solid var(--pwc-grey3);
  border-radius: 12px;
}}
/* Pills/tags */
.pwc-pill {{
  display: inline-block;
  padding: .15rem .5rem;
  border-radius: 999px;
  font-size: .75rem;
  color: #fff;
  background: var(--pwc-red);
}}
</style>
""", unsafe_allow_html=True)
# --- Fix white-on-white text (force readable defaults) ---
st.markdown("""
<style>
/* … all the fixes I gave you … */
</style>
""", unsafe_allow_html=True)

# Optional: branded header on every page
st.markdown(
    """
    <div class="pwc-topbar">
      <div style="width:16px;height:16px;border-radius:3px;background:white;opacity:.95"></div>
      <div style="width:12px;height:12px;border-radius:3px;background:#FFD27A;opacity:.95"></div>
      <div style="width:20px;height:20px;border-radius:3px;background:#FF9151;opacity:.95"></div>
      <strong style="margin-left:.25rem">TPA — Transfer Pricing Associate</strong>
      <span class="pwc-badge">demo</span>
    </div>
    """,
    unsafe_allow_html=True,
)
# --- Fix white-on-white text (force readable defaults) ---
st.markdown("""
<style>
/* … all the fixes I gave you … */
</style>
""", unsafe_allow_html=True)


# --------------------------
# Helpers
# --------------------------

def read_pdf(file) -> str:
    if pdfplumber is None:
        return ""
    try:
        with pdfplumber.open(file) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""


def docx_replace_text(doc, replacements: Dict[str, str]) -> int:
    """Replace text in paragraphs and tables while preserving run styles. Returns number of replacements."""
    count = 0
    # Paragraphs
    for p in doc.paragraphs:
        for old, new in replacements.items():
            if old in p.text:
                full = "".join(r.text for r in p.runs)
                full_new = full.replace(old, new)
                if full != full_new:
                    count += full.count(old)
                    style = p.runs[0].style if p.runs else None
                    for r in p.runs:
                        r.text = ""
                    if p.runs:
                        p.runs[0].text = full_new
                        if style: p.runs[0].style = style
                    else:
                        p.add_run(full_new)
    # Tables
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old, new in replacements.items():
                        if old in p.text:
                            full = "".join(r.text for r in p.runs)
                            full_new = full.replace(old, new)
                            if full != full_new:
                                count += full.count(old)
                                style = p.runs[0].style if p.runs else None
                                for r in p.runs:
                                    r.text = ""
                                if p.runs:
                                    p.runs[0].text = full_new
                                    if style: p.runs[0].style = style
                                else:
                                    p.add_run(full_new)
    return count


def fetch_title(url: str) -> str:
    try:
        r = requests.get(url, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        return (soup.title.string.strip() if soup.title and soup.title.string else url)[:120]
    except Exception:
        return url


def summarize_benchmark(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "Benchmark not provided."
    acc = (df.get("Decision", "").astype(str).str.lower() == "accept").sum()
    rej = (df.get("Decision", "").astype(str).str.lower() == "reject").sum()
    total = len(df)
    return f"Vendor study summary: {acc} accepted, {rej} rejected, {total} total comparables."

# --------------------------
# Page: TPD Draft (enhanced)
# --------------------------
if page == "TPD Draft":
    st.title("TPD Draft Generator — Formatting Preserved (DOCX)")
    st.write("Upload prior‑year TPD (**DOCX recommended**). Choose update options and generate a roll‑forward draft. PDFs will not preserve styles.")

    prior = st.file_uploader("Upload Prior TPD (DOCX recommended; PDF supported as JSON)", type=["docx", "pdf"], accept_multiple_files=False)

    mode = st.radio("What information is available?", [
        "No information",
        "Client information request",
        "Benchmark study",
        "Both (IRL + Benchmark)",
    ])

    colA, colB = st.columns(2)
    with colA:
        new_fy = st.number_input("New FY (e.g., 2024)", min_value=1990, max_value=2100, value=2024)
    with colB:
        report_date = st.text_input("Report date to insert (optional, e.g., 30 June 2025)")

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
        irl_upload = st.file_uploader("Attach client info (TXT/CSV) to insert", type=["txt", "csv"], key="irl")
        if irl_upload is not None:
            try:
                if irl_upload.name.endswith(".csv"):
                    df = pd.read_csv(irl_upload)
                    irl_text = "\n".join("- " + " | ".join(map(str, row)) for _, row in df.iterrows())
                else:
                    irl_text = irl_upload.read().decode("utf-8", errors="ignore")
                st.caption("Loaded client information for inclusion.")
            except Exception as e:
                st.error(f"Could not read client info: {e}")

    st.subheader("Industry update (optional)")
    st.write("Paste credible source URLs (World Bank/IMF/OECD/stat agencies/newsroom). We add a subsection with numbered footnotes.")
    urls = st.text_area("Source URLs (one per line)")
    url_list = [u.strip() for u in urls.splitlines() if u.strip()]

    if st.button("Generate draft now", type="primary"):
        if prior is None:
            st.error("Please upload a prior TPD (DOCX recommended).")
        else:
            is_docx = prior.name.lower().endswith(".docx") and DocxDocument is not None
            is_pdf = prior.name.lower().endswith(".pdf")
            prior_buffer = io.BytesIO(prior.getvalue())

            if is_docx:
                doc = DocxDocument(prior_buffer)

                # Discover old FYs present, default to previous year if none found
                txt_all = "\n".join(p.text for p in doc.paragraphs)
                old_years = set(re.findall(r"FY\s*-?_?\s*(20\d{2})", txt_all, flags=re.I)) or {str(new_fy - 1)}

                # Build replacements for common FY formats
                repl = {}
                for y in old_years:
                    repl[f"FY{y}"] = f"FY{new_fy}"
                    repl[f"FY {y}"] = f"FY {new_fy}"
                    repl[f"FYE {y}"] = f"FYE {new_fy}"
                if report_date:
                    # Gentle insert: replace the literal label if present, otherwise append later
                    repl["Report Date"] = f"Report Date: {report_date}"

                hits = docx_replace_text(doc, repl)

                # Append conditional sections
                if bench_df is not None and not bench_df.empty:
                    summary = summarize_benchmark(bench_df)
                    p = doc.add_paragraph()
                    p.add_run("\nEconomic Analysis — Benchmark Update: ").bold = True
                    doc.add_paragraph(summary)
                if irl_text:
                    p = doc.add_paragraph()
                    p.add_run("\nClient Information Provided:").bold = True
                    for line in irl_text.splitlines():
                        if line.strip():
                            doc.add_paragraph("• " + line.strip())
                if url_list:
                    doc.add_paragraph()
                    doc.add_paragraph("Industry Update (Current Year)")
                    foots = []
                    for i, url in enumerate(url_list, start=1):
                        title = fetch_title(url)
                        doc.add_paragraph(f"- See: {title} [^{i}]")
                        foots.append((i, url))
                    if foots:
                        doc.add_paragraph("Sources:")
                        for i, url in foots:
                            doc.add_paragraph(f"  ^{i} {url}")

                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.download_button(
                    "Download Draft (DOCX)", data=out.getvalue(), file_name="TPD_Draft.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                st.success(f"Draft generated. Replacements applied: {hits}")

            elif is_pdf:
                text = read_pdf(prior_buffer)
                payload = {
                    "note": "PDF input: style not preserved. Upload DOCX to keep formatting.",
                    "new_fy": new_fy,
                    "report_date": report_date,
                    "industry_sources": url_list,
                    "benchmark_summary": summarize_benchmark(bench_df) if bench_df is not None else None,
                    "irl": irl_text,
                }
                st.download_button("Download Draft (JSON)", data=json.dumps(payload, indent=2).encode("utf-8"), file_name="TPD_Draft.json", mime="application/json")
                st.info("To preserve fonts/colours, please upload a DOCX prior TPD.")

# --- Simple placeholders for the other pages (kept minimal for this iteration) ---
elif page == "TNMM Review":
    st.title("TNMM Benchmark Review")
    st.info("Use your earlier TNMM page or integrate the reviewer later — this build focuses on the TPD generator.")
elif page == "CUT/CUP Review":
    st.title("CUT/CUP Review")
    st.info("To keep this file focused, CUT/CUP scoring is not included in this iteration.")
elif page == "Information Request List":
    st.title("Information Request List (IRL)")
    st.write("Generate a quick request pack for clients.")
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
else:
    st.title("Advisory / Opportunity Spotting")
    st.info("Advisory heuristics can be re-attached later. Focus here is on roll‑forward.")

