# app.py — TPA TPD Generator (DOCX formatting preserved, .DOC conversion, industry sources, modes)
from __future__ import annotations
import io, re, json, os, subprocess, tempfile
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
    from docx.text.paragraph import Paragraph
except Exception:
    DocxDocument = None  # will guard later

import requests
from bs4 import BeautifulSoup

st.set_page_config(page_title="TPA — TPD Generator", layout="wide")
st.sidebar.title("TPA (Transfer Pricing Associate)")
st.sidebar.caption("TPD roll-forward with formatting + .doc support + industry sources")

page = st.sidebar.radio(
    "Choose function",
    ["TPD Draft", "TNMM Review", "CUT/CUP Review", "Information Request List", "Advisory / Opportunity Spotting"],
)

# --------------------------
# Helpers: PDF + Industry
# --------------------------
def read_pdf(file_like) -> str:
    if pdfplumber is None:
        return ""
    try:
        with pdfplumber.open(file_like) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""

def fetch_title(url: str) -> str:
    try:
        r = requests.get(url, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        title = soup.title.string.strip() if soup.title and soup.title.string else url
        return title[:120]
    except Exception:
        return url

# --------------------------
# Helpers: .DOC → .DOCX conversion (best-effort)
# --------------------------
def _try_libreoffice_convert(doc_bytes: bytes) -> Optional[bytes]:
    """Use LibreOffice (soffice) to convert .doc → .docx, if available."""
    with tempfile.TemporaryDirectory() as tmpdir:
        in_path = os.path.join(tmpdir, "in.doc")
        out_path = os.path.join(tmpdir, "in.docx")
        with open(in_path, "wb") as f:
            f.write(doc_bytes)
        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "docx", "--outdir", tmpdir, in_path],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            if os.path.exists(out_path):
                with open(out_path, "rb") as f:
                    return f.read()
        except Exception:
            return None
    return None

def _try_pandoc_convert(doc_bytes: bytes) -> Optional[bytes]:
    """Use pypandoc/Pandoc to convert .doc → .docx (requires pandoc)."""
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
    """
    Convert legacy .doc to .docx.
    Try LibreOffice → Pandoc. If both unavailable, raise a helpful error.
    """
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

# --------------------------
# Helpers: DOCX formatting-preserving replacements
# --------------------------
def _iter_all_paragraphs(doc):
    """Yield all paragraphs from body, tables, headers, and footers."""
    # Body
    for p in doc.paragraphs:
        yield p
    # Body tables
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    # Headers/Footers
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
    """Replace text inside a paragraph, keeping the first run's style."""
    if old not in paragraph.text:
        return 0
    runs = paragraph.runs
    full = "".join(r.text for r in runs)
    n = full.count(old)
    full_new = full.replace(old, new)
    if runs:
        style = runs[0].style
        for r in runs:
            r.text = ""
        runs[0].text = full_new
        runs[0].style = style
    else:
        paragraph.add_run(full_new)
    return n

def docx_replace_text_everywhere(doc: "DocxDocument", replacements: Dict[str, str]) -> int:
    """Apply replacements across body, tables, headers, and footers."""
    total = 0
    for p in _iter_all_paragraphs(doc):
        for old, new in replacements.items():
            total += _replace_preserving_style(p, old, new)
    return total

def detect_years(text: str) -> set:
    """Detect FY years and ranges like 2023/24."""
    years = set(re.findall(r"(?:FY\\s*-?_?\\s*)?(20\\d{2})", text, flags=re.I))
    # ranges e.g. 2023/24
    ranges = re.findall(r"(?:FY\\s*)?(20\\d{2})\\s*/\\s*(\\d{2})", text, flags=re.I)
    for y, yy in ranges:
        try:
            y1 = int(y)
            y2 = (y1 // 100) * 100 + int(yy)
            years.add(str(y1))
            years.add(str(y2))
        except Exception:
            pass
    return years

def bump_range_token(token: str, new_start_year: int) -> str:
    """Convert tokens like '2023/24' → '2024/25' (honors 'FY ' prefix if present)."""
    m = re.search(r"(20\\d{2})\\s*/\\s*(\\d{2})", token)
    if not m:
        return token
    y1 = new_start_year
    y2 = (y1 % 100) + 1
    return re.sub(r"(20\\d{2})\\s*/\\s*(\\d{2})", f"{y1}/{y2:02d}", token)

def build_rollforward_replacements(doc: "DocxDocument", new_fy: int, report_date: str) -> Dict[str, str]:
    """
    Build a conservative replacement map:
      FY2023 → FY2024
      FY 2023 → FY 2024
      FYE 2023 → FYE 2024
      2023/24 → 2024/25
      (selected labeled standalone years only)
    """
    text_all = "\n".join(p.text for p in _iter_all_paragraphs(doc))
    years = detect_years(text_all) or {str(new_fy - 1)}

    repl: Dict[str, str] = {}
    for y in years:
        repl[f"FY{y}"] = f"FY{new_fy}"
        repl[f"FY {y}"] = f"FY {new_fy}"
        repl[f"FYE {y}"] = f"FYE {new_fy}"
        # labeled standalone years (avoid global year replace)
        repl[f"Financial Year {y}"] = f"Financial Year {new_fy}"
        repl[f"Fiscal Year {y}"] = f"Fiscal Year {new_fy}"

    # year ranges like 2023/24 (and FY 2023/24)
    tokens = re.findall(r"(?:FY\\s*)?(20\\d{2}\\s*/\\s*\\d{2})", text_all)
    for t in set(tokens):
        repl[t] = bump_range_token(t, new_fy)

    if report_date:
        # Only replace the "Report Date" label if it exists; we don't inject new paragraphs here
        repl["Report Date"] = f"Report Date: {report_date}"

    return repl

def summarize_benchmark(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "Benchmark not provided."
    acc = (df.get("Decision", "").astype(str).str.lower() == "accept").sum()
    rej = (df.get("Decision", "").astype(str).str.lower() == "reject").sum()
    total = len(df)
    return f"Vendor study summary: {acc} accepted, {rej} rejected, {total} total comparables."

# --------------------------
# PAGE: TPD Draft (full)
# --------------------------
if page == "TPD Draft":
    st.title("TPD Draft Generator — Preserve Formatting (Word) + Industry Sources")
    st.write("Upload prior-year TPD as **Microsoft Word (.docx or .doc)** to preserve fonts/colours/sizes. PDFs are supported but styles cannot be preserved.")

    prior = st.file_uploader(
        "Upload Prior TPD (DOCX/DOC preferred; PDF supported as JSON fallback)",
        type=["docx", "doc", "pdf"],
        accept_multiple_files=False
    )

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
        report_date = st.text_input("Report date to insert (optional, e.g., 30 June 2025)", value="")

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

    st.subheader("Industry update (optional)")
    st.write("Paste credible source URLs (World Bank/IMF/OECD/stat agencies/newsroom). We’ll add a subsection with numbered footnotes.")
    urls = st.text_area("Source URLs (one per line)", value="")
    url_list = [u.strip() for u in urls.splitlines() if u.strip()]

    # Advanced: user-defined replacements
    adv = st.expander("Advanced: custom replacements (JSON)", expanded=False)
    with adv:
        st.write('Example: {"{{ENTITY}}": "ABC Pte Ltd", "{{COUNTRY}}": "Singapore"}')
        repl_json = st.text_area("Key-value JSON (optional)", value="")

    if st.button("Generate TPD draft now", type="primary"):
        if prior is None:
            st.error("Please upload a prior TPD (Word .docx/.doc preferred).")
        else:
            name = prior.name.lower()
            is_docx = name.endswith(".docx") and (DocxDocument is not None)
            is_doc = name.endswith(".doc")
            is_pdf = name.endswith(".pdf")

            # Parse advanced replacements
            user_repl: Dict[str, str] = {}
            if repl_json.strip():
                try:
                    user_repl = json.loads(repl_json)
                    if not isinstance(user_repl, dict):
                        st.warning("Custom replacements must be a JSON object (key-value). Ignored.")
                        user_repl = {}
                except Exception:
                    st.warning("Invalid JSON for custom replacements. Ignored.")
                    user_repl = {}

            # Convert legacy .doc to .docx if needed
            prior_buffer = io.BytesIO(prior.getvalue())
            if is_doc:
                try:
                    converted = convert_doc_to_docx_bytes(prior.getvalue())
                    prior_buffer = io.BytesIO(converted)
                    is_docx = True
                    st.info("Converted legacy .doc file to .docx for processing.")
                except Exception as e:
                    st.error(str(e))
                    is_docx = False  # will block DOCX flow and avoid crash

            if is_docx:
                if DocxDocument is None:
                    st.error("python-docx is not available in this environment.")
                else:
                    doc = DocxDocument(prior_buffer)

                    # Build roll-forward replacements and merge user overrides
                    auto_repl = build_rollforward_replacements(doc, int(new_fy), report_date.strip())
                    auto_repl.update(user_repl)

                    # Apply replacements in body/tables/headers/footers
                    hits = docx_replace_text_everywhere(doc, auto_repl)

                    # Conditional inserts
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
                        "Download Draft (DOCX)",
                        data=out.getvalue(),
                        file_name="TPD_Draft.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                    st.success(f"Draft generated. Replacements applied: {hits}")

            elif is_pdf:
                # Style cannot be preserved from PDFs — return a JSON draft
                text = read_pdf(prior_buffer)
                payload = {
                    "note": "PDF input: style not preserved. Upload .docx to keep formatting.",
                    "new_fy": int(new_fy),
                    "report_date": report_date.strip(),
                    "industry_sources": url_list,
                    "irl": irl_text,
                }
                if bench_df is not None and not bench_df.empty:
                    payload["benchmark_summary"] = summarize_benchmark(bench_df)
                st.download_button(
                    "Download Draft (JSON)",
                    data=json.dumps(payload, indent=2).encode("utf-8"),
                    file_name="TPD_Draft.json",
                    mime="application/json",
                )
                st.info("To preserve fonts/colours, please upload a Word .docx file.")
            else:
                st.error("Unsupported file type. Please upload .docx, .doc, or .pdf.")

# --------------------------
# PAGE: TNMM (kept simple)
# --------------------------
elif page == "TNMM Review":
    st.title("TNMM Benchmark Review (quick demo)")
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

# --------------------------
# PAGE: CUT/CUP (placeholder)
# --------------------------
elif page == "CUT/CUP Review":
    st.title("CUT / CUP Agreements Review (demo)")
    st.info("Clause extraction + scoring can be wired next; current focus is the TPD generator.")

# --------------------------
# PAGE: IRL
# --------------------------
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

# --------------------------
# PAGE: Advisory (simple)
# --------------------------
else:
    st.title("Advisory / Opportunity Spotting (demo)")
    st.info("Upload a benchmark on the TNMM page to explore opportunities; simplified here.")

