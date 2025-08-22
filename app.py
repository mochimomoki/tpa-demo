# app.py — TPA Core Demo with real TPD roll-forward
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


st.set_page_config(page_title="TPA (Transfer Pricing Associate)", layout="wide")
st.sidebar.title("TPA (Transfer Pricing Associate)")
st.sidebar.write("AI helper for transfer pricing associates")

page = st.sidebar.radio(
    "Choose function",
    ["TNMM Review", "CUT/CUP Review", "TPD Draft", "Information Request List", "Advisory / Opportunity Spotting"],
)


# --------------------------
# Utility parsers/extractors
# --------------------------
def read_pdf(file) -> str:
    if pdfplumber is None:
        return ""
    try:
        with pdfplumber.open(file) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""


def read_docx(file) -> str:
    """Read docx text using python-docx (simple paragraph join)."""
    if DocxDocument is None:
        return ""
    try:
        doc = DocxDocument(file)
        return "\n".join(p.text for p in doc.paragraphs)
    except Exception:
        return ""


def extract_fallback_context(text: str) -> Dict[str, Any]:
    """Very light heuristics to pull FAR/policy/industry hints from prior TPD text."""
    t = (text or "").lower()

    # crude FAR detection
    is_dist_or_services = any(k in t for k in ["distribution", "distributor", "routine services", "shared services", "support services", "back office"])
    has_manu = any(k in t for k in ["manufactur", "plant", "factory"])
    owns_ip = any(k in t for k in ["intangible", "ip ownership", "patent", "trademark", "r&d"])

    functions = "Routine distribution/services" if is_dist_or_services and not has_manu else (
        "Manufacturing (possible) + distribution" if has_manu else "Services"
    )
    assets = "No unique intangibles" if not owns_ip else "Owns / exploits intangibles"
    risks = "Limited" if "limited risk" in t or "limited-risk" in t or "low risk" in t else "Standard commercial risks"

    # guess method
    method = "TNMM" if "tnmm" in t or "transactional net margin" in t else ("CUP" if "cup" in t or "cut" in t else "TNMM")
    tested_party = "Local entity" if "tested party" in t or "local entity" in t else "Least complex entity"

    # pull a few sentences as industry extract
    lines = [ln.strip() for ln in (text or "").splitlines() if ln and len(ln.strip()) > 30]
    industry_sample = lines[:5]

    return {
        "far": {"functions": functions, "assets": assets, "risks": risks},
        "policy": {"method": method, "tested_party": tested_party, "rationale": "Least complex entity", "markup_or_margin": "TBD"},
        "industry": {"extract": industry_sample},
    }


def generate_rollforward_docx(ctx: Dict[str, Any], comps: pd.DataFrame | None, cutcup: pd.DataFrame | None) -> bytes:
    """Create a simple but presentable DOCX draft using python-docx. JSON fallback if missing."""
    if DocxDocument is None:
        # Fallback JSON blob if docx not available
        payload = {"context": ctx, "accepted": [], "rejected": [], "cutcup": []}
        if isinstance(comps, pd.DataFrame) and not comps.empty:
            payload["accepted"] = comps[comps["Decision"].astype(str).str.lower() == "accept"].to_dict("records")
            payload["rejected"] = comps[comps["Decision"].astype(str).str.lower() == "reject"].to_dict("records")
        if isinstance(cutcup, pd.DataFrame) and not cutcup.empty:
            payload["cutcup"] = cutcup.to_dict("records")
        return json.dumps(payload, indent=2).encode("utf-8")

    doc = DocxDocument()
    doc.add_heading("Transfer Pricing Documentation — Draft (Roll-forward)", level=1)

    # 1. Executive summary
    doc.add_heading("1. Executive Summary", level=2)
    doc.add_paragraph(
        "This draft updates the prior-year transfer pricing documentation using refreshed business context, "
        "policy confirmation, and a preliminary benchmarking and agreement review. Figures and margins are placeholders."
    )

    # 2. Group / Tested Party Overview (FAR)
    doc.add_heading("2. Tested Party Overview (FAR)", level=2)
    far = ctx.get("far", {})
    doc.add_paragraph(f"Functions: {far.get('functions', 'TBD')}")
    doc.add_paragraph(f"Assets: {far.get('assets', 'TBD')}")
    doc.add_paragraph(f"Risks: {far.get('risks', 'TBD')}")

    # 3. TP Policy
    doc.add_heading("3. Transfer Pricing Policy", level=2)
    pol = ctx.get("policy", {})
    doc.add_paragraph(f"Method: {pol.get('method','TNMM')}")
    doc.add_paragraph(f"Tested Party: {pol.get('tested_party','Local entity')}")
    doc.add_paragraph(f"Rationale: {pol.get('rationale','Least complex entity')}")
    doc.add_paragraph(f"Target markup/margin: {pol.get('markup_or_margin','TBD')}")

    # 4. Industry
    doc.add_heading("4. Industry Overview (Extracts)", level=2)
    for ln in ctx.get("industry", {}).get("extract", [])[:6]:
        doc.add_paragraph(f"• {ln}")

    # 5. Benchmarking summary (if any data provided)
    doc.add_heading("5. Economic Analysis — Benchmarking (Preview)", level=2)
    if isinstance(comps, pd.DataFrame) and not comps.empty:
        accepts = comps[comps["Decision"].astype(str).str.lower() == "accept"].to_dict("records")
        rejects = comps[comps["Decision"].astype(str).str.lower() == "reject"].to_dict("records")
        doc.add_paragraph("Accepted comparables (sample):")
        for c in accepts[:8]:
            doc.add_paragraph(f"  - {c.get('Company','?')} — {c.get('Reason','')}")
        doc.add_paragraph("Rejected comparables (sample):")
        for c in rejects[:8]:
            doc.add_paragraph(f"  - {c.get('Company','?')} — {c.get('Reason','')}")

        doc.add_paragraph("Note: full numerical analysis, filters and financial tables to be appended.")
    else:
        doc.add_paragraph("Benchmarking not yet provided in this draft.")

    # 6. Intercompany Agreements — CUT/CUP (if any)
    doc.add_heading("6. Intercompany Agreements (CUT/CUP — Preview)", level=2)
    if isinstance(cutcup, pd.DataFrame) and not cutcup.empty:
        top = cutcup.iloc[0].to_dict()
        doc.add_paragraph(
            f"Top match: {top.get('Agreement','agreement.pdf')} "
            f"({top.get('Score %', '—')}%) — {top.get('Notes','')}"
        )
    else:
        doc.add_paragraph("No agreements uploaded for analysis in this draft.")

    # 7. Compliance and Documentation
    doc.add_heading("7. Compliance & Documentation", level=2)
    doc.add_paragraph(
        "This draft is prepared for management review. Local documentation format and specific "
        "disclosure requirements will be applied in the final version."
    )

    # 8. Conclusion
    doc.add_heading("8. Conclusion", level=2)
    doc.add_paragraph("Please review, confirm any business changes, and provide financials to finalize margins.")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()


# --------------------------
# Page: TNMM Review (simple)
# --------------------------
if page == "TNMM Review":
    st.title("TNMM Benchmark Review")
    st.write("Upload a benchmark file and the app will flag quick inconsistencies.")

    file = st.file_uploader("Upload Benchmark (CSV/XLSX)", type=["csv", "xlsx"])
    df = None
    if file:
        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
            st.dataframe(df)
            if "Decision" in df.columns and "Reason" in df.columns:
                flags = df[(df["Decision"].astype(str).str.lower() == "reject") & (df["Reason"].astype(str).str.strip() == "")]
                st.subheader("⚠️ Potential Issues (rejects with empty reason)")
                st.dataframe(flags)
                if not flags.empty:
                    st.download_button("Download flags (CSV)", flags.to_csv(index=False).encode("utf-8"),
                                       file_name="tnmm_flags.csv", mime="text/csv")
        except Exception as e:
            st.error(f"Could not read file: {e}")


# --------------------------
# Page: CUT/CUP (placeholder)
# --------------------------
elif page == "CUT/CUP Review":
    st.title("CUT / CUP Agreements Review")
    st.write("Upload agreements (PDF) and set parameters to score them (light demo).")
    uploads = st.file_uploader("Upload Agreements (PDF)", type=["pdf"], accept_multiple_files=True)
    st.text_input("Territory must contain", value="apac", key="territory")
    st.checkbox("Require exclusivity", value=False, key="excl")
    st.number_input("Royalty min %", min_value=0.0, value=2.0, step=0.5, key="rmin")
    st.number_input("Royalty max %", min_value=0.0, value=8.0, step=0.5, key="rmax")
    st.number_input("Minimum term (years)", min_value=0, value=3, step=1, key="term")
    st.info("(Light demo — scoring not fully wired in this minimal build.)")


# --------------------------
# Page: TPD Draft (real roll-forward)
# --------------------------
elif page == "TPD Draft":
    st.title("TPD Draft Generator")
    st.write("Upload prior-year TPD (PDF or DOCX). The app will extract context and produce a roll-forward draft.")

    prior = st.file_uploader("Upload Prior TPD (PDF/DOCX)", type=["pdf", "docx"])
    comps_file = st.file_uploader("Optional: Upload Benchmark (CSV/XLSX) to include a short summary",
                                  type=["csv", "xlsx"], accept_multiple_files=False)

    comps_df: Optional[pd.DataFrame] = None
    if comps_file:
        try:
            comps_df = pd.read_csv(comps_file) if comps_file.name.endswith(".csv") else pd.read_excel(comps_file)
            st.caption("Loaded benchmark for summary section.")
        except Exception as e:
            st.error(f"Could not read benchmark: {e}")

    extracted = {}
    if prior is not None:
        ext_text = ""
        if prior.name.lower().endswith(".pdf"):
            ext_text = read_pdf(prior)
        else:
            ext_text = read_docx(prior)

        if not ext_text:
            st.warning("Could not read the uploaded file (ensure it's a readable PDF/DOCX).")
        else:
            extracted = extract_fallback_context(ext_text)
            with st.expander("Extracted context (preview)"):
                st.json(extracted)

    if st.button("Generate draft now", type="primary"):
        if prior is None:
            st.error("Please upload a prior TPD first.")
        else:
            output = generate_rollforward_docx(extracted or extract_fallback_context(""), comps_df, None)
            if DocxDocument is None:
                st.info("DOCX not available here — providing JSON fallback.")
                st.download_button("Download Draft (JSON)", data=output, file_name="TPD_Draft.json", mime="application/json")
            else:
                st.download_button(
                    "Download Draft (DOCX)",
                    data=output,
                    file_name="TPD_Draft.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            st.success("Draft generated.")


# --------------------------
# Page: IRL
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
        st.download_button("Download IRL (JSON)", data=json.dumps(out, indent=2).encode("utf-8"),
                           file_name="info_requests.json", mime="application/json")


# --------------------------
# Page: Advisory
# --------------------------
else:
    st.title("Advisory / Opportunity Spotting")
    st.write("Upload a benchmark to let the app flag quick opportunities.")
    file = st.file_uploader("Upload Benchmark (CSV/XLSX)", type=["csv", "xlsx"])
    if file:
        df = pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)
        rej = df[df.get("Decision", "").astype(str).str.lower() == "reject"]["Reason"].astype(str).str.lower()
        opps: List[Dict[str, str]] = []
        if (rej.str.contains("loss").sum() >= 2) or (rej.str.contains("functional").sum() >= 2):
            opps.append({"Type": "APA", "Severity": "High", "Rationale": "Repeated loss/functional disputes; consider APA."})
        if opps:
            st.dataframe(pd.DataFrame(opps))
        else:
            st.info("No obvious opportunities detected on this quick pass.")
