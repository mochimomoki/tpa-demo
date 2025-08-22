if page == "TPD Draft":
    st.title("TPD Draft Generator")
    st.write("Upload prior-year TPD as **Microsoft Word (.docx or .doc)** to preserve fonts/colours/sizes. PDFs are supported but styles cannot be preserved.")

    # 1) Upload prior TPD
    prior = st.file_uploader(
        "Upload Prior TPD (DOCX/DOC preferred; PDF supported as JSON fallback)",
        type=["docx", "doc", "pdf"],
        accept_multiple_files=False
    )

    # 2) New FY (and report date)
    colA, colB = st.columns(2)
    with colA:
        new_fy = st.number_input("New FY (e.g., 2024)", min_value=1990, max_value=2100, value=2024)
    with colB:
        report_date = st.text_input("Report date (optional, e.g., 30 June 2025)", value="")

    # 3) Type of information available
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

    # 4) Industry analysis mode
    st.subheader("Industry Analysis Mode")
    industry_mode = st.radio(
        "How should we handle Industry Analysis?",
        ["Roll-forward (update facts & stats)", "Full Rewrite"],
        help="Roll-forward: update outdated numbers and citations only, keeping prior narrative. Full Rewrite: rebuild the section from scratch."
    )

    # Detect industry from prior TPD text (uses helpers from earlier in the file)
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

    # 5) Industry sources (country + optional URLs/uploads)
    st.subheader("Industry sources (optional)")
    override_country = st.text_input("Country for auto research", value="Singapore")
    urls = st.text_area("Extra source URLs (one per line, optional)", value="")
    user_url_list = [u.strip() for u in urls.splitlines() if u.strip()]
    user_reports = st.file_uploader("Upload market/industry reports (PDF/DOCX/TXT — optional)", type=["pdf","docx","txt"], accept_multiple_files=True)

    # Advanced text replacements
    adv = st.expander("Advanced: custom replacements (JSON)", expanded=False)
    with adv:
        st.write('Example: {"{{ENTITY}}": "ABC Pte Ltd", "{{COUNTRY}}": "Singapore"}')
        repl_json = st.text_area("Key-value JSON (optional)", value="")

    # Generate
    if st.button("Generate TPD draft now", type="primary"):
        if prior is None:
            st.error("Please upload a prior TPD (Word .docx/.doc preferred).")
        else:
            # Auto sector research (default credible sources)
            sector_pack = auto_sector_research(override_country, industry_choice)
            auto_lines, auto_foots = format_sector_update_text(sector_pack)

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

            # Prepare file routing
            name = prior.name.lower()
            is_docx = name.endswith(".docx") and (DocxDocument is not None)
            is_doc = name.endswith(".doc")
            is_pdf = name.endswith(".pdf")

            prior_buffer = io.BytesIO(prior.getvalue())
            if is_doc:
                try:
                    converted = convert_doc_to_docx_bytes(prior.getvalue())
                    prior_buffer = io.BytesIO(converted)
                    is_docx = True
                    st.info("Converted legacy .doc file to .docx for processing.")
                except Exception as e:
                    st.error(str(e))
                    is_docx = False

            # DOCX flow (preserve formatting)
            if is_docx:
                if DocxDocument is None:
                    st.error("python-docx is not available in this environment.")
                else:
                    doc = DocxDocument(prior_buffer)
                    auto_repl = build_rollforward_replacements(doc, int(new_fy), report_date.strip())
                    auto_repl.update(user_repl)
                    hits = docx_replace_text_everywhere(doc, auto_repl)

                    # Inserts for selected info
                    if bench_df is not None and not bench_df.empty:
                        acc = (bench_df.get("Decision", "").astype(str).str.lower() == "accept").sum()
                        rej = (bench_df.get("Decision", "").astype(str).str.lower() == "reject").sum()
                        summary = f"Vendor study summary: {acc} accepted, {rej} rejected, {len(bench_df)} total comparables."
                        p = doc.add_paragraph()
                        p.add_run("\nEconomic Analysis — Benchmark Update: ").bold = True
                        doc.add_paragraph(summary)

                    if irl_text:
                        p = doc.add_paragraph()
                        p.add_run("\nClient Information Provided:").bold = True
                        for line in irl_text.splitlines():
                            if line.strip():
                                doc.add_paragraph("• " + line.strip())

                    # Industry Update (mode-aware)
                    doc.add_paragraph()
                    doc.add_paragraph(f"Industry Update — {industry_choice}")
                    if industry_mode.startswith("Roll-forward"):
                        doc.add_paragraph("The prior-year narrative is retained. The facts and figures below are refreshed for the current period:")

                    if auto_lines:
                        for ln in auto_lines:
                            doc.add_paragraph(f"- {ln}")

                    foots: List[Tuple[int, str]] = []
                    if auto_foots:
                        foots.extend(auto_foots)

                    if user_url_list:
                        for u in user_url_list:
                            title = fetch_title(u)
                            doc.add_paragraph(f"- See: {title}")
                            foots.append((len(foots) + 1, u))

                    if user_reports:
                        for f in user_reports:
                            label = getattr(f, "name", "uploaded report")
                            doc.add_paragraph(f"- See: {label}")
                            foots.append((len(foots) + 1, f"uploaded://{label}"))

                    if foots:
                        doc.add_paragraph("Sources:")
                        for i, url in foots:
                            doc.add_paragraph(f"  ^{i} {url}")

                    out = io.BytesIO(); doc.save(out); out.seek(0)
                    st.download_button(
                        "Download Draft (DOCX)",
                        data=out.getvalue(),
                        file_name="TPD_Draft.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                    st.success(f"Draft generated. Replacements applied: {hits}")

            # PDF flow (JSON fallback)
            elif is_pdf:
                text = read_pdf(prior_buffer)
                payload = {
                    "note": "PDF input: style not preserved. Upload .docx to keep formatting.",
                    "new_fy": int(new_fy),
                    "report_date": report_date.strip(),
                    "industry": industry_choice,
                    "country": override_country,
                    "analysis_mode": industry_mode,
                    "auto_research": {"lines": auto_lines, "sources": [u for _, u in auto_foots]},
                    "irl": irl_text,
                }
                if bench_df is not None and not bench_df.empty:
                    acc = (bench_df.get("Decision", "").astype(str).str.lower() == "accept").sum()
                    rej = (bench_df.get("Decision", "").astype(str).str.lower() == "reject").sum()
                    payload["benchmark_summary"] = f"{acc} accepted, {rej} rejected, {len(bench_df)} total"
                if user_url_list:
                    payload["user_sources"] = user_url_list
                if user_reports:
                    payload["uploaded_reports"] = [getattr(f, "name", "report") for f in user_reports]

                st.download_button(
                    "Download Draft (JSON)",
                    data=json.dumps(payload, indent=2).encode("utf-8"),
                    file_name="TPD_Draft.json",
                    mime="application/json",
                )
                st.info("To preserve fonts/colours, please upload a Word .docx file.")

            else:
                st.error("Unsupported file type. Please upload .docx, .doc, or .pdf.")




