import streamlit as st
import pandas as pd

# --- Sidebar navigation ---
st.sidebar.title("TPA (Transfer Pricing Associate)")
st.sidebar.write("AI helper for transfer pricing associates")

page = st.sidebar.radio("Choose function", [
    "TNMM Review",
    "CUT/CUP Review",
    "TPD Draft",
    "Information Request List",
    "Advisory / Opportunity Spotting"
])

# --- TNMM Review ---
if page == "TNMM Review":
    st.title("TNMM Benchmark Review")
    st.write("Upload a benchmark file and let AI highlight inconsistent accept/reject decisions.")

    file = st.file_uploader("Upload Benchmark (CSV or Excel)", type=["csv", "xlsx"])
    if file:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        st.dataframe(df.head())

        # Dummy logic: flag rows with "Reject" but missing reason
        if "Decision" in df.columns and "Reason" in df.columns:
            flags = df[(df["Decision"] == "Reject") & (df["Reason"].isna())]
            st.subheader("⚠️ Potential Issues")
            st.dataframe(flags)

# --- CUT/CUP Review ---
elif page == "CUT/CUP Review":
    st.title("CUT / CUP Agreements Review")
    st.write("Upload agreements and set parameters. (Demo placeholder)")
    st.file_uploader("Upload Agreements (PDFs)", type=["pdf"], accept_multiple_files=True)
    st.text_input("Enter parameters (comma separated)")
    st.info("AI would parse agreements, extract clauses, and score comparables here.")

# --- TPD Draft ---
elif page == "TPD Draft":
    st.title("TPD Draft Generator")
    st.write("Upload prior-year TPD or info, AI will roll forward into a draft. (Demo placeholder)")
    st.file_uploader("Upload Prior TPD (PDF/DOCX)", type=["pdf","docx"])
    st.info("AI would generate new draft TPD here.")

# --- Information Request List ---
elif page == "Information Request List":
    st.title("Information Request List (IRL)")
    st.write("Generate questions to send clients for missing info.")
    industry = st.text_input("Enter industry")
    transactions = st.text_area("Enter transaction types")
    if st.button("Generate IRL"):
        st.success("Generated IRL:")
        st.write(f"- Please provide updated financials for {industry}")
        st.write(f"- Provide agreements and details for: {transactions}")

# --- Advisory ---
elif page == "Advisory / Opportunity Spotting":
    st.title("Advisory / Opportunity Spotting")
    st.write("AI will flag potential APA/BAPA/TP modelling opportunities. (Demo placeholder)")
    st.info("Example: Entity X has recurring losses, APA may be advisable.")
