import streamlit as st
import pandas as pd
from openai import OpenAI
import os
from dotenv import load_dotenv
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io

# ===== SETUP =====
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    st.error("OPENAI_API_KEY not found.")
    st.stop()

client = OpenAI(api_key=api_key)

st.set_page_config(
    page_title="ExplainMyData AI",
    page_icon="edai_logo.png",
    layout="wide"
)

# ===== SESSION STATE =====
if "usage_count" not in st.session_state:
    st.session_state.usage_count = 0

if "result" not in st.session_state:
    st.session_state.result = ""

if "translated_result" not in st.session_state:
    st.session_state.translated_result = ""

if "chat_answer" not in st.session_state:
    st.session_state.chat_answer = ""

# ===== STYLING =====
st.markdown("""
<style>
.stapp {
    background-color: #0f141a;
}

/* HEADER BACKGROUND */
.hero-wrap {
    background: linear-gradient(135deg, #a5d8ff 0%, #74c0fc 50%, #4dabf7 100%);
    border-radius: 24px;
    padding: 30px;
    margin-bottom: 25px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.25);
}

/* INNER CARD */
.hero-card {
    background: rgba(15,23,42,0.72);
    border-radius: 18px;
    padding: 20px;
    text-align: center;
}

/* TITLE */
.hero-title {
    font-size: 2.4rem;
    font-weight: 800;
    color: #f3f4f6 !important;
}

/* MOTTO */
.hero-motto {
    font-size: 1.1rem;
    color: #d1d5db !important;
    margin-bottom: 12px;
}

/* BADGES */
.hero-badge {
    display: inline-block;
    background: #e3f2fd;
    color: #0f172a;
    padding: 8px 14px;
    margin: 4px;
    border-radius: 999px;
    font-size: 0.9rem;
    font-weight: 600;
}

/* BUTTONS */
.stButton > button {
    border-radius: 10px !important;
    background-color: #21a366 !important;
    color: white !important;
    font-weight: 700 !important;
}

.stButton > button:hover {
    background-color: #188a54 !important;
    color: white !important;
}

/* TABS */
div[data-baseweb="tab"] {
    background-color: #1c2229 !important;
    color: #e6edf5 !important;
    border-radius: 10px 10px 0 0 !important;
    font-weight: 700 !important;
}

div[data-baseweb="tab"][aria-selected="true"] {
    background-color: #2563eb !important;
    color: white !important;
}

div[data-baseweb="tab-highlight"] {
    background-color: transparent !important;
}
</style>
""", unsafe_allow_html=True)

# ===== HEADER =====
header_left, header_center, header_right = st.columns([1, 2.5, 1])

with header_center:
    st.markdown('<div class="hero-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="hero-card">', unsafe_allow_html=True)

    st.image("edai_logo.png", width=140)

    st.markdown("""
    <div class="hero-title">ExplainMyData AI</div>
    <div class="hero-motto">Your data, clearly explained.</div>

    <div>
        <span class="hero-badge">AI-powered</span>
        <span class="hero-badge">Fast insights</span>
        <span class="hero-badge">Global translation</span>
        <span class="hero-badge">PDF export</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ===== FILE UPLOAD =====
uploaded_file = st.file_uploader("Upload your CSV or Excel file", type=["csv", "xlsx"])
st.info("⚡ For best performance, upload files under 50MB.")

# ===== PDF FUNCTION =====
def create_pdf_report(text):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    y = height - 50

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "ExplainMyData AI Report")
    y -= 30

    c.setFont("Helvetica", 10)
    for line in text.split("\n"):
        if y < 50:
            c.showPage()
            y = height - 50
        c.drawString(50, y, line[:100])
        y -= 14

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# ===== LOAD DATA =====
def load_data(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

# ===== MAIN =====
if uploaded_file is not None:
    df = load_data(uploaded_file)

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    categorical_cols = df.select_dtypes(include="object").columns.tolist()

    st.subheader("📊 Data Overview")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows", df.shape[0])
    c2.metric("Columns", df.shape[1])
    c3.metric("Missing", int(df.isna().sum().sum()))
    c4.metric("Duplicates", int(df.duplicated().sum()))

    st.markdown("---")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Overview", "Charts", "AI Insights", "Translate", "Ask & Export"]
    )

    # ===== OVERVIEW =====
    with tab1:
        st.dataframe(df.head())

    # ===== CHARTS =====
    with tab2:
        if numeric_cols:
            col = st.selectbox("Numeric column", numeric_cols)

            fig, ax = plt.subplots()
            df[col].dropna().plot(kind="hist", ax=ax)
            ax.set_title(col)
            ax.set_xlabel(col)
            st.pyplot(fig)

    # ===== AI INSIGHTS =====
    with tab3:
        st.write(f"Free analyses used: {st.session_state.usage_count}/2")

        if st.session_state.usage_count >= 2:
            st.error("Free limit reached. Upgrade to continue.")

        if st.button("Analyze Data"):
            if st.session_state.usage_count >= 2:
                st.stop()

            summary = df.describe(include="all").fillna("").to_string()

            response = client.responses.create(
                model="gpt-5",
                input=f"Analyze:\n{summary}"
            )

            st.session_state.result = response.output_text
            st.session_state.usage_count += 1

        if st.session_state.result:
            st.write(st.session_state.result)

    # ===== TRANSLATE =====
    with tab4:
        lang = st.selectbox("Language", ["French", "Spanish", "Arabic"])

        if st.button("Translate"):
            if st.session_state.result:
                response = client.responses.create(
                    model="gpt-5",
                    input=f"Translate to {lang}:\n{st.session_state.result}"
                )
                st.session_state.translated_result = response.output_text

        if st.session_state.translated_result:
            st.write(st.session_state.translated_result)

    # ===== ASK & EXPORT =====
    with tab5:
        question = st.text_input("Ask your data")

        if question:
            response = client.responses.create(
                model="gpt-5",
                input=question
            )
            st.write(response.output_text)

        st.markdown("---")

        if st.session_state.result:
            pdf = create_pdf_report(st.session_state.result)
            st.download_button("📄 Download PDF", pdf, "report.pdf")

        st.markdown("---")

        st.markdown("### 💎 Upgrade to Premium")

        premium_url = "PASTE_YOUR_STRIPE_LINK_HERE"
        st.link_button("Upgrade Now", premium_url)

else:
    st.warning("Upload a file to begin.")