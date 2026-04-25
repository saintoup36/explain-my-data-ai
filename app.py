import os
import io
import json
import uuid
import html
import zipfile
import locale
import shutil
import hashlib
import base64
from pathlib import Path
from datetime import datetime
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from dotenv import load_dotenv
from PIL import Image, ImageOps, ImageFilter
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from sklearn.linear_model import LinearRegression

try:
    from docx import Document
except Exception:
    Document = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="ExplainMyData AI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env", override=True)

APP_NAME = "ExplainMyData AI"
TAGLINE = "Your data, clearly explained"
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5")
SUPPORT_EMAIL = os.getenv("SUPPORT_EMAIL", "support@explainmydata.ai")
APP_BASE_URL = os.getenv("APP_BASE_URL", "").strip()
SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY", "").strip()
STRIPE_ONE_TIME_LINK = os.getenv(
    "STRIPE_ONE_TIME_LINK",
    "https://buy.stripe.com/fZucN7gO33mU8MegZIabK00",
).strip()
STRIPE_MONTHLY_LINK = os.getenv(
    "STRIPE_MONTHLY_LINK",
    "https://buy.stripe.com/6oU14papF4qY5A29xgabK02",
).strip()
STRIPE_SUCCESS_URL = os.getenv("STRIPE_SUCCESS_URL", "").strip()
PREMIUM_USERS = {
    email.strip().lower()
    for email in os.getenv("PREMIUM_USERS", "").split(",")
    if email.strip()
}
TESSERACT_PATH = os.getenv(
    "TESSERACT_PATH",
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
LOGO_PATH = BASE_DIR / "edai_logo.png"
FREE_ANALYSIS_LIMIT = 2
MAX_UPLOAD_MB = 200
SUPPORTED_UPLOADS = [
    "csv",
    "xlsx",
    "xls",
    "txt",
    "docx",
    "pdf",
    "png",
    "jpg",
    "jpeg",
    "webp",
    "zip",
]

if not OPENAI_API_KEY:
    try:
        OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", "")
    except Exception:
        OPENAI_API_KEY = ""

if pytesseract is not None and os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

client = OpenAI(api_key=OPENAI_API_KEY) if OpenAI and OPENAI_API_KEY else None

STORAGE_ROOT = BASE_DIR / "app_storage"
USER_UPLOADS_DIR = STORAGE_ROOT / "user_uploads"
USER_REPORTS_DIR = STORAGE_ROOT / "user_reports"
SHARED_REPORTS_DIR = STORAGE_ROOT / "shared_reports"
SUPPORT_DIR = STORAGE_ROOT / "support"

for folder in [
    STORAGE_ROOT,
    USER_UPLOADS_DIR,
    USER_REPORTS_DIR,
    SHARED_REPORTS_DIR,
    SUPPORT_DIR,
]:
    folder.mkdir(parents=True, exist_ok=True)


# =========================================================
# STATE
# =========================================================
def init_state() -> None:
    defaults = {
        "authenticated": False,
        "auth_user": None,
        "auth_mode": "local-demo",
        "stripe_last_payment": "",
        "usage_count": 0,
        "gdpr_consent": False,
        "app_language": "English",
        "result": "",
        "translated_result": "",
        "doctor_result": "",
        "chat_answer": "",
        "ocr_text": "",
        "ocr_preview": None,
        "support_answer": "",
        "last_saved_report_hash": "",
        "use_sample_data": False,
        "developer_premium_override": False,
        "premium_status": "inactive",
        "user_plan": "free",
        "forecast_df": None,
        "cleaned_df": None,
        "loaded_filename": "",
        "raw_text": "",
        "data_mode": "structured",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# =========================================================
# LABELS
# =========================================================
LANGUAGES = {
    "English": "en",
    "French": "fr",
    "Spanish": "es",
    "Haitian Creole": "ht",
    "Portuguese": "pt",
    "Arabic": "ar",
    "German": "de",
    "Mandarin Chinese": "zh",
    "Hindi": "hi",
    "Bengali": "bn",
    "Indonesian": "id",
    "Urdu": "ur",
}

TRANSLATIONS = {
    "en": {
        "upload_title": "Upload your file or document",
        "dashboard": "Dashboard",
        "data_doctor": "Data Doctor",
        "email": "Email",
        "email_placeholder": "Type your email",
        "password": "Password",
        "password_placeholder": "Type your password",
        "ask_data": "Ask Your Data",
        "report": "Consulting Report",
        "forecast": "Forecast",
        "saved": "Saved",
        "account": "Account",
        "ocr": "OCR",
        "decision_engine": "Decision Engine",
        "scenario_simulator": "Scenario Simulator",
        "dashboard_gallery": "Dashboard Gallery",
        "shareable_report": "Shareable Report",
        "boardroom_pdf": "Boardroom PDF",
        "industry_template": "Industry Template",
        "generate_dashboard": "Generate Executive Dashboard",
        "auto_fix_dataset": "Auto-Fix Dataset",
        "generate_decision_report": "Generate Decision Report",
        "create_share_link": "Create Shareable Report Link",
        "save_dashboard": "Save Dashboard Snapshot",
        "compare": "Compare",
        "cleaning": "Cleaning Lab",
        "legal": "Legal",
        "login": "Login",
        "create": "Create",
        "logout": "Logout",
    },

    "fr": {
        "upload_title": "Téléchargez votre fichier ou document",
        "dashboard": "Tableau de bord",
        "data_doctor": "Docteur des données",
        "email": "Email",
        "email_placeholder": "Entrez votre email",
        "password": "Mot de passe",
        "password_placeholder": "Entrez votre mot de passe",
        "ask_data": "Posez une question",
        "report": "Rapport",
        "forecast": "Prévision",
        "saved": "Enregistré",
        "account": "Compte",
        "ocr": "OCR",
        "compare": "Comparer",
        "cleaning": "Nettoyage",
        "legal": "Légal",
        "login": "Connexion",
        "create": "Créer",
        "logout": "Déconnexion",
    },

    "es": {
        "upload_title": "Sube tu archivo o documento",
        "dashboard": "Panel",
        "data_doctor": "Doctor de datos",
        "email": "Correo",
        "email_placeholder": "Ingrese su correo",
        "password": "Contraseña",
        "password_placeholder": "Ingrese su contraseña",
        "ask_data": "Pregunta a tus datos",
        "report": "Informe",
        "forecast": "Pronóstico",
        "saved": "Guardado",
        "account": "Cuenta",
        "ocr": "OCR",
        "compare": "Comparar",
        "cleaning": "Limpieza",
        "legal": "Legal",
        "login": "Iniciar sesión",
        "create": "Crear",
        "logout": "Cerrar sesión",
    },

    "ht": {
        "upload_title": "Telechaje dosye ou",
        "dashboard": "Tablo",
        "data_doctor": "Doktè done",
        "email": "Imèl",
        "email_placeholder": "Antre imèl ou",
        "password": "Modpas",
        "password_placeholder": "Antre modpas ou",
        "ask_data": "Poze kestyon",
        "report": "Rapò",
        "forecast": "Previzyon",
        "saved": "Sove",
        "account": "Kont",
        "ocr": "OCR",
        "compare": "Konpare",
        "cleaning": "Netwayaj",
        "legal": "Legal",
        "login": "Konekte",
        "create": "Kreye",
        "logout": "Dekonekte",
    },
    
    "zh": {},
    "hi": {},
    "bn": {},
    "id": {},
    "ur": {},
}


UI_TEXT = {
    "English": {
        "app_language": "Choose app language",
        "sign_in": "Sign in",
        "email": "Email",
        "email_placeholder": "Type your email",
        "password": "Password",
        "password_placeholder": "Type your password",
        "login": "Login",
        "create": "Create",
        "logout": "Log out",
        "use_sample_dataset": "Use Sample Dataset",
        "upload_title": "Upload your file or document",
        "upload_subtle": "CSV, Excel, TXT, DOCX, PDF, PNG, JPG, JPEG, WEBP, ZIP • Max size: 200MB",
        "rows": "Rows",
        "columns": "Columns",
        "missing": "Missing",
        "duplicates": "Duplicates",
        "generate_ai_insights": "Generate AI Insights",
        "run_ai_data_doctor": "Run AI Data Doctor",
        "download_executive_report": "Download Executive Report",
        "analysis_focus": "Analysis focus",
    },
    "Haitian Creole": {
        "app_language": "Chwazi lang aplikasyon an",
        "sign_in": "Konekte",
        "email": "Imèl",
        "email_placeholder": "Ekri imèl ou",
        "password": "Modpas",
        "password_placeholder": "Ekri modpas ou",
        "login": "Antre",
        "create": "Kreye",
        "logout": "Dekonekte",
        "use_sample_dataset": "Sèvi ak done egzanp",
        "upload_title": "Telechaje fichye oswa dokiman ou",
        "upload_subtle": "CSV, Excel, TXT, DOCX, PDF, PNG, JPG, JPEG, WEBP, ZIP • Gwosè maksimòm: 200MB",
        "rows": "Ranje",
        "columns": "Kolòn",
        "missing": "Ki manke",
        "duplicates": "Diplikasyon",
        "generate_ai_insights": "Jenere Enfòmasyon AI",
        "run_ai_data_doctor": "Lanse AI Data Doctor",
        "download_executive_report": "Telechaje Rapò Egzekitif",
        "analysis_focus": "Fokus analiz la",
    },
}

# =========================================================
# HELPERS
# =========================================================
def auto_translate(text, target_language):
    if not ai_available():
        return text  # fallback if no API

    prompt = f"Translate this UI text into {target_language}. Keep it short:\n{text}"
    try:
        return call_openai(prompt)
    except:
        return text

def t(key):
    lang_name = st.session_state.get("app_language", "English")
    lang_code = LANGUAGES.get(lang_name, "en")

    # 1. Try full translations
    text = TRANSLATIONS.get(lang_code, {}).get(key)

    # 2. Try UI_TEXT
    if not text:
        text = UI_TEXT.get(lang_name, {}).get(key)

    # 3. Fallback to English
    if not text:
        text = UI_TEXT.get("English", {}).get(key) or TRANSLATIONS.get("en", {}).get(key)

    # 4. FINAL fallback → AI translation (THIS is new)
    text = (
        TRANSLATIONS.get(lang_code, {}).get(key)
        or UI_TEXT.get(lang_name, {}).get(key)
        or UI_TEXT.get("English", {}).get(key)
        or TRANSLATIONS.get("en", {}).get(key)
    )

    # 🚀 Add AI translation as final step
    if text and lang_code != "en":
        return auto_translate(text, lang_name)

    return text or key
    
    
def now_iso() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def hash_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def get_base64_image(path: Path) -> str:
    if not path.exists():
        return ""
    try:
        return base64.b64encode(path.read_bytes()).decode("utf-8")
    except Exception:
        return ""


def current_user_email() -> str:
    user = st.session_state.get("auth_user") or {}
    return str(user.get("email", "")).strip().lower()


def current_language() -> str:
    return st.session_state.get("app_language", "English")


def user_key_from_email(email: str) -> str:
    return hash_text(email.strip().lower())[:20]


def user_upload_dir(email: str) -> Path:
    path = USER_UPLOADS_DIR / user_key_from_email(email)
    path.mkdir(parents=True, exist_ok=True)
    return path


def user_report_dir(email: str) -> Path:
    path = USER_REPORTS_DIR / user_key_from_email(email)
    path.mkdir(parents=True, exist_ok=True)
    return path


def ai_available() -> bool:
    return client is not None


def auth_configured() -> bool:
    return bool(SUPABASE_URL and SUPABASE_ANON_KEY)


def detect_user_language() -> str:
    try:
        loc = locale.getdefaultlocale()[0]
        if loc is None:
            return "English"
        loc = loc.lower()
        if loc.startswith("ht"):
            return "Haitian Creole"
        if loc.startswith("fr"):
            return "French"
        if loc.startswith("es"):
            return "Spanish"
        if loc.startswith("de"):
            return "German"
        if loc.startswith("pt"):
            return "Portuguese"
        if loc.startswith("ar"):
            return "Arabic"
        return "English"
    except Exception:
        return "English"


def load_logo_base64() -> str:
    return get_base64_image(LOGO_PATH)


logo_base64 = load_logo_base64()


# =========================================================
# PREMIUM ACCESS CONTROL
# =========================================================
FREE_PLAN = "free"
PREMIUM_PLAN = "premium"

PREMIUM_FEATURES = {
    "ai_insights": True,
    "forecast": True,
    "saved_reports": True,
    "data_doctor_autofix": True,
    "shareable_reports": True,
}


def get_user_plan() -> str:
    email = current_user_email()

    if st.session_state.get("developer_premium_override"):
        return PREMIUM_PLAN
    if email and email in PREMIUM_USERS:
        return PREMIUM_PLAN
    if st.session_state.get("premium_status") == "premium":
        return PREMIUM_PLAN

    return FREE_PLAN


def set_user_plan(plan_name: str) -> None:
    st.session_state["user_plan"] = plan_name
    st.session_state["premium_status"] = (
        "premium" if plan_name == PREMIUM_PLAN else "inactive"
    )


def is_premium_user() -> bool:
    return get_user_plan() == PREMIUM_PLAN


def feature_enabled(feature_name: str) -> bool:
    if feature_name not in PREMIUM_FEATURES:
        return True
    if PREMIUM_FEATURES[feature_name]:
        return is_premium_user()
    return True


def require_feature(feature_name: str, title: str = "Premium Feature") -> bool:
    if feature_enabled(feature_name):
        return True
    render_message(
        f"{title} is available on Premium. Upgrade to unlock full access.",
        "warning",
    )
    return False


def analysis_allowed() -> Tuple[bool, str]:
    if is_premium_user():
        return True, ""
    if st.session_state.get("usage_count", 0) < FREE_ANALYSIS_LIMIT:
        return True, ""
    return False, f"Free analysis limit reached ({FREE_ANALYSIS_LIMIT}). Upgrade to continue."


def increment_usage() -> None:
    st.session_state["usage_count"] = st.session_state.get("usage_count", 0) + 1


def process_stripe_return() -> None:
    params = st.query_params
    payment = str(params.get("payment", "")).lower()
    status = str(params.get("status", "")).lower()
    plan = str(params.get("plan", "")).lower()
    session_id = str(params.get("session_id", "")).strip()

    success = payment == "success" or status == "success"
    if not success:
        return

    payment_key = f"{plan}:{session_id or 'no_session'}"
    if st.session_state.get("stripe_last_payment") == payment_key:
        return

    if plan in ["premium", "monthly", "pro"]:
        set_user_plan(PREMIUM_PLAN)
        st.session_state["stripe_last_payment"] = payment_key
        render_message(
            "Payment confirmed. Premium access has been activated automatically.",
            "success",
        )
    elif plan in ["one_time", "onetime", "single"]:
        st.session_state["stripe_last_payment"] = payment_key
        render_message(
            "Payment confirmed. One-time access has been activated.",
            "success",
        )

                
def build_stripe_success_url(plan_name: str) -> str:
    if STRIPE_SUCCESS_URL:
        base = STRIPE_SUCCESS_URL
    elif APP_BASE_URL:
        base = APP_BASE_URL
    else:
        base = ""
    if not base:
        return ""
    separator = "&" if "?" in base else "?"
    return f"{base}{separator}payment=success&plan={plan_name}&session_id={{CHECKOUT_SESSION_ID}}"


def safe_remove_tree(path: Path) -> None:
    if path.exists() and path.is_dir():
        shutil.rmtree(path, ignore_errors=True)


def file_size_ok(uploaded_file) -> Tuple[bool, str]:
    if uploaded_file is None:
        return False, "No file uploaded."
    size_mb = uploaded_file.size / (1024 * 1024)
    if size_mb > MAX_UPLOAD_MB:
        return False, f"File is {size_mb:.1f}MB. Max allowed size is {MAX_UPLOAD_MB}MB."
    return True, ""


# =========================================================
# SEMANTIC METRICS
# =========================================================
METRIC_DEFINITIONS = {
    "record_count": {
        "label": "Records",
        "description": "Total rows in dataset",
        "format": "integer",
        "required_columns": [],
    },
    "column_count": {
        "label": "Columns",
        "description": "Total fields available",
        "format": "integer",
        "required_columns": [],
    },
    "missing_value_rate": {
        "label": "Missing Rate",
        "description": "Share of missing data",
        "format": "percent",
        "required_columns": [],
    },
    "duplicate_rate": {
        "label": "Duplicate Rate",
        "description": "Share of duplicate rows",
        "format": "percent",
        "required_columns": [],
    },
    "revenue": {
        "label": "Revenue",
        "description": "Total sales value",
        "format": "currency",
        "required_columns": ["Sales"],
    },
    "customers": {
        "label": "Customers",
        "description": "Total customer count",
        "format": "integer",
        "required_columns": ["Customers"],
    },
    "returns": {
        "label": "Returns",
        "description": "Total returns count",
        "format": "integer",
        "required_columns": ["Returns"],
    },
}


def format_metric_value(value, fmt: str) -> str:
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return "N/A"
        if fmt == "currency":
            return f"${float(value):,.2f}"
        if fmt == "percent":
            return f"{float(value) * 100:,.2f}%"
        if fmt == "integer":
            return f"{int(round(float(value))):,}"
        return f"{float(value):,.2f}"
    except Exception:
        return str(value)


def resolve_available_metrics(df: pd.DataFrame) -> dict:
    available = {}
    for metric_key, config in METRIC_DEFINITIONS.items():
        required = config.get("required_columns", [])
        missing = [col for col in required if col not in df.columns]
        available[metric_key] = {
            "available": len(missing) == 0,
            "missing_columns": missing,
            "config": config,
        }
    return available


def compute_semantic_metrics(df: pd.DataFrame) -> list:
    metrics = []
    total_rows = len(df)
    total_cols = len(df.columns)
    total_cells = max(total_rows * max(total_cols, 1), 1)
    available_map = resolve_available_metrics(df)

    def add_metric(metric_key: str, raw_value, source_columns=None):
        config = METRIC_DEFINITIONS[metric_key]
        metrics.append(
            {
                "key": metric_key,
                "label": config["label"],
                "description": config["description"],
                "raw_value": raw_value,
                "formatted_value": format_metric_value(raw_value, config["format"]),
                "source_columns": source_columns or config.get("required_columns", []),
            }
        )

    add_metric("record_count", total_rows, [])
    add_metric("column_count", total_cols, [])
    add_metric(
        "missing_value_rate",
        float(df.isna().sum().sum()) / total_cells,
        list(df.columns),
    )
    add_metric(
        "duplicate_rate",
        float(df.duplicated().sum()) / max(total_rows, 1),
        list(df.columns),
    )

    if available_map["revenue"]["available"]:
        add_metric(
            "revenue",
            pd.to_numeric(df["Sales"], errors="coerce").fillna(0).sum(),
            ["Sales"],
        )
    if available_map["customers"]["available"]:
        add_metric(
            "customers",
            pd.to_numeric(df["Customers"], errors="coerce").fillna(0).sum(),
            ["Customers"],
        )
    if available_map["returns"]["available"]:
        add_metric(
            "returns",
            pd.to_numeric(df["Returns"], errors="coerce").fillna(0).sum(),
            ["Returns"],
        )

    return metrics


def render_semantic_metric_cards(metrics: list, max_items: int = 6) -> None:
    if not metrics:
        return
    selected = metrics[:max_items]
    card_html = '<div class="semantic-metric-grid">'
    for metric in selected:
        card_html += f"""
        <div class="semantic-metric-card">
            <div class="semantic-metric-name">{html.escape(metric['label'])}</div>
            <div class="semantic-metric-value">{html.escape(metric['formatted_value'])}</div>
            <div class="semantic-metric-desc">{html.escape(metric['description'])}</div>
        </div>
        """
    card_html += "</div>"
    st.markdown(card_html, unsafe_allow_html=True)


# =========================================================
# STYLING / CSS
# =========================================================
def inject_global_css() -> None:
    st.markdown(
        """
        <style>
        html, body, [class*="css"] {
            font-family: Arial, sans-serif;
        }

        .stApp {
            background: linear-gradient(135deg, #081018 0%, #0c1622 45%, #111827 100%);
        }

        .block-container {
            padding-top: 0.5rem !important;
            padding-bottom: 0.4rem !important;
        }

        .hero-wrap {
            border: 1px solid rgba(34,197,94,0.25);
            backdrop-filter: blur(6px);
            border-radius: 18px;
            padding: 12px;
            margin-bottom: 10px;
            box-shadow: 0 10px 28px rgba(34,197,94,0.10);
        }

        @keyframes greenFlow {
            0%   { background-position: 0% 50%; }
            100% { background-position: 100% 50%; }
        }

        @keyframes greenGlowPulse {
            0%   { box-shadow: 0 6px 18px rgba(34,197,94,0.35), 0 0 14px rgba(34,197,94,0.25); }
            50%  { box-shadow: 0 8px 24px rgba(34,197,94,0.60), 0 0 28px rgba(34,197,94,0.45); }
            100% { box-shadow: 0 6px 18px rgba(34,197,94,0.35), 0 0 14px rgba(34,197,94,0.25); }
        }
        

        .top-green-bar {
            width: 100%;
            height: 40px;
            background: linear-gradient(
                90deg,
                #001f12,
                #064e3b,
                #00ff7f,
                #22c55e,
                #ccff00,
                #00ff7f,
                #064e3b,
                #001f12                
            );
            background-size: 450% 100%;
            animation: greenFlow 3.5s linear infinite, greenGlowPulse 2.8s ease-in-out infinite;
            border-radius: 0 0 22px 22px;
            margin-top: -4px;
            box-shadow:
                0 8px 24px rgba(34, 197, 94, 0.75),
                0 0 38px rgba(0, 255, 127, 0.55),
                inset 0 -3px 10px rgba(255, 255, 255, 0.28);
            margin-bottom: 12px;
        }
        
        /* Restore original sidebar background */
        section[data-testid="stSidebar"] {
            background-color: #f8fafc !important;  /* light gray */
        }

        .hero-main-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 18px;
            flex-wrap: wrap;
        }
        
        div[data-testid="stTabs"] {
            margin-top: 6px !important;
        }

        .hero-left {
            flex: 1 1 560px;
        }

        .hero-right {
            flex: 0 0 240px;
            display: flex;
            justify-content: center;
            align-items: center;
        }

        .blue-top-bar {
            width: 100%;
            height: 34px;

        background: linear-gradient(90deg, #001f12, #064e3b, #22c55e, #ccff00, #22c55e, #064e3b);
        
        background-size: 300% 100%;
        animation: greenFlow 6s linear infinite;

        border-radius: 0 0 18px 18px;

        margin-top: -7px;
        margin-bottom: 6px;

        box-shadow:
        0 6px 18px rgba(34, 197, 94, 0.45),
        0 0 20px rgba(34, 197, 94, 0.25);
    }

        
        .hero-brand-row {
            display: flex;
            align-items: center;
            gap: 14px;
        }

        .hero-logo img {
            width: 74px;
            height: 74px;
            object-fit: contain;
            border-radius: 14px;
            box-shadow: 0 6px 16px rgba(0,0,0,0.16);
            background: rgba(255,255,255,0.55);
            padding: 6px;
        }

        .hero-brand-line {
            font-size: 2rem;
            font-weight: 900;
            color: #15803d;
            line-height: 1.08;
        }
        
        /* Grey control button */
        button[kind="secondary"],
        button#control_btn {
            background: #e5e7eb !important;
            color: #111827 !important;
            font-weight: 800 !important;
            border-radius: 10px !important;
            border: 1px solid #d1d5db !important;
        }
        .hero-tagline-line {
            font-size: 1.05rem;
            font-weight: 700;
            color: #0c4a6e;
            margin-top: 4px;
        }

        .hero-welcome {
            font-weight: 900;
            color: #14532d;
            text-align: center;
            font-size: 1.18rem;
            background: rgba(255,255,255,0.55);
            border: 1px solid rgba(20,83,45,0.10);
            border-radius: 16px;
            padding: 12px 14px;
            box-shadow: 0 8px 18px rgba(0,0,0,0.12);
            animation: zoomPulse 1.6s ease-in-out infinite;
            transform-origin: center center;
        }

        @keyframes zoomPulse {
            0%   { transform: scale(1); }
            50%  { transform: scale(1.07); }
            100% { transform: scale(1); }
        }

        .feature-pill-wrap {
            display: flex;
            gap: 8px;
            margin: 10px 0 12px 0;
            flex-wrap: wrap;
        }

        .feature-pill {
            background: #dcfce7;
            padding: 6px 12px;
            border-radius: 999px;
            font-weight: 700;
            color: #14532d;
        }

        .trust-bar {
            background: #fef3c7;
            padding: 10px 12px;
            border-radius: 12px;
            font-weight: 700;
            margin: 10px 0 14px 0;
            color: #111827;
        }

        .status-chip-wrap {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin: 10px 0 16px 0;
        }

        .status-chip {
            padding: 7px 12px;
            border-radius: 999px;
            font-weight: 800;
            background: rgba(255,255,255,0.08);
            color: #e2e8f0;
            border: 1px solid rgba(255,255,255,0.08);
        }
        
        /* More space before the tabs row */
        div[data-testid="stTabs"] {
            margin-top: 18px !important;
        }

        .semantic-metric-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 12px;
            margin: 10px 0 16px 0;
        }

        .semantic-metric-card {
            background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(248,250,252,0.96));
            border: 1px solid rgba(15,23,42,0.08);
            border-radius: 18px;
            padding: 14px;
            box-shadow: 0 8px 20px rgba(0,0,0,0.08);
        }

        .semantic-metric-name {
            font-size: 0.85rem;
            color: #475569;
            font-weight: 700;
        }

        .semantic-metric-value {
            font-size: 1.25rem;
            color: #0f172a;
            font-weight: 900;
            margin: 6px 0;
        }

        .semantic-metric-desc {
            font-size: 0.82rem;
            color: #64748b;
        }

        div[data-testid="stFileUploader"] section {
            border: 2px dashed #22c55e !important;
            background: rgba(255,255,255,0.03) !important;
            border-radius: 16px !important;
            margin-bottom: 6px !important;   /* ↓ reduce gap */
        }
                
        /* Drag and drop text → white */
        div[data-testid="stFileUploader"] small,
        div[data-testid="stFileUploader"] span,
        div[data-testid="stFileUploader"] label,
        div[data-testid="stFileUploader"] p {
            color: #ffffff !important;
        }

        div[data-testid="stFileUploader"] button {
            background: #facc15 !important;
            color: #111827 !important;
            border-radius: 10px !important;
            border: none !important;
            font-weight: 800 !important;
        }

        section[data-testid="stSidebar"] div[data-testid="stButton"] button,
        section[data-testid="stSidebar"] div[data-testid="stLinkButton"] a {
            background: #f5f5dc !important;
            color: #111827 !important;
            font-weight: 800 !important;
            border-radius: 10px !important;
            border: none !important;
        }

        div[data-testid="stTextInput"] input::placeholder,
        div[data-testid="stTextInputRootElement"] input::placeholder,
        input::placeholder {
            color: #94a3b8 !important;
            opacity: 1 !important;
        }

        footer { visibility: hidden; }
        
        /* Bigger top tabs: Dashboard → Legal */
        button[data-baseweb="tab"] {
            font-size: 1.08rem !important;
            font-weight: 800 !important;
            padding: 14px 22px !important;
            min-height: 56px !important;
            border-radius: 14px !important;
            cursor: pointer !important;
        }

        button[data-baseweb="tab"] p {
            font-size: 1.08rem !important;
            font-weight: 800 !important;
        }
        
        /* Gradient green divider */
        .green-gradient-divider {
            border: none;
            height: 3px;
            background: linear-gradient(to right, transparent, #22c55e, #16a34a, #22c55e, transparent);
            margin: 14px 0 18px 0;
            border-radius: 999px;
        }
        
        /* Hand cursor on clickable items */
        button,
        button * ,
        a,
        a *,
        div[data-testid="stFileUploader"] button,
        button[data-baseweb="tab"],
        .stSelectbox,
        .stSelectbox * {
        cursor: pointer !important;
    }
/* SAFE readability fix: text only, do not touch buttons */
div[data-testid="stMarkdownContainer"] h1,
div[data-testid="stMarkdownContainer"] h2,
div[data-testid="stMarkdownContainer"] h3,
div[data-testid="stMarkdownContainer"] h4,
div[data-testid="stMarkdownContainer"] p,
div[data-testid="stMarkdownContainer"] li {
    color: #f8fafc !important;
}

/* Dashboard metric labels/values only */
div[data-testid="stMetric"] label,
div[data-testid="stMetric"] div {
    color: #f8fafc !important;
}

/* Dataframe area spacing only */
div[data-testid="stDataFrame"] {
    border-radius: 12px !important;
}

/* Keep sidebar button/link text dark and visible */
section[data-testid="stSidebar"] button,
section[data-testid="stSidebar"] button *,
section[data-testid="stSidebar"] a,
section[data-testid="stSidebar"] a * {
    color: #111827 !important;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label {
    color: #111827 !important;
}

/* Keep main buttons readable */
div[data-testid="stButton"] button,
div[data-testid="stButton"] button *,
div[data-testid="stDownloadButton"] button,
div[data-testid="stDownloadButton"] button * {
    color: #111827 !important;
    font-weight: 800 !important;
}

/* Keep footer visible */
.footer-note {
    text-align: center;
    font-size: 0.85rem;
    color: #c3cfde !important;
    margin-top: 20px;
    margin-bottom: 18px;
}


/* Keep sidebar inputs readable */
section[data-testid="stSidebar"] input {
    color: #111827 !important;
    background: #ffffff !important;
}

/* Keep sidebar buttons readable */
section[data-testid="stSidebar"] button,
section[data-testid="stSidebar"] button *,
section[data-testid="stSidebar"] a,
section[data-testid="stSidebar"] a * {
    color: #111827 !important;
    opacity: 1 !important;
    visibility: visible !important;
}

/* Password eye icon fix */
section[data-testid="stSidebar"] div[data-baseweb="input"] {
    background-color: #ffffff !important;
}

section[data-testid="stSidebar"] div[data-baseweb="base-input"] {
    background-color: #ffffff !important;
}

section[data-testid="stSidebar"] input {
    background-color: #ffffff !important;
    color: #111827 !important;
}

section[data-testid="stSidebar"] svg {
    color: #111827 !important;
    fill: #111827 !important;
}

/* Remove dark patch behind password eye button */
section[data-testid="stSidebar"] div[data-baseweb="input"] > div {
    background-color: #ffffff !important;
}

section[data-testid="stSidebar"] button[aria-label*="password"],
section[data-testid="stSidebar"] button[title*="password"],
section[data-testid="stSidebar"] button {
    background-color: #ffffff !important;
    color: #111827 !important;
}

/* Fix checkbox text under "Use Sample Dataset" */
section[data-testid="stSidebar"] div[data-baseweb="checkbox"] label {
    color: #ffffff !important;
}

/* Fix the check icon */
section[data-testid="stSidebar"] div[data-baseweb="checkbox"] svg {
    fill: #ffffff !important;
    color: #ffffff !important;
}

/* Fix blurry text in upgrade section */
div:has(> a[href*="stripe"]) {
    transform: none !important;
    filter: none !important;
    opacity: 1 !important;
}

/* Password eye icon dark patch fix */
section[data-testid="stSidebar"] div[data-baseweb="input"],
section[data-testid="stSidebar"] div[data-baseweb="base-input"],
section[data-testid="stSidebar"] div[data-baseweb="input"] > div,
section[data-testid="stSidebar"] div[data-baseweb="base-input"] > div {
    background-color: #ffffff !important;
}

section[data-testid="stSidebar"] input {
    background-color: #ffffff !important;
    color: #111827 !important;
}

section[data-testid="stSidebar"] button[aria-label*="password"],
section[data-testid="stSidebar"] button[title*="password"] {
    background-color: #ffffff !important;
    color: #111827 !important;
    border: none !important;
}

.how-card *,
section[data-testid="stSidebar"] .how-card * {
    color: #111827 !important;
    opacity: 1 !important;
    visibility: visible !important;
}

.how-title,
section[data-testid="stSidebar"] .how-title {
    color: #14532d !important;
    font-size: 1.02rem !important;
    font-weight: 900 !important;
    margin-bottom: 8px !important;
}

.how-text,
section[data-testid="stSidebar"] .how-text {
    color: #111827 !important;
    line-height: 1.7 !important;
    font-weight: 700 !important;
}
        </style>
        """,
        unsafe_allow_html=True,
    )

# =========================================================
# RENDERERS
# =========================================================
def render_top_bar() -> None:
    st.markdown('<div class="top-green-bar"></div>', unsafe_allow_html=True)

def render_hero() -> None:
    logo_html = (
        f'<img src="data:image/png;base64,{logo_base64}" alt="ExplainMyData AI Logo">'
        if logo_base64
        else ""
    )
    st.markdown(
        f"""
        <div class="hero-wrap">
            <div class="hero-main-row">
                <div class="hero-left">
                    <div class="hero-brand-row">
                        <div class="hero-logo">{logo_html}</div>
                        <div class="hero-brand-stack">
                            <div class="hero-brand-line">{html.escape(APP_NAME)}</div>
                            <div class="hero-tagline-line">{html.escape(TAGLINE)}</div>
                        </div>
                    </div>
                </div>
                <div class="hero-right">
                    <div class="hero-welcome">🚀 Turn messy data into clean business insights</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_feature_pills() -> None:
    st.markdown(
        """
        <div class="feature-pill-wrap">
            <div class="feature-pill">⚡ Fast Insights</div>
            <div class="feature-pill">📊 AI Powered</div>
            <div class="feature-pill">🌍 Global Access</div>
            <div class="feature-pill">📄 PDF Export</div>
            <div class="feature-pill">🧠 Business Ready</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_trust_bar() -> None:
    st.markdown(
        """
        <div class="trust-bar">
            🔐 Secure file handling • Consent-aware AI • Privacy-first design
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_status_strip() -> None:
    st.markdown(
        """
        <div class="status-chip-wrap">
            <div class="status-chip">🔐 Secure file handling</div>
            <div class="status-chip">⚡ Fast processing</div>
            <div class="status-chip">📊 AI-powered insights</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_report_card(title: str, body: str, accent: str = "#22c55e") -> None:
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, rgba(255,255,255,0.08), rgba(255,255,255,0.03));
            border-left: 5px solid {accent};
            border-radius: 18px;
            padding: 16px;
            margin: 10px 0;
            color: #f8fafc;
            box-shadow: 0 8px 22px rgba(0,0,0,0.16);
        ">
            <div style="font-size:1.05rem;font-weight:900;margin-bottom:8px;">{html.escape(title)}</div>
            <div style="font-size:0.98rem;line-height:1.65;white-space:pre-wrap;">{html.escape(body)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_message(message: str, kind: str = "info") -> None:
    styles = {
        "info": {
            "bg": "linear-gradient(135deg, rgba(37,99,235,0.16), rgba(30,64,175,0.10))",
            "border": "#60a5fa",
            "text": "#eff6ff",
            "icon": "ℹ️",
        },
        "success": {
            "bg": "linear-gradient(135deg, rgba(22,163,74,0.16), rgba(21,128,61,0.10))",
            "border": "#4ade80",
            "text": "#f0fdf4",
            "icon": "✅",
        },
        "warning": {
            "bg": "linear-gradient(135deg, rgba(245,158,11,0.16), rgba(180,83,9,0.10))",
            "border": "#fbbf24",
            "text": "#fffbeb",
            "icon": "⚠️",
        },
        "error": {
            "bg": "linear-gradient(135deg, rgba(239,68,68,0.16), rgba(153,27,27,0.10))",
            "border": "#f87171",
            "text": "#fef2f2",
            "icon": "❌",
        },
    }
    cfg = styles.get(kind, styles["info"])
    st.markdown(
        f"""
        <div style="
            background: {cfg['bg']};
            border: 1px solid {cfg['border']};
            border-left: 5px solid {cfg['border']};
            border-radius: 16px;
            padding: 12px 14px;
            margin: 10px 0;
            color: {cfg['text']};
            font-weight: 700;
            box-shadow: 0 8px 20px rgba(0,0,0,0.12);
        ">
            <span style="margin-right:8px;">{cfg['icon']}</span>{html.escape(str(message))}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_privacy_policy() -> None:
    st.markdown("### Privacy Policy")
    st.markdown(
        """
- Uploaded files are processed to create analysis, charts, OCR results, reports, translations, and user-requested features.
- If you sign in, files and reports may be stored in your user workspace.
- If AI is enabled, selected content may be sent to your configured AI provider.
- Secure deployment still depends on HTTPS, secret management, and proper hosting.
        """
    )


def render_terms_of_service() -> None:
    st.markdown("### Terms of Service")
    st.markdown(
        """
- AI outputs may be incomplete or inaccurate and should be reviewed before business use.
- Users must not upload unlawful content or data they do not have permission to process.
- Premium functionality depends on payment and deployment configuration.
        """
    )


def render_data_handling_notice() -> None:
    st.markdown("### How your data is handled")
    st.markdown(
        """
1. Uploaded files are parsed to create summaries, charts, OCR results, and AI insights.
2. If signed in, files and reports can be saved under your user workspace.
3. If AI analysis is enabled, selected content may be sent to your configured AI provider.
4. You can delete saved local data from the app controls.
        """
    )


# =========================================================
# AUTH
# =========================================================
def local_login(email: str, password: str) -> Tuple[bool, str]:
    if not email or not password:
        return False, "Please enter your email and password."
    st.session_state["authenticated"] = True
    st.session_state["auth_user"] = {"email": email.strip().lower()}
    return True, "Logged in successfully."


def local_create_account(email: str, password: str) -> Tuple[bool, str]:
    if not email or not password:
        return False, "Please enter your email and password."
    st.session_state["authenticated"] = True
    st.session_state["auth_user"] = {"email": email.strip().lower()}
    return True, "Account created successfully."


def logout() -> None:
    st.session_state["authenticated"] = False
    st.session_state["auth_user"] = None
    set_user_plan(FREE_PLAN)


# =========================================================
# FILE LOADERS
# =========================================================
def load_structured_data(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(uploaded_file)
    if suffix in [".xlsx", ".xls"]:
        return pd.read_excel(uploaded_file)
    raise ValueError("Unsupported structured file type.")


def extract_text_from_txt(uploaded_file) -> str:
    return uploaded_file.getvalue().decode("utf-8", errors="ignore")


def extract_text_from_docx(uploaded_file) -> str:
    if Document is None:
        raise RuntimeError("python-docx is not installed.")
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])


def extract_text_from_pdf(uploaded_file) -> str:
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed.")
    text_parts = []
    pdf = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
    for page in pdf:
        text_parts.append(page.get_text())
    return "\n".join(text_parts)


def extract_text_from_image(image: Image.Image) -> str:
    if pytesseract is None:
        return "OCR unavailable because pytesseract is not installed."
    return pytesseract.image_to_string(image)


def preprocess_image_for_ocr(
    image: Image.Image,
    grayscale=True,
    autocontrast=True,
    sharpen=True,
    threshold=False,
    threshold_value=160,
) -> Image.Image:
    processed = image.copy()
    if grayscale:
        processed = ImageOps.grayscale(processed)
    if autocontrast:
        processed = ImageOps.autocontrast(processed)
    if sharpen:
        processed = processed.filter(ImageFilter.SHARPEN)
    if threshold:
        processed = processed.point(lambda p: 255 if p > threshold_value else 0)
    return processed


def extract_text_from_image_with_options(
    image: Image.Image,
    grayscale=True,
    autocontrast=True,
    sharpen=True,
    threshold=False,
    threshold_value=160,
) -> Tuple[str, Image.Image]:
    processed = preprocess_image_for_ocr(
        image,
        grayscale,
        autocontrast,
        sharpen,
        threshold,
        threshold_value,
    )
    text = extract_text_from_image(processed)
    return text, processed


def load_zip_preview(uploaded_file) -> Tuple[pd.DataFrame, str]:
    with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue())) as zf:
        names = zf.namelist()
    df = pd.DataFrame({"zip_members": names})
    raw_text = "ZIP contains:\n" + "\n".join(names[:300])
    return df, raw_text


def build_data_summary(df: pd.DataFrame, mode: str, raw_text: str) -> str:
    if mode == "structured":
        preview = df.head(10).to_string(index=False)
        return f"""
Structured dataset shape: {df.shape[0]} rows x {df.shape[1]} columns
Columns: {', '.join(df.columns.astype(str).tolist()[:50])}
Preview:
{preview}
"""
    if mode == "text":
        return f"""
Text dataset shape: {df.shape[0]} rows x {df.shape[1]} columns
Text preview:
{raw_text[:5000]}
"""
    if mode == "image":
        return f"""
Image metadata table shape: {df.shape[0]} rows x {df.shape[1]} columns
OCR preview:
{raw_text[:3000] if raw_text else 'No OCR text available.'}
"""
    return "No summary available."


def parse_uploaded_file(uploaded_file) -> Tuple[pd.DataFrame, str, str]:
    suffix = Path(uploaded_file.name).suffix.lower().replace(".", "")

    if suffix in ["csv", "xlsx", "xls"]:
        df = load_structured_data(uploaded_file)
        return df, "structured", ""

    if suffix == "txt":
        text = extract_text_from_txt(uploaded_file)
        df = pd.DataFrame({"text": text.splitlines() or [text]})
        return df, "text", text

    if suffix == "docx":
        text = extract_text_from_docx(uploaded_file)
        df = pd.DataFrame({"text": text.splitlines() or [text]})
        return df, "text", text

    if suffix == "pdf":
        text = extract_text_from_pdf(uploaded_file)
        df = pd.DataFrame({"text": text.splitlines() or [text]})
        return df, "text", text

    if suffix in ["png", "jpg", "jpeg", "webp"]:
        image = Image.open(io.BytesIO(uploaded_file.getvalue()))
        text = extract_text_from_image(image)
        df = pd.DataFrame(
            {"image_width": [image.size[0]], "image_height": [image.size[1]]}
        )
        st.session_state["ocr_preview"] = image
        st.session_state["ocr_text"] = text
        return df, "image", text

    if suffix == "zip":
        zip_df, zip_text = load_zip_preview(uploaded_file)
        return zip_df, "text", zip_text

    raise ValueError("Unsupported file type.")


def save_uploaded_file_for_user(email: str, uploaded_file) -> Optional[Path]:
    if not email or uploaded_file is None:
        return None
    target = user_upload_dir(email) / uploaded_file.name
    try:
        target.write_bytes(uploaded_file.getvalue())
        return target
    except Exception:
        return None


# =========================================================
# AI
# =========================================================
def call_openai(prompt: str) -> str:
    if not ai_available():
        raise RuntimeError(
            "OpenAI is not configured. Add OPENAI_API_KEY to .env or Streamlit secrets."
        )
    response = client.responses.create(model=DEFAULT_MODEL, input=prompt)
    return response.output_text.strip()


def generate_ai_analysis(
    df: pd.DataFrame,
    mode: str,
    raw_text: str,
    focus: str,
    language: str,
) -> str:
    industry = st.session_state.get("industry_template", "General Business")
    prompt = f"""
Respond in {language}.
You are a senior business analyst.
The selected industry template is: {industry}.
Tailor insights specifically to this industry.

Use the uploaded content summary below.
Provide these sections clearly:
1. Executive Summary
2. Key Insights
3. What These Insights Mean
4. Risks or Anomalies
5. Recommended Actions
6. Expected Business Impact
7. Questions Worth Exploring Next

Rules:
- Be practical and professional
- Do not invent unsupported facts
- If the input is text or image-based, explain what is actually present
- Keep it concise but useful

Content summary:
{build_data_summary(df, mode, raw_text)}
"""
    return call_openai(prompt)


def generate_translation(text: str, target_language: str) -> str:
    prompt = f"""
Translate the text below into {target_language}.
Keep the meaning accurate, business-friendly, and easy to understand.

Text:
{text}
"""
    return call_openai(prompt)


def answer_followup_question(
    df: pd.DataFrame,
    mode: str,
    raw_text: str,
    question: str,
    language: str,
) -> str:
    prompt = f"""
Respond in {language}.
You are a helpful data analyst.
Use only the information supported by the summary below.

Summary:
{build_data_summary(df, mode, raw_text)}

User question:
{question}
"""
    return call_openai(prompt)


def build_ai_data_doctor_report(
    df: pd.DataFrame,
    mode: str,
    raw_text: str,
    language: str,
) -> str:
    issues: List[str] = []
    recommendations: List[str] = []

    missing_total = int(df.isna().sum().sum())
    dup_total = int(df.duplicated().sum())
    if missing_total > 0:
        issues.append(f"Missing values detected: {missing_total}")
        recommendations.append(
            "Fill missing numeric values with median or business-approved defaults."
        )
        recommendations.append(
            "Fill missing text fields with explicit placeholders like 'Unknown' only when appropriate."
        )
    if dup_total > 0:
        issues.append(f"Duplicate rows detected: {dup_total}")
        recommendations.append(
            "Review duplicate rows before deletion to separate true duplicates from repeated business events."
        )

    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    outlier_notes = []
    for col in numeric_cols[:8]:
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(series) < 8:
            continue
        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)
        iqr = q3 - q1
        if iqr <= 0:
            continue
        mask = (series < q1 - 1.5 * iqr) | (series > q3 + 1.5 * iqr)
        count = int(mask.sum())
        if count:
            outlier_notes.append(f"{col}: {count} possible outliers")

    if outlier_notes:
        issues.append("Possible outliers detected: " + "; ".join(outlier_notes))
        recommendations.append(
            "Inspect outliers to separate real signals from data entry errors."
        )

    constant_cols = [c for c in df.columns if df[c].nunique(dropna=False) <= 1]
    if constant_cols:
        issues.append("Constant columns detected: " + ", ".join(constant_cols[:15]))
        recommendations.append(
            "Review constant columns. They may be placeholders, IDs gone wrong, or fields with no analytical value."
        )

    if not issues:
        issues.append("No major structural issues detected.")
        recommendations.append(
            "Proceed with analysis while validating business context."
        )

    local_report = (
        "AI Data Doctor Findings\n"
        "-----------------------\n"
        + "\n".join(f"- {item}" for item in issues)
        + "\n\nRecommended Fixes\n"
        "-----------------\n"
        + "\n".join(f"- {item}" for item in recommendations)
    )

    if not ai_available():
        return local_report

    prompt = f"""
Respond in {language}.
You are a professional data quality consultant.
Turn the findings below into a polished diagnostic report.

Findings:
{local_report}

Dataset summary:
{build_data_summary(df, mode, raw_text)}
"""
    return call_openai(prompt)


def generate_cleaning_summary(
    original_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
    language: str = "English",
) -> str:
    if original_df is None or cleaned_df is None:
        return ""
    before_missing = int(original_df.isnull().sum().sum())
    after_missing = int(cleaned_df.isnull().sum().sum())
    duplicates_before = int(original_df.duplicated().sum())
    duplicates_after = int(cleaned_df.duplicated().sum())

    if not ai_available():
        return (
            f"Missing values: {before_missing} → {after_missing}. "
            f"Duplicate rows: {duplicates_before} → {duplicates_after}. "
            f"The cleaning steps improved dataset quality for analysis."
        )

    prompt = f"""
Respond in {language}.

You are a data cleaning expert.

Write a short professional summary (3 lines max) explaining:
- what cleaning actions were applied
- why they matter
- how they improve data quality

Before cleaning:
Missing values: {before_missing}
Duplicate rows: {duplicates_before}

After cleaning:
Missing values: {after_missing}
Duplicate rows: {duplicates_after}
"""
    return call_openai(prompt)


# =========================================================
# FORECASTING
# =========================================================
def prepare_forecast_dataframe(
    df: pd.DataFrame,
    date_col: str,
    value_col: str,
) -> pd.DataFrame:
    working = df[[date_col, value_col]].copy()
    working[date_col] = pd.to_datetime(working[date_col], errors="coerce")
    working[value_col] = pd.to_numeric(working[value_col], errors="coerce")
    working = working.dropna()
    if working.empty:
        return working
    working = (
        working.groupby(date_col, as_index=False)[value_col]
        .sum()
        .sort_values(date_col)
    )
    return working


def build_forecast_ml(
    grouped_df: pd.DataFrame,
    periods: int = 6,
) -> Tuple[pd.DataFrame, str, str]:
    if grouped_df.empty or len(grouped_df) < 3:
        return pd.DataFrame(), "Not enough data", ""

    work = grouped_df.copy().reset_index(drop=True)
    work["time_idx"] = np.arange(len(work))

    X = work[["time_idx"]]
    y = work.iloc[:, 1]
    model = LinearRegression()
    model.fit(X, y)

    future_idx = np.arange(len(work), len(work) + periods)
    future_dates = pd.date_range(
        start=work.iloc[-1, 0],
        periods=periods + 1,
        freq="MS",
    )[1:]
    preds = model.predict(pd.DataFrame({"time_idx": future_idx}))

    residuals = y - model.predict(X)
    band = float(np.std(residuals)) if len(residuals) > 1 else 0.0

    forecast_df = pd.DataFrame(
        {
            "Date": future_dates,
            "Forecast": preds,
            "Lower_Bound": preds - band,
            "Upper_Bound": preds + band,
        }
    )

    explanation = (
        "Linear trend forecast generated from the historical pattern in the selected metric."
    )
    return forecast_df, "Forecast model: LinearRegression", explanation


def generate_forecast_interpretation(
    forecast_df: pd.DataFrame,
    metric_name: str,
    language: str = "English",
) -> str:
    if forecast_df.empty:
        return ""
    first_pred = float(forecast_df["Forecast"].iloc[0])
    last_pred = float(forecast_df["Forecast"].iloc[-1])
    trend = "stable"
    if last_pred > first_pred:
        trend = "upward"
    elif last_pred < first_pred:
        trend = "downward"
    return (
        f"The forecast for {metric_name} shows a {trend} direction over the selected horizon.\n"
        f"This helps you anticipate likely business movement if recent patterns continue.\n"
        f"Use it as a planning signal for budget, staffing, inventory, or operations."
    )


def generate_forecast_recommendation(
    forecast_df: pd.DataFrame,
    metric_name: str,
    language: str = "English",
) -> str:
    if forecast_df.empty:
        return ""
    return (
        f"Monitor {metric_name} closely, review the projected direction, "
        f"and align planning decisions with the forecast trend."
    )


# =========================================================
# CLEANING / COMPARE / PDF
# =========================================================
def compare_dataframes(df_a: pd.DataFrame, df_b: pd.DataFrame) -> str:
    lines = [
        "Comparison Summary",
        "------------------",
        f"File A shape: {df_a.shape[0]} rows × {df_a.shape[1]} columns",
        f"File B shape: {df_b.shape[0]} rows × {df_b.shape[1]} columns",
        f"Row difference: {df_b.shape[0] - df_a.shape[0]}",
        f"Column difference: {df_b.shape[1] - df_a.shape[1]}",
    ]
    shared_numeric = sorted(
        set(df_a.select_dtypes(include="number").columns).intersection(
            df_b.select_dtypes(include="number").columns
        )
    )
    if shared_numeric:
        lines.append("")
        lines.append("Shared Numeric Column Changes:")
        for col in shared_numeric[:10]:
            a_mean = pd.to_numeric(df_a[col], errors="coerce").mean()
            b_mean = pd.to_numeric(df_b[col], errors="coerce").mean()
            lines.append(
                f"- {col}: mean {a_mean:.2f} → {b_mean:.2f} (Δ {b_mean - a_mean:.2f})"
            )
    else:
        lines.append("")
        lines.append("No shared numeric columns found for comparison.")
    return "\n".join(lines)


def apply_cleaning_actions(
    df: pd.DataFrame,
    remove_dup: bool,
    fill_num: bool,
    fill_text: bool,
    trim_spaces: bool,
    standardize_dates_flag: bool,
) -> pd.DataFrame:
    cleaned = df.copy()
    if remove_dup:
        cleaned = cleaned.drop_duplicates()
    if fill_num:
        for col in cleaned.select_dtypes(include="number").columns:
            cleaned[col] = cleaned[col].fillna(cleaned[col].median())
    if fill_text:
        for col in cleaned.select_dtypes(include=["object", "category"]).columns:
            cleaned[col] = cleaned[col].fillna("Unknown")
    if trim_spaces:
        for col in cleaned.select_dtypes(include=["object", "category"]).columns:
            cleaned[col] = cleaned[col].astype(str).str.strip()
    if standardize_dates_flag:
        for col in cleaned.columns:
            if "date" in str(col).lower():
                cleaned[col] = pd.to_datetime(cleaned[col], errors="coerce")
    return cleaned


def generate_pdf_report(text: str, filename: str = "executive_report.pdf") -> str:
    out_path = BASE_DIR / filename
    doc = SimpleDocTemplate(str(out_path), pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    for paragraph in str(text).split("\n\n"):
        story.append(
            Paragraph(
                html.escape(paragraph).replace("\n", "<br/>"),
                styles["Normal"],
            )
        )
        story.append(Spacer(1, 12))
    doc.build(story)
    return str(out_path)


# =========================================================
# SAVING / SHARING
# =========================================================
def report_payload(
    file_name: str,
    analysis: str,
    translation: str = "",
    answer: str = "",
    doctor: str = "",
) -> dict:
    return {
        "report_id": uuid.uuid4().hex[:14],
        "file_name": file_name,
        "analysis": analysis,
        "translation": translation,
        "answer": answer,
        "doctor": doctor,
        "created_at": now_iso(),
    }


def save_report_for_user(email: str, payload: dict) -> Optional[Path]:
    if not email or not payload:
        return None
    out_path = user_report_dir(email) / f"{payload['report_id']}.json"
    try:
        out_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return out_path
    except Exception:
        return None


def load_user_reports(email: str) -> List[dict]:
    if not email:
        return []
    folder = user_report_dir(email)
    results = []
    for path in sorted(folder.glob("*.json"), reverse=True):
        try:
            results.append(json.loads(path.read_text(encoding="utf-8")))
        except Exception:
            continue
    return results


def create_shareable_report_link(
    email: str,
    file_name: str,
    analysis: str,
    translation: str = "",
    doctor: str = "",
) -> str:
    share_id = uuid.uuid4().hex[:16]
    payload = {
        "share_id": share_id,
        "file_name": file_name,
        "analysis": analysis,
        "translation": translation,
        "doctor": doctor,
        "shared_by": email,
        "created_at": now_iso(),
    }
    out_path = SHARED_REPORTS_DIR / f"{share_id}.json"
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    if APP_BASE_URL:
        return f"{APP_BASE_URL}?share={share_id}"
    return f"?share={share_id}"


# =========================================================
# SAMPLE DATA
# =========================================================
def sample_dataset() -> pd.DataFrame:
    dates = pd.date_range("2025-01-01", periods=12, freq="MS")
    return pd.DataFrame(
        {
            "Date": dates,
            "Sales": [
                12000,
                13200,
                12800,
                14100,
                15500,
                16200,
                17000,
                16800,
                17600,
                18300,
                19100,
                20500,
            ],
            "Customers": [
                210,
                225,
                218,
                233,
                248,
                255,
                270,
                266,
                278,
                286,
                294,
                310,
            ],
            "Returns": [11, 14, 12, 16, 15, 17, 18, 16, 19, 20, 22, 21],
            "Region": ["South"] * 12,
        }
    )


# =========================================================
# WORLD-CLASS BUSINESS INTELLIGENCE ADD-ONS
# =========================================================
def compute_data_quality_score(df: pd.DataFrame) -> Tuple[int, list]:
    """Simple executive-grade quality score for fast decision confidence."""
    if df is None or df.empty:
        return 0, ["No data loaded."]

    rows = max(len(df), 1)
    cells = max(df.shape[0] * max(df.shape[1], 1), 1)
    missing_rate = float(df.isna().sum().sum()) / cells
    duplicate_rate = float(df.duplicated().sum()) / rows
    constant_cols = [c for c in df.columns if df[c].nunique(dropna=False) <= 1]

    score = 100
    score -= min(35, int(missing_rate * 100))
    score -= min(25, int(duplicate_rate * 100))
    score -= min(20, len(constant_cols) * 4)
    score = max(0, min(100, score))

    notes = []
    if missing_rate > 0:
        notes.append(f"Missing values affect {missing_rate:.1%} of cells.")
    if duplicate_rate > 0:
        notes.append(f"Duplicate rows affect {duplicate_rate:.1%} of records.")
    if constant_cols:
        notes.append("Constant columns detected: " + ", ".join(map(str, constant_cols[:8])))
    if not notes:
        notes.append("Dataset looks structurally healthy for first-pass analysis.")
    return score, notes


def render_executive_command_center(df: pd.DataFrame) -> None:
    """Indispensable industry layer: instant health, readiness, and decision cues."""
    if df is None:
        return

    score, notes = compute_data_quality_score(df)
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    text_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    date_like_cols = [c for c in df.columns if "date" in str(c).lower() or "time" in str(c).lower()]

    st.markdown("### 🌍 Executive Command Center")
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Data Quality Score", f"{score}/100")
    q2.metric("Numeric Fields", len(numeric_cols))
    q3.metric("Text Fields", len(text_cols))
    q4.metric("Date/Time Fields", len(date_like_cols))

    with st.expander("Decision Readiness Notes", expanded=True):
        for note in notes:
            st.write(f"• {note}")
        if score >= 85:
            st.success("This dataset is ready for executive reporting, forecasting, and AI analysis.")
        elif score >= 65:
            st.warning("This dataset is usable, but cleaning is recommended before major business decisions.")
        else:
            st.error("This dataset needs cleaning before it should drive important decisions.")

def generate_executive_dashboard(df: pd.DataFrame):
    if df is None or df.empty:
        return None

    metrics = compute_semantic_metrics(df)

    summary = {
        "rows": df.shape[0],
        "columns": df.shape[1],
        "missing": int(df.isna().sum().sum()),
        "duplicates": int(df.duplicated().sum())
    }

    insights = []

    if summary["missing"] > 0:
        insights.append("Dataset contains missing values that may affect accuracy.")

    if summary["duplicates"] > 0:
        insights.append("Duplicate records detected. Consider cleaning before analysis.")

    if not insights:
        insights.append("Dataset is clean and ready for analysis.")

    return metrics, summary, insights

def auto_fix_dataset(df: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    fixed = df.copy()
    actions = []

    before_rows = len(fixed)

    fixed = fixed.drop_duplicates()
    removed_duplicates = before_rows - len(fixed)

    if removed_duplicates > 0:
        actions.append(f"Removed {removed_duplicates} duplicate rows.")

    for col in fixed.select_dtypes(include="number").columns:
        missing_before = fixed[col].isna().sum()
        if missing_before > 0:
            fixed[col] = fixed[col].fillna(fixed[col].median())
            actions.append(f"Filled missing numeric values in '{col}' with median.")

    for col in fixed.select_dtypes(include=["object", "category"]).columns:
        missing_before = fixed[col].isna().sum()
        if missing_before > 0:
            fixed[col] = fixed[col].fillna("Unknown")
            actions.append(f"Filled missing text values in '{col}' with 'Unknown'.")

        fixed[col] = fixed[col].astype(str).str.strip()

    for col in fixed.columns:
        if "date" in str(col).lower() or "time" in str(col).lower():
            try:
                fixed[col] = pd.to_datetime(fixed[col], errors="coerce")
                actions.append(f"Standardized date/time column '{col}'.")
            except Exception:
                pass

    if not actions:
        actions.append("No major cleaning issues found. Dataset already looks clean.")

    return fixed, actions

def generate_decision_engine_report(
    df: pd.DataFrame,
    mode: str,
    raw_text: str,
    language: str = "English",
) -> str:
    if df is None or df.empty:
        return "No dataset loaded."

    local_summary = build_data_summary(df, mode, raw_text)

    if not ai_available():
        missing_total = int(df.isna().sum().sum())
        dup_total = int(df.duplicated().sum())

        return f"""
Decision Engine Report

1. Immediate Priority
Review dataset quality before making major decisions.

2. Main Risk
Missing values: {missing_total}
Duplicate rows: {dup_total}

3. Recommended Action
Clean the dataset, review trends, and validate key business columns.

4. Business Impact
Better data quality improves forecasting, reporting, and decision confidence.
"""

    prompt = f"""
Respond in {language}.

You are an executive business strategist and senior data analyst.

Analyze the dataset summary below and produce a decision-ready report with these sections:

1. Executive Decision Summary
2. Top 3 Business Opportunities
3. Top 3 Risks
4. Recommended Actions
5. What To Do First
6. Expected Business Impact
7. Questions Leaders Should Ask Next

Rules:
- Be practical.
- Do not invent facts.
- Base recommendations only on the dataset summary.
- Write like a boardroom advisor.

Dataset summary:
{local_summary}
"""

    return call_openai(prompt)

def run_scenario_simulator(base_value: float, change_percent: float) -> dict:
    adjusted_value = base_value * (1 + change_percent / 100)

    return {
        "base_value": base_value,
        "change_percent": change_percent,
        "adjusted_value": adjusted_value,
        "difference": adjusted_value - base_value,
    }
    
def get_industry_kpis(df: pd.DataFrame):
    industry = st.session_state.get("industry_template", "General Business")

    if df is None or df.empty:
        return []

    numeric_cols = df.select_dtypes(include="number").columns.tolist()

    if not numeric_cols:
        return [
            ("Rows", df.shape[0]),
            ("Columns", df.shape[1]),
        ]

    first_num = numeric_cols[0]
    second_num = numeric_cols[1] if len(numeric_cols) > 1 else numeric_cols[0]

    kpis = []

    if industry == "Finance":
        kpis.append(("Total Value", pd.to_numeric(df[first_num], errors="coerce").sum()))
        kpis.append(("Average Value", pd.to_numeric(df[first_num], errors="coerce").mean()))
        kpis.append(("Maximum Value", pd.to_numeric(df[first_num], errors="coerce").max()))

    elif industry == "Marketing":
        kpis.append(("Performance Total", pd.to_numeric(df[first_num], errors="coerce").sum()))
        kpis.append(("Average Performance", pd.to_numeric(df[first_num], errors="coerce").mean()))
        kpis.append(("Growth Signal", pd.to_numeric(df[second_num], errors="coerce").sum()))

    elif industry == "Sales":
        kpis.append(("Sales Signal", pd.to_numeric(df[first_num], errors="coerce").sum()))
        kpis.append(("Average Sale Signal", pd.to_numeric(df[first_num], errors="coerce").mean()))
        kpis.append(("Peak Value", pd.to_numeric(df[first_num], errors="coerce").max()))

    elif industry == "Logistics":
        kpis.append(("Volume Processed", pd.to_numeric(df[first_num], errors="coerce").sum()))
        kpis.append(("Operational Average", pd.to_numeric(df[first_num], errors="coerce").mean()))
        kpis.append(("Capacity Signal", pd.to_numeric(df[second_num], errors="coerce").sum()))

    elif industry == "Healthcare":
        kpis.append(("Activity Volume", pd.to_numeric(df[first_num], errors="coerce").sum()))
        kpis.append(("Average Measure", pd.to_numeric(df[first_num], errors="coerce").mean()))

    else:
        kpis.append(("Rows", df.shape[0]))
        kpis.append(("Columns", df.shape[1]))
        kpis.append(("Missing Values", int(df.isna().sum().sum())))

    return kpis

def save_dashboard_snapshot(df: pd.DataFrame, file_name: str) -> Optional[Path]:
    if df is None or df.empty:
        return None

    email = current_user_email() or "local_user"
    folder = user_report_dir(email)

    snapshot = {
        "dashboard_id": uuid.uuid4().hex[:12],
        "file_name": file_name,
        "created_at": now_iso(),
        "rows": int(df.shape[0]),
        "columns": int(df.shape[1]),
        "missing_values": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "industry_template": st.session_state.get("industry_template", "General Business"),
        "data_quality_score": compute_data_quality_score(df)[0],
    }

    out_path = folder / f"dashboard_{snapshot['dashboard_id']}.json"
    out_path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
    return out_path


def load_dashboard_snapshots() -> List[dict]:
    email = current_user_email() or "local_user"
    folder = user_report_dir(email)

    dashboards = []
    for path in sorted(folder.glob("dashboard_*.json"), reverse=True):
        try:
            dashboards.append(json.loads(path.read_text(encoding="utf-8")))
        except Exception:
            continue

    return dashboards

def generate_auto_dashboard(df: pd.DataFrame) -> dict:
    if df is None or df.empty:
        return {}

    numeric_cols = df.select_dtypes(include="number").columns.tolist()

    if not numeric_cols:
        return {}

    main_col = numeric_cols[0]

    summary = {
        "column": main_col,
        "total": float(df[main_col].sum()),
        "mean": float(df[main_col].mean()),
        "max": float(df[main_col].max()),
        "min": float(df[main_col].min()),
    }

    return summary

FREE_LIMIT = 2

# Compatibility wrappers kept here so the later section does not erase the premium logic above.
def get_user_plan():
    email = current_user_email()
    if st.session_state.get("developer_premium_override"):
        return PREMIUM_PLAN
    if email and email in PREMIUM_USERS:
        return PREMIUM_PLAN
    if st.session_state.get("premium_status") == "premium":
        return PREMIUM_PLAN
    if st.session_state.get("user_plan") == PREMIUM_PLAN:
        return PREMIUM_PLAN
    return FREE_PLAN

def set_user_plan(plan):
    st.session_state["user_plan"] = plan
    st.session_state["premium_status"] = "premium" if plan == PREMIUM_PLAN else "inactive"

def is_premium():
    return get_user_plan() == PREMIUM_PLAN
# =========================================================
# MAIN
# =========================================================
def render_how_it_works() -> None:
    st.markdown(
        """
        <div class="how-card">
            <div class="how-title">How it works</div>
            <div class="how-text">
                1. Upload your file (CSV, Excel, PDF, image, or ZIP)<br/>
                2. Review overview, charts, and quick dashboard<br/>
                3. Generate AI-powered analysis and diagnostics<br/>
                4. Ask follow-up questions and export results<br/>
                5. Save, share, or upgrade for premium features
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def main() -> None:
    init_state()
    inject_global_css()
    process_stripe_return()

    if not st.session_state.get("app_language"):
        st.session_state["app_language"] = detect_user_language()

    render_top_bar()
    render_hero()
    render_feature_pills()
    render_trust_bar()
    render_status_strip()
    st.markdown('<hr class="green-gradient-divider">', unsafe_allow_html=True)

    with st.sidebar:
        plan = "Premium" if is_premium_user() else "Free Plan"
        st.markdown(
            f"""
            <div style="
                margin:8px 0 14px;
                padding:12px 14px;
                border-radius:12px;
                background:#fef3c7;
                color:#92400e;
                font-weight:700;
                border:1px solid rgba(0,0,0,0.08);
            ">
                Current Plan: {plan}
            </div>
            """,
            unsafe_allow_html=True,
        )
        
        lang_names = list(LANGUAGES.keys())
        
        sorted_langs = ["English"] + sorted(
            [lang for lang in lang_names if lang != "English"],
            key=lambda x: x.lower(),
        )

        selected_lang = st.selectbox(
            "🌍 Language",
            sorted_langs,
            index=sorted_langs.index(st.session_state.get("app_language", "English")),
            key="language_selectbox",
        )
        
        st.session_state["app_language"] = selected_lang
        
        # 👇 ADD IT RIGHT HERE
        industry_template = st.selectbox(
            "🏢 Industry Template",
            [
                "General Business",
                "Finance",
                "Marketing",
                "Sales",
                "Healthcare",
                "Logistics",
                "Education",
                "Retail",
                "Operations",
            ],
            key="industry_template",
        )
        
        st.markdown(
            """
            <div style="
                color:#ffffff;
                font-size:1.15rem;
                font-weight:900;
                margin:14px 0 8px 0;
            ">
                🔐 Account Access
            </div>
            """,
            unsafe_allow_html=True,
        )
                    
        if not st.session_state.get("authenticated"):
            email = st.text_input(t("email"), placeholder=t("email_placeholder"), key="sidebar_email")
            password = st.text_input(t("password"), placeholder=t("password_placeholder"), type="password", key="sidebar_password")

            login_col, create_col = st.columns(2)
            with login_col:
                if st.button(t("login"), key="login_btn", use_container_width=True):
                    ok, msg = local_login(email, password)
                    render_message(msg, "success" if ok else "warning")
                    if ok:
                        st.rerun()
            with create_col:
                if st.button(t("create"), key="create_btn", use_container_width=True):
                    ok, msg = local_create_account(email, password)
                    render_message(msg, "success" if ok else "warning")
                    if ok:
                        st.rerun()

            if st.button("Developer Sign In", key="dev_signin_btn", use_container_width=True):
                st.session_state["authenticated"] = True
                st.session_state["auth_user"] = {"email": "developer@explainmydata.ai"}
                st.session_state["developer_premium_override"] = True
                st.session_state["premium_status"] = "premium"
                st.session_state["user_plan"] = "premium"
                st.rerun()
        else:
            st.success(f"Signed in as {current_user_email()}")
            st.checkbox("Developer Premium Override", key="developer_premium_override")
            if st.button(t("logout"), key="logout_btn", use_container_width=True):
                logout()
                st.rerun()

        if st.button(t("use_sample_dataset"), key="sample_btn", use_container_width=True):
            st.session_state["use_sample_data"] = True

        st.checkbox("I agree to the data handling terms for AI analysis", key="gdpr_consent")
        remaining = max(0, FREE_ANALYSIS_LIMIT - st.session_state.get("usage_count", 0))
        st.caption(f"Free analyses remaining: {remaining}")

        st.markdown("---")
        
        st.markdown("### 💎 Upgrade Your Experience")
        st.caption("Unlock advanced AI insights, reports, and decision tools.")

        if STRIPE_MONTHLY_LINK:
            st.link_button("💎 Premium", STRIPE_MONTHLY_LINK, use_container_width=True)

        if STRIPE_ONE_TIME_LINK:
            st.link_button("💳 One-Time", STRIPE_ONE_TIME_LINK, use_container_width=True)

                
    uploaded_file = st.file_uploader(t("upload_title"), type=SUPPORTED_UPLOADS, help=t("upload_subtle"), key="main_uploader")


    df = None
    mode = st.session_state.get("data_mode", "structured")
    raw_text = st.session_state.get("raw_text", "")
    loaded_file_name = st.session_state.get("loaded_filename", "")

    if st.session_state.get("use_sample_data"):
        df = sample_dataset()
        mode = "structured"
        raw_text = ""
        loaded_file_name = "sample_dataset.csv"
        st.session_state["data_mode"] = mode
        st.session_state["raw_text"] = raw_text
        st.session_state["loaded_filename"] = loaded_file_name

    if uploaded_file is not None:
        ok, msg = file_size_ok(uploaded_file)
        if not ok:
            render_message(msg, "error")
        else:
            try:
                df, mode, raw_text = parse_uploaded_file(uploaded_file)
                loaded_file_name = uploaded_file.name
                st.session_state["data_mode"] = mode
                st.session_state["raw_text"] = raw_text
                st.session_state["loaded_filename"] = loaded_file_name
                st.session_state["use_sample_data"] = False
                if st.session_state.get("authenticated"):
                    save_uploaded_file_for_user(current_user_email(), uploaded_file)
            except Exception as exc:
                render_message(f"Upload parse error: {exc}", "error")
                df = None

    control_col, _ = st.columns([1, 8])
    with control_col:
        if st.button("⚙️ Control Panel", key="control_btn"):
            st.session_state["show_controls"] = not st.session_state.get("show_controls", False)

    if st.session_state.get("show_controls"):
        st.markdown("### ⚙️ Control Panel")
        st.write("Plan:", get_user_plan())
        st.write("AI configured:", "Yes" if ai_available() else "No")
        st.write("Auth configured:", "Yes" if auth_configured() else "No")
        st.slider("Usage Count", 0, 10, key="usage_count")

    tabs = st.tabs([
        f"📊 {t('dashboard')}", "🤖 AI Dashboard", f"🧠 {t('data_doctor')}", f"🧭 {t('decision_engine')}", f"🧪 {t('scenario_simulator')}", f"💬 {t('ask_data')}",
        f"📋 {t('report')}", f"📈 {t('forecast')}", f"📂 {t('saved')}", f"🖼️ {t('dashboard_gallery')}",
        f"👤 {t('account')}", f"🖼️ {t('ocr')}", f"🔄 {t('compare')}",
        f"🧹 {t('cleaning')}", f"⚖️ {t('legal')}", 
    ])

    with tabs[0]:
        if df is None:
            render_message("Upload a dataset or click Use Sample Dataset to get started.", "info")
        else:
            st.subheader("Data Overview")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(t("rows"), f"{len(df):,}")
            c2.metric(t("columns"), f"{len(df.columns):,}")
            c3.metric(t("missing"), f"{int(df.isna().sum().sum()):,}")
            c4.metric(t("duplicates"), f"{int(df.duplicated().sum()):,}")
            render_semantic_metric_cards(compute_semantic_metrics(df))
            render_executive_command_center(df)
            
            st.markdown("### Preview")
            
            st.markdown("### 📊 Industry KPIs")
            
            if df is not None:
                kpis = get_industry_kpis(df)

            if not kpis:
                render_message(
                    "No industry KPIs available for this dataset. Try another industry template or upload data with numeric columns.",
                    "warning",
                )
            else:
                cols = st.columns(len(kpis))

                for i, (label, value) in enumerate(kpis):
                    try:
                        cols[i].metric(label, f"{value:,.2f}")
                    except Exception:
                        cols[i].metric(label, str(value))
                
            st.markdown("### 📊 Industry KPIs")
            
            if df is not None:
                kpis = get_industry_kpis(df)

                cols = st.columns(len(kpis))

                for i, (label, value) in enumerate(kpis):
                    try:
                        cols[i].metric(label, f"{value:,.2f}")
                    except:
                        cols[i].metric(label, str(value))
            
            st.markdown("### 📈 Industry Trend")

            numeric_cols = df.select_dtypes(include="number").columns.tolist()

            if numeric_cols:
                col_choice = numeric_cols[0]

                fig, ax = plt.subplots()
                ax.plot(df[col_choice].fillna(0))
                ax.set_title(f"{col_choice} Trend")

                st.pyplot(fig)
            
            st.markdown("### 📈 Executive Dashboard")

            if df is not None:
                if st.button("📈 Generate Executive Dashboard", key="exec_dashboard_btn"):
                    metrics, summary, insights = generate_executive_dashboard(df)

                    col1, col2, col3, col4 = st.columns(4)

                    col1.metric("Rows", summary["rows"])
                    col2.metric("Columns", summary["columns"])
                    col3.metric("Missing", summary["missing"])
                    col4.metric("Duplicates", summary["duplicates"])

                    st.markdown("#### 📊 Key Metrics")
                    render_semantic_metric_cards(metrics)

                    st.markdown("#### 🧠 Insights")
                    for item in insights:
                        st.write(f"• {item}")
            
            st.dataframe(df.head(50), use_container_width=True)
            numeric_cols = df.select_dtypes(include="number").columns.tolist()
            if numeric_cols:
                chart_col = st.selectbox("Quick chart metric", numeric_cols, key="dashboard_chart_col")
                fig, ax = plt.subplots()
                pd.to_numeric(df[chart_col], errors="coerce").dropna().plot(kind="hist", ax=ax)
                ax.set_title(f"Distribution of {chart_col}")
                ax.set_xlabel(chart_col)
                st.pyplot(fig)

    with tabs[1]:
        st.markdown("### 🤖 AI Auto Dashboard")

        if df is None:
            render_message("Upload a dataset first.", "info")
        else:
            if st.button("🤖 Build My Dashboard", key="auto_dashboard_btn"):

                result = generate_auto_dashboard(df)

                if not result:
                    render_message("Not enough numeric data for dashboard.", "warning")
                else:
                    st.markdown("#### 📊 AI Selected Metrics")

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Total", f"{result['total']:,.2f}")
                    c2.metric("Average", f"{result['mean']:,.2f}")
                    c3.metric("Max", f"{result['max']:,.2f}")
                    c4.metric("Min", f"{result['min']:,.2f}")

                    st.markdown("#### 📈 AI Trend")

                    fig, ax = plt.subplots()
                    ax.plot(df[result["column"]].fillna(0))
                    ax.set_title(f"{result['column']} Trend (AI Selected)")
                    st.pyplot(fig)
    
    with tabs[2]:
        st.subheader("AI Data Doctor")
        if df is None:
            render_message("Upload data first so the Data Doctor can inspect it.", "info")
        else:
            if st.button(t("run_ai_data_doctor"), key="doctor_btn"):
                st.session_state["doctor_result"] = build_ai_data_doctor_report(df, mode, raw_text, current_language())
            if st.session_state.get("doctor_result"):
                render_report_card("Data Doctor Report", st.session_state["doctor_result"], "#facc15")

        st.markdown("### 🩺 Auto-Fix Data Doctor")

        if df is None:
            render_message("Upload a dataset first.", "info")
        else:
            if st.button(f"🛠 {t('auto_fix_dataset')}", key="auto_fix_btn"):
                cleaned_df, actions = auto_fix_dataset(df)
                st.session_state["cleaned_df"] = cleaned_df

                render_message("Dataset cleaned successfully.", "success")

                st.markdown("#### Fixes Applied")
                for action in actions:
                    st.write(f"• {action}")

                st.markdown("#### Cleaned Dataset Preview")
                st.dataframe(cleaned_df.head(20), use_container_width=True)

                csv_data = cleaned_df.to_csv(index=False).encode("utf-8")

                st.download_button(
                    "⬇️ Download Cleaned CSV",
                    csv_data,
                    file_name="cleaned_dataset.csv",
                    mime="text/csv",
                    key="download_cleaned_csv_btn",
                )

    with tabs[3]:
        st.markdown("### 🧭 AI Decision Engine")

        if df is None:
            render_message("Upload a dataset first.", "info")
        else:
            st.write(
                "This feature turns your dataset into executive decisions, risks, opportunities, and recommended actions."
            )

            if st.button("🧭 Generate Decision Report", key="decision_engine_btn"):
                if require_feature("ai_insights", "Decision Engine"):
                    decision_report = generate_decision_engine_report(
                        df, mode, raw_text, current_language()
                    )
                    try:
                        decision_report = generate_decision_engine_report(
                            df,
                            mode,
                            raw_text,
                            current_language(),
                        )
                        st.session_state["decision_report"] = decision_report
                        render_report_card("AI Decision Engine Report", decision_report, "#22c55e")
                    except Exception as exc:
                        render_message(f"Decision Engine error: {exc}", "error")

                if st.session_state.get("decision_report"):
                    st.download_button(
                        "⬇️ Download Decision Report",
                        st.session_state["decision_report"],
                        file_name="decision_engine_report.txt",
                        mime="text/plain",
                        key="download_decision_report_btn",
                    )
    
    with tabs[4]:
        st.markdown("### 🧪 Scenario Simulator")

        if df is None:
            render_message("Upload a dataset first.", "info")
        else:
            numeric_cols = df.select_dtypes(include="number").columns.tolist()

            if not numeric_cols:
                render_message("No numeric columns available for simulation.", "warning")
            else:
                selected_metric = st.selectbox(
                    "Choose a numeric column to simulate",
                    numeric_cols,
                    key="scenario_metric_select",
                )

                base_value = float(pd.to_numeric(df[selected_metric], errors="coerce").sum())

                change_percent = st.slider(
                    "What-if change (%)",
                    -100.0,
                    100.0,
                    10.0,
                    step=1.0,
                    key="scenario_change_slider",
                )

                result = run_scenario_simulator(base_value, change_percent)

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Current Total", f"{result['base_value']:,.2f}")
                c2.metric("Change", f"{result['change_percent']}%")
                c3.metric("Projected Total", f"{result['adjusted_value']:,.2f}")
                c4.metric("Difference", f"{result['difference']:,.2f}")

                st.markdown("#### Business Meaning")
                if result["difference"] > 0:
                    st.success("This scenario shows a positive projected movement.")
                elif result["difference"] < 0:
                    st.warning("This scenario shows a negative projected movement.")
                else:
                    st.info("This scenario shows no projected movement.")
    
    with tabs[5]:
        st.subheader("Ask Your Data")
        if df is None:
            render_message("Upload data first to ask questions.", "info")
        else:
            question = st.text_area("Ask a question about this file", key="question_box")
            if st.button("Ask", key="ask_btn"):
                if not question.strip():
                    render_message("Please type a question first.", "warning")
                else:
                    try:
                        st.session_state["chat_answer"] = answer_followup_question(df, mode, raw_text, question, current_language())
                    except Exception as exc:
                        st.session_state["chat_answer"] = f"Could not generate answer: {exc}"
            if st.session_state.get("chat_answer"):
                render_report_card("Answer", st.session_state["chat_answer"], "#38bdf8")

    with tabs[6]:
        st.subheader("Consulting Report")
        if df is None:
            render_message("Upload data first to create a consulting report.", "info")
        else:
            focus = st.selectbox(t("analysis_focus"), ["Executive summary", "Sales growth", "Operations", "Risk", "Customer behavior", "Data quality"], key="analysis_focus_select")
            if st.button(t("generate_ai_insights"), key="insights_btn"):
                allowed, reason = analysis_allowed()
                if not st.session_state.get("gdpr_consent"):
                    render_message("Please agree to the data handling terms before AI analysis.", "warning")
                elif not allowed:
                    render_message(reason, "warning")
                else:
                    try:
                        st.session_state["result"] = generate_ai_analysis(df, mode, raw_text, focus, current_language())
                        increment_usage()
                    except Exception as exc:
                        st.session_state["result"] = f"Could not generate AI report: {exc}"
            if st.session_state.get("result"):
                render_report_card("Executive Consulting Report", st.session_state["result"], "#22c55e")
                pdf_path = generate_pdf_report(st.session_state["result"])
                with open(pdf_path, "rb") as pdf_file:
                    st.download_button(t("download_executive_report"), data=pdf_file.read(), file_name="executive_report.pdf", mime="application/pdf")
                if st.button("Save Report", key="save_report_btn"):
                    if not st.session_state.get("authenticated"):
                        render_message("Please sign in before saving reports.", "warning")
                    else:
                        payload = report_payload(loaded_file_name, st.session_state.get("result", ""), st.session_state.get("translated_result", ""), st.session_state.get("chat_answer", ""), st.session_state.get("doctor_result", ""))
                        saved_path = save_report_for_user(current_user_email(), payload)
                        render_message("Report saved successfully." if saved_path else "Report could not be saved.", "success" if saved_path else "error")
                if st.button("Create Shareable Report Link", key="share_report_btn"):
                    if require_feature("shareable_reports", "Shareable Reports"):
                        link = create_shareable_report_link(current_user_email(), loaded_file_name, st.session_state.get("result", ""), st.session_state.get("translated_result", ""), st.session_state.get("doctor_result", ""))
                        st.code(link)
    
        st.markdown("### 🧾 Boardroom PDF Report")

        if df is None:
            render_message("Upload a dataset first.", "info")
        else:
            if st.button("🧾 Generate Boardroom PDF", key="boardroom_pdf_btn"):
                if require_feature("saved_reports", "Boardroom PDF"):
                    report_text = (
                        "ExplainMyData AI - Boardroom Report\n\n"
                        f"File: {loaded_file_name}\n"
                        f"Rows: {df.shape[0]}\n"
                        f"Columns: {df.shape[1]}\n"
                        f"Missing Values: {int(df.isna().sum().sum())}\n"
                        f"Duplicate Rows: {int(df.duplicated().sum())}\n\n"
                        "Executive Insights\n"
                        "------------------\n"
                        f"{st.session_state.get('result', 'No AI insights generated yet.')}\n\n"
                        "Data Doctor Findings\n"
                        "--------------------\n"
                        f"{st.session_state.get('doctor_result', 'No Data Doctor report generated yet.')}\n\n"
                        "Decision Engine\n"
                        "---------------\n"
                        f"{st.session_state.get('decision_report', 'No Decision Engine report generated yet.')}\n"
                    )

                    pdf_path = generate_pdf_report(report_text, "boardroom_report.pdf")

                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download Boardroom PDF",
                            f,
                            file_name="boardroom_report.pdf",
                            mime="application/pdf",
                            key="download_boardroom_pdf_btn",
                        )
        
        st.markdown("### 🔗 Shareable Report")

        if df is None:
            render_message("Upload a dataset first.", "info")
        else:
            if st.button(f"🔗 {t('create_share_link')}", key="create_share_link_btn"):
                try:
                    analysis_text = st.session_state.get("result", "")
                    translation_text = st.session_state.get("translated_result", "")
                    doctor_text = st.session_state.get("doctor_result", "")

                    if not analysis_text and not doctor_text:
                        render_message(
                            "Generate AI insights or Data Doctor report first before sharing.",
                            "warning",
                        )
                    else:
                        share_link = create_shareable_report_link(
                            current_user_email(),
                            loaded_file_name,
                            analysis_text,
                            translation_text,
                            doctor_text,
                        )

                        st.session_state["share_link"] = share_link
                        render_message("Shareable report link created.", "success")

                except Exception as exc:
                    render_message(f"Share link error: {exc}", "error")

            if st.session_state.get("share_link"):
                st.markdown("#### Your Shareable Link")
                st.code(st.session_state["share_link"])

    with tabs[7]:
        st.subheader("Forecast")
        if df is None:
            render_message("Upload structured data first to forecast.", "info")
        else:
            date_candidates = [c for c in df.columns if "date" in str(c).lower() or "time" in str(c).lower()]
            numeric_cols = df.select_dtypes(include="number").columns.tolist()
            if not date_candidates or not numeric_cols:
                render_message("Forecast needs one date/time column and one numeric column.", "warning")
            else:
                date_col = st.selectbox("Date column", date_candidates, key="forecast_date_col")
                value_col = st.selectbox("Value column", numeric_cols, key="forecast_value_col")
                periods = st.slider("Forecast periods", 3, 24, 6, key="forecast_periods")
                if st.button("Build Forecast", key="forecast_btn"):
                    grouped = prepare_forecast_dataframe(df, date_col, value_col)
                    forecast_df, model_name, explanation = build_forecast_ml(grouped, periods)
                    st.session_state["forecast_df"] = forecast_df
                    st.session_state["forecast_model_name"] = model_name
                    st.session_state["forecast_explanation"] = explanation
                forecast_df = st.session_state.get("forecast_df")
                if isinstance(forecast_df, pd.DataFrame) and not forecast_df.empty:
                    st.write(st.session_state.get("forecast_model_name", ""))
                    st.dataframe(forecast_df, use_container_width=True)
                    fig, ax = plt.subplots()
                    forecast_df.plot(x="Date", y="Forecast", ax=ax)
                    ax.set_title("Forecast")
                    st.pyplot(fig)
                    render_report_card("Forecast Interpretation", generate_forecast_interpretation(forecast_df, value_col, current_language()), "#a78bfa")
                    render_report_card("Recommended Action", generate_forecast_recommendation(forecast_df, value_col, current_language()), "#22c55e")

    with tabs[8]:
        st.subheader("Saved Reports")
        if not st.session_state.get("authenticated"):
            render_message("Sign in to view saved reports.", "info")
        else:
            reports = load_user_reports(current_user_email())
            if not reports:
                render_message("No saved reports yet.", "info")
            for item in reports:
                with st.expander(f"{item.get('file_name', 'Report')} — {item.get('created_at', '')}"):
                    if item.get("analysis"):
                        render_report_card("Analysis", item["analysis"], "#22c55e")
                    if item.get("doctor"):
                        render_report_card("Data Doctor", item["doctor"], "#facc15")
                    if item.get("answer"):
                        render_report_card("Answer", item["answer"], "#38bdf8")

    with tabs[9]:
        st.markdown("### 🖼️ Saved Dashboard Gallery")

        if df is not None:
            if st.button("💾 Save Current Dashboard Snapshot", key="save_dashboard_snapshot_btn"):
                path = save_dashboard_snapshot(df, loaded_file_name)
                if path:
                    render_message("Dashboard snapshot saved.", "success")
                else:
                    render_message("Could not save dashboard snapshot.", "error")

        dashboards = load_dashboard_snapshots()

        if not dashboards:
            render_message("No dashboard snapshots saved yet.", "info")
        else:
            for item in dashboards:
                with st.expander(f"{item.get('file_name', 'Dashboard')} — {item.get('created_at', '')}"):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Rows", item.get("rows", 0))
                    c2.metric("Columns", item.get("columns", 0))
                    c3.metric("Missing", item.get("missing_values", 0))
                    c4.metric("Quality", f"{item.get('data_quality_score', 0)}/100")

                    st.write("Industry:", item.get("industry_template", "General Business"))
    
    with tabs[10]:
        st.subheader("Account")
        st.write("Authentication:", "Signed in" if st.session_state.get("authenticated") else "Not signed in")
        st.write("Email:", current_user_email() or "Not available")
        st.write("Plan:", get_user_plan())
        st.write("Usage count:", st.session_state.get("usage_count", 0))
        if st.session_state.get("authenticated"):
            if st.button("Delete my local saved data", key="delete_account_data_btn"):
                safe_remove_tree(user_upload_dir(current_user_email()))
                safe_remove_tree(user_report_dir(current_user_email()))
                render_message("Local saved files and reports were deleted.", "success")

    with tabs[11]:
        st.subheader("OCR")
        if st.session_state.get("ocr_preview") is not None:
            st.image(st.session_state["ocr_preview"], caption="OCR image preview", use_container_width=True)
        if st.session_state.get("ocr_text"):
            st.text_area("Extracted OCR text", st.session_state["ocr_text"], height=260)
        else:
            render_message("Upload an image file to extract OCR text.", "info")

    with tabs[12]:
        st.subheader("Compare Two Files")
        file_a = st.file_uploader("Upload File A", type=["csv", "xlsx", "xls"], key="compare_a")
        file_b = st.file_uploader("Upload File B", type=["csv", "xlsx", "xls"], key="compare_b")
        if file_a is not None and file_b is not None:
            try:
                df_a = load_structured_data(file_a)
                df_b = load_structured_data(file_b)
                render_report_card("Comparison", compare_dataframes(df_a, df_b), "#38bdf8")
            except Exception as exc:
                render_message(f"Comparison failed: {exc}", "error")

    with tabs[13]:
        st.subheader("Cleaning Lab")
        if df is None:
            render_message("Upload data first to clean it.", "info")
        else:
            remove_dup = st.checkbox("Remove duplicate rows", value=True, key="clean_dup")
            fill_num = st.checkbox("Fill numeric missing values with median", value=True, key="clean_num")
            fill_text = st.checkbox("Fill text missing values with Unknown", value=True, key="clean_text")
            trim_spaces = st.checkbox("Trim text spaces", value=True, key="clean_trim")
            standardize_dates_flag = st.checkbox("Standardize date columns", value=False, key="clean_dates")
            if st.button("Apply Cleaning", key="apply_cleaning_btn"):
                cleaned = apply_cleaning_actions(df, remove_dup, fill_num, fill_text, trim_spaces, standardize_dates_flag)
                st.session_state["cleaned_df"] = cleaned
                st.session_state["cleaning_summary"] = generate_cleaning_summary(df, cleaned, current_language())
            cleaned_df = st.session_state.get("cleaned_df")
            if isinstance(cleaned_df, pd.DataFrame):
                st.dataframe(cleaned_df.head(50), use_container_width=True)
                if st.session_state.get("cleaning_summary"):
                    render_report_card("Cleaning Summary", st.session_state["cleaning_summary"], "#22c55e")
                st.download_button("Download Cleaned CSV", cleaned_df.to_csv(index=False).encode("utf-8"), file_name="cleaned_data.csv", mime="text/csv")

    with tabs[14]:
        render_privacy_policy()
        st.markdown("---")
        render_terms_of_service()
        st.markdown("---")
        render_data_handling_notice()

    st.markdown('<hr class="green-gradient-divider">', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="footer-note">
            Clean data. Clear insights. Better decisions. Built with Streamlit, Python, AI, and business-first analytics.
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
