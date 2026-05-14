import io
import json
import os
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
try:
    from supabase import create_client
except Exception:
    create_client = None

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
        "user_email": "",
        "is_premium": False,
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
       # "developer_premium_override": False,
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
# Global language support. English stays first; other languages are sorted in the sidebar.
LANGUAGES = {
    "English": "en",
    "Arabic": "ar",
    "Bengali": "bn",
    "French": "fr",
    "German": "de",
    "Haitian Creole": "ht",
    "Hindi": "hi",
    "Indonesian": "id",
    "Italian": "it",
    "Japanese": "ja",
    "Korean": "ko",
    "Mandarin Chinese": "zh",
    "Nigerian Pidgin": "pcm",
    "Portuguese": "pt",
    "Russian": "ru",
    "Spanish": "es",
    "Swahili": "sw",
    "Tamil": "ta",
    "Telugu": "te",
    "Turkish": "tr",
    "Urdu": "ur",
    "Vietnamese": "vi",
}

LANGUAGE_ALIASES = {
    "en": "English", "fr": "French", "es": "Spanish", "ht": "Haitian Creole",
    "pt": "Portuguese", "ar": "Arabic", "de": "German", "zh": "Mandarin Chinese",
    "zh_cn": "Mandarin Chinese", "zh_tw": "Mandarin Chinese", "hi": "Hindi",
    "bn": "Bengali", "id": "Indonesian", "ur": "Urdu", "ru": "Russian",
    "ja": "Japanese", "ko": "Korean", "it": "Italian", "tr": "Turkish",
    "vi": "Vietnamese", "sw": "Swahili", "ta": "Tamil", "te": "Telugu", "pcm": "Nigerian Pidgin",
}

BASE_UI_TEXT = {
    "app_language": "Choose app language",
    "language": "Language",
    "current_plan": "Current Plan",
    "free_plan": "Free Plan",
    "premium": "Premium",
    "industry_template": "Industry Template",
    "account_access": "Account Access",
    "sign_in": "Sign in",
    "email": "Email",
    "email_placeholder": "Type your email",
    "password": "Password",
    "password_placeholder": "Type your password",
    "forgot_password": "Forgot password?",
    "reset_password_sent": "Password reset email sent if this account exists.",
    "reset_password_unavailable": "Password reset requires Supabase authentication to be configured.",
    "welcome_back": "Welcome back",
    "login_subtitle": "Sign in to manage reports, premium access, and saved dashboards.",
    "secure_workspace": "Secure workspace",
    "workspace_control_panel": "Workspace Control Panel",
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
    "quality": "Quality",
    "data_overview": "Data Overview",
    "preview": "Preview",
    "generate_ai_insights": "Generate AI Insights",
    "run_ai_data_doctor": "Run AI Data Doctor",
    "download_executive_report": "Download Executive Report",
    "analysis_focus": "Analysis focus",
    "dashboard": "Dashboard",
    "ai_dashboard": "AI Dashboard",
    "data_doctor": "Data Doctor",
    "ask_data": "Ask Your Data",
    "report": "Consulting Report",
    "forecast": "Forecast",
    "saved": "Saved Reports",
    "account": "Account",
    "ocr": "OCR",
    "decision_engine": "Decision Engine",
    "scenario_simulator": "Scenario Simulator",
    "dashboard_gallery": "Dashboard Gallery",
    "shareable_report": "Shareable Report",
    "boardroom_pdf": "Boardroom PDF",
    "industry_template_tab": "Industry Template",
    "generate_dashboard": "Generate Executive Dashboard",
    "auto_fix_dataset": "Auto-Fix Dataset",
    "generate_decision_report": "Generate Decision Report",
    "create_share_link": "Create Shareable Report Link",
    "save_dashboard": "Save Dashboard Snapshot",
    "compare": "Compare",
    "cleaning": "Cleaning Lab",
    "legal": "Legal",
    "control_panel": "Control Panel",
    "data_handling_terms": "I agree to the data handling terms for AI analysis",
    "free_analyses_remaining": "Free analyses remaining",
    "upgrade_title": "Upgrade Your Experience",
    "upgrade_subtitle": "Unlock advanced AI insights, reports, and decision tools.",
    "one_time": "One-Time",
    "how_it_works": "How it works",
    "about_us": "About Us",
    "upload_start": "Upload a dataset or click Use Sample Dataset to get started.",
    "no_data_clean": "Upload data first to clean it.",
    "no_data_forecast": "Upload structured data first to forecast.",
    "no_saved_reports": "No saved reports yet.",
    "sign_in_saved": "Sign in to view saved reports.",
    "footer_slogan": "Clean data. Clear insights. Better decisions.",
    "rights_reserved": "All Rights Reserved.",
    "upgrade_to_premium": "Upgrade to Premium",
    "checkout_created": "Checkout session created. Continue to Stripe below.",
    "continue_to_stripe": "Continue to Stripe Checkout",
    "premium_access_active": "Premium access is active.",
    "privacy_policy": "Privacy Policy",
    "terms_of_service": "Terms of Service",
    "how_data_handled": "How your data is handled",
    "delete_my_local_saved_data": "Delete my local saved data",
    "apply_cleaning": "Apply Cleaning",
    "executive_command_center": "Executive Command Center",
    "dataset_ready": "This dataset is ready for executive reporting, forecasting, and AI analysis.",
    "dataset_cleaning_recommended": "This dataset is usable, but cleaning is recommended before major business decisions.",
    "dataset_needs_cleaning": "This dataset needs cleaning before it should drive important decisions.",
    "please_log_in_first": "Please log in first.",
    "checkout_failed": "Checkout failed.",
    "yes": "Yes",
    "no": "No",
    "control_panel_title": "Control Panel",
    "ai_configured": "AI configured",
    "auth_configured": "Auth configured",
    "industry_kpis": "Industry KPIs",
    "industry_trend": "Industry Trend",
    "executive_dashboard": "Executive Dashboard",
    "key_metrics": "Key Metrics",
    "insights": "Insights",
    "quick_chart_metric": "Quick chart metric",
    "ai_auto_dashboard": "AI Auto Dashboard",
    "build_my_dashboard": "Build My Dashboard",
    "ai_selected_metrics": "AI Selected Metrics",
    "ai_trend": "AI Trend",
    "auto_fix_data_doctor": "Auto-Fix Data Doctor",
    "fixes_applied": "Fixes Applied",
    "cleaned_dataset_preview": "Cleaned Dataset Preview",
    "ai_decision_engine": "AI Decision Engine",
    "ai_decision_engine_report": "AI Decision Engine Report",
    "decision_engine_error": "Decision Engine error",
    "download_decision_report": "Download Decision Report",
    "scenario_simulator_title": "Scenario Simulator",
    "upload_dataset_first": "Upload a dataset first.",
    "no_numeric_simulation": "No numeric columns available for simulation.",
    "choose_numeric_column": "Choose a numeric column to simulate",
    "what_if_change": "What-if change (%)",
    "current_total": "Current Total",
    "change": "Change",
    "projected_total": "Projected Total",
    "difference": "Difference",
    "business_meaning": "Business Meaning",
    "positive_projected_movement": "This scenario shows a positive projected movement.",
    "negative_projected_movement": "This scenario shows a negative projected movement.",
    "no_projected_movement": "This scenario shows no projected movement.",
    "ask": "Ask",
    "ask_question_about_file": "Ask a question about this file",
    "upload_data_ask": "Upload data first to ask questions.",
    "please_type_question": "Please type a question first.",
    "answer": "Answer",
    "could_not_generate_answer": "Could not generate answer",
    "consulting_report": "Consulting Report",
    "upload_data_consulting": "Upload data first to create a consulting report.",
    "executive_summary": "Executive summary",
    "sales_growth": "Sales growth",
    "operations": "Operations",
    "risk": "Risk",
    "customer_behavior": "Customer behavior",
    "data_quality": "Data quality",
    "agree_data_terms_ai": "Please agree to the data handling terms before AI analysis.",
    "could_not_generate_ai_report": "Could not generate AI report",
    "executive_consulting_report": "Executive Consulting Report",
    "save_report": "Save Report",
    "please_sign_in_save": "Please sign in before saving reports.",
    "report_saved_successfully": "Report saved successfully.",
    "report_could_not_saved": "Report could not be saved.",
    "create_shareable_report_link": "Create Shareable Report Link",
    "shareable_reports": "Shareable Reports",
    "boardroom_pdf_report": "Boardroom PDF Report",
    "generate_boardroom_pdf": "Generate Boardroom PDF",
    "download_boardroom_pdf": "Download Boardroom PDF",
    "shareable_report_title": "Shareable Report",
    "generate_first_before_sharing": "Generate AI insights or Data Doctor report first before sharing.",
    "shareable_report_created": "Shareable report link created.",
    "share_link_error": "Share link error",
    "your_shareable_link": "Your Shareable Link",
    "forecast_needs_date_numeric": "Forecast needs one date/time column and one numeric column.",
    "date_column": "Date column",
    "value_column": "Value column",
    "forecast_periods": "Forecast periods",
    "build_forecast": "Build Forecast",
    "forecast_interpretation": "Forecast Interpretation",
    "recommended_action": "Recommended Action",
    "saved_dashboard_gallery": "Saved Dashboard Gallery",
    "save_current_dashboard_snapshot": "Save Current Dashboard Snapshot",
    "dashboard_snapshot_saved": "Dashboard snapshot saved.",
    "dashboard_snapshot_error": "Could not save dashboard snapshot.",
    "no_dashboard_snapshots": "No dashboard snapshots saved yet.",
    "authentication": "Authentication",
    "signed_in": "Signed in",
    "not_signed_in": "Not signed in",
    "not_available": "Not available",
    "plan": "Plan",
    "usage_count": "Usage count",
    "usage_count_label": "Usage Count",
    "supabase_connected": "Supabase connected",
    "local_saved_deleted": "Local saved files and reports were deleted.",
    "ocr_image_preview": "OCR image preview",
    "extracted_ocr_text": "Extracted OCR text",
    "upload_image_ocr": "Upload an image file to extract OCR text.",
    "upload_file_a": "Upload File A",
    "upload_file_b": "Upload File B",
    "comparison": "Comparison",
    "comparison_failed": "Comparison failed",
    "remove_duplicate_rows": "Remove duplicate rows",
    "fill_numeric_missing": "Fill numeric missing values with median",
    "fill_text_missing": "Fill text missing values with Unknown",
    "trim_text_spaces": "Trim text spaces",
    "standardize_date_columns": "Standardize date columns",
    "cleaning_summary": "Cleaning Summary",
    "download_cleaned_csv": "Download Cleaned CSV",
    "turn_messy_data": "Turn messy data into clean business insights",
    "fast_insights": "Fast Insights",
    "ai_powered": "AI Powered",
    "global_access": "Global Access",
    "pdf_export": "PDF Export",
    "business_ready": "Business Ready",
    "secure_file_handling": "Secure file handling",
    "consent_aware_ai": "Consent-aware AI",
    "privacy_first_design": "Privacy-first design",
    "fast_processing": "Fast processing",
    "ai_powered_insights": "AI-powered insights",
    "workspace": "Workspace",
}

TRANSLATIONS = {
    "en": BASE_UI_TEXT,
    "fr": {"language":"Langue","current_plan":"Forfait actuel","free_plan":"Forfait gratuit","industry_template":"Modèle sectoriel","account_access":"Accès au compte","email":"Email","email_placeholder":"Entrez votre email","password":"Mot de passe","password_placeholder":"Entrez votre mot de passe","login":"Connexion","create":"Créer","logout":"Déconnexion","use_sample_dataset":"Utiliser un jeu de données exemple","upload_title":"Téléchargez votre fichier ou document","rows":"Lignes","columns":"Colonnes","missing":"Manquants","duplicates":"Doublons","dashboard":"Tableau de bord","ai_dashboard":"Tableau IA","data_doctor":"Docteur des données","ask_data":"Questionner vos données","report":"Rapport conseil","forecast":"Prévision","saved":"Rapports enregistrés","account":"Compte","decision_engine":"Moteur de décision","scenario_simulator":"Simulateur de scénario","dashboard_gallery":"Galerie de tableaux","shareable_report":"Rapport partageable","compare":"Comparer","cleaning":"Nettoyage","legal":"Légal","generate_ai_insights":"Générer des insights IA","run_ai_data_doctor":"Lancer AI Data Doctor","analysis_focus":"Objectif de l’analyse","upgrade_title":"Améliorez votre expérience","how_it_works":"Comment ça marche","about_us":"À propos","footer_slogan":"Données propres. Insights clairs. Meilleures décisions."},
    "es": {"language":"Idioma","current_plan":"Plan actual","free_plan":"Plan gratis","industry_template":"Plantilla de industria","account_access":"Acceso a la cuenta","email":"Correo","email_placeholder":"Escribe tu correo","password":"Contraseña","password_placeholder":"Escribe tu contraseña","login":"Iniciar sesión","create":"Crear","logout":"Cerrar sesión","use_sample_dataset":"Usar datos de ejemplo","upload_title":"Sube tu archivo o documento","rows":"Filas","columns":"Columnas","missing":"Faltantes","duplicates":"Duplicados","dashboard":"Panel","ai_dashboard":"Panel de IA","data_doctor":"Doctor de datos","ask_data":"Pregunta a tus datos","report":"Informe consultivo","forecast":"Pronóstico","saved":"Informes guardados","account":"Cuenta","decision_engine":"Motor de decisión","scenario_simulator":"Simulador de escenarios","dashboard_gallery":"Galería de paneles","shareable_report":"Informe compartible","compare":"Comparar","cleaning":"Limpieza","legal":"Legal","generate_ai_insights":"Generar insights de IA","run_ai_data_doctor":"Ejecutar AI Data Doctor","analysis_focus":"Enfoque del análisis","upgrade_title":"Mejora tu experiencia","how_it_works":"Cómo funciona","about_us":"Sobre nosotros","footer_slogan":"Datos limpios. Ideas claras. Mejores decisiones."},
    "ht": {"language":"Lang","current_plan":"Plan aktyèl","free_plan":"Plan gratis","industry_template":"Modèl endistri","account_access":"Aksè kont","email":"Imèl","email_placeholder":"Ekri imèl ou","password":"Modpas","password_placeholder":"Ekri modpas ou","login":"Antre","create":"Kreye","logout":"Dekonekte","use_sample_dataset":"Sèvi ak done egzanp","upload_title":"Telechaje fichye oswa dokiman ou","rows":"Ranje","columns":"Kolòn","missing":"Ki manke","duplicates":"Diplikasyon","dashboard":"Tablo","ai_dashboard":"Tablo AI","data_doctor":"Doktè Done","ask_data":"Poze done yo kesyon","report":"Rapò konsiltasyon","forecast":"Previzyon","saved":"Rapò sove","account":"Kont","decision_engine":"Motè desizyon","scenario_simulator":"Similatè senaryo","dashboard_gallery":"Galri tablo","shareable_report":"Rapò pou pataje","compare":"Konpare","cleaning":"Netwayaj","legal":"Legal","generate_ai_insights":"Jenere analiz AI","run_ai_data_doctor":"Lanse AI Data Doctor","analysis_focus":"Fokus analiz la","upgrade_title":"Amelyore eksperyans ou","how_it_works":"Kijan li mache","about_us":"Sou nou","footer_slogan":"Done pwòp. Analiz klè. Pi bon desizyon."},
    "pt": {"language":"Idioma","current_plan":"Plano atual","free_plan":"Plano gratuito","industry_template":"Modelo de setor","account_access":"Acesso à conta","email":"Email","password":"Senha","login":"Entrar","create":"Criar","logout":"Sair","use_sample_dataset":"Usar dados de exemplo","upload_title":"Envie seu arquivo ou documento","rows":"Linhas","columns":"Colunas","missing":"Ausentes","duplicates":"Duplicados","dashboard":"Painel","ai_dashboard":"Painel de IA","data_doctor":"Doutor de Dados","ask_data":"Pergunte aos seus dados","report":"Relatório consultivo","forecast":"Previsão","saved":"Relatórios salvos","account":"Conta","compare":"Comparar","cleaning":"Limpeza","legal":"Legal","generate_ai_insights":"Gerar insights de IA","analysis_focus":"Foco da análise","upgrade_title":"Melhore sua experiência","how_it_works":"Como funciona","about_us":"Sobre nós"},
    "de": {"language":"Sprache","current_plan":"Aktueller Plan","free_plan":"Kostenloser Plan","industry_template":"Branchenvorlage","account_access":"Kontozugang","email":"E-Mail","password":"Passwort","login":"Anmelden","create":"Erstellen","logout":"Abmelden","use_sample_dataset":"Beispieldaten verwenden","upload_title":"Datei oder Dokument hochladen","rows":"Zeilen","columns":"Spalten","missing":"Fehlend","duplicates":"Duplikate","dashboard":"Dashboard","ai_dashboard":"KI-Dashboard","data_doctor":"Daten-Doktor","ask_data":"Daten fragen","report":"Beratungsbericht","forecast":"Prognose","saved":"Gespeicherte Berichte","account":"Konto","compare":"Vergleichen","cleaning":"Bereinigung","legal":"Rechtliches","generate_ai_insights":"KI-Einblicke generieren","analysis_focus":"Analysefokus","upgrade_title":"Erlebnis verbessern","how_it_works":"So funktioniert es","about_us":"Über uns"},
    "ar": {"language":"اللغة","current_plan":"الخطة الحالية","free_plan":"الخطة المجانية","industry_template":"قالب القطاع","account_access":"الوصول إلى الحساب","email":"البريد الإلكتروني","password":"كلمة المرور","login":"تسجيل الدخول","create":"إنشاء","logout":"تسجيل الخروج","use_sample_dataset":"استخدم بيانات تجريبية","upload_title":"حمّل ملفك أو مستندك","rows":"الصفوف","columns":"الأعمدة","missing":"القيم المفقودة","duplicates":"التكرارات","dashboard":"لوحة التحكم","ai_dashboard":"لوحة الذكاء الاصطناعي","data_doctor":"طبيب البيانات","ask_data":"اسأل بياناتك","report":"تقرير استشاري","forecast":"التوقعات","saved":"التقارير المحفوظة","account":"الحساب","compare":"مقارنة","cleaning":"تنظيف البيانات","legal":"قانوني","generate_ai_insights":"إنشاء رؤى بالذكاء الاصطناعي","analysis_focus":"محور التحليل","upgrade_title":"طوّر تجربتك","how_it_works":"كيف يعمل","about_us":"من نحن"},
    "zh": {"language":"语言","current_plan":"当前方案","free_plan":"免费方案","industry_template":"行业模板","account_access":"账户访问","email":"邮箱","password":"密码","login":"登录","create":"创建","logout":"退出登录","use_sample_dataset":"使用示例数据","upload_title":"上传文件或文档","rows":"行","columns":"列","missing":"缺失","duplicates":"重复","dashboard":"仪表板","ai_dashboard":"AI 仪表板","data_doctor":"数据医生","ask_data":"询问你的数据","report":"咨询报告","forecast":"预测","saved":"已保存报告","account":"账户","compare":"比较","cleaning":"清洗实验室","legal":"法律","generate_ai_insights":"生成 AI 洞察","analysis_focus":"分析重点","upgrade_title":"升级你的体验","how_it_works":"使用方式","about_us":"关于我们"},
    "hi": {"language":"भाषा","current_plan":"वर्तमान योजना","free_plan":"मुफ़्त योजना","industry_template":"उद्योग टेम्पलेट","account_access":"खाता पहुँच","email":"ईमेल","password":"पासवर्ड","login":"लॉग इन","create":"बनाएँ","logout":"लॉग आउट","use_sample_dataset":"नमूना डेटा उपयोग करें","upload_title":"अपनी फ़ाइल या दस्तावेज़ अपलोड करें","rows":"पंक्तियाँ","columns":"कॉलम","missing":"अनुपस्थित","duplicates":"डुप्लिकेट","dashboard":"डैशबोर्ड","ai_dashboard":"AI डैशबोर्ड","data_doctor":"डेटा डॉक्टर","ask_data":"अपने डेटा से पूछें","report":"परामर्श रिपोर्ट","forecast":"पूर्वानुमान","saved":"सहेजी गई रिपोर्ट","account":"खाता","compare":"तुलना","cleaning":"क्लीनिंग लैब","legal":"कानूनी","generate_ai_insights":"AI इनसाइट बनाएँ","analysis_focus":"विश्लेषण फ़ोकस","upgrade_title":"अपना अनुभव बेहतर करें","how_it_works":"यह कैसे काम करता है","about_us":"हमारे बारे में"},
    "bn": {"language":"ভাষা","current_plan":"বর্তমান প্ল্যান","free_plan":"ফ্রি প্ল্যান","industry_template":"ইন্ডাস্ট্রি টেমপ্লেট","account_access":"অ্যাকাউন্ট অ্যাক্সেস","email":"ইমেল","password":"পাসওয়ার্ড","login":"লগইন","create":"তৈরি করুন","logout":"লগ আউট","use_sample_dataset":"নমুনা ডেটা ব্যবহার করুন","upload_title":"আপনার ফাইল বা ডকুমেন্ট আপলোড করুন","rows":"সারি","columns":"কলাম","missing":"অনুপস্থিত","duplicates":"ডুপ্লিকেট","dashboard":"ড্যাশবোর্ড","ai_dashboard":"AI ড্যাশবোর্ড","data_doctor":"ডেটা ডাক্তার","ask_data":"আপনার ডেটাকে জিজ্ঞাসা করুন","report":"পরামর্শ প্রতিবেদন","forecast":"পূর্বাভাস","saved":"সংরক্ষিত রিপোর্ট","account":"অ্যাকাউন্ট","compare":"তুলনা","cleaning":"ক্লিনিং ল্যাব","legal":"আইনি","generate_ai_insights":"AI ইনসাইট তৈরি করুন","analysis_focus":"বিশ্লেষণের ফোকাস","upgrade_title":"আপনার অভিজ্ঞতা উন্নত করুন","how_it_works":"এটি কীভাবে কাজ করে","about_us":"আমাদের সম্পর্কে"},
    "id": {"language":"Bahasa","current_plan":"Paket saat ini","free_plan":"Paket gratis","industry_template":"Templat industri","account_access":"Akses akun","email":"Email","password":"Kata sandi","login":"Masuk","create":"Buat","logout":"Keluar","use_sample_dataset":"Gunakan data contoh","upload_title":"Unggah file atau dokumen Anda","rows":"Baris","columns":"Kolom","missing":"Hilang","duplicates":"Duplikat","dashboard":"Dasbor","ai_dashboard":"Dasbor AI","data_doctor":"Dokter Data","ask_data":"Tanya Data Anda","report":"Laporan konsultasi","forecast":"Prakiraan","saved":"Laporan tersimpan","account":"Akun","compare":"Bandingkan","cleaning":"Lab pembersihan","legal":"Legal","generate_ai_insights":"Buat insight AI","analysis_focus":"Fokus analisis","upgrade_title":"Tingkatkan pengalaman Anda","how_it_works":"Cara kerja","about_us":"Tentang kami"},
    "ur": {"language":"زبان","current_plan":"موجودہ پلان","free_plan":"مفت پلان","industry_template":"صنعتی ٹیمپلیٹ","account_access":"اکاؤنٹ رسائی","email":"ای میل","password":"پاس ورڈ","login":"لاگ ان","create":"بنائیں","logout":"لاگ آؤٹ","use_sample_dataset":"نمونہ ڈیٹا استعمال کریں","upload_title":"اپنی فائل یا دستاویز اپ لوڈ کریں","rows":"قطاریں","columns":"کالم","missing":"غائب","duplicates":"نقل","dashboard":"ڈیش بورڈ","ai_dashboard":"AI ڈیش بورڈ","data_doctor":"ڈیٹا ڈاکٹر","ask_data":"اپنے ڈیٹا سے پوچھیں","report":"مشاورتی رپورٹ","forecast":"پیش گوئی","saved":"محفوظ رپورٹس","account":"اکاؤنٹ","compare":"موازنہ","cleaning":"کلیننگ لیب","legal":"قانونی","generate_ai_insights":"AI بصیرت بنائیں","analysis_focus":"تجزیہ کا مرکز","upgrade_title":"اپنا تجربہ بہتر بنائیں","how_it_works":"یہ کیسے کام کرتا ہے","about_us":"ہمارے بارے میں"},
    "ru": {"language":"Язык","current_plan":"Текущий план","free_plan":"Бесплатный план","industry_template":"Отраслевой шаблон","account_access":"Доступ к аккаунту","email":"Email","password":"Пароль","login":"Войти","create":"Создать","logout":"Выйти","use_sample_dataset":"Использовать пример данных","upload_title":"Загрузите файл или документ","rows":"Строки","columns":"Столбцы","missing":"Пропуски","duplicates":"Дубликаты","dashboard":"Панель","ai_dashboard":"AI-панель","data_doctor":"Доктор данных","ask_data":"Спросить данные","report":"Консультационный отчет","forecast":"Прогноз","saved":"Сохраненные отчеты","account":"Аккаунт","compare":"Сравнить","cleaning":"Очистка","legal":"Правовая информация","generate_ai_insights":"Создать AI-инсайты","analysis_focus":"Фокус анализа","upgrade_title":"Улучшите опыт","how_it_works":"Как это работает","about_us":"О нас"},
}

for _code in ["it", "ja", "ko", "pcm", "sw", "ta", "te", "tr", "vi"]:
    TRANSLATIONS.setdefault(_code, {})

UI_TEXT = {
    "English": BASE_UI_TEXT,
    "French": TRANSLATIONS["fr"],
    "Spanish": TRANSLATIONS["es"],
    "Haitian Creole": TRANSLATIONS["ht"],
    "Portuguese": TRANSLATIONS["pt"],
    "German": TRANSLATIONS["de"],
    "Arabic": TRANSLATIONS["ar"],
    "Mandarin Chinese": TRANSLATIONS["zh"],
    "Hindi": TRANSLATIONS["hi"],
    "Bengali": TRANSLATIONS["bn"],
    "Indonesian": TRANSLATIONS["id"],
    "Urdu": TRANSLATIONS["ur"],
    "Russian": TRANSLATIONS["ru"],
}

# =========================================================
# HELPERS
# =========================================================
@st.cache_data(show_spinner=False)
def auto_translate(text: str, target_language: str) -> str:
    """Optional UI fallback translator.

    Keep this disabled by default so the app does not burn API calls on every rerun.
    Turn it on only if needed by adding ENABLE_AI_UI_TRANSLATION=true to .env.
    """
    if not text or not target_language or target_language == "English":
        return text
    if os.getenv("ENABLE_AI_UI_TRANSLATION", "false").strip().lower() != "true":
        return text
    if not ai_available():
        return text

    prompt = f"Translate this short app UI label into {target_language}. Return only the translation:\n{text}"
    try:
        return call_openai(prompt)
    except Exception:
        return text


def t(key: str) -> str:
    """Translate a UI key safely with stable fallbacks."""
    lang_name = st.session_state.get("app_language", "English")
    lang_code = LANGUAGES.get(lang_name, "en")

    text = (
        TRANSLATIONS.get(lang_code, {}).get(key)
        or UI_TEXT.get(lang_name, {}).get(key)
        or BASE_UI_TEXT.get(key)
        or TRANSLATIONS.get("en", {}).get(key)
    )

    if text and lang_code != "en" and key not in TRANSLATIONS.get(lang_code, {}):
        text = auto_translate(text, lang_name)

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
    return bool(SUPABASE_URL and SUPABASE_ANON_KEY and create_client is not None)


@st.cache_resource(show_spinner=False)
def get_supabase_client():
    """Return a cached Supabase client when configuration is available."""
    if not auth_configured():
        return None
    try:
        return create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
    except Exception as exc:
        print("Supabase client error:", exc)
        return None


def detect_user_language() -> str:
    try:
        loc = locale.getdefaultlocale()[0]
        if not loc:
            return "English"
        loc = loc.lower().replace("-", "_")
        parts = loc.split("_")
        candidates = [loc] + parts
        for code in candidates:
            if code in LANGUAGE_ALIASES:
                return LANGUAGE_ALIASES[code]
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

    
def set_user_plan(plan_name: str) -> None:
    st.session_state["user_plan"] = plan_name
    st.session_state["premium_status"] = (
        "premium" if plan_name == PREMIUM_PLAN else "inactive"
    )


def is_premium_user() -> bool:
    return bool(st.session_state.get("is_premium", False))


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
    """No-op on purpose.

    Premium activation should come only from the Stripe webhook updating
    Supabase user_profiles.is_premium. Do not unlock premium from URL
    query parameters inside Streamlit.
    """
    return None

def render_upgrade_checkout_button():
    """Render Stripe Payment Link buttons without calling a local checkout backend."""
    monthly_url = os.getenv("STRIPE_PAYMENT_LINK_SUBSCRIPTION", "").strip()
    one_time_url = os.getenv("STRIPE_PAYMENT_LINK_ONE_TIME", "").strip()

    if is_premium_user():
        st.success(t("premium_access_active"))
        return

    st.caption("Monthly plan: billed monthly. Cancel anytime.")
    if monthly_url:
        st.link_button("💎 Monthly Premium", monthly_url, use_container_width=True)
    else:
        st.warning("Monthly Stripe payment link is missing.")

    st.caption("Lifetime plan: one-time payment. No monthly fees.")
    if one_time_url:
        st.link_button("💳 Lifetime Access", one_time_url, use_container_width=True)
    else:
        st.warning("One-time Stripe payment link is missing.")

    st.caption("Secure payments powered by Stripe.")


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
            height: 14px;
            background: linear-gradient(
            90deg,
            #001f12,
            #064e3b,
            #22c55e,
            #ccff00,
            #22c55e,
            #064e3b,
            #001f12
        );
        background-size: 350% 100%;
        animation: greenFlow 5s linear infinite;
        border-radius: 0 0 14px 14px;
        margin-top: -6px;
        margin-bottom: 6px;
        box-shadow:
            0 4px 12px rgba(34, 197, 94, 0.35),
            0 0 14px rgba(34, 197, 94, 0.25);
    }
        
        /* Restore original sidebar background */
        section[data-testid="stSidebar"] {
            background-color: #f8fafc !important;  /* light gray */
        }
        
/* Sidebar spacing system */
.sidebar-section {
    margin-top: 24px !important;
}

.sidebar-title {
    font-size: 0.85rem;
    font-weight: 700;
    color: #d1d5db;
    margin-bottom: 6px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
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
        
/* Fix text visibility inside premium plan box */
.premium-sidebar-box,
.premium-sidebar-box * {
    color: #92400e !important;
}
        
/* Remove thin vertical divider lines around sidebar selectboxes */
section[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"],
section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] {
border: none !important;
box-shadow: none !important;
}

        /* Remove left/right borders from language and industry dropdown areas */
        section[data-testid="stSidebar"] div[data-baseweb="select"],
        section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        border-left: none !important;
        border-right: none !important;
        box-shadow: none !important;
        
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
            font-size: 1.08rem;
            font-weight: 800;
            color: #e5f4ff !important;
            margin-top: 4px;
            letter-spacing: 0.01em;
            text-shadow: 0 2px 10px rgba(0,0,0,0.35);
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

.how-card {
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
    background: #fef3c7 !important;
    border-radius: 18px !important;
    padding: 14px 16px !important;
    margin: 12px 0 18px 0 !important;
    color: #111827 !important;
}

.how-card * {
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

.block-container {
    padding-top: 0.35rem !important;
    padding-bottom: 0.2rem !important;
}

.hero-wrap {
    margin-bottom: 6px !important;
    padding: 10px !important;
}

.feature-pill-wrap {
    margin: 6px 0 8px 0 !important;
}

.trust-bar {
    margin: 6px 0 8px 0 !important;
    padding: 8px 12px !important;
}

.status-chip-wrap {
    margin: 6px 0 10px 0 !important;
}

/* Fix Premium/Upgrade text in sidebar */
section[data-testid="stSidebar"] .premium-sidebar-box,
section[data-testid="stSidebar"] .premium-sidebar-box * {
    color: #111827 !important;
    opacity: 1 !important;
    visibility: visible !important;
}

/* Push premium sidebar section lower */
.premium-sidebar-box {
    margin-top: 28px !important;
    background: #fef3c7 !important;
    border-radius: 14px !important;
    padding: 14px !important;
    border: 1px solid #facc15 !important;
}

div[data-testid="stTabs"] {
    margin-top: 8px !important;
}


/* FINAL SAFE SIDEBAR POLISH */
section[data-testid="stSidebar"] {
    background-color: #f8fafc !important;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span {
    color: #111827 !important;
    opacity: 1 !important;
    visibility: visible !important;
}

section[data-testid="stSidebar"] div[data-baseweb="checkbox"] label {
    color: #111827 !important;
}

.account-sidebar-box {
    background: #ffffff !important;
    border: 1px solid #e5e7eb !important;
    border-radius: 14px !important;
    padding: 12px !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.06) !important;
}

.account-signed-in-box {
    background: #ecfdf5 !important;
    border: 1px solid #bbf7d0 !important;
    border-radius: 12px !important;
    padding: 10px 12px !important;
    margin-bottom: 8px !important;
}

.account-status-label {
    color: #166534 !important;
    font-size: 0.78rem !important;
    font-weight: 900 !important;
}

.account-email {
    color: #111827 !important;
    font-size: 0.86rem !important;
    font-weight: 700 !important;
    word-break: break-word !important;
}

.premium-sidebar-box {
    margin-top: 24px !important;
    background: #fef3c7 !important;
    border: 1px solid #facc15 !important;
    border-radius: 16px !important;
    padding: 14px !important;
    box-shadow: 0 6px 16px rgba(0,0,0,0.10) !important;
}

.premium-sidebar-box,
.premium-sidebar-box * {
    opacity: 1 !important;
    visibility: visible !important;
    text-shadow: none !important;
}

.plan-card-inner {
    background: transparent !important;
    color: #92400e !important;
}

.plan-row {
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    gap: 8px !important;
    margin-bottom: 6px !important;
}

.plan-title {
    color: #92400e !important;
    font-size: 0.86rem !important;
    font-weight: 900 !important;
}

.plan-badge {
    color: #ffffff !important;
    background: #16a34a !important;
    border-radius: 999px !important;
    padding: 3px 8px !important;
    font-size: 0.68rem !important;
    font-weight: 900 !important;
    letter-spacing: 0.04em !important;
}

.plan-name {
    color: #92400e !important;
    font-size: 1.08rem !important;
    font-weight: 900 !important;
}

.upgrade-subtitle {
    color: #374151 !important;
    font-size: 0.82rem !important;
    font-weight: 700 !important;
    line-height: 1.45 !important;
    margin-top: 8px !important;
    margin-bottom: 8px !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] button,
section[data-testid="stSidebar"] div[data-testid="stLinkButton"] a {
    background: #f5f5dc !important;
    color: #111827 !important;
    font-weight: 800 !important;
    border-radius: 10px !important;
    border: 1px solid #e5e7eb !important;
}

/* Hide Streamlit input helper: "Press Enter to apply" */
div[data-testid="InputInstructions"] {
    display: none !important;
}

/* Backup: hide only input instructions, not every small tag */
div[data-testid="stTextInput"] small,
div[data-testid="stNumberInput"] small,
div[data-testid="stTextArea"] small {
    display: none !important;
}


/* ===== ELITE SAAS SIDEBAR + LOGIN CARD POLISH ===== */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f8fafc 0%, #eef2ff 42%, #ecfdf5 100%) !important;
    border-right: 1px solid rgba(15, 23, 42, 0.08) !important;
}

section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
    padding-top: 1rem !important;
}

.sidebar-control-shell {
    background: linear-gradient(135deg, #0f172a, #064e3b) !important;
    color: #ffffff !important;
    border-radius: 20px !important;
    padding: 16px 14px !important;
    margin: 4px 0 16px 0 !important;
    box-shadow: 0 14px 34px rgba(15, 23, 42, 0.22) !important;
}

.sidebar-control-shell * {
    color: #ffffff !important;
}

.sidebar-control-title {
    font-size: 1rem !important;
    font-weight: 950 !important;
    letter-spacing: -0.02em !important;
}

.sidebar-control-subtitle {
    font-size: 0.78rem !important;
    font-weight: 700 !important;
    color: #d1fae5 !important;
    margin-top: 4px !important;
    line-height: 1.35 !important;
}

.sidebar-section {
    margin-top: 14px !important;
}

.sidebar-title {
    color: #475569 !important;
    font-size: 0.73rem !important;
    font-weight: 950 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    margin-bottom: 8px !important;
}

.account-sidebar-box,
.sidebar-settings-card {
    background: rgba(255, 255, 255, 0.92) !important;
    border: 1px solid rgba(226, 232, 240, 0.95) !important;
    border-radius: 18px !important;
    padding: 14px !important;
    box-shadow: 0 10px 26px rgba(15, 23, 42, 0.08) !important;
    backdrop-filter: blur(10px) !important;
}

.login-card {
    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%) !important;
    border: 1px solid rgba(226, 232, 240, 0.98) !important;
    border-radius: 20px !important;
    padding: 16px 14px !important;
    box-shadow: 0 14px 32px rgba(15, 23, 42, 0.10) !important;
}

.login-title {
    color: #0f172a !important;
    font-size: 1.1rem !important;
    font-weight: 950 !important;
    text-align: center !important;
    margin-bottom: 4px !important;
}

.login-subtitle {
    color: #64748b !important;
    font-size: 0.78rem !important;
    font-weight: 700 !important;
    text-align: center !important;
    margin-bottom: 12px !important;
    line-height: 1.45 !important;
}

.login-card .secure-pill {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 6px !important;
    width: 100% !important;
    color: #166534 !important;
    background: #dcfce7 !important;
    border: 1px solid #bbf7d0 !important;
    border-radius: 999px !important;
    font-size: 0.72rem !important;
    font-weight: 900 !important;
    padding: 6px 8px !important;
    margin-bottom: 12px !important;
}

section[data-testid="stSidebar"] div[data-testid="stTextInput"] input {
    border-radius: 12px !important;
    border: 1px solid #dbe3ef !important;
    background: #ffffff !important;
    color: #111827 !important;
    box-shadow: inset 0 1px 2px rgba(15, 23, 42, 0.03) !important;
}

section[data-testid="stSidebar"] div[data-testid="stTextInput"] label,
section[data-testid="stSidebar"] div[data-testid="stSelectbox"] label {
    color: #334155 !important;
    font-weight: 850 !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] button {
    border-radius: 12px !important;
    min-height: 38px !important;
    font-weight: 900 !important;
    border: 1px solid rgba(15, 23, 42, 0.08) !important;
    box-shadow: 0 6px 14px rgba(15, 23, 42, 0.08) !important;
}

section[data-testid="stSidebar"] div[data-testid="stButton"] button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 10px 20px rgba(15, 23, 42, 0.13) !important;
}

.login-card div[data-testid="stButton"] button,
.premium-sidebar-box div[data-testid="stButton"] button {
    background: linear-gradient(90deg, #16a34a, #22c55e) !important;
    color: #ffffff !important;
    border: none !important;
}

.login-card div[data-testid="stButton"] button *,
.premium-sidebar-box div[data-testid="stButton"] button * {
    color: #ffffff !important;
}

.forgot-password-note {
    color: #2563eb !important;
    text-align: center !important;
    font-size: 0.76rem !important;
    font-weight: 850 !important;
    margin: 6px 0 0 0 !important;
}

.premium-sidebar-box {
    background: linear-gradient(135deg, #fffbeb, #fef3c7) !important;
    border: 1px solid #facc15 !important;
    border-radius: 20px !important;
    padding: 16px !important;
    box-shadow: 0 14px 30px rgba(146, 64, 14, 0.15) !important;
}

.plan-badge {
    color: #ffffff !important;
    background: linear-gradient(90deg, #16a34a, #22c55e) !important;
}

.plan-name, .plan-title, .plan-card-inner {
    color: #92400e !important;
}


/* ===== COMPACT STRIPE-STYLE ACCOUNT / PLAN / USAGE POLISH ===== */
section[data-testid="stSidebar"] hr,
section[data-testid="stSidebar"] .green-gradient-divider {
    display: none !important;
    height: 0 !important;
    margin: 0 !important;
    border: 0 !important;
}

section[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"],
section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] {
    border: none !important;
    box-shadow: none !important;
}

.sidebar-user-mini-card {
    display: flex !important;
    align-items: center !important;
    gap: 10px !important;
    background: #ffffff !important;
    border: 1px solid #e5e7eb !important;
    border-radius: 16px !important;
    padding: 10px 12px !important;
    margin: 10px 0 12px 0 !important;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.07) !important;
}

.sidebar-avatar {
    width: 38px !important;
    height: 38px !important;
    min-width: 38px !important;
    border-radius: 999px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    background: linear-gradient(135deg, #064e3b, #22c55e) !important;
    color: #ffffff !important;
    font-size: 0.95rem !important;
    font-weight: 950 !important;
    box-shadow: 0 8px 16px rgba(22, 163, 74, 0.24) !important;
}

.sidebar-user-details {
    min-width: 0 !important;
    flex: 1 !important;
}

.sidebar-user-label {
    color: #64748b !important;
    font-size: 0.68rem !important;
    font-weight: 950 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
    line-height: 1.1 !important;
}

.sidebar-user-email {
    color: #0f172a !important;
    font-size: 0.82rem !important;
    font-weight: 850 !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    max-width: 180px !important;
}

.sidebar-plan-pill-row {
    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 8px !important;
    margin-top: 8px !important;
}

.stripe-plan-badge {
    color: #ffffff !important;
    background: #111827 !important;
    border-radius: 999px !important;
    padding: 4px 10px !important;
    font-size: 0.68rem !important;
    font-weight: 950 !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    box-shadow: inset 0 0 0 1px rgba(255,255,255,0.10) !important;
}

.stripe-plan-badge.free {
    background: linear-gradient(90deg, #64748b, #334155) !important;
}

.stripe-plan-badge.premium {
    background: linear-gradient(90deg, #16a34a, #22c55e) !important;
}

.usage-compact-card {
    background: rgba(255, 255, 255, 0.88) !important;
    border: 1px solid #e5e7eb !important;
    border-radius: 16px !important;
    padding: 12px !important;
    margin-top: 16px !important;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.06) !important;
}

.usage-compact-top {
    display: flex !important;
    justify-content: space-between !important;
    align-items: center !important;
    gap: 8px !important;
    margin-bottom: 8px !important;
}

.usage-compact-label {
    color: #334155 !important;
    font-size: 0.76rem !important;
    font-weight: 950 !important;
}

.usage-compact-count {
    color: #0f172a !important;
    font-size: 0.74rem !important;
    font-weight: 950 !important;
    background: #f1f5f9 !important;
    border-radius: 999px !important;
    padding: 3px 8px !important;
}

.usage-progress-track {
    width: 100% !important;
    height: 8px !important;
    border-radius: 999px !important;
    background: #e5e7eb !important;
    overflow: hidden !important;
}

.usage-progress-fill {
    height: 100% !important;
    border-radius: 999px !important;
    background: linear-gradient(90deg, #16a34a, #22c55e) !important;
}

.usage-compact-note {
    color: #64748b !important;
    font-size: 0.70rem !important;
    font-weight: 750 !important;
    margin-top: 7px !important;
    line-height: 1.35 !important;
}

.forgot-password-note {
    display: none !important;
}

.upgrade-subtitle {
    color: #475569 !important;
}

section[data-testid="stSidebar"] div[data-testid="stCheckbox"] label,
section[data-testid="stSidebar"] div[data-baseweb="checkbox"] label {
    color: #111827 !important;
}

/* Hide only Streamlit input helper text such as 'Press Enter to apply'. Do not hide all small text. */
div[data-testid="InputInstructions"] {
    display: none !important;
}


/* FINAL SIDEBAR CLEANUP: no separator cards above Language/Account Access */
.sidebar-control-shell {
    display: none !important;
}

.sidebar-settings-card {
    margin-top: 6px !important;
}

.account-sidebar-box {
    margin-top: 12px !important;
}

.sidebar-user-mini-card-inside {
    margin-top: 0 !important;
    margin-bottom: 12px !important;
    box-shadow: none !important;
    background: #f8fafc !important;
}

section[data-testid="stSidebar"] hr,
section[data-testid="stSidebar"] .green-gradient-divider,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] hr {
    display: none !important;
    height: 0 !important;
    margin: 0 !important;
    border: 0 !important;
}



/* ===== FINAL HERO READABILITY + ELITE POLISH ===== */
.hero-tagline-line {
    color: #e5f4ff !important;
    font-weight: 850 !important;
    text-shadow: 0 2px 12px rgba(0,0,0,0.38) !important;
}

.hero-wrap {
    background: linear-gradient(135deg, rgba(15,23,42,0.92), rgba(6,78,59,0.74)) !important;
    border: 1px solid rgba(34,197,94,0.30) !important;
}

.hero-brand-line {
    color: #7ee787 !important;
    text-shadow: 0 2px 14px rgba(0,0,0,0.35) !important;
}

.status-chip {
    background: rgba(255,255,255,0.11) !important;
    border-color: rgba(255,255,255,0.14) !important;
}

/* Keep the small FREE/PREMIUM badge white */
section[data-testid="stSidebar"] .stripe-plan-badge,
section[data-testid="stSidebar"] .stripe-plan-badge.free,
section[data-testid="stSidebar"] .stripe-plan-badge.premium,
section[data-testid="stSidebar"] .stripe-plan-badge * {
    color: #ffffff !important;
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
                    <div class="hero-welcome">🚀 {html.escape(t("turn_messy_data"))}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_feature_pills() -> None:
    st.markdown(
        f"""
        <div class="feature-pill-wrap">
            <div class="feature-pill">⚡ {html.escape(t("fast_insights"))}</div>
            <div class="feature-pill">📊 {html.escape(t("ai_powered"))}</div>
            <div class="feature-pill">🌍 {html.escape(t("global_access"))}</div>
            <div class="feature-pill">📄 {html.escape(t("pdf_export"))}</div>
            <div class="feature-pill">🧠 {html.escape(t("business_ready"))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_trust_bar() -> None:
    st.markdown(
        f"""
        <div class="trust-bar">
            🔐 {html.escape(t("secure_file_handling"))} • {html.escape(t("consent_aware_ai"))} • {html.escape(t("privacy_first_design"))}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_status_strip() -> None:
    st.markdown(
        f"""
        <div class="status-chip-wrap">
            <div class="status-chip">🔐 {html.escape(t("secure_file_handling"))}</div>
            <div class="status-chip">⚡ {html.escape(t("fast_processing"))}</div>
            <div class="status-chip">📊 {html.escape(t("ai_powered_insights"))}</div>
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
    st.markdown(f"### {t('privacy_policy')}")
    st.markdown(
        """
- Uploaded files are processed to create analysis, charts, OCR results, reports, translations, and user-requested features.
- If you sign in, files and reports may be stored in your user workspace.
- If AI is enabled, selected content may be sent to your configured AI provider.
- Secure deployment still depends on HTTPS, secret management, and proper hosting.
        """
    )


def render_terms_of_service() -> None:
    st.markdown(f"### {t('terms_of_service')}")
    st.markdown(
        """
- AI outputs may be incomplete or inaccurate and should be reviewed before business use.
- Users must not upload unlawful content or data they do not have permission to process.
- Premium functionality depends on payment and deployment configuration.
        """
    )


def render_data_handling_notice() -> None:
    st.markdown(f"### {t('how_data_handled')}")
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

    email = email.strip().lower()
    supabase_client = get_supabase_client()

    if supabase_client is None:
        return False, "Login is not configured. Please check Supabase settings."

    try:
        response = supabase_client.auth.sign_in_with_password(
            {"email": email, "password": password}
        )
        user = getattr(response, "user", None)
        session = getattr(response, "session", None)

        if user is None:
            return False, "Login failed. Please check your email and password."

        st.session_state["authenticated"] = True
        st.session_state["auth_user"] = {
            "email": email,
            "id": getattr(user, "id", None),
        }
        st.session_state["auth_session"] = session
        st.session_state["auth_mode"] = "supabase"
        st.session_state["user_email"] = email
        st.session_state["sidebar_auth_email"] = ""
        st.session_state["sidebar_auth_password"] = ""

        ensure_user_profile(email)
        load_user_profile(email)

        return True, "Logged in successfully."

    except Exception:
        st.session_state["sidebar_auth_password"] = ""
        return False, "Login failed. Please check your email and password."


def local_create_account(email: str, password: str) -> Tuple[bool, str]:
    if not email or not password:
        return False, "Please enter your email and password."

    email = email.strip().lower()
    supabase_client = get_supabase_client()

    if supabase_client is None:
        return False, "Account creation is not configured. Please check Supabase settings."

    try:
        response = supabase_client.auth.sign_up(
            {"email": email, "password": password}
        )
        user = getattr(response, "user", None)
        session = getattr(response, "session", None)

        if user is None:
            st.session_state["sidebar_auth_password"] = ""
            return False, "Account creation failed. Please try again."

        st.session_state["authenticated"] = True
        st.session_state["auth_user"] = {
            "email": email,
            "id": getattr(user, "id", None),
        }
        st.session_state["auth_session"] = session
        st.session_state["auth_mode"] = "supabase"
        st.session_state["user_email"] = email
        st.session_state["sidebar_auth_email"] = ""
        st.session_state["sidebar_auth_password"] = ""

        ensure_user_profile(email)
        load_user_profile(email)

        return True, "Account created successfully."

    except Exception:
        st.session_state["sidebar_auth_password"] = ""
        return False, "Account creation failed. This email may already exist or the password may be too weak."


def local_create_account(email: str, password: str) -> Tuple[bool, str]:
    if not email or not password:
        return False, "Please enter your email and password."

    email = email.strip().lower()
    supabase_client = get_supabase_client()

    if supabase_client is None:
        return False, "Account creation is not configured. Please check Supabase settings."

    try:
        response = supabase_client.auth.sign_up(
            {"email": email, "password": password}
        )
        user = getattr(response, "user", None)
        session = getattr(response, "session", None)

        if user is None:
            st.session_state["sidebar_auth_password"] = ""
            return False, "Account creation failed. Please try again."

        st.session_state["authenticated"] = True
        st.session_state["auth_user"] = {
            "email": email,
            "id": getattr(user, "id", None),
        }
        st.session_state["auth_session"] = session
        st.session_state["auth_mode"] = "supabase"
        st.session_state["user_email"] = email
        st.session_state["sidebar_auth_email"] = ""
        st.session_state["sidebar_auth_password"] = ""

        ensure_user_profile(email)
        load_user_profile(email)

        return True, "Account created successfully."

    except Exception:
        st.session_state["sidebar_auth_password"] = ""
        return False, "Account creation failed. This email may already exist or the password may be too weak."

def send_password_reset(email: str) -> Tuple[bool, str]:
    """Send a Supabase password reset email when Supabase auth is configured."""
    email = (email or "").strip().lower()
    if not email:
        return False, t("please_log_in_first")

    supabase_client = get_supabase_client()
    if supabase_client is None:
        return False, t("reset_password_unavailable")

    try:
        redirect_to = APP_BASE_URL or None
        if redirect_to:
            try:
                supabase_client.auth.reset_password_email(
                    email,
                    options={"redirect_to": redirect_to},
                )
            except TypeError:
                supabase_client.auth.reset_password_email(email)
        else:
            supabase_client.auth.reset_password_email(email)
        return True, t("reset_password_sent")
    except Exception as exc:
        return False, f"Password reset failed: {exc}"


def ensure_user_profile(email: str) -> None:
    """Create a Supabase profile row for the signed-in user if it does not already exist."""
    email = (email or st.session_state.get("user_email", "")).strip().lower()
    if not email:
        return

    st.session_state["user_email"] = email

    supabase_client = get_supabase_client()
    if supabase_client is None:
        return

    try:
        existing = (
            supabase_client.table("user_profiles")
            .select("email")
            .eq("email", email)
            .execute()
        )

        if not existing.data:
            supabase_client.table("user_profiles").insert(
                {
                    "email": email,
                    "is_premium": False,
                }
            ).execute()
    except Exception as exc:
        print("Error ensuring user profile:", exc)


def activate_premium_for_user(email: str) -> bool:
    """Turn premium on for a user in Supabase and in the current Streamlit session."""
    email = (email or st.session_state.get("user_email", "") or current_user_email()).strip().lower()

    if not email:
        return False

    st.session_state["user_email"] = email
    st.session_state["is_premium"] = True
    set_user_plan(PREMIUM_PLAN)

    supabase_client = get_supabase_client()
    if supabase_client is None:
        return False

    try:
        ensure_user_profile(email)
        supabase_client.table("user_profiles") \
            .update({"is_premium": True}) \
            .eq("email", email) \
            .execute()
        return True
    except Exception as exc:
        print("Error activating premium user:", exc)
        return False


def load_user_profile(email: str = "") -> None:
    """Load the premium flag from Supabase into session_state."""
    email = (email or st.session_state.get("user_email", "") or current_user_email()).strip().lower()

    if not email:
        st.session_state["is_premium"] = False
        set_user_plan(FREE_PLAN)
        return

    st.session_state["user_email"] = email
    st.session_state["is_premium"] = False

    # Local override list remains useful for owner/testing access.
    if email in PREMIUM_USERS:
        st.session_state["is_premium"] = True
        set_user_plan(PREMIUM_PLAN)
        return

    supabase_client = get_supabase_client()
    if supabase_client is None:
        set_user_plan(FREE_PLAN)
        return

    try:
        profile = (
            supabase_client.table("user_profiles")
            .select("is_premium")
            .eq("email", email)
            .single()
            .execute()
        )

        if profile.data:
            st.session_state["is_premium"] = profile.data.get("is_premium", False)
        else:
            st.session_state["is_premium"] = False

        set_user_plan(PREMIUM_PLAN if st.session_state.get("is_premium", False) else FREE_PLAN)
    except Exception as exc:
        # If the profile row is missing, create it and keep user as free.
        print("Error loading user profile:", exc)
        st.session_state["is_premium"] = False
        set_user_plan(FREE_PLAN)
        ensure_user_profile(email)

def logout() -> None:
    supabase_client = get_supabase_client()
    if supabase_client is not None:
        try:
            supabase_client.auth.sign_out()
        except Exception:
            pass
    st.session_state["authenticated"] = False
    st.session_state["auth_user"] = None
    st.session_state["auth_session"] = None
    st.session_state["user_email"] = ""
    st.session_state["is_premium"] = False
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
        df = df.loc[:, ~df.columns.duplicated()].copy()
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
def prepare_forecast_dataframe(df, date_col, value_col):
    working = df.copy()

    # Remove duplicate column names safely
    working = working.loc[:, ~working.columns.duplicated()].copy()

    if date_col not in working.columns:
        st.error(f"Date column '{date_col}' was not found. Please choose another date column.")
        return pd.DataFrame()

    if value_col not in working.columns:
        st.error(f"Value column '{value_col}' was not found. Please choose another value column.")
        return pd.DataFrame()

    working[date_col] = pd.to_datetime(working[date_col], errors="coerce")
    working[value_col] = pd.to_numeric(working[value_col], errors="coerce")

    working = working.dropna(subset=[date_col, value_col])

    if working.empty:
        return pd.DataFrame()

    grouped = (
        working.groupby(date_col, as_index=False)[value_col]
        .sum()
        .sort_values(date_col)
    )

    return grouped

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

def create_checkout_session(plan_type):
    if plan_type == "monthly":
        url = os.getenv("STRIPE_PAYMENT_LINK_SUBSCRIPTION")

    elif plan_type == "one_time":
        url = os.getenv("STRIPE_PAYMENT_LINK_ONE_TIME")

    else:
        st.error("Invalid plan.")
        return

    if not url:
        st.error("Stripe payment link missing.")
        return

    st.link_button("Continue to Stripe Checkout", url)

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

    st.markdown(f"### 🌍 {t('executive_command_center')}")
    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Data Quality Score", f"{score}/100")
    q2.metric("Numeric Fields", len(numeric_cols))
    q3.metric("Text Fields", len(text_cols))
    q4.metric("Date/Time Fields", len(date_like_cols))

    with st.expander("Decision Readiness Notes", expanded=True):
        for note in notes:
            st.write(f"• {note}")
        if score >= 85:
            st.success(t("dataset_ready"))
        elif score >= 65:
            st.warning(t("dataset_cleaning_recommended"))
        else:
            st.error(t("dataset_needs_cleaning"))

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


supabase = get_supabase_client()

# Compatibility wrappers kept here so the later section does not erase the premium logic above.
def get_user_plan():
    email = current_user_email() or st.session_state.get("user_email", "")
    email = str(email).strip().lower()

    if email:
        st.session_state["user_email"] = email

    if email and email in PREMIUM_USERS:
        st.session_state["is_premium"] = True
        return PREMIUM_PLAN

    return PREMIUM_PLAN if st.session_state.get("is_premium", False) else FREE_PLAN


def set_user_plan(plan):
    st.session_state["user_plan"] = plan
    st.session_state["premium_status"] = "premium" if plan == PREMIUM_PLAN else "inactive"
    st.session_state["is_premium"] = plan == PREMIUM_PLAN


def is_premium():
    return st.session_state.get("is_premium", False)



# =========================================================
# MAIN
# =========================================================
def main() -> None:
    init_state()
    inject_global_css()
    process_stripe_return()

    email = current_user_email() or st.session_state.get("user_email", "")
    if email:
        st.session_state["user_email"] = str(email).strip().lower()
        ensure_user_profile(st.session_state["user_email"])
        load_user_profile(st.session_state["user_email"])

    if not st.session_state.get("app_language"):
        st.session_state["app_language"] = detect_user_language()

    render_top_bar()
    render_hero()
    render_feature_pills()
    render_trust_bar()
    render_status_strip()
    st.markdown('<hr class="green-gradient-divider">', unsafe_allow_html=True)

    with st.sidebar:
# -------------------
# SETTINGS
# -------------------
        st.markdown('<div class="sidebar-section sidebar-settings-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="sidebar-title">{html.escape(t("language"))}</div>', unsafe_allow_html=True)

        language_options = ["English"] + sorted(
            [lang for lang in LANGUAGES.keys() if lang != "English"],
            key=lambda x: x.lower(),
        )
        current_lang = st.session_state.get("app_language", "English")
        if current_lang not in language_options:
            current_lang = "English"

        selected_language = st.selectbox(
            t("app_language"),
            language_options,
            index=language_options.index(current_lang),
            key="language_selectbox",
        )
        st.session_state["app_language"] = selected_language

        industry_options = [
            "General Business",
            "Finance",
            "Marketing",
            "Sales",
            "Healthcare",
            "Logistics",
            "Education",
            "Retail",
            "Operations",
        ]
        industry_template = st.selectbox(
            t("industry_template"),
            industry_options,
            key="industry_template",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # -------------------
        # ACCOUNT ACCESS
        # -------------------
        st.markdown('<div class="sidebar-section account-sidebar-box">', unsafe_allow_html=True)
        st.markdown(f'<div class="sidebar-title">{html.escape(t("account_access"))}</div>', unsafe_allow_html=True)

        sidebar_email_display = current_user_email() or st.session_state.get("user_email", "") or "Guest workspace"
        sidebar_initial = (sidebar_email_display[:1] or "G").upper()
        st.markdown(
            f"""
            <div class="sidebar-user-mini-card sidebar-user-mini-card-inside">
                <div class="sidebar-avatar">{html.escape(sidebar_initial)}</div>
                <div class="sidebar-user-details">
                    <div class="sidebar-user-label">{html.escape(t('workspace'))}</div>
                    <div class="sidebar-user-email">{html.escape(sidebar_email_display)}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.session_state.get("authenticated"):
            signed_email = current_user_email() or st.session_state.get("user_email", "")
            st.markdown(
                f"""
                <div class="account-signed-in-box">
                    <div class="account-status-label">{html.escape(t("signed_in"))}</div>
                    <div class="account-email">{html.escape(signed_email or t("not_available"))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(t("logout"), key="sidebar_logout_btn", use_container_width=True):
                logout()
                st.rerun()
        else:
            st.markdown(
                f"""
                <div class="login-card">
                    <div class="login-title">{html.escape(t('welcome_back'))}</div>
                    <div class="login-subtitle">{html.escape(t('login_subtitle'))}</div>
                    <div class="secure-pill">🔐 {html.escape(t('secure_workspace'))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            sidebar_email = st.text_input(
                t("email"),
                placeholder=t("email_placeholder"),
                key="sidebar_auth_email",
            )
            sidebar_password = st.text_input(
                t("password"),
                placeholder=t("password_placeholder"),
                type="password",
                key="sidebar_auth_password",
            )
            login_col, create_col = st.columns(2)
            with login_col:
                if st.button(t("login"), key="sidebar_login_btn", use_container_width=True):
                    ok, msg = local_login(sidebar_email, sidebar_password)
                    if ok:
                        render_message(msg, "success")
                        st.rerun()
                    else:
                        render_message(msg, "error")
            with create_col:
                if st.button(t("create"), key="sidebar_create_btn", use_container_width=True):
                    ok, msg = local_create_account(sidebar_email, sidebar_password)
                    if ok:
                        render_message(msg, "success")
                        st.rerun()
                    else:
                        render_message(msg, "error")
            if st.button(t("forgot_password"), key="sidebar_forgot_password_btn", use_container_width=True):
                ok, msg = send_password_reset(sidebar_email)
                if ok:
                    render_message(msg, "success")
                else:
                    render_message(msg, "warning")
        st.markdown('</div>', unsafe_allow_html=True)

        # -------------------
        # PREMIUM
        # -------------------
        st.markdown('<div class="sidebar-section premium-sidebar-box">', unsafe_allow_html=True)
        plan = t("premium") if is_premium_user() else t("free_plan")
        plan_badge = "PREMIUM" if is_premium_user() else "FREE"
        st.markdown(
            f"""
            <div class="plan-card-inner">
                <div class="plan-row">
                    <span class="plan-title">💎 {html.escape(t("current_plan"))}</span>
                    <span class="stripe-plan-badge {'premium' if is_premium_user() else 'free'}">{html.escape(plan_badge)}</span>
                </div>
                <div class="sidebar-plan-pill-row">
                    <div class="plan-name">{html.escape(plan)}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div class='upgrade-subtitle'>{html.escape(t('upgrade_subtitle'))}</div>",
            unsafe_allow_html=True,
        )
        render_upgrade_checkout_button()
        st.markdown('</div>', unsafe_allow_html=True)

        if st.button(t("use_sample_dataset"), key="sample_btn", use_container_width=True):
            st.session_state["use_sample_data"] = True

        st.checkbox(t("data_handling_terms"), key="gdpr_consent")

        used_count = int(st.session_state.get("usage_count", 0) or 0)
        used_capped = min(used_count, FREE_ANALYSIS_LIMIT)
        remaining = max(0, FREE_ANALYSIS_LIMIT - used_count)
        usage_pct = int((used_capped / max(FREE_ANALYSIS_LIMIT, 1)) * 100)
        st.markdown(
            f"""
            <div class="usage-compact-card">
                <div class="usage-compact-top">
                    <span class="usage-compact-label">{html.escape(t('free_analyses_remaining'))}</span>
                    <span class="usage-compact-count">{used_capped}/{FREE_ANALYSIS_LIMIT}</span>
                </div>
                <div class="usage-progress-track">
                    <div class="usage-progress-fill" style="width:{usage_pct}%;"></div>
                </div>
                <div class="usage-compact-note">{html.escape(str(remaining))} {html.escape(t('free_analyses_remaining')).lower()}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.expander(f"📘 {t('how_it_works')}", expanded=False):
            st.markdown(
                """
                <div style="color:#111827; font-size:0.84rem; line-height:1.55; font-weight:700;">
                    1. Upload your dataset or document.<br>
                    2. Let AI analyze the content.<br>
                    3. Review insights, risks, forecasts, and reports.<br>
                    4. Save, share, or upgrade features.<br>
                    5. Upgrade to unlock advanced tools.
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.expander(f"ℹ️ {t('about_us')}", expanded=False):
            st.markdown(
                """
                <div style="color:#111827; font-size:0.82rem; line-height:1.55; font-weight:700;">
                    ExplainMyData AI helps users turn raw files into clear business insights,
                    forecasts, data quality checks, and executive-style reports.
                </div>
                """,
                unsafe_allow_html=True,
            )

    
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
        if st.button(f"⚙️ {t('control_panel')}", key="control_btn"):
            st.session_state["show_controls"] = not st.session_state.get("show_controls", False)

    if st.session_state.get("show_controls"):
        st.markdown(f"### ⚙️ {t('control_panel_title')}")
        st.write(f"{t('plan')}:", get_user_plan())
        st.write(f"{t('ai_configured')}:", t("yes") if ai_available() else t("no"))
        st.write(f"{t('auth_configured')}:", t("yes") if auth_configured() else t("no"))
        st.slider(t("usage_count_label"), 0, 10, key="usage_count")
        st.write("SUPABASE URL:", os.getenv("SUPABASE_URL"))

    tabs = st.tabs([
        f"📊 {t('dashboard')}", f"🤖 {t('ai_dashboard')}", f"🧠 {t('data_doctor')}", f"🧭 {t('decision_engine')}", f"🧪 {t('scenario_simulator')}", f"💬 {t('ask_data')}",
        f"📋 {t('report')}", f"📈 {t('forecast')}", f"📂 {t('saved')}", f"🖼️ {t('dashboard_gallery')}",
        f"👤 {t('account')}", f"🖼️ {t('ocr')}", f"🔄 {t('compare')}",
        f"🧹 {t('cleaning')}", f"⚖️ {t('legal')}", 
    ])

    with tabs[0]:
        if df is None:
            render_message(t("upload_start"), "info")
        else:
            st.subheader(t("data_overview"))
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(t("rows"), f"{len(df):,}")
            c2.metric(t("columns"), f"{len(df.columns):,}")
            c3.metric(t("missing"), f"{int(df.isna().sum().sum()):,}")
            c4.metric(t("duplicates"), f"{int(df.duplicated().sum()):,}")
            render_semantic_metric_cards(compute_semantic_metrics(df))
            render_executive_command_center(df)
            
            st.markdown(f"### {t('preview')}")
            
            st.markdown(f"### 📊 {t('industry_kpis')}")
            
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
                
            st.markdown(f"### 📊 {t('industry_kpis')}")
            
            if df is not None:
                kpis = get_industry_kpis(df)

                cols = st.columns(len(kpis))

                for i, (label, value) in enumerate(kpis):
                    try:
                        cols[i].metric(label, f"{value:,.2f}")
                    except:
                        cols[i].metric(label, str(value))
            
            st.markdown(f"### 📈 {t('industry_trend')}")

            numeric_cols = df.select_dtypes(include="number").columns.tolist()

            if numeric_cols:
                col_choice = numeric_cols[0]

                fig, ax = plt.subplots()
                ax.plot(df[col_choice].fillna(0))
                ax.set_title(f"{col_choice} Trend")

                st.pyplot(fig)
            
            st.markdown(f"### 📈 {t('executive_dashboard')}")

            if df is not None:
                if st.button(f"📈 {t('generate_dashboard')}", key="exec_dashboard_btn"):
                    metrics, summary, insights = generate_executive_dashboard(df)

                    col1, col2, col3, col4 = st.columns(4)

                    col1.metric("Rows", summary["rows"])
                    col2.metric("Columns", summary["columns"])
                    col3.metric("Missing", summary["missing"])
                    col4.metric("Duplicates", summary["duplicates"])

                    st.markdown(f"#### 📊 {t('key_metrics')}")
                    render_semantic_metric_cards(metrics)

                    st.markdown(f"#### 🧠 {t('insights')}")
                    for item in insights:
                        st.write(f"• {item}")
            
            st.dataframe(df.head(50), use_container_width=True)
            numeric_cols = df.select_dtypes(include="number").columns.tolist()
            if numeric_cols:
                chart_col = st.selectbox(t("quick_chart_metric"), numeric_cols, key="dashboard_chart_col")
                fig, ax = plt.subplots()
                pd.to_numeric(df[chart_col], errors="coerce").dropna().plot(kind="hist", ax=ax)
                ax.set_title(f"Distribution of {chart_col}")
                ax.set_xlabel(chart_col)
                st.pyplot(fig)

    with tabs[1]:
        st.markdown(f"### 🤖 {t('ai_auto_dashboard')}")

        if df is None:
            render_message(t("upload_dataset_first"), "info")
        else:
            if st.button(f"🤖 {t('build_my_dashboard')}", key="auto_dashboard_btn"):

                result = generate_auto_dashboard(df)

                if not result:
                    render_message("Not enough numeric data for dashboard.", "warning")
                else:
                    st.markdown(f"#### 📊 {t('ai_selected_metrics')}")

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Total", f"{result['total']:,.2f}")
                    c2.metric("Average", f"{result['mean']:,.2f}")
                    c3.metric("Max", f"{result['max']:,.2f}")
                    c4.metric("Min", f"{result['min']:,.2f}")

                    st.markdown(f"#### 📈 {t('ai_trend')}")

                    fig, ax = plt.subplots()
                    ax.plot(df[result["column"]].fillna(0))
                    ax.set_title(f"{result['column']} Trend (AI Selected)")
                    st.pyplot(fig)
    
    with tabs[2]:
        st.subheader(t("AI Data Doctor"))
        if df is None:
            render_message("Upload data first so the Data Doctor can inspect it.", "info")
        else:
            if st.button(t("run_ai_data_doctor"), key="doctor_btn"):
                st.session_state["doctor_result"] = build_ai_data_doctor_report(df, mode, raw_text, current_language())
            if st.session_state.get("doctor_result"):
                render_report_card("Data Doctor Report", st.session_state["doctor_result"], "#facc15")

        st.markdown(f"### 🩺 {t('auto_fix_data_doctor')}")

        if df is None:
            render_message(t("upload_dataset_first"), "info")
        else:
            if st.button(f"🛠 {t('auto_fix_dataset')}", key="auto_fix_btn"):
                cleaned_df, actions = auto_fix_dataset(df)
                st.session_state["cleaned_df"] = cleaned_df

                render_message("Dataset cleaned successfully.", "success")

                st.markdown(f"#### {t('fixes_applied')}")
                for action in actions:
                    st.write(f"• {action}")

                st.markdown(f"#### {t('cleaned_dataset_preview')}")
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
        st.markdown(f"### 🧭 {t('ai_decision_engine')}")

        if df is None:
            render_message(t("upload_dataset_first"), "info")
        else:
            st.write(
                "This feature turns your dataset into executive decisions, risks, opportunities, and recommended actions."
            )

            if st.button(f"🧭 {t('generate_decision_report')}", key="decision_engine_btn"):
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
                        render_report_card(t("ai_decision_engine_report"), decision_report, "#22c55e")
                    except Exception as exc:
                        render_message(f"{t('decision_engine_error')}: {exc}", "error")

                if st.session_state.get("decision_report"):
                    st.download_button(
                        f"⬇️ {t('download_decision_report')}",
                        st.session_state["decision_report"],
                        file_name="decision_engine_report.txt",
                        mime="text/plain",
                        key="download_decision_report_btn",
                    )
    
    with tabs[4]:
        st.markdown(f"### 🧪 {t('scenario_simulator_title')}")

        if df is None:
            render_message(t("upload_dataset_first"), "info")
        else:
            numeric_cols = df.select_dtypes(include="number").columns.tolist()

            if not numeric_cols:
                render_message(t("no_numeric_simulation"), "warning")
            else:
                selected_metric = st.selectbox(
                    t("choose_numeric_column"),
                    numeric_cols,
                    key="scenario_metric_select",
                )

                base_value = float(pd.to_numeric(df[selected_metric], errors="coerce").sum())

                change_percent = st.slider(
                    t("what_if_change"),
                    -100.0,
                    100.0,
                    10.0,
                    step=1.0,
                    key="scenario_change_slider",
                )

                result = run_scenario_simulator(base_value, change_percent)

                c1, c2, c3, c4 = st.columns(4)
                c1.metric(t("current_total"), f"{result['base_value']:,.2f}")
                c2.metric(t("change"), f"{result['change_percent']}%")
                c3.metric(t("projected_total"), f"{result['adjusted_value']:,.2f}")
                c4.metric(t("difference"), f"{result['difference']:,.2f}")

                st.markdown(f"#### {t('business_meaning')}")
                if result["difference"] > 0:
                    st.success(t("positive_projected_movement"))
                elif result["difference"] < 0:
                    st.warning(t("negative_projected_movement"))
                else:
                    st.info(t("no_projected_movement"))
    
    with tabs[5]:
        st.subheader(t("ask_data"))
        if df is None:
            render_message(t("upload_data_ask"), "info")
        else:
            question = st.text_area(t("ask_question_about_file"), key="question_box")
            if st.button(t("ask"), key="ask_btn"):
                if not question.strip():
                    render_message(t("please_type_question"), "warning")
                else:
                    try:
                        st.session_state["chat_answer"] = answer_followup_question(df, mode, raw_text, question, current_language())
                    except Exception as exc:
                        st.session_state["chat_answer"] = f"{t('could_not_generate_answer')}: {exc}"
            if st.session_state.get("chat_answer"):
                render_report_card(t("answer"), st.session_state["chat_answer"], "#38bdf8")

    with tabs[6]:
        st.subheader(t("consulting_report"))
        if df is None:
            render_message(t("upload_data_consulting"), "info")
        else:
            focus_options = [t("executive_summary"), t("sales_growth"), t("operations"), t("risk"), t("customer_behavior"), t("data_quality")]
            focus = st.selectbox(t("analysis_focus"), focus_options, key="analysis_focus_select")
            if st.button(t("generate_ai_insights"), key="insights_btn"):
                allowed, reason = analysis_allowed()
                if not st.session_state.get("gdpr_consent"):
                    render_message(t("agree_data_terms_ai"), "warning")
                elif not allowed:
                    render_message(reason, "warning")
                else:
                    try:
                        st.session_state["result"] = generate_ai_analysis(df, mode, raw_text, focus, current_language())
                        increment_usage()
                    except Exception as exc:
                        st.session_state["result"] = f"{t('could_not_generate_ai_report')}: {exc}"
            if st.session_state.get("result"):
                render_report_card(t("executive_consulting_report"), st.session_state["result"], "#22c55e")
                pdf_path = generate_pdf_report(st.session_state["result"])
                with open(pdf_path, "rb") as pdf_file:
                    st.download_button(t("download_executive_report"), data=pdf_file.read(), file_name="executive_report.pdf", mime="application/pdf")
                if st.button(t("save_report"), key="save_report_btn"):
                    if not st.session_state.get("authenticated"):
                        render_message(t("please_sign_in_save"), "warning")
                    else:
                        payload = report_payload(loaded_file_name, st.session_state.get("result", ""), st.session_state.get("translated_result", ""), st.session_state.get("chat_answer", ""), st.session_state.get("doctor_result", ""))
                        saved_path = save_report_for_user(current_user_email(), payload)
                        render_message(t("report_saved_successfully") if saved_path else t("report_could_not_saved"), "success" if saved_path else "error")
                if st.button(t("create_shareable_report_link"), key="share_report_btn"):
                    if require_feature("shareable_reports", t("shareable_reports")):
                        link = create_shareable_report_link(current_user_email(), loaded_file_name, st.session_state.get("result", ""), st.session_state.get("translated_result", ""), st.session_state.get("doctor_result", ""))
                        st.code(link)
    
        st.markdown(f"### 🧾 {t('boardroom_pdf_report')}")

        if df is None:
            render_message(t("upload_dataset_first"), "info")
        else:
            if st.button(f"🧾 {t('generate_boardroom_pdf')}", key="boardroom_pdf_btn"):
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
                            f"⬇️ {t('download_boardroom_pdf')}",
                            f,
                            file_name="boardroom_report.pdf",
                            mime="application/pdf",
                            key="download_boardroom_pdf_btn",
                        )
        
        st.markdown(f"### 🔗 {t('shareable_report_title')}")

        if df is None:
            render_message(t("upload_dataset_first"), "info")
        else:
            if st.button(f"🔗 {t('create_share_link')}", key="create_share_link_btn"):
                try:
                    analysis_text = st.session_state.get("result", "")
                    translation_text = st.session_state.get("translated_result", "")
                    doctor_text = st.session_state.get("doctor_result", "")

                    if not analysis_text and not doctor_text:
                        render_message(
                            t("generate_first_before_sharing"),
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
                        render_message(t("shareable_report_created"), "success")

                except Exception as exc:
                    render_message(f"{t('share_link_error')}: {exc}", "error")

            if st.session_state.get("share_link"):
                st.markdown(f"#### {t('your_shareable_link')}")
                st.code(st.session_state["share_link"])

    with tabs[7]:
        st.subheader(t("forecast"))
        if df is None:
            render_message(t("no_data_forecast"), "info")
        else:
            date_candidates = [c for c in df.columns if "date" in str(c).lower() or "time" in str(c).lower()]
            numeric_cols = df.select_dtypes(include="number").columns.tolist()
            if not date_candidates or not numeric_cols:
                render_message(t("forecast_needs_date_numeric"), "warning")
            else:
                date_col = st.selectbox(t("date_column"), date_candidates, key="forecast_date_col")
                value_col = st.selectbox(t("value_column"), numeric_cols, key="forecast_value_col")
                periods = st.slider(t("forecast_periods"), 3, 24, 6, key="forecast_periods")
                if st.button(t("build_forecast"), key="forecast_btn"):
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
                    render_report_card(t("forecast_interpretation"), generate_forecast_interpretation(forecast_df, value_col, current_language()), "#a78bfa")
                    render_report_card(t("recommended_action"), generate_forecast_recommendation(forecast_df, value_col, current_language()), "#22c55e")

    with tabs[8]:
        st.subheader(t("saved"))
        if not st.session_state.get("authenticated"):
            render_message(t("sign_in_saved"), "info")
        else:
            reports = load_user_reports(current_user_email())
            if not reports:
                render_message(t("no_saved_reports"), "info")
            for item in reports:
                with st.expander(f"{item.get('file_name', 'Report')} — {item.get('created_at', '')}"):
                    if item.get("analysis"):
                        render_report_card(t("analysis_focus"), item["analysis"], "#22c55e")
                    if item.get("doctor"):
                        render_report_card(t("data_doctor"), item["doctor"], "#facc15")
                    if item.get("answer"):
                        render_report_card(t("answer"), item["answer"], "#38bdf8")

    with tabs[9]:
        st.markdown(f"### 🖼️ {t('saved_dashboard_gallery')}")

        if df is not None:
            if st.button(f"💾 {t('save_current_dashboard_snapshot')}", key="save_dashboard_snapshot_btn"):
                path = save_dashboard_snapshot(df, loaded_file_name)
                if path:
                    render_message(t("dashboard_snapshot_saved"), "success")
                else:
                    render_message(t("dashboard_snapshot_error"), "error")

        dashboards = load_dashboard_snapshots()

        if not dashboards:
            render_message(t("no_dashboard_snapshots"), "info")
        else:
            for item in dashboards:
                with st.expander(f"{item.get('file_name', 'Dashboard')} — {item.get('created_at', '')}"):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric(t("rows"), item.get("rows", 0))
                    c2.metric(t("columns"), item.get("columns", 0))
                    c3.metric(t("missing"), item.get("missing_values", 0))
                    c4.metric(t("quality"), f"{item.get('data_quality_score', 0)}/100")

                    st.write(f"{t('industry_template')}:", item.get("industry_template", "General Business"))
        
        st.write(f"{t('supabase_connected')}:", bool(SUPABASE_URL and SUPABASE_ANON_KEY))
    
    with tabs[10]:
        st.subheader(t("account"))
        st.write(f"{t('authentication')}:", t("signed_in") if st.session_state.get("authenticated") else t("not_signed_in"))
        st.write(f"{t('email')}:", current_user_email() or t("not_available"))
        st.write(f"{t('plan')}:", get_user_plan())
        st.write(f"{t('usage_count')}:", st.session_state.get("usage_count", 0))
        if st.session_state.get("authenticated"):
            if st.button(t("delete_my_local_saved_data"), key="delete_account_data_btn"):
                safe_remove_tree(user_upload_dir(current_user_email()))
                safe_remove_tree(user_report_dir(current_user_email()))
                render_message(t("local_saved_deleted"), "success")

    with tabs[11]:
        st.subheader(t("ocr"))
        if st.session_state.get("ocr_preview") is not None:
            st.image(st.session_state["ocr_preview"], caption=t("ocr_image_preview"), use_container_width=True)
        if st.session_state.get("ocr_text"):
            st.text_area(t("extracted_ocr_text"), st.session_state["ocr_text"], height=260)
        else:
            render_message(t("upload_image_ocr"), "info")

    with tabs[12]:
        st.subheader(t("compare"))
        file_a = st.file_uploader(t("upload_file_a"), type=["csv", "xlsx", "xls"], key="compare_a")
        file_b = st.file_uploader(t("upload_file_b"), type=["csv", "xlsx", "xls"], key="compare_b")
        if file_a is not None and file_b is not None:
            try:
                df_a = load_structured_data(file_a)
                df_b = load_structured_data(file_b)
                render_report_card(t("comparison"), compare_dataframes(df_a, df_b), "#38bdf8")
            except Exception as exc:
                render_message(f"{t('comparison_failed')}: {exc}", "error")

    with tabs[13]:
        st.subheader(t("cleaning"))
        if df is None:
            render_message(t("no_data_clean"), "info")
        else:
            remove_dup = st.checkbox(t("remove_duplicate_rows"), value=True, key="clean_dup")
            fill_num = st.checkbox(t("fill_numeric_missing"), value=True, key="clean_num")
            fill_text = st.checkbox(t("fill_text_missing"), value=True, key="clean_text")
            trim_spaces = st.checkbox(t("trim_text_spaces"), value=True, key="clean_trim")
            standardize_dates_flag = st.checkbox(t("standardize_date_columns"), value=False, key="clean_dates")
            if st.button(t("apply_cleaning"), key="apply_cleaning_btn"):
                cleaned = apply_cleaning_actions(df, remove_dup, fill_num, fill_text, trim_spaces, standardize_dates_flag)
                st.session_state["cleaned_df"] = cleaned
                st.session_state["cleaning_summary"] = generate_cleaning_summary(df, cleaned, current_language())
            cleaned_df = st.session_state.get("cleaned_df")
            if isinstance(cleaned_df, pd.DataFrame):
                st.dataframe(cleaned_df.head(50), use_container_width=True)
                if st.session_state.get("cleaning_summary"):
                    render_report_card(t("cleaning_summary"), st.session_state["cleaning_summary"], "#22c55e")
                st.download_button(t("download_cleaned_csv"), cleaned_df.to_csv(index=False).encode("utf-8"), file_name="cleaned_data.csv", mime="text/csv")

    with tabs[14]:
        render_privacy_policy()
        st.markdown("---")
        render_terms_of_service()
        st.markdown("---")
        render_data_handling_notice()

    st.markdown('<hr class="green-gradient-divider">', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="footer-note">
            {t("footer_slogan")}<br>
            <span style="font-size:0.65rem;color:#94a3b8;">© 2026 ExplainMyData AI. {t("rights_reserved")}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
def render_footer():
    current_year = datetime.now().year
    st.markdown(
        f"""
        <div style="
            width:100%;
            text-align:center;
            font-size:0.66rem;
            color:#94a3b8;
            margin-top:24px;
            margin-bottom:2px;
            padding-bottom:2px;
            opacity:0.85;
        ">
            © {current_year} ExplainMyData AI. {t("rights_reserved")}
        </div>
        """,
        unsafe_allow_html=True,
    )

if __name__ == "__main__":
    main()
