import io
import os

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="ExplainMyData AI",
    page_icon="edai_logo.png",
    layout="wide"
)

# =========================================================
# UI TRANSLATIONS
# =========================================================
TRANSLATIONS = {
    "English": {
        "language_selector": "App language",
        "app_name": "ExplainMyData AI",
        "tagline": "Turn raw files into clear business insights.",
        "features": "What you can do:",
        "feature_1": "- Upload CSV / Excel / TXT / Word",
        "feature_2": "- Explore KPIs and charts",
        "feature_3": "- Generate AI insights",
        "feature_4": "- Ask follow-up questions",
        "feature_5": "- Export reports",
        "ai_ready": "AI Ready Check",
        "ai_ready_ok": "The AI assistant is ready. You can upload a file, analyze it, ask questions, and export results.",
        "ai_ready_fail": "AI service is not responding right now. You can still upload a file and try again in a moment.",
        "header_title": "ExplainMyData AI",
        "header_motto": "Your data, clearly explained.",
        "badge_1": "AI-powered",
        "badge_2": "Fast insights",
        "badge_3": "Global access",
        "badge_4": "PDF export",
        "badge_5": "Text & Word support",
        "upload_label": "Upload your file",
        "supported_formats": "⚡ Supported formats: CSV, XLSX, XLS, TXT, DOCX",
        "data_overview": "📊 Data Overview",
        "rows": "Rows",
        "columns": "Columns",
        "missing": "Missing",
        "duplicates": "Duplicates",
        "overview_tab": "Overview",
        "insight_tab": "Insight Board",
        "charts_tab": "Charts",
        "ai_tab": "AI Insights",
        "translate_tab": "Translate",
        "export_tab": "Ask & Export",
        "upload_begin": "Upload a file to begin.",
        "how_it_works": "How it works",
        "step_1": "1. Upload a CSV, Excel, TXT, or DOCX file.",
        "step_2": "2. Review your overview, charts, and insight board.",
        "step_3": "3. Generate AI insights in your selected language.",
        "step_4": "4. Ask follow-up questions to explore patterns, risks, and opportunities.",
        "step_5": "5. Download reports and share findings with teammates or clients.",
        "workspace_loaded": "Analysis workspace loaded successfully.",
        "text_mode_warning": "This file was converted into a text-based table for analysis. Charts may be limited depending on the extracted content.",
        "snapshot": "🔍 Dataset Snapshot",
        "text_preview": "Extracted Text Preview",
        "preview": "Preview",
        "column_summary": "Column Summary",
        "data_quality_notes": "Data Quality Notes",
        "cleaning_assistant": "🧹 Data Cleaning Assistant",
        "show_cleaning": "Show Cleaning Suggestions",
        "missing_cols": "Columns with missing values:",
        "no_missing": "No missing values found.",
        "dup_review": "duplicate rows to review.",
        "no_dup": "No duplicate rows found.",
        "insight_board_title": "📌 Insight Board",
        "numeric_spotlight": "Numeric Feature Spotlight",
        "category_spotlight": "Category Feature Spotlight",
        "choose_numeric": "Choose a numeric column",
        "choose_categorical": "Choose a categorical column",
        "no_numeric": "No numeric columns available.",
        "no_categorical": "No categorical columns available.",
        "quality_snapshot": "Data Quality Snapshot",
        "contains_missing": "This dataset contains missing values.",
        "no_missing_detected": "No missing values detected.",
        "contains_dup": "This dataset contains duplicate rows.",
        "no_dup_detected": "No duplicate rows detected.",
        "correlation_map": "Correlation Map",
        "need_two_numeric": "Need at least two numeric columns for a correlation map.",
        "quick_text_insights": "Quick Text Insights",
        "charts_title": "📈 Interactive Dashboard / Charts",
        "chart_type": "Choose chart type",
        "histogram": "Histogram",
        "box_plot": "Box Plot",
        "bar_chart": "Bar Chart",
        "line_chart": "Line Chart",
        "choose_numeric_chart": "Choose numeric column",
        "choose_cat_filter": "Choose categorical filter",
        "none": "None",
        "select_value": "Select value from",
        "select_cat_for_bar": "Select a categorical column to use the Bar Chart.",
        "choose_date": "Choose date/year column",
        "no_date": "No date/year column found for line chart.",
        "ai_insights_title": "🧠 Smart AI Insights",
        "analysis_focus": "Choose analysis focus",
        "focus_general": "General Analysis",
        "focus_sales": "Sales Analysis",
        "focus_marketing": "Marketing Analysis",
        "focus_customer": "Customer Analysis",
        "focus_operations": "Operations Analysis",
        "focus_financial": "Financial Analysis",
        "focus_document": "Document Summary Analysis",
        "free_used": "Free analyses used",
        "free_limit": "Free limit reached. Upgrade to continue.",
        "analyze_data": "Analyze Data",
        "generating": "Generating insights...",
        "analysis_results": "Analysis Results",
        "how_read_analysis": "🧭 How to Read This Analysis",
        "download_insights": "📥 Download Insights (.txt)",
        "translate_title": "🌍 Translate Results",
        "translation_language": "Choose translation language",
        "generate_first": "Generate AI insights first, then translate them here.",
        "translate_insights": "Translate Insights",
        "no_translate_yet": "No insights to translate yet.",
        "translating": "Translating to",
        "translated_results": "Translated Results",
        "download_translation": "📥 Download Translation",
        "ask_export_title": "💬 Ask & Export",
        "ask_export_text": "Ask follow-up questions in plain language, then export your results.",
        "question_placeholder": "Example: Which trend matters most in this dataset or document?",
        "thinking": "Thinking...",
        "answer": "Answer",
        "export_results": "📥 Export Results",
        "download_csv": "📥 Download Processed Data as CSV",
        "download_pdf": "📄 Download PDF Report",
        "download_translated": "🌍 Download Translated Results",
        "premium_title": "💎 Upgrade to Premium",
        "premium_text": "Unlock unlimited analysis, deeper insights, advanced exports, translations, and premium reports.",
        "premium_button": "💎 Upgrade to Premium",
        "next_steps": "Suggested Next Steps",
        "next_1": "- Review missing values and duplicates.",
        "next_2": "- Explore important columns.",
        "next_3": "- Generate AI insights based on your business goal.",
        "next_4": "- Ask follow-up questions for deeper understanding.",
        "next_5": "- Download and share your report.",
    },
    "French": {
        "language_selector": "Langue de l’application",
        "app_name": "ExplainMyData AI",
        "tagline": "Transformez des fichiers bruts en informations claires.",
        "features": "Ce que vous pouvez faire :",
        "feature_1": "- Téléverser CSV / Excel / TXT / Word",
        "feature_2": "- Explorer KPI et graphiques",
        "feature_3": "- Générer des analyses IA",
        "feature_4": "- Poser des questions de suivi",
        "feature_5": "- Exporter des rapports",
        "ai_ready": "Vérifier l’IA",
        "ai_ready_ok": "L’assistant IA est prêt. Vous pouvez téléverser un fichier, l’analyser, poser des questions et exporter les résultats.",
        "ai_ready_fail": "Le service IA ne répond pas pour le moment. Vous pouvez quand même téléverser un fichier et réessayer dans un instant.",
        "header_title": "ExplainMyData AI",
        "header_motto": "Vos données, clairement expliquées.",
        "badge_1": "Propulsé par l’IA",
        "badge_2": "Analyses rapides",
        "badge_3": "Accès mondial",
        "badge_4": "Export PDF",
        "badge_5": "Support texte et Word",
        "upload_label": "Téléversez votre fichier",
        "supported_formats": "⚡ Formats pris en charge : CSV, XLSX, XLS, TXT, DOCX",
        "data_overview": "📊 Vue d’ensemble des données",
        "rows": "Lignes",
        "columns": "Colonnes",
        "missing": "Manquants",
        "duplicates": "Doublons",
        "overview_tab": "Aperçu",
        "insight_tab": "Tableau d’insights",
        "charts_tab": "Graphiques",
        "ai_tab": "Analyses IA",
        "translate_tab": "Traduire",
        "export_tab": "Questions & Export",
        "upload_begin": "Téléversez un fichier pour commencer.",
        "how_it_works": "Comment ça marche",
        "step_1": "1. Téléversez un fichier CSV, Excel, TXT ou DOCX.",
        "step_2": "2. Consultez l’aperçu, les graphiques et le tableau d’insights.",
        "step_3": "3. Générez des analyses IA dans la langue sélectionnée.",
        "step_4": "4. Posez des questions de suivi pour explorer les tendances, risques et opportunités.",
        "step_5": "5. Téléchargez des rapports et partagez vos conclusions avec collègues ou clients.",
        "workspace_loaded": "Espace d’analyse chargé avec succès.",
        "text_mode_warning": "Ce fichier a été converti en tableau textuel pour l’analyse. Les graphiques peuvent être limités selon le contenu extrait.",
        "snapshot": "🔍 Aperçu du jeu de données",
        "text_preview": "Aperçu du texte extrait",
        "preview": "Aperçu",
        "column_summary": "Résumé des colonnes",
        "data_quality_notes": "Notes sur la qualité des données",
        "cleaning_assistant": "🧹 Assistant de nettoyage des données",
        "show_cleaning": "Afficher les suggestions de nettoyage",
        "missing_cols": "Colonnes avec des valeurs manquantes :",
        "no_missing": "Aucune valeur manquante trouvée.",
        "dup_review": "lignes en double à examiner.",
        "no_dup": "Aucune ligne en double trouvée.",
        "insight_board_title": "📌 Tableau d’insights",
        "numeric_spotlight": "Mise en avant numérique",
        "category_spotlight": "Mise en avant catégorielle",
        "choose_numeric": "Choisissez une colonne numérique",
        "choose_categorical": "Choisissez une colonne catégorielle",
        "no_numeric": "Aucune colonne numérique disponible.",
        "no_categorical": "Aucune colonne catégorielle disponible.",
        "quality_snapshot": "Instantané de qualité des données",
        "contains_missing": "Ce jeu de données contient des valeurs manquantes.",
        "no_missing_detected": "Aucune valeur manquante détectée.",
        "contains_dup": "Ce jeu de données contient des lignes en double.",
        "no_dup_detected": "Aucune ligne en double détectée.",
        "correlation_map": "Carte de corrélation",
        "need_two_numeric": "Au moins deux colonnes numériques sont nécessaires pour une carte de corrélation.",
        "quick_text_insights": "Observations rapides",
        "charts_title": "📈 Tableau de bord / Graphiques",
        "chart_type": "Choisissez le type de graphique",
        "histogram": "Histogramme",
        "box_plot": "Boîte à moustaches",
        "bar_chart": "Graphique en barres",
        "line_chart": "Graphique en ligne",
        "choose_numeric_chart": "Choisissez une colonne numérique",
        "choose_cat_filter": "Choisissez un filtre catégoriel",
        "none": "Aucun",
        "select_value": "Sélectionnez une valeur de",
        "select_cat_for_bar": "Sélectionnez une colonne catégorielle pour utiliser le graphique en barres.",
        "choose_date": "Choisissez une colonne date/année",
        "no_date": "Aucune colonne date/année trouvée pour le graphique en ligne.",
        "ai_insights_title": "🧠 Analyses IA intelligentes",
        "analysis_focus": "Choisissez le type d’analyse",
        "focus_general": "Analyse générale",
        "focus_sales": "Analyse des ventes",
        "focus_marketing": "Analyse marketing",
        "focus_customer": "Analyse client",
        "focus_operations": "Analyse opérationnelle",
        "focus_financial": "Analyse financière",
        "focus_document": "Analyse de résumé de document",
        "free_used": "Analyses gratuites utilisées",
        "free_limit": "Limite gratuite atteinte. Passez à la version Premium pour continuer.",
        "analyze_data": "Analyser les données",
        "generating": "Génération des analyses...",
        "analysis_results": "Résultats de l’analyse",
        "how_read_analysis": "🧭 Comment lire cette analyse",
        "download_insights": "📥 Télécharger les analyses (.txt)",
        "translate_title": "🌍 Traduire les résultats",
        "translation_language": "Choisissez la langue de traduction",
        "generate_first": "Générez d’abord des analyses IA, puis traduisez-les ici.",
        "translate_insights": "Traduire les analyses",
        "no_translate_yet": "Aucune analyse à traduire pour le moment.",
        "translating": "Traduction vers",
        "translated_results": "Résultats traduits",
        "download_translation": "📥 Télécharger la traduction",
        "ask_export_title": "💬 Questions & Export",
        "ask_export_text": "Posez des questions de suivi en langage naturel, puis exportez vos résultats.",
        "question_placeholder": "Exemple : quelle tendance est la plus importante dans ce jeu de données ou document ?",
        "thinking": "Réflexion...",
        "answer": "Réponse",
        "export_results": "📥 Exporter les résultats",
        "download_csv": "📥 Télécharger les données traitées en CSV",
        "download_pdf": "📄 Télécharger le rapport PDF",
        "download_translated": "🌍 Télécharger les résultats traduits",
        "premium_title": "💎 Passer à Premium",
        "premium_text": "Débloquez des analyses illimitées, des insights plus profonds, des exports avancés, des traductions et des rapports premium.",
        "premium_button": "💎 Passer à Premium",
        "next_steps": "Étapes recommandées",
        "next_1": "- Vérifiez les valeurs manquantes et les doublons.",
        "next_2": "- Explorez les colonnes importantes.",
        "next_3": "- Générez des analyses IA selon votre objectif métier.",
        "next_4": "- Posez des questions de suivi pour approfondir la compréhension.",
        "next_5": "- Téléchargez et partagez votre rapport.",
    },
    "Spanish": {
        "language_selector": "Idioma de la aplicación",
        "app_name": "ExplainMyData AI",
        "tagline": "Convierte archivos sin procesar en ideas claras de negocio.",
        "features": "Lo que puedes hacer:",
        "feature_1": "- Subir CSV / Excel / TXT / Word",
        "feature_2": "- Explorar KPI y gráficos",
        "feature_3": "- Generar análisis con IA",
        "feature_4": "- Hacer preguntas de seguimiento",
        "feature_5": "- Exportar informes",
        "ai_ready": "Verificar IA",
        "ai_ready_ok": "El asistente de IA está listo. Puedes subir un archivo, analizarlo, hacer preguntas y exportar resultados.",
        "ai_ready_fail": "El servicio de IA no responde en este momento. Aún puedes subir un archivo e intentarlo de nuevo en un momento.",
        "header_title": "ExplainMyData AI",
        "header_motto": "Tus datos, claramente explicados.",
        "badge_1": "Impulsado por IA",
        "badge_2": "Ideas rápidas",
        "badge_3": "Acceso global",
        "badge_4": "Exportación PDF",
        "badge_5": "Soporte de texto y Word",
        "upload_label": "Sube tu archivo",
        "supported_formats": "⚡ Formatos compatibles: CSV, XLSX, XLS, TXT, DOCX",
        "data_overview": "📊 Resumen de datos",
        "rows": "Filas",
        "columns": "Columnas",
        "missing": "Faltantes",
        "duplicates": "Duplicados",
        "overview_tab": "Resumen",
        "insight_tab": "Panel de insights",
        "charts_tab": "Gráficos",
        "ai_tab": "Insights IA",
        "translate_tab": "Traducir",
        "export_tab": "Preguntar y exportar",
        "upload_begin": "Sube un archivo para comenzar.",
        "how_it_works": "Cómo funciona",
        "step_1": "1. Sube un archivo CSV, Excel, TXT o DOCX.",
        "step_2": "2. Revisa la vista general, los gráficos y el panel de insights.",
        "step_3": "3. Genera análisis con IA en tu idioma seleccionado.",
        "step_4": "4. Haz preguntas de seguimiento para explorar patrones, riesgos y oportunidades.",
        "step_5": "5. Descarga informes y comparte hallazgos con colegas o clientes.",
        "workspace_loaded": "Espacio de análisis cargado correctamente.",
        "text_mode_warning": "Este archivo se convirtió en una tabla basada en texto para el análisis. Los gráficos pueden ser limitados según el contenido extraído.",
        "snapshot": "🔍 Vista previa del conjunto de datos",
        "text_preview": "Vista previa del texto extraído",
        "preview": "Vista previa",
        "column_summary": "Resumen de columnas",
        "data_quality_notes": "Notas sobre la calidad de los datos",
        "cleaning_assistant": "🧹 Asistente de limpieza de datos",
        "show_cleaning": "Mostrar sugerencias de limpieza",
        "missing_cols": "Columnas con valores faltantes:",
        "no_missing": "No se encontraron valores faltantes.",
        "dup_review": "filas duplicadas para revisar.",
        "no_dup": "No se encontraron filas duplicadas.",
        "insight_board_title": "📌 Panel de insights",
        "numeric_spotlight": "Enfoque numérico",
        "category_spotlight": "Enfoque categórico",
        "choose_numeric": "Elige una columna numérica",
        "choose_categorical": "Elige una columna categórica",
        "no_numeric": "No hay columnas numéricas disponibles.",
        "no_categorical": "No hay columnas categóricas disponibles.",
        "quality_snapshot": "Resumen de calidad de datos",
        "contains_missing": "Este conjunto de datos contiene valores faltantes.",
        "no_missing_detected": "No se detectaron valores faltantes.",
        "contains_dup": "Este conjunto de datos contiene filas duplicadas.",
        "no_dup_detected": "No se detectaron filas duplicadas.",
        "correlation_map": "Mapa de correlación",
        "need_two_numeric": "Se necesitan al menos dos columnas numéricas para un mapa de correlación.",
        "quick_text_insights": "Insights rápidos",
        "charts_title": "📈 Panel interactivo / Gráficos",
        "chart_type": "Elige el tipo de gráfico",
        "histogram": "Histograma",
        "box_plot": "Diagrama de caja",
        "bar_chart": "Gráfico de barras",
        "line_chart": "Gráfico de líneas",
        "choose_numeric_chart": "Elige una columna numérica",
        "choose_cat_filter": "Elige un filtro categórico",
        "none": "Ninguno",
        "select_value": "Selecciona un valor de",
        "select_cat_for_bar": "Selecciona una columna categórica para usar el gráfico de barras.",
        "choose_date": "Elige una columna de fecha/año",
        "no_date": "No se encontró una columna de fecha/año para el gráfico de líneas.",
        "ai_insights_title": "🧠 Insights inteligentes de IA",
        "analysis_focus": "Elige el enfoque del análisis",
        "focus_general": "Análisis general",
        "focus_sales": "Análisis de ventas",
        "focus_marketing": "Análisis de marketing",
        "focus_customer": "Análisis de clientes",
        "focus_operations": "Análisis de operaciones",
        "focus_financial": "Análisis financiero",
        "focus_document": "Análisis de resumen de documento",
        "free_used": "Análisis gratuitos usados",
        "free_limit": "Se alcanzó el límite gratuito. Actualiza a Premium para continuar.",
        "analyze_data": "Analizar datos",
        "generating": "Generando insights...",
        "analysis_results": "Resultados del análisis",
        "how_read_analysis": "🧭 Cómo leer este análisis",
        "download_insights": "📥 Descargar insights (.txt)",
        "translate_title": "🌍 Traducir resultados",
        "translation_language": "Elige el idioma de traducción",
        "generate_first": "Primero genera insights con IA y luego tradúcelos aquí.",
        "translate_insights": "Traducir insights",
        "no_translate_yet": "Todavía no hay insights para traducir.",
        "translating": "Traduciendo a",
        "translated_results": "Resultados traducidos",
        "download_translation": "📥 Descargar traducción",
        "ask_export_title": "💬 Preguntar y exportar",
        "ask_export_text": "Haz preguntas de seguimiento en lenguaje natural y luego exporta tus resultados.",
        "question_placeholder": "Ejemplo: ¿qué tendencia importa más en este conjunto de datos o documento?",
        "thinking": "Pensando...",
        "answer": "Respuesta",
        "export_results": "📥 Exportar resultados",
        "download_csv": "📥 Descargar datos procesados como CSV",
        "download_pdf": "📄 Descargar informe PDF",
        "download_translated": "🌍 Descargar resultados traducidos",
        "premium_title": "💎 Actualizar a Premium",
        "premium_text": "Desbloquea análisis ilimitados, insights más profundos, exportaciones avanzadas, traducciones y reportes premium.",
        "premium_button": "💎 Actualizar a Premium",
        "next_steps": "Siguientes pasos sugeridos",
        "next_1": "- Revisa valores faltantes y duplicados.",
        "next_2": "- Explora columnas importantes.",
        "next_3": "- Genera insights con IA según tu objetivo de negocio.",
        "next_4": "- Haz preguntas de seguimiento para una comprensión más profunda.",
        "next_5": "- Descarga y comparte tu informe.",
    }
}

# =========================================================
# CSS / STYLING
# =========================================================
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: Arial, sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(29, 78, 137, 0.18) 0%, rgba(29, 78, 137, 0.00) 28%),
        radial-gradient(circle at bottom right, rgba(34, 197, 94, 0.10) 0%, rgba(34, 197, 94, 0.00) 30%),
        linear-gradient(135deg, #081018 0%, #0c1622 45%, #111827 100%);
    background-attachment: fixed;
}

.block-container {
    padding-top: 0.7rem !important;
    padding-bottom: 0.15rem !important;
    max-width: 96% !important;
}

main {
    padding-bottom: 0 !important;
    margin-bottom: 0 !important;
}

footer {
    visibility: hidden;
    height: 0 !important;
}

section[data-testid="stSidebar"] {
    background: #f3f4f6 !important;
    width: 210px !important;
    min-width: 210px !important;
    border-right: 1px solid rgba(15, 23, 42, 0.08);
}

section[data-testid="stSidebar"] .block-container {
    padding-top: 0.9rem !important;
    padding-bottom: 0.8rem !important;
    padding-left: 0.9rem !important;
    padding-right: 0.9rem !important;
}

section[data-testid="stSidebar"] * {
    color: #163a6b !important;
}

section[data-testid="stSidebar"] .stButton > button {
    background: linear-gradient(135deg, #67e8f9 0%, #22d3ee 100%) !important;
    color: #000000 !important;
    font-weight: 800 !important;
    border: 1px solid #06b6d4 !important;
    border-radius: 10px !important;
}

section[data-testid="stSidebar"] .stButton > button:hover {
    background: linear-gradient(135deg, #a5f3fc 0%, #67e8f9 100%) !important;
    color: #000000 !important;
}

.hero-wrap {
    background: linear-gradient(135deg, #9fdfb4 0%, #82d39d 50%, #63c48a 100%);
    border-radius: 18px;
    padding: 24px 24px;
    margin-top: 26px !important;
    margin-bottom: 12px !important;
    box-shadow: 0 6px 18px rgba(0,0,0,0.18);
    border: 1px solid rgba(255,255,255,0.18);
    position: relative;
    overflow: hidden;
}

.hero-wrap::before {
    content: "";
    position: absolute;
    top: -16px;
    right: 20px;
    width: 120px;
    height: 120px;
    background: radial-gradient(circle, rgba(255,255,255,0.18) 0%, rgba(255,255,255,0.00) 72%);
    pointer-events: none;
}

.hero-wrap::after {
    content: "";
    display: block;
    margin-top: 14px;
    border-bottom: 1px solid rgba(20, 83, 45, 0.18);
}

.hero-title {
    font-size: 2.2rem;
    font-weight: 800;
    color: #14532d !important;
    margin-bottom: 6px;
}

.hero-motto {
    font-size: 1.08rem;
    color: #f8fffb !important;
    font-weight: 700;
    margin-bottom: 14px;
}

.hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.78);
    color: #14532d;
    padding: 7px 12px;
    margin: 4px 6px 0 0;
    border-radius: 999px;
    font-size: 0.88rem;
    font-weight: 600;
}

.section-wrap {
    background: linear-gradient(
        135deg,
        rgba(230, 255, 242, 0.94) 0%,
        rgba(217, 251, 232, 0.92) 45%,
        rgba(199, 249, 212, 0.90) 100%
    );
    backdrop-filter: blur(6px);
    -webkit-backdrop-filter: blur(6px);
    border-radius: 22px;
    padding: 20px;
    margin-top: 8px;
    margin-bottom: 8px;
    box-shadow: 0 10px 24px rgba(0, 0, 0, 0.18);
    border: 2px solid rgba(20, 83, 45, 0.18);
    position: relative;
    overflow: hidden;
}

.section-wrap::before {
    content: "";
    position: absolute;
    top: -30px;
    right: -30px;
    width: 140px;
    height: 140px;
    background: radial-gradient(circle, rgba(255,255,255,0.22) 0%, rgba(255,255,255,0.00) 72%);
    pointer-events: none;
}

.section-wrap::after {
    content: "";
    position: absolute;
    bottom: -35px;
    left: -35px;
    width: 160px;
    height: 160px;
    background: radial-gradient(circle, rgba(34,197,94,0.08) 0%, rgba(34,197,94,0.00) 72%);
    pointer-events: none;
}

.section-title {
    color: #14532d !important;
    font-size: 1.4rem;
    font-weight: 800;
    margin-bottom: 8px;
}

.section-subtle {
    color: #166534 !important;
    font-size: 0.98rem;
    margin-bottom: 14px;
}

button[data-baseweb="tab"] {
    border: 1px solid rgba(255,255,255,0.16) !important;
    border-radius: 12px 12px 0 0 !important;
    margin-right: 6px !important;
    background: rgba(255,255,255,0.03) !important;
}

button[data-baseweb="tab"] * {
    color: #f3f4f6 !important;
    font-weight: 600 !important;
}

button[data-baseweb="tab"][aria-selected="true"] {
    border: 1px solid rgba(116, 198, 157, 0.55) !important;
    background: rgba(116, 198, 157, 0.10) !important;
}

/* =========================
   FINAL UPLOADER STYLING
========================= */
.custom-upload-card {
    background: #d8d2c8 !important;
    border: 1.5px solid rgba(90, 82, 72, 0.25);
    border-radius: 16px;
    padding: 16px 18px !important;
    margin-bottom: 6px;
    box-shadow: 0 4px 10px rgba(0,0,0,0.08);
}

.custom-upload-title {
    color: #2f855a !important;
    font-size: 1.3rem !important;   /* 🔥 bigger */
    font-weight: 800 !important;
    margin-bottom: 6px !important;
}

.custom-upload-drop {
    color: #8b5e3c !important;
    font-size: 1.1rem !important;   /* 🔥 bigger */
    font-weight: 700 !important;
    margin-bottom: 6px !important;
}

i.custom-upload-limit {
    background: linear-gradient(135deg, #c2410c 0%, #ea580c 50%, #f97316 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 1rem !important;
    font-weight: 800 !important;
}

[data-testid="stFileUploader"] {
    background: linear-gradient(135deg, #d8d2c8 0%, #cfc8bd 55%, #c4bdb2 100%) !important;
    border: 1.5px solid rgba(90, 82, 72, 0.22) !important;
    border-top: none !important;
    border-radius: 0 0 16px 16px !important;
    padding: 2px 10px 4px 10px !important;
    margin-top: 0 !important;
    box-shadow:
        0 4px 10px rgba(0,0,0,0.08),
        inset 0 0 0 2px rgba(120, 108, 96, 0.22) !important;
}

/* hide native uploader title and prompt text */
[data-testid="stFileUploader"] label,
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] section div div {
    display: none !important;
}

/* Browse files button */
[data-testid="stFileUploader"] button {
    background: #fbbf24 !important;
    color: #000000 !important;
    border: 1px solid #f59e0b !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    box-shadow: none !important;
}

[data-testid="stFileUploader"] button:hover {
    background: #f59e0b !important;
    color: #000000 !important;
    border: 1px solid #f59e0b !important;
}

.stButton > button {
    border-radius: 10px !important;
    background-color: #21a366 !important;
    color: white !important;
    font-weight: 700 !important;
    border: none !important;
}

.stButton > button:hover {
    background-color: #188a54 !important;
    color: white !important;
}

.stButton > button:focus:not(:active) {
    background-color: #21a366 !important;
    color: white !important;
    border: none !important;
    box-shadow: none !important;
}

.stDownloadButton > button {
    border-radius: 10px !important;
    background-color: #2563eb !important;
    color: white !important;
    font-weight: 700 !important;
    border: none !important;
}

.stDownloadButton > button:hover {
    background-color: #1d4ed8 !important;
    color: white !important;
}

[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# SETUP
# =========================================================
env_path = r"C:\\Users\\Saintelus\\Documents\\explain-my-data\\.env"
load_dotenv(dotenv_path=env_path, override=True)

api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
    except Exception:
        api_key = None

if not api_key:
    st.error("OPENAI_API_KEY not found in .env or Streamlit secrets.")
    st.stop()

client = OpenAI(api_key=api_key)

# =========================================================
# SESSION STATE
# =========================================================
if "usage_count" not in st.session_state:
    st.session_state.usage_count = 0
if "result" not in st.session_state:
    st.session_state.result = ""
if "translated_result" not in st.session_state:
    st.session_state.translated_result = ""
if "chat_answer" not in st.session_state:
    st.session_state.chat_answer = ""
if "loaded_mode" not in st.session_state:
    st.session_state.loaded_mode = ""
if "raw_text_content" not in st.session_state:
    st.session_state.raw_text_content = ""
if "app_language" not in st.session_state:
    st.session_state.app_language = "English"

# =========================================================
# HELPERS
# =========================================================
def create_pdf_report(text: str) -> bytes:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    _, height = letter

    text_object = pdf.beginText(40, height - 50)
    text_object.setFont("Helvetica", 11)

    lines = text.split("\\n")
    for line in lines:
        wrapped_lines = [line[i:i+95] for i in range(0, len(line), 95)] if line else [""]
        for wrapped_line in wrapped_lines:
            if text_object.getY() < 50:
                pdf.drawText(text_object)
                pdf.showPage()
                text_object = pdf.beginText(40, height - 50)
                text_object.setFont("Helvetica", 11)
            text_object.textLine(wrapped_line)

    pdf.drawText(text_object)
    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()

def load_structured_data(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file)
    if file.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(file)
    raise ValueError("Unsupported structured file format.")

def load_text_file(file):
    return file.read().decode("utf-8", errors="ignore")

def load_docx_file(file):
    doc = Document(file)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\\n".join(paragraphs)

def text_to_dataframe(text, source_name="text_document"):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return pd.DataFrame({"source": [source_name], "content": ["No readable text found."]})

    return pd.DataFrame({
        "source": [source_name] * len(lines),
        "line_number": list(range(1, len(lines) + 1)),
        "content": lines
    })

def load_any_file(file):
    name = file.name.lower()

    if name.endswith((".csv", ".xlsx", ".xls")):
        df = load_structured_data(file)
        return df, "table", ""

    if name.endswith(".txt"):
        raw_text = load_text_file(file)
        df = text_to_dataframe(raw_text, source_name=file.name)
        return df, "text", raw_text

    if name.endswith(".docx"):
        raw_text = load_docx_file(file)
        df = text_to_dataframe(raw_text, source_name=file.name)
        return df, "text", raw_text

    raise ValueError("Unsupported file format. Please upload CSV, XLSX, XLS, TXT, or DOCX.")

def build_data_summary(df, mode, raw_text):
    if mode == "table":
        return df.describe(include="all").fillna("").to_string()

    preview_lines = raw_text.splitlines()[:80]
    preview_text = "\\n".join(preview_lines)
    return f"""
Document/Text Summary
---------------------
Rows created from text: {df.shape[0]}
Columns: {df.shape[1]}
Source columns: {', '.join(df.columns)}

Preview of extracted text:
{preview_text}
"""

# =========================================================
# SIDEBAR + LANGUAGE
# =========================================================
with st.sidebar:
    language = st.selectbox(
        TRANSLATIONS["English"]["language_selector"],
        list(TRANSLATIONS.keys()),
        index=list(TRANSLATIONS.keys()).index(st.session_state.app_language)
    )
    st.session_state.app_language = language
    T = TRANSLATIONS[language]

    st.markdown(f"## {T['app_name']}")
    st.write(T["tagline"])
    st.markdown("---")
    st.write(f"**{T['features']}**")
    st.write(T["feature_1"])
    st.write(T["feature_2"])
    st.write(T["feature_3"])
    st.write(T["feature_4"])
    st.write(T["feature_5"])
    st.markdown("---")

    if st.button(T["ai_ready"], use_container_width=True):
        try:
            client.responses.create(
                model="gpt-5",
                input=f"Respond in {language}. Confirm in 2 short sentences that the AI assistant is ready and the user can upload, analyze, ask questions, and export results."
            )
            st.success(T["ai_ready_ok"])
        except Exception:
            st.warning(T["ai_ready_fail"])

    st.markdown("---")
    st.caption("Built with Streamlit, Python, Pandas, Matplotlib, ReportLab, python-docx, and OpenAI.")

T = TRANSLATIONS[st.session_state.app_language]
language = st.session_state.app_language

# =========================================================
# HEADER
# =========================================================
st.markdown('<div class="hero-wrap">', unsafe_allow_html=True)

logo_col, text_col = st.columns([1, 4])

with logo_col:
    if os.path.exists("edai_logo.png"):
        st.image("edai_logo.png", width=180)

with text_col:
    st.markdown(f'<div class="hero-title">{T["header_title"]}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="hero-motto">{T["header_motto"]}</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div>
        <span class="hero-badge">{T["badge_1"]}</span>
        <span class="hero-badge">{T["badge_2"]}</span>
        <span class="hero-badge">{T["badge_3"]}</span>
        <span class="hero-badge">{T["badge_4"]}</span>
        <span class="hero-badge">{T["badge_5"]}</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# CUSTOM UPLOADER TEXT + REAL UPLOADER
# =========================================================
st.markdown(
    """
    <div class="custom-upload-card">
        <div class="custom-upload-title">Upload your data or document</div>
        <div class="custom-upload-drop">Start by selecting a file below to generate insights</div>
        <div class="custom-upload-limit">
            Supported formats: CSV, Excel, TXT, DOCX • Max size: 200MB
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "File uploader",
    type=["csv", "xlsx", "xls", "txt", "docx"],
    label_visibility="collapsed"
)

st.markdown(
    f"""
    <div style="
        background: linear-gradient(135deg, #0b1f3a 0%, #12345a 50%, #1d4e89 100%);
        border-radius: 12px;
        padding: 10px 14px;
        border: 1px solid rgba(255,255,255,0.08);
        color: #f5f9ff;
        font-weight: 600;
        margin-top: 6px;
        margin-bottom: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.18);
    ">
        {T["supported_formats"]}
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    '<div style="height:1px; background:rgba(255,255,255,0.12); margin:10px 0 6px 0;"></div>',
    unsafe_allow_html=True
)

# =========================================================
# MAIN
# =========================================================
if uploaded_file is not None:
    try:
        df, mode, raw_text = load_any_file(uploaded_file)
        st.session_state.loaded_mode = mode
        st.session_state.raw_text_content = raw_text

        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        categorical_cols = df.select_dtypes(include=["object", "category", "bool"]).columns.tolist()

        st.subheader(T["data_overview"])
        m1, m2, m3, m4 = st.columns(4)
        m1.metric(T["rows"], df.shape[0])
        m2.metric(T["columns"], df.shape[1])
        m3.metric(T["missing"], int(df.isna().sum().sum()))
        m4.metric(T["duplicates"], int(df.duplicated().sum()))

        if mode == "text":
            st.warning(T["text_mode_warning"])

        st.markdown("---")

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
            [
                T["overview_tab"],
                T["insight_tab"],
                T["charts_tab"],
                T["ai_tab"],
                T["translate_tab"],
                T["export_tab"],
            ]
        )

        with tab1:
            st.markdown(f"""
            <div class="section-wrap">
                <div class="section-title">{T["overview_tab"]}</div>
                <div class="section-subtle">{T["tagline"]}</div>
            """, unsafe_allow_html=True)

            st.subheader(T["snapshot"])
            st.dataframe(df.head(20))

            if mode == "text" and raw_text:
                st.markdown(f"### {T['text_preview']}")
                st.text_area(T["preview"], raw_text[:4000], height=220)

            left_col, right_col = st.columns(2)

            with left_col:
                st.markdown(f"### {T['column_summary']}")
                column_info = pd.DataFrame({
                    "Column": df.columns,
                    "Type": [str(df[col].dtype) for col in df.columns],
                    "Missing Values": [int(df[col].isna().sum()) for col in df.columns],
                    "Unique Values": [int(df[col].nunique(dropna=True)) for col in df.columns],
                })
                st.dataframe(column_info)

            with right_col:
                st.markdown(f"### {T['data_quality_notes']}")
                st.write(f"- {T['missing']}: {int(df.isna().sum().sum())}")
                st.write(f"- {T['duplicates']}: {int(df.duplicated().sum())}")
                if len(numeric_cols) == 0:
                    st.write(f"- {T['no_numeric']}")
                if len(categorical_cols) == 0:
                    st.write(f"- {T['no_categorical']}")

            st.markdown(f"### {T['cleaning_assistant']}")
            if st.button(T["show_cleaning"]):
                missing_cols = df.columns[df.isna().sum() > 0].tolist()
                if missing_cols:
                    st.write(T["missing_cols"])
                    for col in missing_cols:
                        st.write(f"- {col}: {int(df[col].isna().sum())}")
                else:
                    st.success(T["no_missing"])

                duplicates = int(df.duplicated().sum())
                if duplicates > 0:
                    st.warning(f"{duplicates} {T['dup_review']}")
                else:
                    st.success(T["no_dup"])

            st.markdown("</div>", unsafe_allow_html=True)

        with tab2:
            st.markdown(f"""
            <div class="section-wrap">
                <div class="section-title">{T["insight_tab"]}</div>
                <div class="section-subtle">{T["insight_board_title"]}</div>
            """, unsafe_allow_html=True)

            st.subheader(T["insight_board_title"])

            c1, c2, c3, c4 = st.columns(4)
            c1.metric(T["rows"], df.shape[0])
            c2.metric(T["columns"], df.shape[1])
            c3.metric(T["missing"], int(df.isna().sum().sum()))
            c4.metric(T["duplicates"], int(df.duplicated().sum()))

            st.markdown("---")

            left_insight, right_insight = st.columns(2)

            with left_insight:
                st.markdown(f"### {T['numeric_spotlight']}")
                if numeric_cols:
                    selected_num_spotlight = st.selectbox(
                        T["choose_numeric"],
                        numeric_cols,
                        key="spotlight_num"
                    )
                    st.write(f"**Mean:** {df[selected_num_spotlight].mean():.2f}")
                    st.write(f"**Median:** {df[selected_num_spotlight].median():.2f}")
                    st.write(f"**Min:** {df[selected_num_spotlight].min():.2f}")
                    st.write(f"**Max:** {df[selected_num_spotlight].max():.2f}")
                    st.write(f"**Std Dev:** {df[selected_num_spotlight].std():.2f}")
                else:
                    st.info(T["no_numeric"])

            with right_insight:
                st.markdown(f"### {T['category_spotlight']}")
                if categorical_cols:
                    selected_cat_spotlight = st.selectbox(
                        T["choose_categorical"],
                        categorical_cols,
                        key="spotlight_cat"
                    )
                    value_counts_df = (
                        df[selected_cat_spotlight]
                        .astype(str)
                        .value_counts(dropna=False)
                        .head(10)
                        .reset_index()
                    )
                    value_counts_df.columns = [selected_cat_spotlight, "Count"]
                    st.dataframe(value_counts_df)
                else:
                    st.info(T["no_categorical"])

            st.markdown("---")
            st.markdown(f"### {T['quality_snapshot']}")

            if df.isna().sum().sum() > 0:
                st.warning(T["contains_missing"])
            else:
                st.success(T["no_missing_detected"])

            if df.duplicated().sum() > 0:
                st.warning(T["contains_dup"])
            else:
                st.success(T["no_dup_detected"])

            st.markdown("---")
            st.markdown(f"### {T['correlation_map']}")

            if len(numeric_cols) >= 2:
                corr = df[numeric_cols].corr(numeric_only=True)
                fig, ax = plt.subplots(figsize=(8, 5))
                cax = ax.matshow(corr, cmap="Greens")
                fig.colorbar(cax)
                ax.set_xticks(range(len(corr.columns)))
                ax.set_yticks(range(len(corr.columns)))
                ax.set_xticklabels(corr.columns, rotation=45, ha="left")
                ax.set_yticklabels(corr.columns)
                st.pyplot(fig)
            else:
                st.info(T["need_two_numeric"])

            st.markdown(f"### {T['quick_text_insights']}")
            st.write("- Review the KPI cards first to understand size and quality.")
            st.write("- Use the numeric spotlight to detect spread and outliers.")
            st.write("- Use the category spotlight to spot dominant groups or imbalances.")
            st.write("- Use the correlation map to identify relationships worth investigating.")

            st.markdown("</div>", unsafe_allow_html=True)

        with tab3:
            st.markdown(f"""
            <div class="section-wrap">
                <div class="section-title">{T["charts_tab"]}</div>
                <div class="section-subtle">{T["charts_title"]}</div>
            """, unsafe_allow_html=True)

            st.subheader(T["charts_title"])

            chart_type = st.selectbox(
                T["chart_type"],
                [T["histogram"], T["box_plot"], T["bar_chart"], T["line_chart"]]
            )

            selected_num = st.selectbox(
                T["choose_numeric_chart"],
                numeric_cols if numeric_cols else [T["no_numeric"]],
                key="chart_num"
            )

            selected_cat = st.selectbox(
                T["choose_cat_filter"],
                [T["none"]] + categorical_cols,
                key="chart_cat"
            )

            filtered_df = df.copy()

            if selected_cat != T["none"]:
                category_values = (
                    filtered_df[selected_cat]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
                if category_values:
                    category_value = st.selectbox(
                        f"{T['select_value']} {selected_cat}",
                        category_values
                    )
                    filtered_df = filtered_df[filtered_df[selected_cat].astype(str) == category_value]

            if numeric_cols and selected_num != T["no_numeric"]:
                fig, ax = plt.subplots()

                if chart_type == T["histogram"]:
                    filtered_df[selected_num].dropna().plot(kind="hist", bins=20, ax=ax)
                    ax.set_title(f"Distribution of {selected_num}")

                elif chart_type == T["box_plot"]:
                    filtered_df.boxplot(column=selected_num, ax=ax)
                    ax.set_title(f"Box Plot of {selected_num}")

                elif chart_type == T["bar_chart"]:
                    if selected_cat != T["none"]:
                        grouped = (
                            filtered_df.groupby(selected_cat)[selected_num]
                            .mean()
                            .sort_values(ascending=False)
                            .head(10)
                        )
                        grouped.plot(kind="bar", ax=ax)
                        ax.set_title(f"Average {selected_num} by {selected_cat}")
                    else:
                        st.info(T["select_cat_for_bar"])

                elif chart_type == T["line_chart"]:
                    date_candidates = [
                        c for c in filtered_df.columns
                        if "date" in c.lower() or "year" in c.lower()
                    ]
                    if date_candidates:
                        date_col = st.selectbox(T["choose_date"], date_candidates, key="line_date")
                        temp_df = filtered_df.copy()
                        temp_df[date_col] = pd.to_datetime(temp_df[date_col], errors="coerce")
                        temp_df = temp_df.dropna(subset=[date_col, selected_num]).sort_values(date_col)
                        ax.plot(temp_df[date_col], temp_df[selected_num])
                        ax.set_title(f"{selected_num} Over Time")
                        plt.xticks(rotation=45, ha="right")
                    else:
                        st.info(T["no_date"])

                st.pyplot(fig)
            else:
                st.info(T["no_numeric"])

            st.markdown("</div>", unsafe_allow_html=True)

        with tab4:
            st.markdown(f"""
            <div class="section-wrap">
                <div class="section-title">{T["ai_tab"]}</div>
                <div class="section-subtle">{T["ai_insights_title"]}</div>
            """, unsafe_allow_html=True)

            st.subheader(T["ai_insights_title"])

            focus_options = [
                T["focus_general"],
                T["focus_sales"],
                T["focus_marketing"],
                T["focus_customer"],
                T["focus_operations"],
                T["focus_financial"],
                T["focus_document"],
            ]

            analysis_goal = st.selectbox(T["analysis_focus"], focus_options)

            st.write(f"{T['free_used']}: {st.session_state.usage_count}/2")

            if st.session_state.usage_count >= 2:
                st.error(T["free_limit"])

            if st.button(T["analyze_data"]):
                if st.session_state.usage_count >= 2:
                    st.stop()

                with st.spinner(T["generating"]):
                    summary = build_data_summary(df, mode, raw_text)

                    response = client.responses.create(
                        model="gpt-5",
                        input=f"""
Respond in {language}.

You are a senior business data analyst.

The user selected this analysis focus: {analysis_goal}

Analyze the uploaded content summary below and provide these sections clearly:

1. Executive Summary
2. Key Insights
3. What These Insights Mean
4. Risks or Anomalies
5. Recommended Actions
6. Expected Business Impact
7. Questions Worth Exploring Next

Instructions:
- Write in clear, simple language for a non-technical user
- Do not invent unsupported facts
- Be practical, specific, and decision-oriented
- If the uploaded content is text or document-based, summarize and interpret it intelligently
- Keep the tone professional and concise

Uploaded content summary:
{summary}
"""
                    )

                    st.session_state.result = response.output_text
                    st.session_state.translated_result = ""
                    st.session_state.chat_answer = ""
                    st.session_state.usage_count += 1

            if st.session_state.result:
                st.markdown(f"### {T['analysis_results']}")
                st.write(st.session_state.result)

                st.markdown(f"### {T['how_read_analysis']}")
                st.info("""
- Start with the Executive Summary for a quick overview
- Focus on Recommended Actions for decisions
- Use Expected Business Impact to understand why it matters
- Explore Questions Worth Exploring Next to deepen analysis
""")

                st.download_button(
                    label=T["download_insights"],
                    data=st.session_state.result,
                    file_name="insights.txt",
                    mime="text/plain"
                )

            st.markdown("</div>", unsafe_allow_html=True)

        with tab5:
            st.markdown(f"""
            <div class="section-wrap">
                <div class="section-title">{T["translate_tab"]}</div>
                <div class="section-subtle">{T["translate_title"]}</div>
            """, unsafe_allow_html=True)

            st.subheader(T["translate_title"])

            translation_language = st.selectbox(
                T["translation_language"],
                ["English", "French", "Spanish", "Portuguese", "German", "Arabic", "Haitian Creole", "Chinese", "Japanese"]
            )

            if not st.session_state.result:
                st.info(T["generate_first"])

            if st.button(T["translate_insights"]):
                if not st.session_state.result:
                    st.warning(T["no_translate_yet"])
                else:
                    with st.spinner(f"{T['translating']} {translation_language}..."):
                        response = client.responses.create(
                            model="gpt-5",
                            input=f"""
Translate the following business analysis into {translation_language}.
Keep the meaning clear, professional, and easy to understand.

Text:
{st.session_state.result}
"""
                        )
                        st.session_state.translated_result = response.output_text

            if st.session_state.translated_result:
                st.markdown(f"### {T['translated_results']} ({translation_language})")
                st.write(st.session_state.translated_result)

                st.download_button(
                    label=f"{T['download_translation']} ({translation_language})",
                    data=st.session_state.translated_result,
                    file_name="translated_insights.txt",
                    mime="text/plain"
                )

            st.markdown("</div>", unsafe_allow_html=True)

        with tab6:
            st.markdown(f"""
            <div class="section-wrap">
                <div class="section-title">{T["export_tab"]}</div>
                <div class="section-subtle">{T["ask_export_title"]}</div>
            """, unsafe_allow_html=True)

            st.subheader(T["ask_export_title"])
            st.write(T["ask_export_text"])

            user_question = st.text_input(T["question_placeholder"])

            if user_question:
                with st.spinner(T["thinking"]):
                    context = build_data_summary(df, mode, raw_text)

                    response = client.responses.create(
                        model="gpt-5",
                        input=f"""
Respond in {language}.

You are a helpful data analyst.

Here is the uploaded content summary:
{context}

User question:
{user_question}

Answer clearly, simply, and directly.
"""
                    )

                    st.session_state.chat_answer = response.output_text

            if st.session_state.chat_answer:
                st.markdown(f"### {T['answer']}")
                st.write(st.session_state.chat_answer)

            st.markdown("---")
            st.subheader(T["export_results"])

            csv_data = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label=T["download_csv"],
                data=csv_data,
                file_name="processed_data.csv",
                mime="text/csv"
            )

            if st.session_state.result:
                pdf_bytes = create_pdf_report(st.session_state.result)
                st.download_button(
                    label=T["download_pdf"],
                    data=pdf_bytes,
                    file_name="ExplainMyData_Report.pdf",
                    mime="application/pdf"
                )

            if st.session_state.translated_result:
                st.download_button(
                    label=T["download_translated"],
                    data=st.session_state.translated_result,
                    file_name="translated_insights.txt",
                    mime="text/plain"
                )

            st.markdown("---")
            st.markdown(f"### {T['premium_title']}")
            st.write(T["premium_text"])

            premium_url = "PASTE_YOUR_STRIPE_LINK_HERE"
            st.link_button(T["premium_button"], premium_url)

            st.markdown(f"### {T['next_steps']}")
            st.write(T["next_1"])
            st.write(T["next_2"])
            st.write(T["next_3"])
            st.write(T["next_4"])
            st.write(T["next_5"])

            st.markdown("</div>", unsafe_allow_html=True)

        st.success(T["workspace_loaded"])

    except Exception as e:
        st.error(f"Something went wrong: {e}")

else:
    st.markdown(
        f'<p style="color:#fff3a3; font-weight:700; font-size:1.08rem; margin-bottom:4px;">{T["upload_begin"]}</p>',
        unsafe_allow_html=True
    )

    st.markdown(
        f'<h3 style="color:#f3f4f6; margin-top:0; margin-bottom:4px;">{T["how_it_works"]}</h3>',
        unsafe_allow_html=True
    )
    st.markdown(f'<p style="color:#e5e7eb; margin-bottom:2px;">{T["step_1"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p style="color:#e5e7eb; margin-bottom:2px;">{T["step_2"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p style="color:#e5e7eb; margin-bottom:2px;">{T["step_3"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p style="color:#e5e7eb; margin-bottom:2px;">{T["step_4"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p style="color:#e5e7eb; margin-bottom:2px;">{T["step_5"]}</p>', unsafe_allow_html=True)