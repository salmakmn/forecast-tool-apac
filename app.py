
"""
================================================================================
SECTION 1 — GUIDE D'INSTALLATION
================================================================================

FORECAST SUPPORT TOOL — APAC | Small Molecules EM&S Finance
Application Streamlit interactive — app.py

--------------------------------------------------------------------------------
PRÉREQUIS
--------------------------------------------------------------------------------
- Python 3.8 ou supérieur
- pip (gestionnaire de paquets Python)
- Le fichier data_apac_pilote.xlsx dans le même dossier que app.py

--------------------------------------------------------------------------------
ÉTAPE 1 — INSTALLATION DES DÉPENDANCES
--------------------------------------------------------------------------------
Ouvre un terminal (VS Code : Terminal > New Terminal) et exécute :

    pip install streamlit pandas openpyxl plotly numpy xlsxwriter

Vérification de l'installation :
    streamlit --version

--------------------------------------------------------------------------------
ÉTAPE 2 — STRUCTURE DES FICHIERS
--------------------------------------------------------------------------------
Ton dossier de projet doit ressembler à ceci :

    mon_projet/
    ├── app.py                    ← ce fichier
    └── data_apac_pilote.xlsx     ← ton fichier de données

--------------------------------------------------------------------------------
ÉTAPE 3 — LANCER L'APPLICATION
--------------------------------------------------------------------------------
Dans le terminal, depuis le dossier contenant app.py :

    streamlit run app.py

L'application s'ouvrira automatiquement dans ton navigateur à l'adresse :
    http://localhost:8501

Pour arrêter l'application : Ctrl+C dans le terminal

--------------------------------------------------------------------------------
DÉPANNAGE COURANT
--------------------------------------------------------------------------------
Problème : "ModuleNotFoundError: No module named 'streamlit'"
Solution  : pip install streamlit

Problème : "FileNotFoundError: data_apac_pilote.xlsx"
Solution  : Vérifie que le fichier Excel est dans le même dossier que app.py

Problème : "Port 8501 already in use"
Solution  : streamlit run app.py --server.port 8502

Problème : Graphiques qui ne s'affichent pas
Solution  : pip install --upgrade plotly

--------------------------------------------------------------------------------
DÉVELOPPÉ PAR : Salma Kamoun — EM&S Finance APAC
VERSION       : 1.0
================================================================================
"""


# ==============================================================================
# SECTION 2 — CODE COMPLET DE L'APPLICATION
# ==============================================================================

# ------------------------------------------------------------------------------
# IMPORTS
# ------------------------------------------------------------------------------
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

# ==============================================================================
# CONFIGURATION DE LA PAGE — doit être le premier appel Streamlit
# ==============================================================================
st.set_page_config(
    page_title="Forecast Support Tool — APAC",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# CSS PERSONNALISÉ — Style professionnel Sanofi
# ==============================================================================
st.markdown("""
<style>
    /* --- Couleurs et police globale --- */
    html, body, [class*="css"] {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* --- En-tête principal --- */
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #0066CC;
        border-bottom: 3px solid #0066CC;
        padding-bottom: 0.5rem;
        margin-bottom: 1.5rem;
    }

    /* --- Sous-titre --- */
    .sub-header {
        font-size: 0.95rem;
        color: #666666;
        margin-top: -1rem;
        margin-bottom: 1.5rem;
    }

    /* --- Cartes métriques personnalisées --- */
    .metric-card {
        background-color: #f0f7ff;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        border-left: 4px solid #0066CC;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }

    .metric-card-highlight {
        background-color: #fff8f0;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        border-left: 4px solid #FF6B35;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }

    .metric-card-success {
        background-color: #f0fff4;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        border-left: 4px solid #28A745;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }

    /* --- Titres de section --- */
    .section-title {
        font-size: 1.2rem;
        font-weight: 600;
        color: #0066CC;
        border-bottom: 2px solid #e0eeff;
        padding-bottom: 0.4rem;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }

    /* --- Badge MAPE --- */
    .badge-green {
        background-color: #d4edda;
        color: #155724;
        padding: 3px 10px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-orange {
        background-color: #fff3cd;
        color: #856404;
        padding: 3px 10px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.85rem;
    }
    .badge-red {
        background-color: #f8d7da;
        color: #721c24;
        padding: 3px 10px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 0.85rem;
    }

    /* --- Footer --- */
    .footer {
        text-align: center;
        color: #999999;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #eeeeee;
    }

    /* --- Sidebar --- */
    .sidebar-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0066CC;
    }

    /* --- Masquer le menu hamburger Streamlit --- */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONSTANTES — Noms des colonnes et labels d'affichage
# ==============================================================================

# Fichier source
FILEPATH = "data_apac_pilote.xlsx"
SHEET    = "Data_Historique"

# Colonnes identifiantes — noms exacts du fichier Excel
PRODUCT_COL = "Description"
FAMILY_COL  = "Product Family"
CMO_COL     = "CMO (Supplier)"

# Colonnes Volume (noms exacts dans le fichier Excel)
VOLUME_COLS = [
    "Volume A25", "Volume B26",
    "Volume Q1_26", "Volume Q2_26", "Volume Q3_26",
    "Volume H27", "Volume B27", "Volume H28"
]

# Colonnes Spend (noms exacts dans le fichier Excel)
SPEND_COLS = [
    "Spend A25", "Spend B26",
    "Spend Q1_26", "Spend Q2_26", "Spend Q3_26",
    "Spend H27", "Spend B27", "Spend H28"
]

# Colonnes AV (noms exacts dans le fichier Excel)
AV_COLS = [
    "AV A25", "AV B26",
    "AV Q1_26", "AV Q2_26", "AV Q3_26",
    "AV H27", "AV B27", "AV H28"
]

# Labels d'affichage pour les 8 périodes (dans l'ordre)
PERIOD_LABELS = ["A25", "B26", "Q1 26", "Q2 26", "Q3 26", "H27", "B27", "H28"]

# Mapping label → clé interne (pour le sélecteur de période)
PERIOD_KEY_MAP = {
    "A25"   : "A25",
    "B26"   : "B26",
    "Q1 26" : "Q1_26",
    "Q2 26" : "Q2_26",
    "Q3 26" : "Q3_26",
    "H27"   : "H27",
    "B27"   : "B27",
    "H28"   : "H28",
}

# Mapping clé interne → colonnes exactes du fichier Excel
PERIOD_COL_MAP = {
    "A25"   : {"volume": "Volume A25",   "spend": "Spend A25",   "av": "AV A25"},
    "B26"   : {"volume": "Volume B26",   "spend": "Spend B26",   "av": "AV B26"},
    "Q1_26" : {"volume": "Volume Q1_26", "spend": "Spend Q1_26", "av": "AV Q1_26"},
    "Q2_26" : {"volume": "Volume Q2_26", "spend": "Spend Q2_26", "av": "AV Q2_26"},
    "Q3_26" : {"volume": "Volume Q3_26", "spend": "Spend Q3_26", "av": "AV Q3_26"},
    "H27"   : {"volume": "Volume H27",   "spend": "Spend H27",   "av": "AV H27"},
    "B27"   : {"volume": "Volume B27",   "spend": "Spend B27",   "av": "AV B27"},
    "H28"   : {"volume": "Volume H28",   "spend": "Spend H28",   "av": "AV H28"},
}

# Périodes réelles (ligne pleine) vs budget/forecast (ligne pointillée)
ACTUAL_PERIODS  = {"A25", "Q1_26", "Q2_26", "Q3_26"}
FORECAST_PERIODS = {"B26", "H27", "B27", "H28"}

# Couleurs Sanofi
COLOR_BLUE   = "#0066CC"
COLOR_ORANGE = "#FF6B35"
COLOR_GREEN  = "#28A745"
COLOR_LIGHT  = "#f0f7ff"


# ==============================================================================
# FONCTION 1 — CHARGEMENT ET NETTOYAGE DES DONNÉES
# ==============================================================================
@st.cache_data
def load_data(filepath: str, sheet: str) -> pd.DataFrame:
    """
    Charge et nettoie le fichier Excel source.
    - Vérifie que toutes les colonnes requises sont présentes
    - Supprime les virgules dans les colonnes numériques (format '2,857.17')
    - Convertit en float64 avec gestion des erreurs
    - Remplace les NaN par 0
    """
    try:
        df = pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl")

        # --- Vérification défensive des colonnes requises ---
        required_cols = (
            [PRODUCT_COL, FAMILY_COL, CMO_COL]
            + VOLUME_COLS
            + SPEND_COLS
            + AV_COLS
        )
        missing_cols = [c for c in required_cols if c not in df.columns]
        if missing_cols:
            st.error(
                f"❌ **Colonnes manquantes dans le fichier Excel :**\n\n"
                f"{missing_cols}\n\n"
                f"**Colonnes disponibles dans le fichier :**\n\n"
                f"{df.columns.tolist()}"
            )
            st.stop()

        # --- Nettoyage des colonnes Volume (format texte avec virgules) ---
        for col in VOLUME_COLS:
            df[col] = (
                df[col].astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # --- Nettoyage des colonnes Spend (même traitement) ---
        for col in SPEND_COLS:
            df[col] = (
                df[col].astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # --- Nettoyage des colonnes AV (déjà en float, mais on sécurise) ---
        for col in AV_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        return df

    except FileNotFoundError:
        st.error(
            f"❌ Fichier introuvable : **{filepath}**\n\n"
            "Vérifie que `data_apac_pilote.xlsx` est dans le même dossier que `app.py`."
        )
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des données : {e}")
        return pd.DataFrame()


# ==============================================================================
# FONCTION 2 — CALCUL DES KPIs POUR UNE PÉRIODE DONNÉE
# ==============================================================================
@st.cache_data
def calculate_kpis_for_period(df: pd.DataFrame, period_key: str) -> dict:
    """
    Calcule les KPIs pour UNE période spécifique (pas de cumul entre périodes).
    period_key : clé interne de la période (ex: "B26", "Q1_26", etc.)
    Retourne un dictionnaire avec volume, spend, AV et nb_produits pour cette période.
    """
    if df.empty or period_key not in PERIOD_COL_MAP:
        return {}

    cols = PERIOD_COL_MAP[period_key]

    volume    = df[cols["volume"]].sum()
    spend     = df[cols["spend"]].sum()
    av        = df[cols["av"]].sum()
    nb_produits = df[PRODUCT_COL].nunique()

    return {
        "period_key"  : period_key,
        "volume"      : volume,
        "spend"       : spend,
        "av"          : av,
        "nb_produits" : nb_produits,
    }


# ==============================================================================
# FONCTION 3 — CALCUL DES SCÉNARIOS 2027
# ==============================================================================
@st.cache_data
def calculate_scenarios(df: pd.DataFrame) -> dict:
    """
    Calcule les 4 scénarios de forecast pour 2027 :
      1. Tendance      : extrapolation du taux de croissance A25 → 2026
      2. Budget ajusté : B27 corrigé par le facteur d'erreur historique (B26 vs réel)
      3. Pondéré       : 60% H27 + 40% B27
      4. Recommandé    : moyenne des 3 scénarios ci-dessus
    """
    if df.empty:
        return {}

    # --- Volumes de référence (chaque période prise individuellement) ---
    actual_2025_vol = df["Volume A25"].sum()
    b26_vol         = df["Volume B26"].sum()
    q1_vol          = df["Volume Q1_26"].sum()
    q2_vol          = df["Volume Q2_26"].sum()
    q3_vol          = df["Volume Q3_26"].sum()
    h27_vol         = df["Volume H27"].sum()
    b27_vol         = df["Volume B27"].sum()
    h28_vol         = df["Volume H28"].sum()

    # Réel 2026 = somme des 3 trimestres réels (Q1+Q2+Q3)
    actual_2026_vol = q1_vol + q2_vol + q3_vol

    # --- Scénario 1 : Tendance ---
    growth_rate    = (
        (actual_2026_vol - actual_2025_vol) / actual_2025_vol
        if actual_2025_vol > 0 else 0.0
    )
    scenario_trend = actual_2026_vol * (1 + growth_rate)

    # --- Scénario 2 : Budget ajusté (B27 corrigé par l'erreur historique B26) ---
    mape_factor         = (
        abs(actual_2026_vol - b26_vol) / actual_2026_vol
        if actual_2026_vol > 0 else 0.0
    )
    scenario_budget_adj = (
        b27_vol * (1 - mape_factor) if b27_vol > 0 else scenario_trend
    )

    # --- Scénario 3 : Moyenne pondérée (60% H27 + 40% B27) ---
    scenario_weighted = 0.6 * h27_vol + 0.4 * b27_vol

    # --- Scénario 4 : Recommandé (moyenne des 3 scénarios) ---
    scenario_recommended = (
        scenario_trend + scenario_budget_adj + scenario_weighted
    ) / 3

    # --- Écart Recommandé vs B27 ---
    gap_vs_b27 = scenario_recommended - b27_vol
    gap_pct    = (gap_vs_b27 / b27_vol * 100) if b27_vol > 0 else 0.0

    return {
        "actual_2026_vol"       : actual_2026_vol,
        "actual_2025_vol"       : actual_2025_vol,
        "b26_vol"               : b26_vol,
        "h27_vol"               : h27_vol,
        "b27_vol"               : b27_vol,
        "h28_vol"               : h28_vol,
        "growth_rate"           : growth_rate * 100,
        "mape_factor"           : mape_factor * 100,
        "scenario_trend"        : scenario_trend,
        "scenario_budget_adj"   : scenario_budget_adj,
        "scenario_weighted"     : scenario_weighted,
        "scenario_recommended"  : scenario_recommended,
        "gap_vs_b27"            : gap_vs_b27,
        "gap_pct"               : gap_pct,
    }


# ==============================================================================
# FONCTION 4 — CALCUL DU MAPE (Mean Absolute Percentage Error)
# ==============================================================================
@st.cache_data
def calculate_mape(df: pd.DataFrame) -> dict:
    """
    Calcule le MAPE entre le budget B26 et le réel 2026 (Q1+Q2+Q3)
    pour les 3 métriques : Volume, Spend, AV.
    Ajoute une note qualitative (🟢 Bon / 🟡 Correct / 🔴 À revoir).
    """
    if df.empty:
        return {}

    def _mape(actual: float, budget: float) -> float:
        """Calcule le MAPE en pourcentage."""
        return abs(actual - budget) / actual * 100 if actual > 0 else 0.0

    def _rating(mape_val: float) -> str:
        """Attribue une note qualitative selon le seuil MAPE."""
        if mape_val < 10:
            return "🟢 Bon"
        elif mape_val < 20:
            return "🟡 Correct"
        else:
            return "🔴 À revoir"

    # --- Volume : réel 2026 = Q1+Q2+Q3 (chaque trimestre pris séparément) ---
    actual_vol = (
        df["Volume Q1_26"].sum() +
        df["Volume Q2_26"].sum() +
        df["Volume Q3_26"].sum()
    )
    budget_vol  = df["Volume B26"].sum()
    mape_volume = _mape(actual_vol, budget_vol)

    # --- Spend ---
    actual_spend = (
        df["Spend Q1_26"].sum() +
        df["Spend Q2_26"].sum() +
        df["Spend Q3_26"].sum()
    )
    budget_spend = df["Spend B26"].sum()
    mape_spend   = _mape(actual_spend, budget_spend)

    # --- AV ---
    actual_av = (
        df["AV Q1_26"].sum() +
        df["AV Q2_26"].sum() +
        df["AV Q3_26"].sum()
    )
    budget_av = df["AV B26"].sum()
    mape_av   = _mape(actual_av, budget_av)

    # --- Moyenne globale ---
    average_mape = (mape_volume + mape_spend + mape_av) / 3

    return {
        "mape_volume"   : mape_volume,
        "mape_spend"    : mape_spend,
        "mape_av"       : mape_av,
        "average_mape"  : average_mape,
        "rating_volume" : _rating(mape_volume),
        "rating_spend"  : _rating(mape_spend),
        "rating_av"     : _rating(mape_av),
        "rating_avg"    : _rating(average_mape),
    }


# ==============================================================================
# FONCTION 5 — TOP 10 PRODUITS (classés par AV)
# ==============================================================================
@st.cache_data
def get_top10(
    df: pd.DataFrame,
    period_key: str = "B26",
    family_filter: tuple = None,
    cmo_filter: tuple = None
) -> pd.DataFrame:
    """
    Retourne le top 10 des produits par AV pour la période sélectionnée.
    L'AV est le driver principal de l'analyse (rentabilité).
    Applique un filtre par famille et/ou par CMO si fournis.
    Ajoute une colonne '% du total AV'.
    """
    if df.empty:
        return pd.DataFrame()

    # --- Filtre par famille ---
    if family_filter and len(family_filter) > 0:
        df = df[df[FAMILY_COL].isin(family_filter)].copy()

    # --- Filtre par CMO ---
    if cmo_filter and len(cmo_filter) > 0:
        df = df[df[CMO_COL].isin(cmo_filter)].copy()

    if df.empty:
        return pd.DataFrame()

    # --- Récupération des colonnes pour la période sélectionnée ---
    if period_key not in PERIOD_COL_MAP:
        period_key = "B26"
    cols = PERIOD_COL_MAP[period_key]

    # --- Calcul des métriques par produit pour la période ---
    df_top = df.copy()
    df_top["_av"]     = df_top[cols["av"]]
    df_top["_volume"] = df_top[cols["volume"]]
    df_top["_spend"]  = df_top[cols["spend"]]

    # --- Agrégation par produit ---
    df_agg = (
        df_top.groupby([PRODUCT_COL, FAMILY_COL], as_index=False)
        .agg(
            AV    =("_av",     "sum"),
            Volume=("_volume", "sum"),
            Spend =("_spend",  "sum"),
        )
    )

    # --- Tri par AV décroissant et top 10 ---
    df_agg = df_agg.sort_values("AV", ascending=False).head(10)

    # --- Pourcentage du total AV ---
    total_av = df_agg["AV"].sum()
    df_agg["pct_of_total"] = (
        df_agg["AV"] / total_av * 100
        if total_av > 0 else 0.0
    )

    return df_agg.reset_index(drop=True)


# ==============================================================================
# FONCTION 5b — TOP 10 CMO (classés par AV)
# ==============================================================================
@st.cache_data
def get_top10_cmo(
    df: pd.DataFrame,
    period_key: str = "B26",
    family_filter: tuple = None,
    cmo_filter: tuple = None
) -> pd.DataFrame:
    """
    Retourne le top 10 des CMO (Suppliers) par AV pour la période sélectionnée.
    L'AV est le driver principal de l'analyse (rentabilité).
    Applique un filtre par famille et/ou par CMO si fournis.
    Ajoute une colonne '% du total AV' et le nombre de produits distincts par CMO.
    """
    if df.empty:
        return pd.DataFrame()

    # --- Filtre par famille ---
    if family_filter and len(family_filter) > 0:
        df = df[df[FAMILY_COL].isin(family_filter)].copy()

    # --- Filtre par CMO ---
    if cmo_filter and len(cmo_filter) > 0:
        df = df[df[CMO_COL].isin(cmo_filter)].copy()

    if df.empty:
        return pd.DataFrame()

    # --- Récupération des colonnes pour la période sélectionnée ---
    if period_key not in PERIOD_COL_MAP:
        period_key = "B26"
    cols = PERIOD_COL_MAP[period_key]

    # --- Calcul des métriques par ligne pour la période ---
    df_cmo = df.copy()
    df_cmo["_av"]     = df_cmo[cols["av"]]
    df_cmo["_volume"] = df_cmo[cols["volume"]]
    df_cmo["_spend"]  = df_cmo[cols["spend"]]

    # --- Agrégation par CMO ---
    df_agg = (
        df_cmo.groupby(CMO_COL, as_index=False)
        .agg(
            AV         =("_av",       "sum"),
            Volume     =("_volume",   "sum"),
            Spend      =("_spend",    "sum"),
            Nb_Produits=(PRODUCT_COL, "nunique"),
        )
    )

    # --- Tri par AV décroissant et top 10 ---
    df_agg = df_agg.sort_values("AV", ascending=False).head(10)

    # --- Pourcentage du total AV ---
    total_av = df_agg["AV"].sum()
    df_agg["pct_of_total"] = (
        df_agg["AV"] / total_av * 100
        if total_av > 0 else 0.0
    )

    return df_agg.reset_index(drop=True)


# ==============================================================================
# FONCTION 6 — EXPORT EXCEL MULTI-ONGLETS
# ==============================================================================
def export_to_excel(
    df: pd.DataFrame,
    kpis: dict,
    scenarios: dict,
    mape_data: dict,
    top10: pd.DataFrame
) -> bytes:
    """
    Génère un fichier Excel avec 4 onglets :
      - KPIs période sélectionnée
      - Scénarios 2027
      - Top 10 Produits
      - Analyse MAPE
    Retourne les bytes du fichier pour le téléchargement Streamlit.
    """
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book

        # --- Formats ---
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#0066CC",
            "font_color": "white", "border": 1
        })
        fmt_number = workbook.add_format({"num_format": "#,##0.00", "border": 1})
        fmt_pct    = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_label  = workbook.add_format({"bold": True, "border": 1})

        # ---- Onglet 1 : KPIs période ----
        period_key   = kpis.get("period_key", "B26")
        ws1 = workbook.add_worksheet(f"KPIs {period_key}")
        ws1.set_column("A:A", 30)
        ws1.set_column("B:B", 20)
        rows_kpi = [
            ("Métrique", "Valeur"),
            (f"Volume {period_key} (k unités)", round(kpis.get("volume", 0) / 1000, 2)),
            (f"Spend {period_key} (k€)",        round(kpis.get("spend", 0) / 1000, 2)),
            (f"AV {period_key} (k€)",           round(kpis.get("av", 0) / 1000, 2)),
            ("Nb Produits distincts",            kpis.get("nb_produits", 0)),
        ]
        for r, (label, val) in enumerate(rows_kpi):
            ws1.write(r, 0, label, fmt_header if r == 0 else fmt_label)
            ws1.write(r, 1, val,   fmt_header if r == 0 else fmt_number)

        # ---- Onglet 2 : Scénarios 2027 ----
        ws2 = workbook.add_worksheet("Scénarios 2027")
        ws2.set_column("A:A", 30)
        ws2.set_column("B:C", 20)
        b27 = scenarios.get("b27_vol", 0)
        rows_sc = [
            ("Scénario", "Volume (unités)", "Écart vs B27 (%)"),
            ("Budget B27 (référence)",    b27,                                          0),
            ("Scénario Tendance",         scenarios.get("scenario_trend", 0),
             (scenarios.get("scenario_trend", 0) - b27) / b27 * 100 if b27 > 0 else 0),
            ("Scénario Budget Ajusté",    scenarios.get("scenario_budget_adj", 0),
             (scenarios.get("scenario_budget_adj", 0) - b27) / b27 * 100 if b27 > 0 else 0),
            ("Scénario Pondéré",          scenarios.get("scenario_weighted", 0),
             (scenarios.get("scenario_weighted", 0) - b27) / b27 * 100 if b27 > 0 else 0),
            ("⭐ Scénario Recommandé",    scenarios.get("scenario_recommended", 0),
             scenarios.get("gap_pct", 0)),
        ]
        for r, row in enumerate(rows_sc):
            for c, val in enumerate(row):
                ws2.write(r, c, val, fmt_header if r == 0 else (fmt_label if c == 0 else fmt_number))

        # ---- Onglet 3 : Top 10 Produits ----
        if not top10.empty:
            top10.to_excel(writer, sheet_name="Top 10 Produits", index=False)
            ws3 = writer.sheets["Top 10 Produits"]
            ws3.set_column("A:B", 25)
            ws3.set_column("C:E", 18)

        # ---- Onglet 4 : Analyse MAPE ----
        ws4 = workbook.add_worksheet("Analyse MAPE")
        ws4.set_column("A:A", 25)
        ws4.set_column("B:C", 20)
        rows_mape = [
            ("Métrique", "MAPE (%)", "Note"),
            ("Volume",   round(mape_data.get("mape_volume", 0), 2),  mape_data.get("rating_volume", "")),
            ("Spend",    round(mape_data.get("mape_spend", 0), 2),   mape_data.get("rating_spend", "")),
            ("AV",       round(mape_data.get("mape_av", 0), 2),      mape_data.get("rating_av", "")),
            ("Moyenne",  round(mape_data.get("average_mape", 0), 2), mape_data.get("rating_avg", "")),
        ]
        for r, (label, val, note) in enumerate(rows_mape):
            ws4.write(r, 0, label, fmt_header if r == 0 else fmt_label)
            ws4.write(r, 1, val,   fmt_header if r == 0 else fmt_number)
            ws4.write(r, 2, note,  fmt_header if r == 0 else fmt_label)

    return buffer.getvalue()


# ==============================================================================
# FONCTION UTILITAIRE — Formatage des nombres
# ==============================================================================
def fmt_k(val: float, decimals: int = 1) -> str:
    """Formate un nombre en milliers avec suffixe 'k'."""
    return f"{val / 1000:,.{decimals}f} k"


def fmt_pct(val: float, decimals: int = 1) -> str:
    """Formate un pourcentage."""
    sign = "+" if val > 0 else ""
    return f"{sign}{val:.{decimals}f}%"


# ==============================================================================
# PAGE 1 — VUE D'ENSEMBLE
# ==============================================================================
def page_overview(df: pd.DataFrame, mape_data: dict):
    """
    Affiche la page Vue d'ensemble — Performance APAC.
    L'utilisateur choisit UNE période via un sélecteur.
    Les KPIs et graphiques reflètent UNIQUEMENT cette période (pas de cumul).
    """

    st.markdown('<div class="main-header">📊 Vue d\'ensemble — Performance APAC</div>',
                unsafe_allow_html=True)
    st.markdown(
        f'<div class="sub-header">Données au {datetime.now().strftime("%d %B %Y")} '
        f'— Small Molecules EM&S Finance</div>',
        unsafe_allow_html=True
    )

    if df.empty:
        st.warning("⚠️ Aucune donnée disponible. Vérifiez le fichier source.")
        return

    # --------------------------------------------------------------------------
    # SÉLECTEUR DE PÉRIODE — une seule période à la fois, jamais de cumul
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🗓️ Sélection de la période</div>',
                unsafe_allow_html=True)

    selected_label = st.radio(
        "Choisir une période :",
        options=PERIOD_LABELS,
        index=PERIOD_LABELS.index("B26"),   # défaut : B26 (référence budget)
        horizontal=True,
        help="Sélectionne une période pour afficher ses indicateurs. Aucune période n'est cumulée avec une autre."
    )

    # Conversion label → clé interne
    period_key = PERIOD_KEY_MAP[selected_label]

    # Calcul des KPIs pour la période sélectionnée uniquement
    kpis = calculate_kpis_for_period(df, period_key)

    if not kpis:
        st.warning("⚠️ Impossible de calculer les indicateurs pour cette période.")
        return

    st.markdown("---")

    # --------------------------------------------------------------------------
    # LIGNE 1 — 4 cartes métriques (période sélectionnée uniquement)
    # --------------------------------------------------------------------------
    st.markdown(
        f'<div class="section-title">📌 Indicateurs — Période : {selected_label}</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            label=f"📦 Volume {selected_label}",
            value=fmt_k(kpis["volume"]) + " unités",
        )

    with col2:
        st.metric(
            label=f"💶 Spend {selected_label}",
            value=fmt_k(kpis["spend"]) + "€",
        )

    with col3:
        st.metric(
            label=f"⭐ AV {selected_label}",
            value=fmt_k(kpis["av"]) + "€",
        )

    with col4:
        st.metric(
            label="🏷️ Nb Produits",
            value=f"{kpis['nb_produits']}",
            delta="produits distincts"
        )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # GRAPHIQUES — Une valeur par période (jamais de cumul)
    # Chaque barre = somme de tous les produits filtrés pour CETTE période uniquement
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📈 Vue par période (toutes les périodes)</div>',
                unsafe_allow_html=True)

    # Calcul d'une valeur par période (somme des produits filtrés, pas de cumul)
    vol_by_period   = [df[c].sum() / 1000 for c in VOLUME_COLS]
    av_by_period    = [df[c].sum() / 1000 for c in AV_COLS]
    spend_by_period = [df[c].sum() / 1000 for c in SPEND_COLS]

    # Couleurs : période sélectionnée en bleu Sanofi, autres en gris clair
    bar_colors_vol = [
        COLOR_BLUE if lbl == selected_label else "#D0D8E4"
        for lbl in PERIOD_LABELS
    ]
    bar_colors_av = [
        COLOR_BLUE if lbl == selected_label else "#D0D8E4"
        for lbl in PERIOD_LABELS
    ]

    col_left, col_right = st.columns(2)

    # --- Graphique gauche : Volume par période ---
    with col_left:
        fig_vol = go.Figure(go.Bar(
            x=PERIOD_LABELS,
            y=vol_by_period,
            marker_color=bar_colors_vol,
            text=[f"{v:,.1f} k" for v in vol_by_period],
            textposition="outside",
        ))
        fig_vol.update_layout(
            title="Volume par Période",
            xaxis_title="Période",
            yaxis_title="Volume (k unités)",
            plot_bgcolor="white",
            paper_bgcolor="white",
            showlegend=False,
            margin=dict(l=40, r=20, t=60, b=40),
            height=360,
        )
        fig_vol.update_xaxes(showgrid=False)
        fig_vol.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
        st.plotly_chart(fig_vol, use_container_width=True)

    # --- Graphique droit : AV par période ---
    with col_right:
        fig_av = go.Figure(go.Bar(
            x=PERIOD_LABELS,
            y=av_by_period,
            marker_color=bar_colors_av,
            text=[f"{v:,.1f} k" for v in av_by_period],
            textposition="outside",
        ))
        fig_av.update_layout(
            title="AV par Période (k€)",
            xaxis_title="Période",
            yaxis_title="AV (k€)",
            plot_bgcolor="white",
            paper_bgcolor="white",
            showlegend=False,
            margin=dict(l=40, r=20, t=60, b=40),
            height=360,
        )
        fig_av.update_xaxes(showgrid=False)
        fig_av.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
        st.plotly_chart(fig_av, use_container_width=True)

    # --- Graphique pleine largeur : Spend par période (ligne continue/pointillée) ---
    st.markdown('<div class="section-title">💶 Spend par Période (k€)</div>',
                unsafe_allow_html=True)

    # Séparation réel (ligne pleine) vs budget/forecast (ligne pointillée)
    # On trace deux segments : réel et forecast, reliés au point de jonction
    actual_indices  = [i for i, k in enumerate(PERIOD_KEY_MAP.values()) if k in ACTUAL_PERIODS]
    forecast_indices = [i for i, k in enumerate(PERIOD_KEY_MAP.values()) if k in FORECAST_PERIODS]

    fig_spend = go.Figure()

    # Segment réel — ligne bleue pleine
    # Inclut A25 (index 0), Q1_26 (2), Q2_26 (3), Q3_26 (4)
    actual_x = [PERIOD_LABELS[i] for i in actual_indices]
    actual_y = [spend_by_period[i] for i in actual_indices]

    fig_spend.add_trace(go.Scatter(
        x=actual_x,
        y=actual_y,
        mode="lines+markers",
        name="Réel",
        line=dict(color=COLOR_BLUE, width=2.5),
        marker=dict(size=9, color=COLOR_BLUE),
    ))

    # Segment budget/forecast — ligne orange pointillée
    # Inclut B26 (1), H27 (5), B27 (6), H28 (7)
    forecast_x = [PERIOD_LABELS[i] for i in forecast_indices]
    forecast_y = [spend_by_period[i] for i in forecast_indices]

    fig_spend.add_trace(go.Scatter(
        x=forecast_x,
        y=forecast_y,
        mode="lines+markers",
        name="Budget / Forecast",
        line=dict(color=COLOR_ORANGE, width=2.5, dash="dash"),
        marker=dict(size=9, color=COLOR_ORANGE),
    ))

    # Mise en évidence de la période sélectionnée
    sel_idx = PERIOD_LABELS.index(selected_label)
    fig_spend.add_trace(go.Scatter(
        x=[selected_label],
        y=[spend_by_period[sel_idx]],
        mode="markers",
        name=f"Sélectionné : {selected_label}",
        marker=dict(size=14, color=COLOR_BLUE, symbol="star"),
    ))

    fig_spend.update_layout(
        title="Spend par Période (k€) — Réel vs Budget/Forecast",
        xaxis_title="Période",
        yaxis_title="Spend (k€)",
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        margin=dict(l=40, r=20, t=60, b=40),
        height=380,
    )
    fig_spend.update_xaxes(showgrid=True, gridcolor="#f0f0f0")
    fig_spend.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
    st.plotly_chart(fig_spend, use_container_width=True)

    st.markdown("---")

    # --------------------------------------------------------------------------
    # RÉSUMÉ MAPE (fiabilité du budget B26 vs réel 2026)
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🎯 Fiabilité du Budget 2026 (MAPE B26 vs Réel)</div>',
                unsafe_allow_html=True)

    if mape_data:
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f'<div class="metric-card">'
                f'<b>📦 MAPE Volume</b><br>'
                f'<span style="font-size:1.8rem;font-weight:700;color:{COLOR_BLUE}">'
                f'{mape_data["mape_volume"]:.1f}%</span><br>'
                f'{mape_data["rating_volume"]}</div>',
                unsafe_allow_html=True
            )
        with c2:
            st.markdown(
                f'<div class="metric-card">'
                f'<b>💶 MAPE Spend</b><br>'
                f'<span style="font-size:1.8rem;font-weight:700;color:{COLOR_BLUE}">'
                f'{mape_data["mape_spend"]:.1f}%</span><br>'
                f'{mape_data["rating_spend"]}</div>',
                unsafe_allow_html=True
            )
        with c3:
            st.markdown(
                f'<div class="metric-card">'
                f'<b>⭐ MAPE AV</b><br>'
                f'<span style="font-size:1.8rem;font-weight:700;color:{COLOR_BLUE}">'
                f'{mape_data["mape_av"]:.1f}%</span><br>'
                f'{mape_data["rating_av"]}</div>',
                unsafe_allow_html=True
            )


# ==============================================================================
# PAGE 2 — SCÉNARIOS 2027
# ==============================================================================
def page_scenarios(df: pd.DataFrame, scenarios: dict):
    """Affiche la page d'analyse des scénarios 2027."""

    st.markdown('<div class="main-header">📈 Analyse des Scénarios 2027</div>',
                unsafe_allow_html=True)

    if not scenarios:
        st.warning("⚠️ Impossible de calculer les scénarios. Vérifiez les données.")
        return

    # --- Explication des scénarios ---
    with st.expander("ℹ️ Comment lire cette page", expanded=False):
        st.markdown("""
        **Cette page compare 4 approches de forecast pour 2027 :**

        | Scénario | Logique |
        |---|---|
        | 🔵 **Tendance** | Extrapole le taux de croissance observé entre A25 et le réel 2026 |
        | 🟢 **Budget Ajusté** | Corrige B27 par le facteur d'erreur historique (écart B26 vs réel 2026) |
        | 🟠 **Pondéré** | Compromis entre H27 (60%) et B27 (40%) |
        | ⭐ **Recommandé** | Moyenne des 3 scénarios — notre meilleure estimation |

        **B27** est affiché comme ligne de référence (budget officiel 2027).
        L'écart entre le Recommandé et B27 indique si le budget doit être révisé.
        """)

    # --------------------------------------------------------------------------
    # LIGNE 1 — 4 cartes scénarios
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📊 Valeurs des scénarios (Volume 2027)</div>',
                unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    b27 = scenarios["b27_vol"]

    with col1:
        delta_trend = fmt_pct(
            (scenarios["scenario_trend"] - b27) / b27 * 100 if b27 > 0 else 0
        )
        st.metric(
            "🔵 Tendance",
            fmt_k(scenarios["scenario_trend"]) + " unités",
            delta=f"{delta_trend} vs B27"
        )

    with col2:
        delta_adj = fmt_pct(
            (scenarios["scenario_budget_adj"] - b27) / b27 * 100 if b27 > 0 else 0
        )
        st.metric(
            "🟢 Budget Ajusté",
            fmt_k(scenarios["scenario_budget_adj"]) + " unités",
            delta=f"{delta_adj} vs B27"
        )

    with col3:
        delta_w = fmt_pct(
            (scenarios["scenario_weighted"] - b27) / b27 * 100 if b27 > 0 else 0
        )
        st.metric(
            "🟠 Pondéré (60/40)",
            fmt_k(scenarios["scenario_weighted"]) + " unités",
            delta=f"{delta_w} vs B27"
        )

    with col4:
        st.markdown(
            f'<div class="metric-card-highlight">'
            f'<b>⭐ Recommandé</b><br>'
            f'<span style="font-size:1.5rem;font-weight:700;color:{COLOR_ORANGE}">'
            f'{fmt_k(scenarios["scenario_recommended"])} unités</span><br>'
            f'<span style="color:#666">{fmt_pct(scenarios["gap_pct"])} vs B27</span>'
            f'</div>',
            unsafe_allow_html=True
        )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # GRAPHIQUE — Comparaison des 4 scénarios vs B27
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📊 Comparaison des scénarios vs Budget B27</div>',
                unsafe_allow_html=True)

    scenario_names  = ["Tendance", "Budget Ajusté", "Pondéré", "Recommandé"]
    scenario_values = [
        scenarios["scenario_trend"],
        scenarios["scenario_budget_adj"],
        scenarios["scenario_weighted"],
        scenarios["scenario_recommended"],
    ]
    bar_colors = [COLOR_BLUE, COLOR_GREEN, COLOR_ORANGE, "#9B59B6"]

    fig_sc = go.Figure()

    # Barres des scénarios
    fig_sc.add_trace(go.Bar(
        x=scenario_names,
        y=[v / 1000 for v in scenario_values],
        marker_color=bar_colors,
        text=[f"{v/1000:,.1f} k" for v in scenario_values],
        textposition="outside",
        name="Scénarios",
    ))

    # Ligne de référence B27
    fig_sc.add_hline(
        y=b27 / 1000,
        line_dash="dash",
        line_color="red",
        line_width=2,
        annotation_text=f"B27 = {b27/1000:,.1f} k",
        annotation_position="top right",
        annotation_font_color="red",
    )

    fig_sc.update_layout(
        title="Scénarios de Volume 2027 vs Budget Officiel (B27)",
        xaxis_title="Scénario",
        yaxis_title="Volume (k unités)",
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        margin=dict(l=40, r=20, t=60, b=40),
        height=420,
    )
    fig_sc.update_xaxes(showgrid=False)
    fig_sc.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
    st.plotly_chart(fig_sc, use_container_width=True)

    # --------------------------------------------------------------------------
    # ANALYSE DE L'ÉCART
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🔍 Analyse de l\'écart — Recommandé vs B27</div>',
                unsafe_allow_html=True)

    gap_pct = scenarios["gap_pct"]
    gap_abs = scenarios["gap_vs_b27"]

    # Couleur selon l'amplitude de l'écart
    if abs(gap_pct) < 5:
        gap_color = COLOR_GREEN
        gap_msg   = "✅ Écart faible — le budget B27 est cohérent avec notre estimation."
    elif abs(gap_pct) < 15:
        gap_color = COLOR_ORANGE
        gap_msg   = "⚠️ Écart modéré — une révision partielle de B27 est recommandée."
    else:
        gap_color = "#DC3545"
        gap_msg   = "🔴 Écart significatif — B27 devrait être revu en profondeur."

    col_gap1, col_gap2 = st.columns(2)
    with col_gap1:
        st.markdown(
            f'<div class="metric-card" style="border-left-color:{gap_color}">'
            f'<b>Écart absolu (Recommandé - B27)</b><br>'
            f'<span style="font-size:1.6rem;font-weight:700;color:{gap_color}">'
            f'{fmt_k(gap_abs)} unités</span>'
            f'</div>',
            unsafe_allow_html=True
        )
    with col_gap2:
        st.markdown(
            f'<div class="metric-card" style="border-left-color:{gap_color}">'
            f'<b>Écart en pourcentage</b><br>'
            f'<span style="font-size:1.6rem;font-weight:700;color:{gap_color}">'
            f'{fmt_pct(gap_pct)}</span>'
            f'</div>',
            unsafe_allow_html=True
        )

    st.info(gap_msg)

    # --------------------------------------------------------------------------
    # TABLEAU RÉCAPITULATIF DES SCÉNARIOS
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📋 Tableau récapitulatif</div>',
                unsafe_allow_html=True)

    df_table = pd.DataFrame({
        "Scénario"          : ["Budget B27 (réf.)", "Tendance", "Budget Ajusté", "Pondéré", "⭐ Recommandé"],
        "Volume (unités)"   : [b27] + scenario_values,
        "Volume (k unités)" : [round(v / 1000, 1) for v in [b27] + scenario_values],
        "Écart vs B27 (%)"  : [0.0] + [
            round((v - b27) / b27 * 100, 1) if b27 > 0 else 0.0
            for v in scenario_values
        ],
    })
    st.dataframe(df_table, use_container_width=True, hide_index=True)


# ==============================================================================
# PAGE 3 — TOP PRODUITS (classés par AV)
# ==============================================================================
def page_top_products(df: pd.DataFrame, family_filter: list, cmo_filter: list):
    """
    Affiche la page Top 10 Produits par AV pour la période sélectionnée.
    L'AV est le driver principal (rentabilité), le volume vient en second.
    """

    st.markdown('<div class="main-header">🏆 Top 10 Produits — AV par Période</div>',
                unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # SÉLECTEUR DE PÉRIODE
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🗓️ Sélection de la période</div>',
                unsafe_allow_html=True)

    selected_label = st.radio(
        "Choisir une période :",
        options=PERIOD_LABELS,
        index=PERIOD_LABELS.index("B26"),
        horizontal=True,
        key="top_products_period",
        help="Le classement est basé sur l'AV de la période sélectionnée."
    )
    period_key = PERIOD_KEY_MAP[selected_label]

    st.markdown("---")

    # --- Calcul du top 10 avec filtres famille, CMO et période ---
    top10 = get_top10(
        df,
        period_key=period_key,
        family_filter=tuple(family_filter) if family_filter else None,
        cmo_filter=tuple(cmo_filter) if cmo_filter else None,
    )

    if top10.empty:
        st.warning("⚠️ Aucun produit trouvé avec les filtres sélectionnés.")
        return

    # --------------------------------------------------------------------------
    # GRAPHIQUE — Top 10 horizontal bar chart (axe X = AV)
    # --------------------------------------------------------------------------
    st.markdown(
        f'<div class="section-title">📊 Top 10 par AV — Période : {selected_label}</div>',
        unsafe_allow_html=True
    )

    # Palette de couleurs par famille
    families      = top10[FAMILY_COL].unique().tolist()
    palette       = [COLOR_BLUE, COLOR_ORANGE, COLOR_GREEN, "#9B59B6",
                     "#E74C3C", "#1ABC9C", "#F39C12", "#2ECC71", "#3498DB", "#E67E22"]
    family_colors = {fam: palette[i % len(palette)] for i, fam in enumerate(families)}
    bar_colors    = [family_colors[f] for f in top10[FAMILY_COL]]

    fig_top = go.Figure(go.Bar(
        x=top10["AV"] / 1000,
        y=top10[PRODUCT_COL],
        orientation="h",
        marker_color=bar_colors,
        text=[f"{v/1000:,.1f} k€ ({p:.1f}%)"
              for v, p in zip(top10["AV"], top10["pct_of_total"])],
        textposition="outside",
    ))

    fig_top.update_layout(
        title=f"Top 10 Produits par AV — {selected_label}",
        xaxis_title="AV (k€)",
        yaxis_title="",
        yaxis=dict(autorange="reversed"),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=200, r=100, t=60, b=40),
        height=420,
    )
    fig_top.update_xaxes(showgrid=True, gridcolor="#f0f0f0")
    st.plotly_chart(fig_top, use_container_width=True)

    # --------------------------------------------------------------------------
    # TABLEAU INTERACTIF
    # Ordre des colonnes : Produit, Famille, AV, Volume, Spend, % du total AV
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📋 Détail des 10 premiers produits</div>',
                unsafe_allow_html=True)

    df_display = top10[[PRODUCT_COL, FAMILY_COL, "AV", "Volume", "Spend", "pct_of_total"]].rename(columns={
        PRODUCT_COL    : "Produit",
        FAMILY_COL     : "Famille",
        "AV"           : f"AV {selected_label} (€)",
        "Volume"       : f"Volume {selected_label} (unités)",
        "Spend"        : f"Spend {selected_label} (€)",
        "pct_of_total" : "% du total AV",
    })
    df_display["% du total AV"] = df_display["% du total AV"].round(1)

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            f"AV {selected_label} (€)"            : st.column_config.NumberColumn(format="%,.0f"),
            f"Volume {selected_label} (unités)"   : st.column_config.NumberColumn(format="%,.0f"),
            f"Spend {selected_label} (€)"         : st.column_config.NumberColumn(format="%,.0f"),
            "% du total AV"                        : st.column_config.NumberColumn(format="%.1f%%"),
        }
    )

    # --------------------------------------------------------------------------
    # GRAPHIQUE — Concentration par famille (pie chart sur AV)
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🥧 Concentration AV par Famille</div>',
                unsafe_allow_html=True)

    # Agrégation par famille sur tout le dataframe filtré pour la période
    df_fam = df.copy()
    if family_filter:
        df_fam = df_fam[df_fam[FAMILY_COL].isin(family_filter)]
    if cmo_filter:
        df_fam = df_fam[df_fam[CMO_COL].isin(cmo_filter)]

    cols = PERIOD_COL_MAP[period_key]
    df_fam["_av"] = df_fam[cols["av"]]
    fam_agg = df_fam.groupby(FAMILY_COL)["_av"].sum().reset_index()
    fam_agg = fam_agg.sort_values("_av", ascending=False)

    fig_pie = go.Figure(go.Pie(
        labels=fam_agg[FAMILY_COL],
        values=fam_agg["_av"],
        hole=0.4,
        marker_colors=palette[:len(fam_agg)],
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>AV: %{value:,.0f} €<br>Part: %{percent}<extra></extra>",
    ))
    fig_pie.update_layout(
        title=f"Distribution de l'AV {selected_label} par Famille",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=20),
        height=380,
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    # --------------------------------------------------------------------------
    # EXPORT CSV
    # --------------------------------------------------------------------------
    st.markdown("---")
    csv_data = top10.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
    st.download_button(
        label="⬇️ Télécharger le Top 10 (CSV)",
        data=csv_data,
        file_name=f"top10_produits_av_{selected_label}_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )


# ==============================================================================
# PAGE 4 — ANALYSE CMO (SUPPLIER) (classée par AV)
# ==============================================================================
def page_cmo_analysis(df: pd.DataFrame, family_filter: list, cmo_filter: list):
    """
    Affiche la page d'analyse par CMO (Supplier).
    Classement par AV pour la période sélectionnée.
    """

    st.markdown(
        '<div class="main-header">🏭 Analyse par CMO (Supplier) — AV par Période</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        f'<div class="sub-header">Données au {datetime.now().strftime("%d %B %Y")} '
        f'— Small Molecules EM&S Finance</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------------------------
    # SÉLECTEUR DE PÉRIODE
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🗓️ Sélection de la période</div>',
                unsafe_allow_html=True)

    selected_label = st.radio(
        "Choisir une période :",
        options=PERIOD_LABELS,
        index=PERIOD_LABELS.index("B26"),
        horizontal=True,
        key="cmo_period",
        help="Le classement est basé sur l'AV de la période sélectionnée."
    )
    period_key = PERIOD_KEY_MAP[selected_label]

    st.markdown("---")

    # --- Calcul du top 10 CMO avec filtres ---
    top10_cmo = get_top10_cmo(
        df,
        period_key=period_key,
        family_filter=tuple(family_filter) if family_filter else None,
        cmo_filter=tuple(cmo_filter) if cmo_filter else None,
    )

    if top10_cmo.empty:
        st.warning("⚠️ Aucun CMO trouvé avec les filtres sélectionnés.")
        return

    # --------------------------------------------------------------------------
    # CARTES MÉTRIQUES RÉSUMÉ
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📌 Résumé CMO — Période : {}</div>'.format(selected_label),
                unsafe_allow_html=True)

    # Calcul des totaux sur le périmètre filtré pour la période sélectionnée
    df_scope = df.copy()
    if family_filter:
        df_scope = df_scope[df_scope[FAMILY_COL].isin(family_filter)]
    if cmo_filter:
        df_scope = df_scope[df_scope[CMO_COL].isin(cmo_filter)]

    cols = PERIOD_COL_MAP[period_key]
    total_av_scope  = df_scope[cols["av"]].sum()
    nb_cmo_total    = df_scope[CMO_COL].nunique()
    nb_produits     = df_scope[PRODUCT_COL].nunique()

    # AV couverte par le top 10 CMO
    av_top10  = top10_cmo["AV"].sum()
    pct_top10 = av_top10 / total_av_scope * 100 if total_av_scope > 0 else 0.0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🏭 Nombre de CMO", f"{nb_cmo_total}")
    with col2:
        st.metric("📦 Produits distincts", f"{nb_produits}")
    with col3:
        st.metric(f"⭐ AV totale {selected_label}", fmt_k(total_av_scope) + "€")
    with col4:
        st.metric(
            "🔝 Concentration Top 10",
            f"{pct_top10:.1f}%",
            delta="de l'AV totale"
        )

    st.markdown("---")

    # --------------------------------------------------------------------------
    # GRAPHIQUE — Top 10 CMO horizontal bar chart (axe X = AV)
    # --------------------------------------------------------------------------
    st.markdown(
        f'<div class="section-title">📊 Top 10 CMO par AV — Période : {selected_label}</div>',
        unsafe_allow_html=True
    )

    palette    = [COLOR_BLUE, COLOR_ORANGE, COLOR_GREEN, "#9B59B6",
                  "#E74C3C", "#1ABC9C", "#F39C12", "#2ECC71", "#3498DB", "#E67E22"]
    bar_colors = [palette[i % len(palette)] for i in range(len(top10_cmo))]

    fig_cmo = go.Figure(go.Bar(
        x=top10_cmo["AV"] / 1000,
        y=top10_cmo[CMO_COL],
        orientation="h",
        marker_color=bar_colors,
        text=[
            f"{v/1000:,.1f} k€ ({p:.1f}%) — {n} produit(s)"
            for v, p, n in zip(
                top10_cmo["AV"],
                top10_cmo["pct_of_total"],
                top10_cmo["Nb_Produits"]
            )
        ],
        textposition="outside",
    ))

    fig_cmo.update_layout(
        title=f"Top 10 CMO (Suppliers) par AV — {selected_label}",
        xaxis_title="AV (k€)",
        yaxis_title="",
        yaxis=dict(autorange="reversed"),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=220, r=120, t=60, b=40),
        height=420,
    )
    fig_cmo.update_xaxes(showgrid=True, gridcolor="#f0f0f0")
    st.plotly_chart(fig_cmo, use_container_width=True)

    # --------------------------------------------------------------------------
    # TABLEAU INTERACTIF
    # Ordre des colonnes : CMO, AV, Volume, Spend, % du total AV, Nb Produits
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📋 Détail des 10 premiers CMO</div>',
                unsafe_allow_html=True)

    df_display = top10_cmo[[CMO_COL, "AV", "Volume", "Spend", "pct_of_total", "Nb_Produits"]].rename(columns={
        CMO_COL        : "CMO (Supplier)",
        "AV"           : f"AV {selected_label} (€)",
        "Volume"       : f"Volume {selected_label} (unités)",
        "Spend"        : f"Spend {selected_label} (€)",
        "pct_of_total" : "% du total AV",
        "Nb_Produits"  : "Nb Produits",
    })
    df_display["% du total AV"] = df_display["% du total AV"].round(1)

    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            f"AV {selected_label} (€)"          : st.column_config.NumberColumn(format="%,.0f"),
            f"Volume {selected_label} (unités)" : st.column_config.NumberColumn(format="%,.0f"),
            f"Spend {selected_label} (€)"       : st.column_config.NumberColumn(format="%,.0f"),
            "% du total AV"                      : st.column_config.NumberColumn(format="%.1f%%"),
            "Nb Produits"                        : st.column_config.NumberColumn(format="%d"),
        }
    )

    # --------------------------------------------------------------------------
    # GRAPHIQUE — Concentration par CMO (pie chart sur AV)
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">🥧 Distribution de l\'AV par CMO</div>',
                unsafe_allow_html=True)

    # Agrégation complète par CMO sur le périmètre filtré
    df_scope["_av"] = df_scope[cols["av"]]
    cmo_agg = df_scope.groupby(CMO_COL)["_av"].sum().reset_index()
    cmo_agg = cmo_agg.sort_values("_av", ascending=False)

    # Regrouper les CMO hors top 10 dans "Autres" pour lisibilité
    if len(cmo_agg) > 10:
        top_cmo_names = cmo_agg.head(10)[CMO_COL].tolist()
        autres_av     = cmo_agg[~cmo_agg[CMO_COL].isin(top_cmo_names)]["_av"].sum()
        cmo_pie       = cmo_agg[cmo_agg[CMO_COL].isin(top_cmo_names)].copy()
        if autres_av > 0:
            cmo_pie = pd.concat([
                cmo_pie,
                pd.DataFrame({CMO_COL: ["Autres"], "_av": [autres_av]})
            ], ignore_index=True)
    else:
        cmo_pie = cmo_agg.copy()

    fig_pie_cmo = go.Figure(go.Pie(
        labels=cmo_pie[CMO_COL],
        values=cmo_pie["_av"],
        hole=0.4,
        marker_colors=palette[:len(cmo_pie)],
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>AV: %{value:,.0f} €<br>Part: %{percent}<extra></extra>",
    ))
    fig_pie_cmo.update_layout(
        title=f"Distribution de l'AV {selected_label} par CMO (Supplier)",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=20),
        height=400,
    )
    st.plotly_chart(fig_pie_cmo, use_container_width=True)

    # --------------------------------------------------------------------------
    # EXPORT CSV
    # --------------------------------------------------------------------------
    st.markdown("---")
    csv_data = top10_cmo.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
    st.download_button(
        label="⬇️ Télécharger l'analyse CMO (CSV)",
        data=csv_data,
        file_name=f"top10_cmo_av_{selected_label}_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )


# ==============================================================================
# PAGE 5 — QUALITÉ DES DONNÉES
# ==============================================================================
def page_data_quality(df: pd.DataFrame, mape_data: dict, kpis: dict, scenarios: dict, top10):
    """Affiche la page Qualité & Fiabilité des Données."""

    st.markdown('<div class="main-header">🔍 Qualité & Fiabilité des Données</div>',
                unsafe_allow_html=True)

    if not mape_data:
        st.warning("⚠️ Impossible de calculer les métriques de qualité.")
        return

    # --------------------------------------------------------------------------
    # TABLEAU MAPE
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📊 Analyse MAPE — Budget B26 vs Réel 2026</div>',
                unsafe_allow_html=True)

    df_mape = pd.DataFrame({
        "Métrique"  : ["Volume", "Spend", "AV", "Moyenne globale"],
        "MAPE (%)"  : [
            round(mape_data["mape_volume"], 2),
            round(mape_data["mape_spend"],  2),
            round(mape_data["mape_av"],     2),
            round(mape_data["average_mape"],2),
        ],
        "Note"      : [
            mape_data["rating_volume"],
            mape_data["rating_spend"],
            mape_data["rating_av"],
            mape_data["rating_avg"],
        ],
    })
    st.dataframe(df_mape, use_container_width=True, hide_index=True)

    # --------------------------------------------------------------------------
    # GRAPHIQUE MAPE — Barres avec lignes de référence
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📈 MAPE par Métrique</div>',
                unsafe_allow_html=True)

    metrics     = ["Volume", "Spend", "AV"]
    mape_values = [
        mape_data["mape_volume"],
        mape_data["mape_spend"],
        mape_data["mape_av"],
    ]
    bar_cols = [
        COLOR_GREEN  if v < 10 else
        COLOR_ORANGE if v < 20 else
        "#DC3545"
        for v in mape_values
    ]

    fig_mape = go.Figure()

    fig_mape.add_trace(go.Bar(
        x=metrics,
        y=mape_values,
        marker_color=bar_cols,
        text=[f"{v:.1f}%" for v in mape_values],
        textposition="outside",
        name="MAPE",
    ))

    # Ligne de référence 10% (seuil Bon)
    fig_mape.add_hline(
        y=10, line_dash="dot", line_color=COLOR_GREEN, line_width=1.5,
        annotation_text="Seuil Bon (10%)",
        annotation_position="top right",
        annotation_font_color=COLOR_GREEN,
    )
    # Ligne de référence 20% (seuil À revoir)
    fig_mape.add_hline(
        y=20, line_dash="dot", line_color="#DC3545", line_width=1.5,
        annotation_text="Seuil À revoir (20%)",
        annotation_position="top right",
        annotation_font_color="#DC3545",
    )

    fig_mape.update_layout(
        title="MAPE par Métrique — Budget B26 vs Réel 2026",
        xaxis_title="Métrique",
        yaxis_title="MAPE (%)",
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False,
        margin=dict(l=40, r=20, t=60, b=40),
        height=380,
    )
    fig_mape.update_xaxes(showgrid=False)
    fig_mape.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
    st.plotly_chart(fig_mape, use_container_width=True)

    # --------------------------------------------------------------------------
    # COMPLÉTUDE DES DONNÉES
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">✅ Complétude des données par colonne</div>',
                unsafe_allow_html=True)

    all_cols    = VOLUME_COLS + SPEND_COLS + AV_COLS
    completeness_data = []
    for col in all_cols:
        if col in df.columns:
            non_zero = (df[col] != 0).sum()
            pct      = non_zero / len(df) * 100 if len(df) > 0 else 0
            completeness_data.append({
                "Colonne"           : col,
                "Valeurs non-nulles": non_zero,
                "Total lignes"      : len(df),
                "Complétude (%)"    : round(pct, 1),
            })

    df_comp = pd.DataFrame(completeness_data)
    st.dataframe(
        df_comp,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Complétude (%)": st.column_config.ProgressColumn(
                format="%.1f%%", min_value=0, max_value=100
            )
        }
    )

    # --------------------------------------------------------------------------
    # GUIDE D'INTERPRÉTATION
    # --------------------------------------------------------------------------
    st.markdown('<div class="section-title">📖 Guide d\'interprétation</div>',
                unsafe_allow_html=True)

    st.info(
        "**🟢 MAPE < 10% — Bon**\n\n"
        "Le budget est fiable. Les prévisions B27 peuvent être utilisées avec confiance "
        "sans correction majeure. L'écart historique est dans la marge d'erreur normale."
    )
    st.warning(
        "**🟡 MAPE entre 10% et 20% — Correct**\n\n"
        "Écart notable mais gérable. Il est recommandé d'appliquer une correction partielle "
        "au budget (scénario Budget Ajusté) et de surveiller les produits les plus volatils."
    )
    st.error(
        "**🔴 MAPE > 20% — À revoir**\n\n"
        "Le processus budgétaire présente un problème structurel. Le budget B27 ne devrait "
        "pas être utilisé tel quel. Privilégier le scénario Tendance ou Recommandé, "
        "et investiguer les causes racines des écarts (hypothèses de marché, données sources)."
    )

    # --------------------------------------------------------------------------
    # EXPORT COMPLET
    # --------------------------------------------------------------------------
    st.markdown("---")
    st.markdown('<div class="section-title">⬇️ Export de l\'analyse complète</div>',
                unsafe_allow_html=True)

    excel_bytes = export_to_excel(df, kpis, scenarios, mape_data, top10)
    st.download_button(
        label="📥 Télécharger l'analyse complète (Excel)",
        data=excel_bytes,
        file_name=f"forecast_apac_analyse_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ==============================================================================
# FONCTION PRINCIPALE
# ==============================================================================
def main():
    """Point d'entrée de l'application Streamlit."""

    # --------------------------------------------------------------------------
    # CHARGEMENT DES DONNÉES
    # --------------------------------------------------------------------------
    with st.spinner("⏳ Chargement des données APAC..."):
        df_raw = load_data(FILEPATH, SHEET)

    if df_raw.empty:
        st.stop()

    # --------------------------------------------------------------------------
    # SIDEBAR — Navigation et filtres
    # --------------------------------------------------------------------------
    st.sidebar.markdown(
        '<div class="sidebar-title">📊 Forecast Support Tool</div>',
        unsafe_allow_html=True
    )
    st.sidebar.markdown("**APAC — Small Molecules EM&S**")
    st.sidebar.markdown("---")

    # --- Navigation ---
    page = st.sidebar.radio(
        "Navigation",
        options=[
            "🏠 Vue d'ensemble",
            "📈 Scénarios 2027",
            "🏆 Top Produits",
            "🏭 Analyse CMO",
            "🔍 Qualité des données",
        ],
        label_visibility="collapsed"
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("**🔽 Filtres**")

    # --- Filtre par famille (Product Family) ---
    all_families  = sorted(df_raw[FAMILY_COL].dropna().unique().tolist())
    family_filter = st.sidebar.multiselect(
        "Famille (Product Family)",
        options=all_families,
        default=all_families,
        help="Sélectionne une ou plusieurs familles de produits"
    )

    # --- Filtre par CMO (Supplier) ---
    all_cmos   = sorted(df_raw[CMO_COL].dropna().unique().tolist())
    cmo_filter = st.sidebar.multiselect(
        "CMO (Supplier)",
        options=all_cmos,
        default=all_cmos,
        help="Sélectionne un ou plusieurs CMO (fournisseurs)"
    )

    # --- Application des deux filtres ---
    df_filtered = df_raw.copy()
    if family_filter:
        df_filtered = df_filtered[df_filtered[FAMILY_COL].isin(family_filter)]
    if cmo_filter:
        df_filtered = df_filtered[df_filtered[CMO_COL].isin(cmo_filter)]

    st.sidebar.markdown("---")

    # --- Informations sur les données ---
    st.sidebar.markdown("**ℹ️ Informations**")
    st.sidebar.markdown(f"📦 **{len(df_filtered)}** lignes sélectionnées")
    st.sidebar.markdown(f"🏭 **{df_filtered[CMO_COL].nunique()}** CMO actifs")
    st.sidebar.markdown(f"📅 Mise à jour : {datetime.now().strftime('%d/%m/%Y')}")

    # --------------------------------------------------------------------------
    # CALCUL DES MÉTRIQUES (sur données filtrées)
    # Les KPIs de la vue d'ensemble sont calculés à la volée selon la période choisie
    # --------------------------------------------------------------------------
    scenarios = calculate_scenarios(df_filtered)
    mape_data = calculate_mape(df_filtered)

    # KPIs pour l'export : on utilise B26 comme référence par défaut
    kpis_export = calculate_kpis_for_period(df_filtered, "B26")

    top10 = get_top10(
        df_filtered,
        period_key="B26",
        family_filter=tuple(family_filter) if family_filter else None,
        cmo_filter=tuple(cmo_filter) if cmo_filter else None,
    )

    # --------------------------------------------------------------------------
    # ROUTAGE DES PAGES
    # --------------------------------------------------------------------------
    if page == "🏠 Vue d'ensemble":
        page_overview(df_filtered, mape_data)

    elif page == "📈 Scénarios 2027":
        page_scenarios(df_filtered, scenarios)

    elif page == "🏆 Top Produits":
        page_top_products(df_filtered, family_filter, cmo_filter)

    elif page == "🏭 Analyse CMO":
        page_cmo_analysis(df_filtered, family_filter, cmo_filter)

    elif page == "🔍 Qualité des données":
        page_data_quality(df_filtered, mape_data, kpis_export, scenarios, top10)

    # --------------------------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------------------------
    st.markdown(
        '<div class="footer">'
        'Forecast Support Tool v1.0 — EM&S Finance APAC | '
        'Développé par Salma Kamoun'
        '</div>',
        unsafe_allow_html=True
    )


# ==============================================================================
# POINT D'ENTRÉE
# ==============================================================================
if __name__ == "__main__":
    main()
