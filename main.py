import streamlit as st
import math
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta

st.set_page_config(
    page_title="MedOil Supply Chain",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# PALETTE DE COULEURS MEDOIL (basée sur medoil.com.tn)
# ============================================================
COLORS = {
    "primary": "#003366",      # Bleu foncé MedOil
    "secondary": "#FF8C00",    # Orange MedOil
    "accent": "#00A3E0",       # Bleu clair
    "success": "#28A745",      # Vert
    "warning": "#FFC107",      # Jaune
    "danger": "#DC3545",       # Rouge
    "dark": "#1A1A2E",         # Fond sombre
    "light": "#F8F9FA",        # Fond clair
    "gray": "#6C757D",
    "white": "#FFFFFF",
    "sidebar_bg": "#0D2137",   # Fond sidebar
}

# ============================================================
# CONFIGURATION DES DÉLAIS PAR SOURCE
# ============================================================
SOURCE_DELAYS = {
    "Export": {"mois": 4, "jours": 120, "label": "Export (4 mois)"},
    "Local": {"mois": 0.5, "jours": 15, "label": "Local (2 semaines)"},
    "BM": {"mois": 0.75, "jours": 21, "label": "BM (3 semaines)"}
}

# ============================================================
# STYLES CSS POUR L'INTERFACE MEDOIL
# ============================================================
st.markdown(f"""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Base */
    html, body, [class*="css"] {{
        font-family: 'Inter', sans-serif;
    }}
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {COLORS["sidebar_bg"]} 0%, {COLORS["primary"]} 100%);
    }}
    
    [data-testid="stSidebar"] * {{
        color: {COLORS["white"]} !important;
    }}
    
    [data-testid="stSidebar"] .stMarkdown {{
        color: {COLORS["white"]};
    }}
    
    /* Bouton de navigation dans sidebar */
    [data-testid="stSidebar"] button {{
        background: rgba(255,255,255,0.1);
        border-radius: 10px;
        margin: 5px 0;
        transition: all 0.3s ease;
    }}
    
    [data-testid="stSidebar"] button:hover {{
        background: {COLORS["secondary"]};
        transform: translateX(5px);
    }}
    
    /* Cards */
    .medoil-card {{
        background: {COLORS["white"]};
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        border-top: 4px solid {COLORS["secondary"]};
        margin-bottom: 1rem;
    }}
    
    .medoil-card-dark {{
        background: linear-gradient(135deg, {COLORS["primary"]}, {COLORS["sidebar_bg"]});
        border-radius: 16px;
        padding: 1.5rem;
        color: {COLORS["white"]};
        margin-bottom: 1rem;
    }}
    
    /* Metrics */
    .metric-value {{
        font-size: 2rem;
        font-weight: 700;
        color: {COLORS["primary"]};
    }}
    
    .metric-label {{
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: {COLORS["gray"]};
    }}
    
    /* Alertes */
    .alert-critical {{
        background: linear-gradient(135deg, #FFF5F5, #FED7D7);
        border-left: 4px solid {COLORS["danger"]};
        border-radius: 12px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.75rem;
        color: #C53030;
    }}
    
    .alert-warning {{
        background: linear-gradient(135deg, #FFFFF0, #FEFCBF);
        border-left: 4px solid {COLORS["warning"]};
        border-radius: 12px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.75rem;
        color: #975A16;
    }}
    
    .alert-success {{
        background: linear-gradient(135deg, #F0FFF4, #C6F6D5);
        border-left: 4px solid {COLORS["success"]};
        border-radius: 12px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.75rem;
        color: #22543D;
    }}
    
    /* Boutons */
    .stButton > button {{
        background: {COLORS["secondary"]};
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
        transition: all 0.3s ease;
    }}
    
    .stButton > button:hover {{
        background: {COLORS["primary"]};
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }}
    
    /* Dataframe */
    [data-testid="stDataframe"] {{
        border-radius: 12px;
        overflow: hidden;
    }}
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 2rem;
    }}
    
    .stTabs [data-baseweb="tab"] {{
        background: transparent;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }}
    
    .stTabs [aria-selected="true"] {{
        background: {COLORS["secondary"]};
        color: white;
    }}
    
    /* Header */
    .medoil-header {{
        background: linear-gradient(135deg, {COLORS["primary"]}, {COLORS["sidebar_bg"]});
        border-radius: 20px;
        padding: 2rem;
        margin-bottom: 2rem;
        color: white;
    }}
    
    h1, h2, h3 {{
        color: {COLORS["primary"]};
    }}
    
    hr {{
        border-color: {COLORS["secondary"]};
        border-width: 2px;
    }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

def fmt(n, d=1):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))):
        return "—"
    return f"{n:,.{d}f}"

def fmtInt(n):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))):
        return "—"
    return f"{round(n):,}"

def clean_num(v):
    """Nettoie une valeur numérique (gère virgules, espaces, #DIV/0!)"""
    if v is None:
        return 0.0
    s = str(v).strip().replace(" ", "").replace(",", ".")
    if s in ("#DIV/0!", "#N/A", "#VALEUR!", "#REF!", "", "-", "—"):
        return 0.0
    try:
        return float(s)
    except:
        return 0.0

def calc_ss(conso_mensuelle, z=1.65, lt_mois=1.0, variab=0.15):
    """Calcule le stock de sécurité avec variabilité"""
    sigma_d = conso_mensuelle * variab
    lt_j = lt_mois * 30
    d_j = conso_mensuelle / 30
    return round(z * math.sqrt(lt_j * sigma_d**2 + d_j**2 * (lt_j * 0.05)**2))

def calc_eoq(conso_an, cout_unit, sc_cost=150, taux=0.20):
    """Calcule la quantité économique (Wilson)"""
    if cout_unit <= 0 or conso_an <= 0:
        return 0, 0, 0
    h = taux * cout_unit
    eoq = math.sqrt(2 * conso_an * sc_cost / h)
    nb = conso_an / eoq
    ct = nb * sc_cost + (eoq / 2) * h
    return round(eoq), round(nb, 1), round(ct, 2)

def calc_pr(conso_mois, lt_mois, ss):
    """Calcule le point de commande"""
    return round((conso_mois / 30) * (lt_mois * 30) + ss)

def get_delay_from_source(source):
    """Retourne le délai en fonction de la source"""
    return SOURCE_DELAYS.get(source, SOURCE_DELAYS["Local"])

def calculate_auto_stats(consumptions):
    """Calcule automatiquement la moyenne et l'écart-type des consommations"""
    if len(consumptions) == 0:
        return 0, 0
    mean = np.mean(consumptions)
    std = np.std(consumptions, ddof=1) if len(consumptions) > 1 else 0
    return mean, std

# ============================================================
# FONCTIONS D'EXPORT EXCEL AVEC 2 FEUILLES
# ============================================================

def build_excel_complete(df_calculs, df_evolution=None):
    """Crée un fichier Excel avec les résultats des calculs et l'évolution"""
    wb = Workbook()
    
    # Feuille 1: Résultats des calculs
    ws1 = wb.active
    ws1.title = "Calculs SC"
    
    thin = Side(style="thin", color="BFBFBF")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    def hdr(cell, val, bg):
        cell.value = val
        cell.font = Font(bold=True, size=9, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd
    
    def dat(cell, val, nf=None):
        cell.value = val
        cell.font = Font(size=9, name="Arial")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = brd
        if nf:
            cell.number_format = nf
    
    # En-têtes
    hdrs = [
        "Code article", "Description", "Classe", "Source",
        "Délai (jours)", "Consommation/mois", "Coût unitaire (TND)",
        "EOQ", "Stock Sécurité", "Point de commande",
        "Coût SS (TND)", "Date calcul"
    ]
    
    for col, label in enumerate(hdrs, 1):
        hdr(ws1.cell(1, col), label, COLORS["primary"])
    
    for i, row in df_calculs.iterrows():
        r = i + 2
        vals = [
            row.get("Code article", ""),
            row.get("Description", ""),
            row.get("Classe", ""),
            row.get("Source", ""),
            row.get("Délai (jours)", 0),
            row.get("Conso mensuelle", 0),
            row.get("Coût unitaire", 0),
            row.get("EOQ", 0),
            row.get("Stock sécurité", 0),
            row.get("Point commande", 0),
            row.get("Coût SS", 0),
            datetime.now().strftime("%Y-%m-%d")
        ]
        for col, val in enumerate(vals, 1):
            dat(ws1.cell(r, col), val)
    
    # Ajuster les largeurs
    for col, w in enumerate([15, 25, 10, 12, 10, 15, 15, 12, 12, 15, 15, 15], 1):
        ws1.column_dimensions[get_column_letter(col)].width = w
    
    # Feuille 2: Évolution (3 mois)
    ws2 = wb.create_sheet("Évolution 3 mois")
    
    evo_hdrs = ["Code article", "Description", "Mois", "Stock projeté", "SS actuel", "Alerte"]
    for col, label in enumerate(evo_hdrs, 1):
        hdr(ws2.cell(1, col), label, COLORS["secondary"])
    
    if df_evolution is not None and len(df_evolution) > 0:
        for i, row in df_evolution.iterrows():
            r = i + 2
            vals = [
                row.get("Code article", ""),
                row.get("Description", ""),
                row.get("Mois", ""),
                row.get("Stock projeté", 0),
                row.get("SS actuel", 0),
                row.get("Alerte", "")
            ]
            for col, val in enumerate(vals, 1):
                dat(ws2.cell(r, col), val)
    
    for col, w in enumerate([15, 25, 12, 15, 12, 20], 1):
        ws2.column_dimensions[get_column_letter(col)].width = w
    
    ws2.freeze_panes = "A2"
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# PAGE 1: ACCUEIL
# ============================================================

def page_accueil():
    st.markdown(f"""
    <div class='medoil-header'>
        <h1 style='color:white; margin-bottom:0.5rem'>🛢️ MedOil Supply Chain Manager</h1>
        <p style='color:rgba(255,255,255,0.8); font-size:1.1rem'>
            Optimisez votre gestion des stocks avec des calculs automatisés EOQ, SS et alertes intelligentes
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div style='font-size:2rem'>📥</div>
            <div class='metric-label'>Import</div>
            <div class='metric-value'>Excel</div>
            <div style='font-size:0.8rem;color:{COLORS["gray"]}'>Chargement automatique</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div style='font-size:2rem'>⚡</div>
            <div class='metric-label'>Calcul</div>
            <div class='metric-value'>Auto EOQ/SS</div>
            <div style='font-size:0.8rem;color:{COLORS["gray"]}'>Temps réel</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div style='font-size:2rem'>📊</div>
            <div class='metric-label'>Évolution</div>
            <div class='metric-value'>3 mois</div>
            <div style='font-size:0.8rem;color:{COLORS["gray"]}'>Projection stock</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div style='font-size:2rem'>🔔</div>
            <div class='metric-label'>Alertes</div>
            <div class='metric-value'>Commandes</div>
            <div style='font-size:0.8rem;color:{COLORS["gray"]}'>Automatiques</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col_left, col_right = st.columns([2, 1])
    
    with col_left:
        st.markdown(f"""
        <div class='medoil-card'>
            <h3>🎯 Comment ça marche ?</h3>
            <ol style='margin-top:1rem; padding-left:1.5rem'>
                <li><strong>Importez</strong> votre fichier Excel avec les données produits</li>
                <li>Les <strong>délais sont automatiques</strong> en fonction de la source (Export/Local/BM)</li>
                <li><strong>Calculez</strong> automatiquement EOQ, Stock Sécurité et Point de commande</li>
                <li><strong>Évaluez l'évolution</strong> sur 3 mois et les gains potentiels</li>
                <li><strong>Générez les alertes</strong> pour déclencher les commandes</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col_right:
        st.markdown(f"""
        <div class='medoil-card-dark' style='text-align:center'>
            <div style='font-size:1.2rem;margin-bottom:1rem'>📈 Gains estimés</div>
            <div style='font-size:2rem;color:{COLORS["secondary"]}'>15-25%</div>
            <div>Réduction stock</div>
            <hr style='margin:1rem 0'>
            <div style='font-size:2rem;color:{COLORS["secondary"]}'>30-40%</div>
            <div>Réduction ruptures</div>
        </div>
        """, unsafe_allow_html=True)

# ============================================================
# PAGE 2: CALCUL AUTOMATIQUE
# ============================================================

def page_calcul_auto():
    st.markdown(f"""
    <div class='medoil-card'>
        <h2>📥 Import & Calcul Automatique</h2>
        <p>Importez votre fichier Excel - Les délais sont définis automatiquement par source</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Paramètres globaux
    with st.expander("⚙️ Paramètres de calcul", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            niveau_service = st.selectbox(
                "Niveau de service",
                ["90% (Z=1.28)", "95% (Z=1.65)", "97.5% (Z=1.96)", "99% (Z=2.33)"],
                index=1
            )
        with col2:
            cout_passation = st.number_input("Coût de passation (TND)", value=150.0, min_value=1.0)
        with col3:
            taux_stockage = st.number_input("Taux de stockage (%)", value=20.0, min_value=1.0)
    
    Z = {"90% (Z=1.28)": 1.28, "95% (Z=1.65)": 1.65, "97.5% (Z=1.96)": 1.96, "99% (Z=2.33)": 2.33}[niveau_service]
    
    # Upload fichier
    uploaded = st.file_uploader(
        "📂 Déposez votre fichier Excel (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="calc_upload"
    )
    
    use_demo = st.checkbox("Utiliser les données de démonstration", value=not bool(uploaded))
    
    df_source = None
    
    if uploaded:
        try:
            all_sheets = pd.read_excel(uploaded, sheet_name=None, header=None)
            sheet_names = list(all_sheets.keys())
            
            if len(sheet_names) > 1:
                st.info(f"📑 {len(sheet_names)} feuilles détectées")
                selected_sheets = st.multiselect(
                    "Sélectionnez les feuilles à traiter",
                    sheet_names,
                    default=[sheet_names[0]]
                )
            else:
                selected_sheets = sheet_names
            
            all_dfs = []
            for sheet in selected_sheets:
                raw = all_sheets[sheet]
                
                # Détection de l'en-tête
                hrow = 0
                for i, row in raw.iterrows():
                    row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                    if any(k in row_str for k in ["code", "description", "article", "consommation"]):
                        hrow = i
                        break
                
                df_raw = pd.read_excel(uploaded, sheet_name=sheet, header=hrow)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                df_raw = df_raw.dropna(how="all")
                
                # Mapping des colonnes
                cmap = {}
                for col in df_raw.columns:
                    cl = col.lower().strip()
                    if "code" in cl and "article" in cl:
                        cmap["Code article"] = col
                    elif "description" in cl or "désignation" in cl:
                        cmap["Description"] = col
                    elif cl == "classe":
                        cmap["Classe"] = col
                    elif "consommation" in cl or "cons" in cl:
                        if "mois" in cl:
                            cmap["Conso mensuelle"] = col
                        else:
                            cmap["Conso mensuelle"] = col
                    elif "coût" in cl or "cout" in cl:
                        cmap["Coût unitaire"] = col
                    elif "source" in cl:
                        cmap["Source"] = col
                
                df_mapped = df_raw.rename(columns={v: k for k, v in cmap.items()})
                df_mapped["_feuille"] = sheet
                all_dfs.append(df_mapped)
            
            df_source = pd.concat(all_dfs, ignore_index=True)
            
            # Nettoyage
            for c in ["Conso mensuelle", "Coût unitaire"]:
                if c in df_source.columns:
                    df_source[c] = df_source[c].apply(clean_num)
            
            if "Code article" in df_source.columns:
                df_source = df_source[df_source["Code article"].notna()]
                df_source = df_source[df_source["Code article"].astype(str).str.strip() != ""]
            
            # Défaut source
            if "Source" not in df_source.columns:
                df_source["Source"] = "Local"
            
            st.success(f"✅ {len(df_source)} articles chargés depuis {len(selected_sheets)} feuille(s)")
            
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")
    
    if use_demo and df_source is None:
        df_source = pd.DataFrame({
            "Code article": ["A001", "A002", "A003", "A004", "A005"],
            "Description": ["Catalyseur Nickel", "Terre Tonsil", "Papier Filtre", "Toile Filtrante", "Perlite"],
            "Classe": ["MP", "MP", "SF", "SF", "MP"],
            "Source": ["Export", "Local", "BM", "Local", "Export"],
            "Conso mensuelle": [324, 5886, 509, 32460, 959],
            "Coût unitaire": [59.23, 1.71, 3.82, 49.79, 1.70]
        })
        st.info("📊 Données de démonstration chargées")
    
    if df_source is not None and len(df_source) > 0:
        with st.expander("👁️ Aperçu des données", expanded=True):
            st.dataframe(df_source, use_container_width=True, hide_index=True)
        
        # Calculs automatiques
        results = []
        for _, row in df_source.iterrows():
            conso = float(row.get("Conso mensuelle", 0) or 0)
            cu = float(row.get("Coût unitaire", 0) or 0)
            source = row.get("Source", "Local")
            
            # Délai selon source
            delay_info = get_delay_from_source(source)
            lt_mois = delay_info["mois"]
            lt_jours = delay_info["jours"]
            
            # Calculs
            conso_an = conso * 12
            ss = calc_ss(conso, z=Z, lt_mois=lt_mois, variab=0.15)
            eoq, _, _ = calc_eoq(conso_an, cu, sc_cost=cout_passation, taux=taux_stockage/100)
            pr = calc_pr(conso, lt_mois, ss)
            cout_ss = round(ss * cu, 2) if cu > 0 else 0
            
            results.append({
                "Code article": row.get("Code article", ""),
                "Description": row.get("Description", ""),
                "Classe": row.get("Classe", ""),
                "Source": source,
                "Délai (jours)": lt_jours,
                "Conso mensuelle": conso,
                "Coût unitaire": cu,
                "Conso annuelle": conso_an,
                "EOQ": eoq,
                "Stock sécurité": ss,
                "Point commande": pr,
                "Coût SS": cout_ss,
                "_feuille": row.get("_feuille", "")
            })
        
        df_results = pd.DataFrame(results)
        
        # Affichage des résultats
        st.markdown("---")
        st.markdown("#### 📊 Résultats des calculs")
        
        col_display = ["Code article", "Description", "Source", "Délai (jours)", 
                       "Conso mensuelle", "Coût unitaire", "EOQ", "Stock sécurité", "Point commande", "Coût SS"]
        
        st.dataframe(
            df_results[col_display],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Coût unitaire": st.column_config.NumberColumn("Coût unitaire (TND)", format="%.2f"),
                "Coût SS": st.column_config.NumberColumn("Coût SS (TND)", format="%.2f"),
                "EOQ": st.column_config.NumberColumn("EOQ", format="%d"),
                "Stock sécurité": st.column_config.NumberColumn("Stock sécurité", format="%d"),
                "Point commande": st.column_config.NumberColumn("Point commande", format="%d"),
            }
        )
        
        # KPIs
        st.markdown("---")
        st.markdown("#### 📈 Indicateurs synthétiques")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Articles traités</div>
                <div class='metric-value'>{len(df_results)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Coût total SS</div>
                <div class='metric-value'>{fmtInt(df_results['Coût SS'].sum())} TND</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>EOQ moyen</div>
                <div class='metric-value'>{fmtInt(df_results['EOQ'].mean())}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>SS total</div>
                <div class='metric-value'>{fmtInt(df_results['Stock sécurité'].sum())}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Graphique
        st.markdown("---")
        st.markdown("#### 📊 Visualisation des indicateurs")
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="EOQ",
            x=df_results["Code article"].astype(str),
            y=df_results["EOQ"],
            marker_color=COLORS["primary"],
            text=df_results["EOQ"].apply(lambda v: f"{int(v):,}" if v > 0 else ""),
            textposition="outside"
        ))
        fig.add_trace(go.Bar(
            name="Stock Sécurité",
            x=df_results["Code article"].astype(str),
            y=df_results["Stock sécurité"],
            marker_color=COLORS["secondary"],
            text=df_results["Stock sécurité"].apply(lambda v: f"{int(v):,}"),
            textposition="outside"
        ))
        fig.add_trace(go.Scatter(
            name="Point commande",
            x=df_results["Code article"].astype(str),
            y=df_results["Point commande"],
            mode="markers+lines",
            marker=dict(color=COLORS["danger"], size=9, symbol="diamond"),
            line=dict(color=COLORS["danger"], width=2, dash="dash")
        ))
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color=COLORS["gray"]),
            legend=dict(bgcolor="rgba(0,0,0,0)"),
            height=400,
            barmode="group",
            xaxis=dict(title="Article", gridcolor="rgba(0,0,0,0.05)"),
            yaxis=dict(title="Unités", gridcolor="rgba(0,0,0,0.05)")
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Export
        st.markdown("---")
        st.markdown("#### 💾 Export")
        
        st.info("Le fichier exporté contient 2 feuilles : Calculs SC et Évolution 3 mois")
        
        excel_buf = build_excel_complete(df_results)
        st.download_button(
            label="⬇️ Télécharger le tableau Excel complet",
            data=excel_buf,
            file_name=f"medoil_sc_calculs_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        # Stocker dans session_state pour la page alertes
        st.session_state['df_calculs'] = df_results

# ============================================================
# PAGE 3: CALCULATEURS
# ============================================================

def page_calculateurs():
    st.markdown(f"""
    <div class='medoil-card'>
        <h2>⚙️ Calculateurs Supply Chain</h2>
        <p>Outils interactifs pour vos calculs EOQ, Stock de sécurité et Point de commande</p>
    </div>
    """, unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["📦 Stock de Sécurité", "📐 EOQ (Wilson)", "🔄 Point de commande"])
    
    with tab1:
        st.markdown("""
        <div style='background:#F0F9FF; padding:1rem; border-radius:12px; margin-bottom:1rem'>
            <strong>Formule :</strong> SS = Z × √(LT × σD² + D² × σLT²)
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### 📊 Consommation")
            conso_mensuelle = st.number_input("Consommation mensuelle (u)", value=500.0, step=50.0)
            variabilite = st.slider("Variabilité de la demande (%)", 5, 30, 15)
        
        with col2:
            st.markdown("##### ⏱️ Délais")
            source_delay = st.selectbox("Source", list(SOURCE_DELAYS.keys()))
            lt_info = SOURCE_DELAYS[source_delay]
            st.caption(f"Délai : {lt_info['label']} ({lt_info['jours']} jours)")
            
            niveau_service = st.selectbox(
                "Niveau de service",
                ["90%", "95%", "97.5%", "99%"],
                index=1
            )
        
        Z_map = {"90%": 1.28, "95%": 1.65, "97.5%": 1.96, "99%": 2.33}
        Z = Z_map[niveau_service]
        
        ss = calc_ss(conso_mensuelle, z=Z, lt_mois=lt_info["mois"], variab=variabilite/100)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Stock de Sécurité</div>
                <div class='metric-value'>{fmtInt(ss)} u</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            cout_unitaire = st.number_input("Coût unitaire (TND)", value=50.0, step=10.0)
            cout_ss = ss * cout_unitaire
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Coût immobilisé SS</div>
                <div class='metric-value'>{fmtInt(cout_ss)} TND</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            pr = calc_pr(conso_mensuelle, lt_info["mois"], ss)
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Point de commande</div>
                <div class='metric-value'>{fmtInt(pr)} u</div>
            </div>
            """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown("""
        <div style='background:#F0F9FF; padding:1rem; border-radius:12px; margin-bottom:1rem'>
            <strong>Formule EOQ :</strong> Q* = √(2 × D × S / H)
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            demande_annuelle = st.number_input("Demande annuelle (u)", value=12000.0, step=1000.0)
            cout_passation = st.number_input("Coût de passation (TND)", value=150.0, step=10.0)
        
        with col2:
            cout_unitaire = st.number_input("Coût unitaire (TND)", value=80.0, step=10.0, key="eoq_cu")
            taux_stockage = st.number_input("Taux de stockage (%)", value=20.0, step=1.0)
        
        eoq, nb_cmd, ct_min = calc_eoq(demande_annuelle, cout_unitaire, cout_passation, taux_stockage/100)
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>EOQ</div>
                <div class='metric-value'>{fmtInt(eoq)} u</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Nb commandes/an</div>
                <div class='metric-value'>{fmt(nb_cmd, 1)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            periodicite = 250 / nb_cmd if nb_cmd > 0 else 0
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Périodicité</div>
                <div class='metric-value'>{fmtInt(periodicite)} jours</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Coût total min</div>
                <div class='metric-value'>{fmtInt(ct_min)} TND</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Graphique de sensibilité
        qty_range = np.linspace(max(10, eoq * 0.2), eoq * 2.5, 100)
        cout_pass = [demande_annuelle / q * cout_passation for q in qty_range]
        cout_stock = [q / 2 * (taux_stockage/100) * cout_unitaire for q in qty_range]
        cout_total = [cout_pass[i] + cout_stock[i] for i in range(len(qty_range))]
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=qty_range, y=cout_pass, name="Coût passation", line=dict(color=COLORS["primary"], width=2)))
        fig.add_trace(go.Scatter(x=qty_range, y=cout_stock, name="Coût stockage", line=dict(color=COLORS["secondary"], width=2)))
        fig.add_trace(go.Scatter(x=qty_range, y=cout_total, name="Coût total", line=dict(color=COLORS["success"], width=2.5)))
        fig.add_vline(x=eoq, line_dash="dash", line_color=COLORS["danger"], annotation_text=f"EOQ = {fmtInt(eoq)}")
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=350,
            xaxis_title="Quantité commandée (u)",
            yaxis_title="Coût annuel (TND)"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.markdown("""
        <div style='background:#F0F9FF; padding:1rem; border-radius:12px; margin-bottom:1rem'>
            <strong>Formule :</strong> Point de commande = Demande journalière × Délai + Stock sécurité
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            conso_mensuelle = st.number_input("Consommation mensuelle (u)", value=500.0, step=50.0, key="pr_conso")
            demande_journaliere = conso_mensuelle / 30
        
        with col2:
            source = st.selectbox("Source", list(SOURCE_DELAYS.keys()), key="pr_source")
            lt_info = SOURCE_DELAYS[source]
            delai_jours = lt_info["jours"]
        
        with col3:
            ss = st.number_input("Stock de sécurité (u)", value=100.0, step=10.0, key="pr_ss")
        
        pr = demande_journaliere * delai_jours + ss
        
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center; margin-top:1rem'>
            <div class='metric-label'>Point de commande</div>
            <div class='metric-value'>{fmtInt(pr)} unités</div>
            <div style='font-size:0.85rem; margin-top:0.5rem'>
                Commander lorsque le stock atteint ce seuil
            </div>
        </div>
        """, unsafe_allow_html=True)

# ============================================================
# PAGE 4: ALERTES
# ============================================================

def page_alertes():
    st.markdown(f"""
    <div class='medoil-card'>
        <h2>🔔 Alertes & Surveillance</h2>
        <p>Importez votre fichier de suivi pour générer les alertes de rupture et de réapprovisionnement</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Upload du fichier de suivi
    uploaded = st.file_uploader(
        "📂 Déposez votre fichier Excel de suivi des stocks",
        type=["xlsx", "xls"],
        key="alert_upload"
    )
    
    # Ou utiliser les calculs de la page précédente
    if 'df_calculs' in st.session_state and st.checkbox("Utiliser les résultats des calculs précédents"):
        df_calculs = st.session_state['df_calculs']
        
        # Simuler le suivi
        df_suivi = df_calculs[["Code article", "Description", "Stock sécurité", "Point commande"]].copy()
        df_suivi["Stock actuel"] = df_suivi["Stock sécurité"] * np.random.uniform(0.5, 1.5, len(df_suivi))
        df_suivi["Consommation mois"] = df_calculs["Conso mensuelle"]
        
        st.success(f"✅ Utilisation des {len(df_suivi)} articles depuis les calculs précédents")
    else:
        df_suivi = None
    
    if uploaded:
        try:
            df_raw = pd.read_excel(uploaded)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            
            # Mapping
            cmap = {}
            for col in df_raw.columns:
                cl = col.lower().strip()
                if "code" in cl and "article" in cl:
                    cmap["Code article"] = col
                elif "description" in cl or "designation" in cl:
                    cmap["Description"] = col
                elif "stock sécurité" in cl or "ss" in cl:
                    cmap["Stock sécurité"] = col
                elif "point commande" in cl or "pr" in cl:
                    cmap["Point commande"] = col
                elif "stock actuel" in cl or "stk" in cl:
                    cmap["Stock actuel"] = col
                elif "consommation" in cl or "cons" in cl:
                    cmap["Consommation mois"] = col
            
            df_suivi = df_raw.rename(columns={v: k for k, v in cmap.items()})
            
            for c in ["Stock sécurité", "Point commande", "Stock actuel", "Consommation mois"]:
                if c in df_suivi.columns:
                    df_suivi[c] = df_suivi[c].apply(clean_num)
            
            st.success(f"✅ {len(df_suivi)} articles chargés")
            
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")
    
    if df_suivi is not None and len(df_suivi) > 0:
        # Générer les alertes
        def generate_alert(row):
            ss = row.get("Stock sécurité", 0)
            pr = row.get("Point commande", 0)
            stock = row.get("Stock actuel", 0)
            conso = row.get("Consommation mois", 0)
            
            if stock <= ss:
                return "🔴 ALERTE CRITIQUE - Stock inférieur au SS", "Commander immédiatement"
            elif stock <= pr:
                return "🟡 ALERTE - Point de commande atteint", "Déclencher une commande"
            else:
                jours_restants = (stock - pr) / (conso / 30) if conso > 0 else 999
                if jours_restants < 15:
                    return "🟠 Attention - Stock faible", f"Commander sous {fmtInt(jours_restants)} jours"
                return "🟢 Normal", f"Stock OK pour {fmtInt(jours_restants)} jours"
        
        alertes = df_suivi.apply(lambda r: generate_alert(r), axis=1)
        df_suivi["Alerte niveau"] = [a[0] for a in alertes]
        df_suivi["Action"] = [a[1] for a in alertes]
        
        # Synthèse
        st.markdown("---")
        st.markdown("#### 📊 Synthèse des alertes")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            n_critique = len(df_suivi[df_suivi["Alerte niveau"] == "🔴 ALERTE CRITIQUE - Stock inférieur au SS"])
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center; border-top-color:{COLORS["danger"]}'>
                <div class='metric-label'>🔴 Critique</div>
                <div class='metric-value' style='color:{COLORS["danger"]}'>{n_critique}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            n_alerte = len(df_suivi[df_suivi["Alerte niveau"] == "🟡 ALERTE - Point de commande atteint"])
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center; border-top-color:{COLORS["warning"]}'>
                <div class='metric-label'>🟡 À commander</div>
                <div class='metric-value' style='color:{COLORS["warning"]}'>{n_alerte}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            n_attention = len(df_suivi[df_suivi["Alerte niveau"] == "🟠 Attention - Stock faible"])
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center; border-top-color:{COLORS["accent"]}'>
                <div class='metric-label'>🟠 Attention</div>
                <div class='metric-value'>{n_attention}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            n_normal = len(df_suivi[df_suivi["Alerte niveau"] == "🟢 Normal"])
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center; border-top-color:{COLORS["success"]}'>
                <div class='metric-label'>🟢 Normal</div>
                <div class='metric-value'>{n_normal}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Tableau des alertes
        st.markdown("---")
        st.markdown("#### 📋 Détail des alertes")
        
        filtre = st.multiselect(
            "Filtrer par niveau d'alerte",
            options=df_suivi["Alerte niveau"].unique(),
            default=["🔴 ALERTE CRITIQUE - Stock inférieur au SS", "🟡 ALERTE - Point de commande atteint", "🟠 Attention - Stock faible"]
        )
        
        if filtre:
            df_show = df_suivi[df_suivi["Alerte niveau"].isin(filtre)]
        else:
            df_show = df_suivi
        
        display_cols = ["Code article", "Description", "Stock actuel", "Stock sécurité", "Point commande", "Alerte niveau", "Action"]
        display_cols = [c for c in display_cols if c in df_show.columns]
        
        st.dataframe(
            df_show[display_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Stock actuel": st.column_config.NumberColumn("Stock actuel", format="%.0f"),
                "Stock sécurité": st.column_config.NumberColumn("SS", format="%.0f"),
                "Point commande": st.column_config.NumberColumn("PR", format="%.0f"),
            }
        )
        
        # Liste des actions urgentes
        st.markdown("---")
        st.markdown("#### 🚨 Actions prioritaires")
        
        urgentes = df_suivi[df_suivi["Alerte niveau"].isin(["🔴 ALERTE CRITIQUE - Stock inférieur au SS", "🟡 ALERTE - Point de commande atteint"])]
        
        if len(urgentes) == 0:
            st.markdown(f"""
            <div class='alert-success'>
                ✅ Aucune alerte critique - Tous les stocks sont suffisants
            </div>
            """, unsafe_allow_html=True)
        
        for _, row in urgentes.iterrows():
            color = COLORS["danger"] if "CRITIQUE" in row["Alerte niveau"] else COLORS["warning"]
            bg_class = "alert-critical" if "CRITIQUE" in row["Alerte niveau"] else "alert-warning"
            
            st.markdown(f"""
            <div class='{bg_class}'>
                <strong>{row['Alerte niveau']}</strong><br/>
                <strong>{row.get('Code article', '')} - {row.get('Description', '')}</strong><br/>
                Stock actuel : {fmtInt(row.get('Stock actuel', 0))} | SS : {fmtInt(row.get('Stock sécurité', 0))} | PR : {fmtInt(row.get('Point commande', 0))}<br/>
                <span style='font-size:0.85rem'>➡️ {row.get('Action', '')}</span>
            </div>
            """, unsafe_allow_html=True)

# ============================================================
# PAGE 5: ÉVALUATION DES GAINS
# ============================================================

def page_evaluation_gains():
    st.markdown(f"""
    <div class='medoil-card'>
        <h2>📈 Évaluation des gains</h2>
        <p>Évaluez les gains potentiels générés par l'optimisation de vos stocks de sécurité</p>
    </div>
    """, unsafe_allow_html=True)
    
    if 'df_calculs' not in st.session_state:
        st.info("Veuillez d'abord effectuer les calculs dans l'onglet 'Import & Calcul Automatique'")
        return
    
    df = st.session_state['df_calculs'].copy()
    
    with st.expander("📊 Configuration des paramètres", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            reduction_ss = st.slider("Réduction du stock sécurité cible (%)", 0, 50, 25)
        with col2:
            taux_financier = st.number_input("Taux de coût financier (%)", value=8.0, step=0.5)
    
    # Calculs des gains
    df["SS actuel"] = df["Stock sécurité"]
    df["SS cible"] = (df["SS actuel"] * (1 - reduction_ss/100)).round(0)
    df["Gain unités"] = df["SS actuel"] - df["SS cible"]
    df["Gain financier"] = df["Gain unités"] * df["Coût unitaire"] * (taux_financier/100)
    
    gain_total_unites = df["Gain unités"].sum()
    gain_total_financier = df["Gain financier"].sum()
    
    # KPIs
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div class='metric-label'>Réduction SS</div>
            <div class='metric-value' style='color:{COLORS["success"]}'>-{reduction_ss}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div class='metric-label'>Gain en unités</div>
            <div class='metric-value'>{fmtInt(gain_total_unites)} u</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class='medoil-card' style='text-align:center'>
            <div class='metric-label'>Gain financier annuel</div>
            <div class='metric-value'>{fmtInt(gain_total_financier)} TND</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Tableau détaillé
    st.markdown("---")
    st.markdown("#### 📋 Détail par article")
    
    detail_cols = ["Code article", "Description", "SS actuel", "SS cible", "Gain unités", "Coût unitaire", "Gain financier"]
    detail_cols = [c for c in detail_cols if c in df.columns]
    
    st.dataframe(
        df[detail_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "SS actuel": st.column_config.NumberColumn("SS actuel", format="%d"),
            "SS cible": st.column_config.NumberColumn("SS cible", format="%d"),
            "Gain unités": st.column_config.NumberColumn("Gain unités", format="%d"),
            "Coût unitaire": st.column_config.NumberColumn("Coût unitaire (TND)", format="%.2f"),
            "Gain financier": st.column_config.NumberColumn("Gain financier (TND)", format="%.0f"),
        }
    )
    
    # Graphique
    st.markdown("---")
    st.markdown("#### 📊 Comparaison SS actuel vs SS cible")
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="SS actuel",
        x=df["Code article"].astype(str),
        y=df["SS actuel"],
        marker_color=COLORS["primary"]
    ))
    fig.add_trace(go.Bar(
        name=f"SS cible (-{reduction_ss}%)",
        x=df["Code article"].astype(str),
        y=df["SS cible"],
        marker_color=COLORS["success"]
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        barmode="group",
        height=400,
        xaxis_title="Article",
        yaxis_title="Unités"
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Simulation sur 3 mois
    st.markdown("---")
    st.markdown("#### 📅 Simulation sur 3 mois")
    
    df_evolution = []
    mois = ["Mois 1", "Mois 2", "Mois 3"]
    
    for _, row in df.iterrows():
        conso_mensuelle = row.get("Conso mensuelle", 0)
        ss_actuel = row.get("SS actuel", 0)
        
        for i, m in enumerate(mois):
            stock_projete = ss_actuel - (conso_mensuelle * (i + 1))
            df_evolution.append({
                "Code article": row["Code article"],
                "Description": row["Description"],
                "Mois": m,
                "Stock projeté": max(0, stock_projete),
                "SS actuel": ss_actuel,
                "Alerte": "RUPTURE" if stock_projete < 0 else "OK"
            })
    
    df_evolution = pd.DataFrame(df_evolution)
    st.dataframe(df_evolution, use_container_width=True, hide_index=True)

# ============================================================
# SIDEBAR & NAVIGATION
# ============================================================

with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding:1rem 0'>
        <div style='font-size:2rem'>🛢️</div>
        <h2 style='color:white; margin:0'>MedOil</h2>
        <p style='color:rgba(255,255,255,0.7); font-size:0.8rem'>Supply Chain Manager</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    # Navigation
    pages = {
        "🏠 Accueil": page_accueil,
        "📥 Calcul Auto": page_calcul_auto,
        "⚙️ Calculateurs": page_calculateurs,
        "🔔 Alertes": page_alertes,
        "📈 Évaluation Gains": page_evaluation_gains,
    }
    
    selected = st.radio(
        "Navigation",
        list(pages.keys()),
        label_visibility="collapsed"
    )
    
    st.divider()
    
    st.markdown(f"""
    <div style='text-align:center; margin-top:2rem'>
        <p style='color:rgba(255,255,255,0.5); font-size:0.7rem'>
            MedOil Supply Chain<br>
            Version 2.0
        </p>
    </div>
    """, unsafe_allow_html=True)

# Affichage de la page sélectionnée
pages[selected]()
