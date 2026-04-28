import streamlit as st
import math
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date
import calendar

st.set_page_config(page_title="MedOil SC", page_icon="🌿", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700&family=DM+Mono:wght@400;500&family=DM+Serif+Display:ital@0;1&display=swap');

:root {
    --green-dark:   #1a3d1a;
    --green-mid:    #2d6a2d;
    --green-light:  #4a8c3f;
    --green-pale:   #c8e6c9;
    --gold:         #f5c518;
    --gold-dark:    #c9a000;
    --gold-mid:     #e8b400;
    --bg-page:      #f4f4ee;
    --bg-card:      #ffffff;
    --bg-sidebar:   #152415;
    --text-dark:    #1a2e1a;
    --text-mid:     #3d5c3d;
    --text-muted:   #7a9a7a;
    --border:       #d4e0d4;
}

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
    background-color: var(--bg-page) !important;
    color: var(--text-dark) !important;
}

[data-testid="stSidebar"] {
    background-color: var(--bg-sidebar) !important;
    border-right: 1px solid #2d4a2d !important;
}
[data-testid="stSidebar"] * { color: #c8e6c9 !important; }
[data-testid="stSidebar"] h1 { color: #ffffff !important; font-size: 1.2rem !important; }
[data-testid="stSidebar"] hr { border-color: #2d4a2d !important; }

[data-testid="stSidebarNavLink"] {
    border-radius: 8px !important;
    color: #a5c8a5 !important;
    font-size: .88rem !important;
    padding: .5rem .9rem !important;
    margin: 2px 0 !important;
    transition: background .15s;
}
[data-testid="stSidebarNavLink"]:hover {
    background: rgba(245,197,24,.12) !important;
    color: var(--gold) !important;
}
[data-testid="stSidebarNavLink"][aria-current="page"] {
    background: rgba(245,197,24,.18) !important;
    color: var(--gold) !important;
    border-left: 3px solid var(--gold) !important;
    font-weight: 600 !important;
}

.main .block-container {
    background-color: var(--bg-page) !important;
    padding-top: 1.5rem !important;
}

h1, h2, h3, h4 { color: var(--text-dark) !important; }

.formula-box {
    background: #eef4ee;
    border-left: 3px solid var(--green-light);
    border-radius: 8px;
    padding: .9rem 1.1rem;
    font-family: 'DM Mono', monospace;
    font-size: .82rem;
    color: var(--text-mid);
    margin-bottom: 1rem;
    line-height: 1.7;
    border: 1px solid var(--border);
}
.f-eq { color: var(--green-mid); font-weight: 600; }

.alert-critical {
    background: rgba(220,53,53,.07);
    border: 1px solid rgba(220,53,53,.3);
    border-left: 4px solid #dc3535;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: .7rem;
    color: #8b1a1a;
}
.alert-warning {
    background: rgba(245,197,24,.1);
    border: 1px solid rgba(200,160,0,.3);
    border-left: 4px solid var(--gold-dark);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: .7rem;
    color: #7a5c00;
}
.alert-ok {
    background: rgba(74,140,63,.08);
    border: 1px solid rgba(74,140,63,.3);
    border-left: 4px solid var(--green-light);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: .7rem;
    color: var(--green-mid);
}
.alert-info {
    background: rgba(45,106,45,.07);
    border: 1px solid rgba(45,106,45,.25);
    border-left: 4px solid var(--green-mid);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: .7rem;
    color: var(--text-mid);
}

.stSelectbox > div > div,
.stNumberInput > div > div > input,
.stTextInput > div > div > input {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    color: var(--text-dark) !important;
    border-radius: 8px !important;
}

[data-testid="stWidgetLabel"] p { color: var(--text-mid) !important; font-size: .85rem !important; }

.stButton > button, .stDownloadButton > button {
    background: var(--gold) !important;
    color: var(--green-dark) !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-family: 'Syne', sans-serif !important;
    padding: .55rem 1.4rem !important;
    transition: background .15s, transform .1s !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    background: var(--gold-dark) !important;
    transform: translateY(-1px) !important;
}

[data-testid="stTabs"] [role="tablist"] {
    border-bottom: 2px solid var(--border) !important;
    gap: .5rem !important;
}
[data-testid="stTabs"] [role="tab"] {
    color: var(--text-muted) !important;
    background: transparent !important;
    border-radius: 8px 8px 0 0 !important;
    padding: .5rem 1.1rem !important;
    font-weight: 500 !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: var(--green-dark) !important;
    border-bottom: 3px solid var(--gold) !important;
    font-weight: 700 !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}

[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
}

.stInfo    { background: rgba(45,106,45,.07) !important; border-color: var(--green-light) !important; color: var(--green-mid) !important; }
.stSuccess { background: rgba(74,140,63,.08) !important; border-color: var(--green-light) !important; }
.stWarning { background: rgba(245,197,24,.1) !important; border-color: var(--gold-dark) !important; }
hr { border-color: var(--border) !important; }

[data-testid="stMultiSelect"] span {
    background: var(--green-pale) !important;
    color: var(--green-dark) !important;
    border-radius: 4px !important;
}
</style>""", unsafe_allow_html=True)

# ── Plotly theme ──────────────────────────────────────────────────────────────
def light_layout():
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(248,250,248,0.6)",
        font=dict(color="#3d5c3d", family="Syne"),
        legend=dict(bgcolor="rgba(255,255,255,.9)", bordercolor="#d4e0d4", borderwidth=1),
        margin=dict(t=20, b=10, l=0, r=0)
    )

C_GREEN  = "#2d6a2d"
C_GOLD   = "#f5c518"
C_GREEN2 = "#4a8c3f"
C_RED    = "#dc3535"
C_PURPLE = "#7a5c9a"
C_MUTED  = "#7a9a7a"

def fmt(n, d=1):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))): return "—"
    return f"{n:,.{d}f}"

def fmtInt(n):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))): return "—"
    return f"{round(n):,}"

def mcard(label, value, unit, delta=None, color="#2d6a2d"):
    dh = f"<div style='font-size:.7rem;color:{color};margin-top:4px'>{delta}</div>" if delta else ""
    return f"""<div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;
        padding:1rem 1.2rem;text-align:center;border-top:3px solid {color};
        box-shadow:0 2px 8px rgba(26,61,26,.06)'>
        <div style='font-size:.65rem;color:#7a9a7a;text-transform:uppercase;letter-spacing:.1em;margin-bottom:.4rem'>{label}</div>
        <div style='font-family:"DM Mono",monospace;font-size:1.6rem;font-weight:500;color:#1a2e1a'>{value}</div>
        <div style='font-size:.68rem;color:#7a9a7a;margin-top:3px'>{unit}</div>{dh}</div>"""

def fbox(title, formula, details):
    return (f"<div class='formula-box'>"
            f"<div style='font-size:.65rem;color:#7a9a7a;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px'>{title}</div>"
            f"<span class='f-eq'>{formula}</span><br/>"
            f"<span style='color:#7a9a7a'>{details}</span></div>")

# ── Niveau de service automatique par classe ABC ──────────────────────────────
ABC_SERVICE_LEVEL = {
    "A": {"pct": 0.99, "Z": 2.33, "label": "99%"},
    "B": {"pct": 0.95, "Z": 1.65, "label": "95%"},
    "C": {"pct": 0.90, "Z": 1.28, "label": "90%"},
}

# ── Délais par source (en jours) ──────────────────────────────────────────────
SOURCE_DELAIS = {
    "Export":  4 * 30,   # 4 mois
    "Local":   14,        # 2 semaines
    "BM":      21,        # 3 semaines
}

def get_lt_jours(source):
    return SOURCE_DELAIS.get(source, 30)

def calc_ss(sigma_d, lt_j, z=1.65):
    """SS = Z × σD × √L   (σD = écart-type demande u/j, L = délai en jours)."""
    return round(z * sigma_d * math.sqrt(lt_j))

def calc_pr(d_j, lt_j, ss):
    return round(d_j * lt_j + ss)

def calc_eoq(conso_an, cout_unit, sc_cost=150, taux=0.20):
    if cout_unit <= 0 or conso_an <= 0: return 0, 0, 0
    h   = taux * cout_unit
    eoq = math.sqrt(2 * conso_an * sc_cost / h)
    nb  = conso_an / eoq
    ct  = nb * sc_cost + (eoq / 2) * h
    return round(eoq), round(nb, 1), round(ct, 2)

# ── Export Excel complet (sans CBN) ──────────────────────────────────────────
def build_excel_complet(df_res):
    wb  = Workbook()
    thin = Side(style="thin", color="BFBFBF")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr(cell, val, bg):
        cell.value = val
        cell.font  = Font(bold=True, size=9, name="Arial")
        cell.fill  = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd

    def dat(cell, val, nf=None, bold=False):
        cell.value = val
        cell.font  = Font(size=9, name="Arial", bold=bold)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = brd
        if nf: cell.number_format = nf

    ws1 = wb.active
    ws1.title = "Indicateurs SC"

    headers_row1 = [
        ("Code article",           "C8E6C9"), ("Description",             "C8E6C9"),
        ("Source approv.",          "FFF9C4"), ("Délai (j)",               "FFF9C4"),
        ("Classe ABC",             "DCEDC8"), ("Classe XYZ",              "DCEDC8"),
        ("Niveau service",          "FFF9C4"), ("Z",                       "FFF9C4"),
        ("Consommation/mois (moy)","DCEDC8"), ("Sigma demande (u/j)",     "DCEDC8"),
        ("Stock sécurité (u)",     "FFE0B2"), ("Coût unitaire (TND)",     "FFE0B2"),
        ("Coût passation (TND)",   "FFE0B2"), ("Taux stockage (%)",       "FFE0B2"),
        ("Coût SS total (TND)",    "FFE0B2"), ("EOQ (u)",                 "B3E5FC"),
        ("Nb cmd/an",              "B3E5FC"), ("Point de commande (u)",   "FFE0B2"),
        ("Stock actuel (u)",       "E8F5E9"), ("Commande en cours (u)",   "E8F5E9"),
    ]

    for col, (label, bg) in enumerate(headers_row1, 1):
        hdr(ws1.cell(1, col), label, bg)

    fmts1 = [None, None, None, "0", None, None, "0%", "0.00",
              "#,##0.0", "#,##0.00",
              "#,##0", "#,##0.00", "#,##0.00", "0.0%",
              "#,##0.00", "#,##0", "0.0", "#,##0", "#,##0", "#,##0"]

    for i, row in df_res.iterrows():
        r = i + 2
        vals = [
            row.get("Code article", ""),
            row.get("Description Article", ""),
            row.get("Source", ""),
            row.get("LT jours", 0),
            row.get("Classe ABC", ""),
            row.get("Classe XYZ", ""),
            row.get("Niveau service pct", 0.95),
            row.get("Z val", 1.65),
            round(row.get("Conso mois moy", 0), 1),
            round(row.get("Sigma D (j)", 0), 4),
            row.get("Stock sécurité", 0),
            row.get("Coût unitaire", 0),
            row.get("Coût passation", 0),
            row.get("Taux stockage", 0),
            row.get("Coût SS", 0),
            row.get("EOQ", 0),
            row.get("Nb cmd an", 0),
            row.get("Point de commande", 0),
            row.get("Stock actuel", 0),
            row.get("Commande en cours", 0),
        ]
        for col, (val, nf) in enumerate(zip(vals, fmts1), 1):
            dat(ws1.cell(r, col), val, nf)

    for col, w in enumerate([14, 30, 10, 10, 8, 8, 12, 8, 18, 16, 14, 16, 16, 14, 16, 12, 10, 18, 16, 18], 1):
        ws1.column_dimensions[get_column_letter(col)].width = w
    ws1.row_dimensions[1].height = 42
    ws1.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════════════════════════════
# PAGES
# ════════════════════════════════════════════════════════════════════════════════

def accueil():
    st.markdown("""
    <div style='background:linear-gradient(135deg,#1a3d1a,#2d6a2d);padding:2.2rem 2.5rem;border-radius:16px;
         border:1px solid rgba(245,197,24,.2);margin-bottom:1.5rem;position:relative;overflow:hidden'>
        <div style='position:absolute;top:-30px;right:-30px;width:200px;height:200px;
             background:radial-gradient(circle,rgba(245,197,24,.12),transparent 70%);border-radius:50%'></div>
        <div style='font-size:.7rem;color:#f5c518;letter-spacing:.14em;text-transform:uppercase;margin-bottom:12px;
             background:rgba(245,197,24,.15);display:inline-block;padding:.3rem .8rem;border-radius:20px;
             border:1px solid rgba(245,197,24,.3)'>✦ Outil supply chain manager · v4.0</div>
        <h1 style='font-family:"DM Serif Display",serif;font-size:2.3rem;color:#ffffff;margin:.4rem 0;line-height:1.25'>
            Pilotez votre <em style='color:#f5c518'>supply chain</em><br/>avec précision</h1>
        <p style='color:#a5c8a5;font-size:.95rem;margin:.8rem 0 1.4rem'>
            Importez vos données de consommation — SS et EOQ calculés automatiquement par article, avec niveau de service adapté à la classe ABC.</p>
    </div>""", unsafe_allow_html=True)

    # Délais
    st.markdown("""
    <div style='display:grid;grid-template-columns:repeat(3,1fr);gap:1rem;margin-bottom:1.5rem'>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;border-top:3px solid #2d6a2d'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:.5rem'>🕐 Délai Export</div>
            <div style='font-size:1.6rem;font-weight:700;color:#1a2e1a;font-family:"DM Mono",monospace'>4 mois</div>
            <div style='font-size:.75rem;color:#7a9a7a'>Approvisionnement international</div>
        </div>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;border-top:3px solid #f5c518'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:.5rem'>🕐 Délai Local</div>
            <div style='font-size:1.6rem;font-weight:700;color:#1a2e1a;font-family:"DM Mono",monospace'>2 semaines</div>
            <div style='font-size:.75rem;color:#7a9a7a'>Fournisseurs locaux</div>
        </div>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;border-top:3px solid #4a8c3f'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:.5rem'>🕐 Délai BM</div>
            <div style='font-size:1.6rem;font-weight:700;color:#1a2e1a;font-family:"DM Mono",monospace'>3 semaines</div>
            <div style='font-size:.75rem;color:#7a9a7a'>Bureau de marché</div>
        </div>
    </div>""", unsafe_allow_html=True)

    # Niveau de service par classe
    st.markdown("""
    <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;margin-bottom:1.5rem;border-top:3px solid #f5c518'>
        <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:.8rem;font-weight:600'>⚙️ Niveau de service automatique par classe ABC</div>
        <div style='display:grid;grid-template-columns:repeat(3,1fr);gap:1rem'>
            <div style='text-align:center;background:#c8e6c9;border-radius:8px;padding:.8rem'>
                <div style='font-size:1.4rem;font-weight:800;color:#1a3d1a'>A</div>
                <div style='font-size:1.1rem;font-weight:700;color:#1a3d1a'>Z = 2.33</div>
                <div style='font-size:.78rem;color:#2d6a2d'>Niveau de service 99%</div>
                <div style='font-size:.72rem;color:#4a8c3f'>Articles stratégiques</div>
            </div>
            <div style='text-align:center;background:#fff9c4;border-radius:8px;padding:.8rem'>
                <div style='font-size:1.4rem;font-weight:800;color:#7a5c00'>B</div>
                <div style='font-size:1.1rem;font-weight:700;color:#7a5c00'>Z = 1.65</div>
                <div style='font-size:.78rem;color:#7a5c00'>Niveau de service 95%</div>
                <div style='font-size:.72rem;color:#c9a000'>Articles intermédiaires</div>
            </div>
            <div style='text-align:center;background:#f5f5f5;border-radius:8px;padding:.8rem'>
                <div style='font-size:1.4rem;font-weight:800;color:#555'>C</div>
                <div style='font-size:1.1rem;font-weight:700;color:#555'>Z = 1.28</div>
                <div style='font-size:.78rem;color:#555'>Niveau de service 90%</div>
                <div style='font-size:.72rem;color:#7a9a7a'>Articles courants</div>
            </div>
        </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<div style='font-size:1rem;font-weight:700;color:#1a2e1a;margin-bottom:1rem;border-left:3px solid #f5c518;padding-left:10px'>Modules disponibles</div>", unsafe_allow_html=True)
    cols = st.columns(3)
    mods = [
        ("📥", "#2d6a2d", "Calcul automatique",    "Importez consommations → SS · EOQ par article → Export XLSX"),
        ("📐", "#c9a000", "Calculateurs SC",         "EOQ · SS · KPIs · Point de réappro. avec courbes"),
        ("🔔", "#dc3535", "Alertes & Stocks",         "Alertes rupture · Déclenchement commandes automatique"),
        ("📈", "#4a8c3f", "Évolution & Gains",        "Suivi SS sur période · évaluation économies réalisées"),
    ]
    for col, (ic, accent, ti, de) in zip(cols, mods[:3]):
        with col:
            st.markdown(
                f"<div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.3rem;"
                f"border-top:3px solid {accent};box-shadow:0 2px 8px rgba(26,61,26,.05)'>"
                f"<div style='font-size:1.6rem;margin-bottom:10px'>{ic}</div>"
                f"<div style='font-size:.9rem;font-weight:700;color:#1a2e1a;margin-bottom:6px'>{ti}</div>"
                f"<div style='font-size:.78rem;color:#7a9a7a;line-height:1.55'>{de}</div></div>",
                unsafe_allow_html=True)
    st.divider()
    st.info("💡 **Flux recommandé :** Calcul automatique → Export XLSX → Import dans Alertes & Stocks")


# ── Classification ABC/XYZ ────────────────────────────────────────────────────
def classify_abc(valeurs_an):
    total = sum(valeurs_an)
    if total == 0:
        return ["C"] * len(valeurs_an)
    idx_sort = sorted(range(len(valeurs_an)), key=lambda i: valeurs_an[i], reverse=True)
    classes = [""] * len(valeurs_an)
    cumul = 0
    for i in idx_sort:
        cumul += valeurs_an[i] / total
        if cumul <= 0.80:
            classes[i] = "A"
        elif cumul <= 0.95:
            classes[i] = "B"
        else:
            classes[i] = "C"
    return classes

def classify_xyz(cv_list):
    classes = []
    for cv in cv_list:
        if cv < 0.20:
            classes.append("X")
        elif cv < 0.50:
            classes.append("Y")
        else:
            classes.append("Z")
    return classes

MOIS_ABREV = ["Jan","Fév","Mar","Avr","Mai","Jun","Jul","Aoû","Sep","Oct","Nov","Déc"]
MOIS_FULL  = ["janvier","février","mars","avril","mai","juin",
               "juillet","août","septembre","octobre","novembre","décembre"]

def detect_month_cols(columns):
    month_cols = []
    for col in columns:
        cl = str(col).lower().strip()
        if cl in [f"m{i}" for i in range(1, 13)]:
            month_cols.append(col)
        elif any(cl.startswith(m.lower()) for m in MOIS_ABREV):
            month_cols.append(col)
        elif any(cl.startswith(m) for m in MOIS_FULL):
            month_cols.append(col)
    return month_cols


# ─────────────────────────────────────────────────────────────────────────────
def import_calcul():
    st.markdown("## 📥 Calcul automatique — Stock de sécurité & EOQ par article")
    st.markdown(
        "<p style='color:#7a9a7a'>Importez votre tableau de consommations sur 12 mois — "
        "le niveau de service est automatiquement déterminé selon la classe ABC de chaque article. "
        "Le coût de passation et le taux de stockage sont saisis <strong>par article</strong>.</p>",
        unsafe_allow_html=True)

    # Légende niveaux de service
    st.markdown("""
    <div style='background:#eef4ee;border:1px solid #d4e0d4;border-radius:10px;padding:.9rem 1.2rem;margin-bottom:1rem'>
        <div style='font-size:.7rem;color:#7a9a7a;text-transform:uppercase;font-weight:600;margin-bottom:.5rem'>Niveaux de service appliqués automatiquement</div>
        <div style='display:flex;gap:1.5rem;font-size:.82rem'>
            <span><span style='background:#c8e6c9;color:#1a3d1a;padding:2px 8px;border-radius:4px;font-weight:700'>A</span> → 99% (Z=2.33) · Articles stratégiques</span>
            <span><span style='background:#fff9c4;color:#7a5c00;padding:2px 8px;border-radius:4px;font-weight:700'>B</span> → 95% (Z=1.65) · Articles intermédiaires</span>
            <span><span style='background:#f5f5f5;color:#555;padding:2px 8px;border-radius:4px;font-weight:700'>C</span> → 90% (Z=1.28) · Articles courants</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("📋 Format attendu du fichier Excel"):
        st.markdown("""
**Le fichier peut avoir 1 ou 2 feuilles — l'application traite chaque feuille séparément.**

**Format recommandé — tableau de consommations 12 mois :**

| Code article | Description | Source | Coût unitaire | Coût passation | Taux stockage | Jan | Fév | … | Déc | Stock actuel | Commande en cours |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 40007 | NYSOSEL | Export | 59.23 | 150 | 0.20 | 310 | 324 | … | 318 | 1200 | 0 |

- Les colonnes mois peuvent être **Jan/Fév…**, **M1/M2…** ou **Janvier/Février…**
- `Source` : **Export** (4 mois) · **Local** (2 semaines) · **BM** (3 semaines)
- `Coût passation` : coût de lancement de commande par article (TND)
- `Taux stockage` : taux de possession annuel par article (ex: 0.20 pour 20%)
- Si ces colonnes sont absentes, vous pourrez les saisir manuellement
""")

    st.divider()
    st.markdown("#### 📂 Chargement du fichier")
    uploaded  = st.file_uploader("Déposez votre fichier Excel (.xlsx / .xls)",
                                  type=["xlsx", "xls"], key="ic_upload")
    use_demo  = st.checkbox("Utiliser les données de démo", value=not bool(uploaded), key="ic_demo")

    all_sheets_data = {}

    if uploaded:
        try:
            all_raw    = pd.read_excel(uploaded, sheet_name=None, header=None)
            sheet_names = list(all_raw.keys())

            if len(sheet_names) > 1:
                st.info(f"📋 **{len(sheet_names)} feuilles détectées :** {', '.join(sheet_names)}")
                feuille_active    = st.radio("Feuille à analyser", sheet_names, horizontal=True, key="ic_sheet_sel")
                sheets_to_process = [feuille_active]
            else:
                sheets_to_process = sheet_names

            for sheet_sel in sheets_to_process:
                raw  = all_raw[sheet_sel]
                hrow = 0
                for i, row in raw.iterrows():
                    row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                    if any(k in row_str for k in ["code", "article", "jan", "m1", "source"]):
                        hrow = i
                        break

                df_raw = pd.read_excel(uploaded, sheet_name=sheet_sel, header=hrow)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                df_raw = df_raw.dropna(how="all")

                cmap = {}
                for col in df_raw.columns:
                    cl = col.lower().strip()
                    if "code" in cl and "article" in cl:                                    cmap["Code article"] = col
                    elif "description" in cl or "désignation" in cl:                        cmap["Description Article"] = col
                    elif cl == "source":                                                     cmap["Source"] = col
                    elif "coût unitaire" in cl or "cout unitaire" in cl or cl == "prix":    cmap["Coût unitaire"] = col
                    elif "coût passation" in cl or "cout passation" in cl or "passation" in cl: cmap["Coût passation"] = col
                    elif "taux stockage" in cl or "taux de stockage" in cl:                 cmap["Taux stockage"] = col
                    elif "stock actuel" in cl or ("stock" in cl and "actuel" in cl):        cmap["Stock actuel"] = col
                    elif "commande en cours" in cl or "cmd en cours" in cl:                 cmap["Commande en cours"] = col

                df_s = df_raw.rename(columns={v: k for k, v in cmap.items()})
                month_cols = detect_month_cols(df_s.columns)

                for c in ["Coût unitaire", "Coût passation", "Taux stockage", "Stock actuel", "Commande en cours"] + month_cols:
                    if c in df_s.columns:
                        df_s[c] = pd.to_numeric(df_s[c], errors="coerce").fillna(0)

                if "Code article" in df_s.columns:
                    df_s = df_s[df_s["Code article"].notna()]
                    df_s = df_s[df_s["Code article"].astype(str).str.strip() != ""]

                if "Source" not in df_s.columns:
                    df_s["Source"] = "Export"

                df_s.attrs["month_cols"] = month_cols
                all_sheets_data[sheet_sel] = df_s
                st.success(f"✅ Feuille **{sheet_sel}** : {len(df_s)} articles · "
                           f"{len(month_cols)} mois détectés ({', '.join(month_cols[:3])}{'…' if len(month_cols)>3 else ''})")

        except Exception as e:
            st.error(f"Erreur de lecture : {e}")

    # ── Démo ─────────────────────────────────────────────────────────────────
    if use_demo and not all_sheets_data:
        rng = np.random.default_rng(42)
        base = [324, 5886, 509, 32460, 959, 3747, 3147, 4432, 16, 22]
        month_data = {f"M{i+1}": [int(b * (0.85 + 0.30 * rng.random())) for b in base]
                      for i in range(12)}
        demo_df = pd.DataFrame({
            "Code article":        [40007,40010,40011,40014,40015,40036,40042,40055,40062,40063],
            "Description Article": ["NYSOSEL","TERRE TONSIL","PAPIER FILTRE","TOILE FILTRANTE",
                                    "PERLITE PAF1","SEL MARIN FIN","MYVEROL 18-04","TRISYL",
                                    "VITAMINE A","VITAMINE E"],
            "Source":              ["Export","Export","Local","Export","BM","Local","Export","BM","Export","Local"],
            "Coût unitaire":       [59.23,1.709,3.823,49.787,1.700,0.85,2.10,8.189,12.5,18.0],
            "Coût passation":      [200,150,100,250,120,80,130,110,180,90],
            "Taux stockage":       [0.22,0.18,0.15,0.25,0.20,0.15,0.18,0.20,0.22,0.18],
            "Stock actuel":        [1200,8000,500,50000,2000,5000,8000,9000,30,50],
            "Commande en cours":   [0,5000,0,0,1000,0,3000,0,0,0],
            **month_data
        })
        demo_df.attrs["month_cols"] = [f"M{i+1}" for i in range(12)]
        all_sheets_data["Démo"] = demo_df
        st.info("📊 Données de démo chargées — coût passation et taux stockage inclus par article")

    if not all_sheets_data:
        return

    for sheet_name, df_source in all_sheets_data.items():
        if len(all_sheets_data) > 1:
            st.markdown(f"---\n### 📄 Feuille : **{sheet_name}**")

        month_cols = df_source.attrs.get("month_cols", [])

        # ── Saisie coût unitaire si absent ───────────────────────────────────
        if "Coût unitaire" not in df_source.columns or df_source["Coût unitaire"].sum() == 0:
            st.warning("⚠️ Colonne **Coût unitaire** non détectée — saisissez les valeurs :")
            cu_vals   = {}
            cols_cu   = st.columns(min(4, len(df_source)))
            for i, (_, row) in enumerate(df_source.iterrows()):
                with cols_cu[i % len(cols_cu)]:
                    cu_vals[str(row.get("Code article", i))] = st.number_input(
                        f"Coût unit. {row.get('Code article','Art'+str(i))}",
                        value=0.0, min_value=0.0, step=0.1,
                        key=f"cu_{sheet_name}_{i}")
            df_source = df_source.copy()
            df_source["Coût unitaire"] = df_source["Code article"].astype(str).map(cu_vals).fillna(0)

        # ── Aperçu tableau source ────────────────────────────────────────────
        with st.expander(f"👁 Tableau de consommations — {sheet_name}", expanded=True):
            disp_cols = (["Code article", "Description Article", "Source", "Coût unitaire"]
                         + (["Coût passation"] if "Coût passation" in df_source.columns else [])
                         + (["Taux stockage"]  if "Taux stockage"  in df_source.columns else [])
                         + month_cols
                         + [c for c in ["Stock actuel", "Commande en cours"] if c in df_source.columns])
            st.dataframe(df_source[[c for c in disp_cols if c in df_source.columns]],
                         use_container_width=True, hide_index=True)

        # ── Première passe : classification ABC pour déterminer Z ─────────────
        # Calcul préliminaire de la valeur annuelle pour ABC
        prelim_vals = []
        for _, row in df_source.iterrows():
            cu = float(row.get("Coût unitaire", 0) or 0)
            if month_cols:
                conso_mois = [float(row.get(mc, 0) or 0) for mc in month_cols]
            else:
                conso_mois = [float(row.get("Consommation", 0) or 0)]
            d_moy = np.mean(conso_mois)
            prelim_vals.append(d_moy * 12 * cu)

        abc_classes_prelim = classify_abc(prelim_vals)

        # ── Saisie coût passation & taux stockage si absents dans le fichier ──
        has_sc  = "Coût passation" in df_source.columns and df_source["Coût passation"].sum() > 0
        has_taux = "Taux stockage"  in df_source.columns and df_source["Taux stockage"].sum() > 0

        df_source = df_source.copy()

        if not has_sc or not has_taux:
            st.markdown("---")
            st.warning("⚠️ **Coût de passation et/ou taux de stockage non détectés** — "
                       "saisissez les valeurs par article ci-dessous.")
            st.markdown("""
<div style='background:#fff9c4;border:1px solid #c9a000;border-radius:8px;padding:.7rem 1rem;margin-bottom:.8rem;font-size:.82rem;color:#7a5c00'>
💡 <strong>Ces paramètres varient selon les articles</strong> : le coût de passation dépend du fournisseur et des procédures d'achat, 
le taux de stockage dépend du coût de capital immobilisé, des coûts d'entreposage et de l'obsolescence propres à chaque produit.
</div>""", unsafe_allow_html=True)

            sc_vals   = {}
            taux_vals = {}
            n_art     = len(df_source)
            cols_per_row = min(4, n_art)

            for chunk_start in range(0, n_art, cols_per_row):
                chunk = df_source.iloc[chunk_start:chunk_start + cols_per_row]
                cols_row = st.columns(cols_per_row)
                for col_i, (_, row) in enumerate(chunk.iterrows()):
                    art_key = str(row.get("Code article", col_i))
                    abc_cls = abc_classes_prelim[list(df_source.index).index(row.name)] if row.name in df_source.index else "B"
                    abc_colors_map = {"A": "#c8e6c9", "B": "#fff9c4", "C": "#f5f5f5"}
                    with cols_row[col_i]:
                        st.markdown(
                            f"<div style='background:{abc_colors_map.get(abc_cls,'#f5f5f5')};"
                            f"border-radius:6px;padding:.4rem .6rem;margin-bottom:.3rem;"
                            f"font-size:.78rem;font-weight:700;color:#1a2e1a'>"
                            f"{row.get('Code article','')} <span style='font-weight:400'>[Classe {abc_cls}]</span></div>",
                            unsafe_allow_html=True)
                        if not has_sc:
                            sc_vals[art_key] = st.number_input(
                                f"Passation (TND)",
                                value=150.0, min_value=0.0, step=10.0,
                                key=f"sc_{sheet_name}_{art_key}",
                                help="Coût de lancement d'une commande pour cet article")
                        if not has_taux:
                            taux_vals[art_key] = st.number_input(
                                f"Taux stockage (%)",
                                value=20.0, min_value=0.0, step=1.0,
                                key=f"tx_{sheet_name}_{art_key}",
                                help="Taux de possession annuel en % du coût unitaire") / 100

            if not has_sc:
                df_source["Coût passation"] = df_source["Code article"].astype(str).map(sc_vals).fillna(150)
            if not has_taux:
                df_source["Taux stockage"]  = df_source["Code article"].astype(str).map(taux_vals).fillna(0.20)

        # ── Calculs par article ───────────────────────────────────────────────
        results = []
        for idx, (_, row) in enumerate(df_source.iterrows()):
            source = str(row.get("Source", "Export")).strip()
            cu     = float(row.get("Coût unitaire", 0) or 0)
            sc_art = float(row.get("Coût passation", 150) or 150)
            tx_art = float(row.get("Taux stockage", 0.20) or 0.20)
            stk    = float(row.get("Stock actuel", 0) or 0)
            cmd    = float(row.get("Commande en cours", 0) or 0)
            lt_j   = get_lt_jours(source)

            if month_cols:
                conso_mois = [float(row.get(mc, 0) or 0) for mc in month_cols]
            else:
                conso_mois = [float(row.get("Consommation", 0) or 0)]

            n_mois   = len(conso_mois)
            d_moy    = np.mean(conso_mois)
            d_moy_j  = d_moy / 30
            d_min    = min(conso_mois)
            d_max    = max(conso_mois)
            sigma_m  = float(np.std(conso_mois, ddof=1)) if n_mois > 1 else d_moy * 0.15
            sigma_j  = sigma_m / 30
            cv       = sigma_m / d_moy if d_moy > 0 else 0
            conso_an = d_moy * 12

            # Classe ABC déterminée en première passe
            abc_cls = abc_classes_prelim[idx]
            # Niveau de service automatique selon classe ABC
            ns_info = ABC_SERVICE_LEVEL[abc_cls]
            Z       = ns_info["Z"]
            NS_pct  = ns_info["pct"]

            ss       = calc_ss(sigma_j, lt_j, Z)
            pr       = calc_pr(d_moy_j, lt_j, ss)
            eoq, nb_cmd, _ = calc_eoq(conso_an, cu, sc_cost=sc_art, taux=tx_art)
            css      = round(ss * cu, 2) if cu > 0 else 0

            results.append({
                "Code article":        row.get("Code article", ""),
                "Description Article": row.get("Description Article", ""),
                "Source":              source,
                "LT jours":            lt_j,
                "Conso mois moy":      round(d_moy, 1),
                "Conso min":           round(d_min, 1),
                "Conso max":           round(d_max, 1),
                "Sigma mois":          round(sigma_m, 2),
                "Sigma D (j)":         round(sigma_j, 4),
                "CV":                  round(cv, 3),
                "D jour":              round(d_moy_j, 4),
                "Conso annuelle":      round(conso_an, 1),
                "Niveau service pct":  NS_pct,
                "Niveau service label": ns_info["label"],
                "Z val":               Z,
                "Coût passation":      sc_art,
                "Taux stockage":       tx_art,
                "Stock sécurité":      ss,
                "Coût unitaire":       cu,
                "Coût SS":             css,
                "EOQ":                 eoq,
                "Nb cmd an":           nb_cmd,
                "Point de commande":   pr,
                "Stock actuel":        stk,
                "Commande en cours":   cmd,
                "Valeur annuelle":     round(conso_an * cu, 2),
                "_conso_mois":         conso_mois,
            })

        df_res = pd.DataFrame(results)

        # Reclasser ABC (avec les vraies valeurs annuelles maintenant)
        df_res["Classe ABC"] = classify_abc(df_res["Valeur annuelle"].tolist())
        # Mettre à jour Z et SS selon la classe ABC finale
        for i, row in df_res.iterrows():
            abc = row["Classe ABC"]
            ns_info = ABC_SERVICE_LEVEL[abc]
            Z_final = ns_info["Z"]
            sigma_j = row["Sigma D (j)"]
            lt_j    = row["LT jours"]
            d_moy_j = row["D jour"]
            cu      = row["Coût unitaire"]
            sc_art  = row["Coût passation"]
            tx_art  = row["Taux stockage"]
            conso_an = row["Conso annuelle"]

            ss_final = calc_ss(sigma_j, lt_j, Z_final)
            pr_final = calc_pr(d_moy_j, lt_j, ss_final)
            eoq_final, nb_final, _ = calc_eoq(conso_an, cu, sc_cost=sc_art, taux=tx_art)

            df_res.at[i, "Z val"]              = Z_final
            df_res.at[i, "Niveau service pct"] = ns_info["pct"]
            df_res.at[i, "Niveau service label"] = ns_info["label"]
            df_res.at[i, "Stock sécurité"]    = ss_final
            df_res.at[i, "Coût SS"]           = round(ss_final * cu, 2) if cu > 0 else 0
            df_res.at[i, "Point de commande"] = pr_final
            df_res.at[i, "EOQ"]               = eoq_final
            df_res.at[i, "Nb cmd an"]         = nb_final

        df_res["Classe XYZ"] = classify_xyz(df_res["CV"].tolist())
        df_res["Classe"]     = df_res["Classe ABC"] + df_res["Classe XYZ"]

        # ════════════════════════════════════════════════════════════════════
        # SECTION 1 — ANALYSE STATISTIQUE DES CONSOMMATIONS
        # ════════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("### 📊 Analyse des consommations sur 12 mois")

        df_stats = df_res[[
            "Code article", "Description Article", "Source",
            "Conso mois moy", "Conso min", "Conso max",
            "Sigma mois", "CV", "Classe ABC", "Classe XYZ", "Classe",
            "Niveau service label"
        ]].copy()

        st.dataframe(df_stats, use_container_width=True, hide_index=True,
            column_config={
                "Conso mois moy":      st.column_config.NumberColumn("Moy. mensuelle (u)", format="%.1f"),
                "Conso min":           st.column_config.NumberColumn("Min (u)",  format="%.1f"),
                "Conso max":           st.column_config.NumberColumn("Max (u)",  format="%.1f"),
                "Sigma mois":          st.column_config.NumberColumn("σ mensuel (u)", format="%.2f"),
                "CV":                  st.column_config.NumberColumn("CV (%)", format="%.1%%"),
                "Classe ABC":          st.column_config.TextColumn("ABC"),
                "Classe XYZ":          st.column_config.TextColumn("XYZ"),
                "Classe":              st.column_config.TextColumn("Classe"),
                "Niveau service label": st.column_config.TextColumn("Niveau service"),
            })

        st.markdown("""
<div style='display:grid;grid-template-columns:1fr 1fr;gap:.8rem;margin:.8rem 0'>
  <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:10px;padding:.9rem 1.1rem'>
    <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:.5rem;font-weight:600'>Classification ABC — Valeur annuelle → Niveau service</div>
    <div style='font-size:.82rem;line-height:1.9'>
      <span style='background:#c8e6c9;color:#1a3d1a;padding:2px 8px;border-radius:4px;font-weight:700'>A</span>&nbsp; Top 80% de la valeur → <strong>99% (Z=2.33)</strong><br/>
      <span style='background:#fff9c4;color:#7a5c00;padding:2px 8px;border-radius:4px;font-weight:700'>B</span>&nbsp; 80–95% → <strong>95% (Z=1.65)</strong><br/>
      <span style='background:#f5f5f5;color:#7a9a7a;padding:2px 8px;border-radius:4px;font-weight:700'>C</span>&nbsp; 95–100% → <strong>90% (Z=1.28)</strong>
    </div>
  </div>
  <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:10px;padding:.9rem 1.1rem'>
    <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:.5rem;font-weight:600'>Classification XYZ — Variabilité (CV)</div>
    <div style='font-size:.82rem;line-height:1.9'>
      <span style='background:#e3f2fd;color:#1a237e;padding:2px 8px;border-radius:4px;font-weight:700'>X</span>&nbsp; CV &lt; 20% — demande stable<br/>
      <span style='background:#fce4ec;color:#880e4f;padding:2px 8px;border-radius:4px;font-weight:700'>Y</span>&nbsp; CV 20–50% — demande variable<br/>
      <span style='background:#fff3e0;color:#e65100;padding:2px 8px;border-radius:4px;font-weight:700'>Z</span>&nbsp; CV &gt; 50% — demande très irrégulière
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

        # ════════════════════════════════════════════════════════════════════
        # SECTION 2 — DÉTAIL PAR ARTICLE
        # ════════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("### 🔍 Détail par article — sélectionner un code")

        art_labels = [
            f"{row['Code article']} — {row['Description Article']} [{row['Classe']}]"
            for _, row in df_res.iterrows()
        ]
        sel_label = st.selectbox("Choisir un article", art_labels, key=f"sel_art_{sheet_name}")
        sel_idx   = art_labels.index(sel_label)
        art       = df_res.iloc[sel_idx]
        conso_mois_art = art["_conso_mois"]
        n_mois_art     = len(conso_mois_art)
        mois_x         = MOIS_ABREV[:n_mois_art] if n_mois_art <= 12 else [f"M{i+1}" for i in range(n_mois_art)]

        abc_colors_disp = {"A": C_GREEN, "B": C_GOLD, "C": C_MUTED}
        xyz_colors_disp = {"X": "#1565c0", "Y": "#880e4f", "Z": "#e65100"}
        abc_col = abc_colors_disp.get(art["Classe ABC"], C_GREEN)
        xyz_col = xyz_colors_disp.get(art["Classe XYZ"], C_MUTED)

        st.markdown(
            f"<div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;"
            f"padding:1.1rem 1.4rem;margin-bottom:1rem;border-left:4px solid {abc_col}'>"
            f"<span style='font-size:1.1rem;font-weight:700;color:#1a2e1a'>"
            f"{art['Code article']} — {art['Description Article']}</span>"
            f"&nbsp;&nbsp;<span style='background:{abc_col};color:#fff;padding:3px 10px;"
            f"border-radius:20px;font-size:.8rem;font-weight:700'>{art['Classe ABC']}</span>"
            f"&nbsp;<span style='background:{xyz_col};color:#fff;padding:3px 10px;"
            f"border-radius:20px;font-size:.8rem;font-weight:700'>{art['Classe XYZ']}</span>"
            f"&nbsp;&nbsp;<span style='font-size:.82rem;color:#7a9a7a'>Source : {art['Source']} "
            f"· Délai : {art['LT jours']}j"
            f"· Niveau service : <strong>{art['Niveau service label']}</strong> (Z={art['Z val']})"
            f"· Passation : {art['Coût passation']:.0f} TND"
            f"· Taux stockage : {art['Taux stockage']*100:.0f}%</span></div>",
            unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.markdown(mcard("Demande moy./mois", fmt(art["Conso mois moy"], 1), "u/mois", color=C_GREEN),  unsafe_allow_html=True)
        with c2: st.markdown(mcard("Demande moy./jour", fmt(art["D jour"], 3),          "u/j",    color=C_GREEN2), unsafe_allow_html=True)
        with c3: st.markdown(mcard("σ mensuelle",       fmt(art["Sigma mois"], 2),       "u/mois", color=C_PURPLE), unsafe_allow_html=True)
        with c4: st.markdown(mcard("σ journalière",     fmt(art["Sigma D (j)"], 4),      "u/j",    color=C_PURPLE), unsafe_allow_html=True)
        with c5: st.markdown(mcard("Coeff. variation",  f"{art['CV']*100:.1f}%",         "CV",     color=xyz_col),  unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:.8rem'></div>", unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.markdown(mcard("Stock de sécurité", fmtInt(art["Stock sécurité"]),    "unités",   color=C_GOLD),   unsafe_allow_html=True)
        with c2: st.markdown(mcard("Point de commande", fmtInt(art["Point de commande"]), "unités",   color=C_RED),    unsafe_allow_html=True)
        with c3: st.markdown(mcard("EOQ",               fmtInt(art["EOQ"]),               "u/cmd",    color=C_GREEN),  unsafe_allow_html=True)
        with c4: st.markdown(mcard("Coût SS",           f"{art['Coût SS']:,.2f}",          "TND",      color=C_GOLD),   unsafe_allow_html=True)
        with c5: st.markdown(mcard("Conso annuelle",    fmtInt(art["Conso annuelle"]),     "u/an",     color=C_GREEN2), unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:.8rem'></div>", unsafe_allow_html=True)

        st.markdown(
            fbox(f"Calcul SS — {art['Code article']} [Classe {art['Classe ABC']} → NS {art['Niveau service label']}]",
                 f"SS = Z × σD × √L = {art['Z val']} × {art['Sigma D (j)']} × √{art['LT jours']}",
                 f"= {art['Z val']} × {art['Sigma D (j)']} × {round(math.sqrt(art['LT jours']),3)}"
                 f" = <strong>{art['Stock sécurité']}</strong> unités"),
            unsafe_allow_html=True)

        st.markdown(
            fbox(f"Calcul EOQ — {art['Code article']}",
                 f"EOQ = √(2 × D × Sc / (h × Cu)) = √(2 × {art['Conso annuelle']:.0f} × {art['Coût passation']:.0f} / ({art['Taux stockage']*100:.0f}% × {art['Coût unitaire']:.3f}))",
                 f"Passation : {art['Coût passation']:.0f} TND · Taux : {art['Taux stockage']*100:.0f}% · EOQ = <strong>{art['EOQ']}</strong> u"),
            unsafe_allow_html=True)

        # Graphique consommations
        fig_art = go.Figure()
        fig_art.add_trace(go.Bar(
            x=mois_x, y=conso_mois_art,
            name="Consommation",
            marker_color=[C_RED if c < art["Stock sécurité"] / 30 else C_GREEN for c in conso_mois_art],
        ))
        fig_art.add_hline(y=art["Conso mois moy"],
                           line_dash="dash", line_color=C_GREEN2, line_width=2,
                           annotation_text=f"Moy. = {art['Conso mois moy']:.1f} u",
                           annotation_font_color=C_GREEN2)
        fig_art.add_hline(y=art["Conso mois moy"] + art["Sigma mois"],
                           line_dash="dot", line_color=C_PURPLE, line_width=1.5,
                           annotation_text=f"Moy+σ = {art['Conso mois moy']+art['Sigma mois']:.1f}",
                           annotation_font_color=C_PURPLE)
        fig_art.add_hline(y=max(0, art["Conso mois moy"] - art["Sigma mois"]),
                           line_dash="dot", line_color=C_PURPLE, line_width=1.5,
                           annotation_text=f"Moy−σ = {max(0,art['Conso mois moy']-art['Sigma mois']):.1f}",
                           annotation_font_color=C_PURPLE)
        fig_art.update_layout(**light_layout(), height=320,
            xaxis=dict(title="Mois", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="Consommation (u)", gridcolor="rgba(0,0,0,.06)"),
            title=f"Historique consommations — {art['Code article']}")
        st.plotly_chart(fig_art, use_container_width=True)

        # ════════════════════════════════════════════════════════════════════
        # SECTION 3 — INDICATEURS SC COMPLETS
        # ════════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("### 🧮 Indicateurs SC — tous articles")
        disp_sc = df_res[[
            "Code article", "Description Article", "Source", "Classe",
            "Niveau service label", "LT jours", "Conso mois moy",
            "Coût passation", "Taux stockage",
            "Stock sécurité", "Coût SS", "EOQ", "Point de commande",
            "Stock actuel", "Commande en cours"
        ]].copy()
        st.dataframe(disp_sc, use_container_width=True, hide_index=True,
            column_config={
                "Niveau service label": st.column_config.TextColumn("NS"),
                "LT jours":           st.column_config.NumberColumn("Délai (j)", format="%d"),
                "Conso mois moy":     st.column_config.NumberColumn("Conso moy./mois", format="%.1f"),
                "Coût passation":     st.column_config.NumberColumn("Passation (TND)", format="%.0f"),
                "Taux stockage":      st.column_config.NumberColumn("Taux stk (%)", format="%.0%%"),
                "Stock sécurité":     st.column_config.NumberColumn("SS (u)", format="%d"),
                "Coût SS":            st.column_config.NumberColumn("Coût SS (TND)", format="%.2f"),
                "EOQ":                st.column_config.NumberColumn("EOQ (u)", format="%d"),
                "Point de commande":  st.column_config.NumberColumn("Point commande (u)", format="%d"),
                "Stock actuel":       st.column_config.NumberColumn("Stock actuel (u)", format="%d"),
                "Commande en cours":  st.column_config.NumberColumn("Cmd en cours (u)", format="%d"),
            })

        labels_g = df_res["Code article"].astype(str).tolist()
        fig_g = go.Figure()
        fig_g.add_trace(go.Bar(name="Stock actuel", x=labels_g, y=df_res["Stock actuel"],   marker_color=C_GREEN2))
        fig_g.add_trace(go.Bar(name="EOQ",          x=labels_g, y=df_res["EOQ"],            marker_color=C_GREEN))
        fig_g.add_trace(go.Scatter(name="Stock sécu.", x=labels_g, y=df_res["Stock sécurité"],
            mode="markers+lines", marker=dict(color=C_GOLD, size=9, symbol="diamond"),
            line=dict(color=C_GOLD, width=2, dash="dash")))
        fig_g.add_trace(go.Scatter(name="Point commande", x=labels_g, y=df_res["Point de commande"],
            mode="markers+lines", marker=dict(color=C_RED, size=8, symbol="x"),
            line=dict(color=C_RED, width=1.5, dash="dot")))
        fig_g.update_layout(**light_layout(), barmode="group", height=380,
            xaxis=dict(title="Article", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="Unités",  gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig_g, use_container_width=True)

        # ════════════════════════════════════════════════════════════════════
        # SECTION 4 — SYNTHÈSE
        # ════════════════════════════════════════════════════════════════════
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        abc_A = (df_res["Classe ABC"] == "A").sum()
        abc_B = (df_res["Classe ABC"] == "B").sum()
        abc_C = (df_res["Classe ABC"] == "C").sum()
        with c1: st.markdown(mcard("Articles traités",   str(len(df_res)),               "articles",        color=C_GREEN),  unsafe_allow_html=True)
        with c2: st.markdown(mcard("Coût total SS",      f"{df_res['Coût SS'].sum():,.0f}", "TND immobilisés", color=C_GOLD),   unsafe_allow_html=True)
        with c3: st.markdown(mcard("EOQ moyen",          fmtInt(df_res[df_res['EOQ']>0]['EOQ'].mean()), "u / commande", color=C_GREEN2), unsafe_allow_html=True)
        with c4: st.markdown(mcard("Répartition ABC",    f"A:{abc_A} B:{abc_B} C:{abc_C}", "articles",       color=C_PURPLE), unsafe_allow_html=True)

        # ════════════════════════════════════════════════════════════════════
        # SECTION 5 — EXPORT
        # ════════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("### 💾 Export XLSX — Indicateurs SC")

        df_res_exp = df_res.drop(columns=["_conso_mois"], errors="ignore")
        buf = build_excel_complet(df_res_exp)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel complet",
            data=buf,
            file_name=f"medoil_sc_{sheet_name}_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        if "df_resultats" not in st.session_state:
            st.session_state["df_resultats"] = {}
        st.session_state["df_resultats"][sheet_name] = df_res_exp


# ─────────────────────────────────────────────────────────────────────────────
def calculateurs():
    st.markdown("## ⚙️ Calculateurs Supply Chain")
    tab_ss, tab_eoq, tab_kpi, tab_rp = st.tabs(
        ["📦 Stock de sécurité", "📐 EOQ (Wilson)", "📊 KPIs stock", "🔄 Point de réappro."])

    with tab_ss:
        st.markdown(fbox("Formule", "SS = Z × σD × √L",
                         "Z déterminé par classe ABC : A→2.33 (99%) · B→1.65 (95%) · C→1.28 (90%) · σD = écart-type demande (u/j) · L = délai (j)"), unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            source_ss = st.selectbox("Source approvisionnement", ["Export", "Local", "BM"], key="ss_source")
            lt_auto   = get_lt_jours(source_ss)
            st.info(f"Délai automatique : **{lt_auto} jours**")
            sd   = st.number_input("Demande moy. (u/j)",          value=50.0, step=1.0, key="ss_sd")
            ss2  = st.number_input("Écart-type demande σD (u/j)", value=8.0,  step=0.5, key="ss_sigma")
        with c2:
            scu      = st.number_input("Coût unitaire (TND)",  value=25.0, step=1.0, key="ss_cu")
            abc_sel  = st.selectbox("Classe ABC", ["A", "B", "C"], key="ss_abc",
                                    help="Détermine automatiquement le niveau de service")
        with c3:
            ns_auto = ABC_SERVICE_LEVEL[abc_sel]
            st.markdown(
                f"<div style='background:#eef4ee;border:1px solid #d4e0d4;border-radius:8px;padding:.9rem;margin-top:1.5rem'>"
                f"<div style='font-size:.65rem;color:#7a9a7a;text-transform:uppercase;margin-bottom:4px'>Niveau service appliqué</div>"
                f"<div style='font-size:1.4rem;font-weight:700;color:#1a2e1a'>{ns_auto['label']}</div>"
                f"<div style='font-size:.82rem;color:#3d5c3d'>Z = <strong>{ns_auto['Z']}</strong> · Classe <strong>{abc_sel}</strong></div>"
                f"</div>",
                unsafe_allow_html=True)

        Z2      = ns_auto["Z"]
        SS_calc = Z2 * ss2 * math.sqrt(lt_auto)

        c = st.columns(3)
        for i, (l, v, u, col) in enumerate([
            ("SS calculé",         fmtInt(SS_calc),       "unités", C_GREEN),
            ("Z × σD × √L",        f"{Z2} × {ss2} × √{lt_auto}", "détail calcul", C_GREEN2),
            ("Coût immobilisé SS", fmtInt(SS_calc * scu), "TND",    C_GOLD),
        ]):
            with c[i]: st.markdown(mcard(l, v, u, color=col), unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:14px'></div>", unsafe_allow_html=True)
        z_vals  = [1.28, 1.65, 2.33]
        ns_vals = [90, 95, 99]
        ns_labs = ["C → 90%", "B → 95%", "A → 99%"]
        ss_v    = [round(z * ss2 * math.sqrt(lt_auto)) for z in z_vals]
        bar_cols = [C_MUTED, C_GOLD, C_GREEN]
        fig = go.Figure(go.Bar(
            x=ns_labs, y=ss_v,
            marker_color=[C_GOLD if z == Z2 else C_GREEN if z == 2.33 else C_MUTED for z in z_vals],
            text=ss_v, textposition="outside"))
        fig.update_layout(**light_layout(), height=260,
            xaxis=dict(title="Classe ABC → Niveau de service", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="SS (unités)", gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)

    with tab_eoq:
        st.markdown(fbox("Modèle de Wilson", "EOQ = √(2×D×Sc / (Sh×Cu))",
                         "D=demande annuelle · Sc=coût passation (propre à l'article) · Sh=taux stockage (propre à l'article) · Cu=coût unitaire"), unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            eD  = st.number_input("Demande annuelle (u)",        value=12000, step=100,  key="eoq_D")
            eSc = st.number_input("Coût passation article (TND)", value=150.0, step=5.0,  key="eoq_Sc",
                                  help="Coût spécifique à cet article/fournisseur")
        with c2:
            eCu = st.number_input("Coût unitaire (TND)",  value=80.0, step=1.0,  key="eoq_Cu")
            eSh = st.number_input("Taux stockage article (%)", value=20.0, step=1.0, key="eoq_Sh",
                                  help="Taux de possession propre à cet article")
        with c3:
            eLT = st.number_input("Délai fourni. (j)", value=10,  step=1, key="eoq_LT")
            eDy = st.number_input("Jours ouvr./an",   value=250, step=1, key="eoq_Dy")

        h   = (eSh / 100) * eCu
        EOQ = math.sqrt(2 * eD * eSc / h) if h > 0 else 0
        nb  = eD / EOQ if EOQ > 0 else 0
        T   = eDy / nb if nb > 0 else 0
        Cp  = nb * eSc; Cs = (EOQ / 2) * h

        st.markdown("---")
        c = st.columns(3)
        data = [("EOQ",               fmtInt(EOQ),     "u/cmd",  C_GREEN),
                ("Nb cmd/an",         fmt(nb, 1),      "cmds",   C_MUTED),
                ("Périodicité",       fmtInt(T),       "jours",  C_MUTED),
                ("Coût passation/an", fmtInt(Cp),      "TND",    C_GOLD),
                ("Coût stockage/an",  fmtInt(Cs),      "TND",    C_GOLD),
                ("Coût total min.",   fmtInt(Cp + Cs), "TND",    C_GREEN2)]
        for i, (l, v, u, col) in enumerate(data):
            with c[i % 3]:
                st.markdown(mcard(l, v, u, color=col), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px'></div>", unsafe_allow_html=True)

        qty_r = np.linspace(max(10, EOQ * .2), EOQ * 2.5, 200)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=qty_r, y=[(eD / q) * eSc for q in qty_r],             name="Passation", line=dict(color=C_GREEN, width=2)))
        fig.add_trace(go.Scatter(x=qty_r, y=[(q / 2) * h    for q in qty_r],             name="Stockage",  line=dict(color=C_GOLD, width=2)))
        fig.add_trace(go.Scatter(x=qty_r, y=[(eD / q) * eSc + (q / 2) * h for q in qty_r], name="Total", line=dict(color=C_GREEN2, width=2.5)))
        fig.add_vline(x=EOQ, line_dash="dash", line_color=C_RED,
                      annotation_text=f"EOQ={fmtInt(EOQ)}", annotation_font_color=C_RED)
        fig.update_layout(**light_layout(), height=300,
            xaxis=dict(title="Quantité (u)", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="Coût/an (TND)", gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)

    with tab_kpi:
        c1, c2 = st.columns(2)
        with c1:
            ks   = st.number_input("Stock moyen (u)",       value=3000,    step=100,   key="kpi_ks")
            kv   = st.number_input("Ventes annuelles (u)",  value=18000,   step=100,   key="kpi_kv")
            kval = st.number_input("Valeur stock (TND)",    value=240000,  step=1000,  key="kpi_kval")
            kca  = st.number_input("CA annuel (TND)",       value=2400000, step=10000, key="kpi_kca")
        with c2:
            kot  = st.number_input("Livrées à temps",       value=465,     step=1,     key="kpi_kot")
            kto  = st.number_input("Total livrées",         value=490,     step=1,     key="kpi_kto")
            klc  = st.number_input("Coût logistique (TND)", value=168000,  step=1000,  key="kpi_klc")
            ksh  = st.number_input("Taux stockage (%)",     value=22.0,    step=1.0,   key="kpi_ksh")

        rot = kv / ks if ks > 0 else 0
        cov = (ks / kv) * 365 if kv > 0 else 0
        srv = (kot / kto) * 100 if kto > 0 else 0
        crt = (klc / kca) * 100 if kca > 0 else 0

        st.markdown("---")
        c = st.columns(3)
        kdata = [
            ("Rotation",         fmt(rot, 1) + "×",       "/an · >6×=excellent", C_GREEN2 if rot >= 6 else C_GOLD),
            ("Couverture",       fmtInt(cov) + "j",        "15–30j=optimal",      C_MUTED),
            ("Taux service OTD", fmt(srv, 1) + "%",        ">95%=cible",           C_GREEN2 if srv >= 95 else C_GOLD),
            ("Coût log/CA",      fmt(crt, 1) + "%",        "objectif <8%",         C_GREEN2 if crt < 8 else C_RED),
            ("Coût stockage/an", fmtInt(kval * (ksh / 100)), "TND",               C_GOLD),
            ("DSI",              fmt(ks / (kv / 365), 0) + "j", "jours de stock", C_MUTED),
        ]
        for i, (l, v, u, col) in enumerate(kdata):
            with c[i % 3]:
                st.markdown(mcard(l, v, u, color=col), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px'></div>", unsafe_allow_html=True)

    with tab_rp:
        st.markdown(fbox("Formule", "PR = (Dmoy × LT) + SS",
                         "Dmoy = demande/jour · LT = délai (j) · SS = stock de sécurité"), unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            src_rp = st.selectbox("Source", ["Export", "Local", "BM"], key="rp_src")
            lt_rp  = get_lt_jours(src_rp)
            st.info(f"Délai auto : **{lt_rp}j**")
            rd  = st.number_input("Demande moy. (u/j)", value=50.0, step=1.0,  key="rp_rd")
        with c2:
            rss = st.number_input("Stock sécu. (u)",  value=132.0, step=10.0, key="rp_rss")
            rcu = st.number_input("Stock actuel (u)", value=500.0, step=10.0, key="rp_rcu")
        with c3:
            rdm  = st.number_input("Demande max (u/j)", value=70.0,  step=1.0,  key="rp_rdm")

        PR = rd * lt_rp + rss
        dL = (rcu - rss) / rd if rd > 0 else 0
        if rcu <= rss:
            st.markdown("<div class='alert-critical'>🔴 RUPTURE IMMINENTE — Commander immédiatement !</div>", unsafe_allow_html=True)
        elif rcu <= PR:
            st.markdown("<div class='alert-warning'>🟡 COMMANDER MAINTENANT — Point de réappro. atteint</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='alert-ok'>🟢 Stock suffisant — {fmtInt(rcu)} u > PR ({fmtInt(PR)} u)</div>", unsafe_allow_html=True)

        c = st.columns(3)
        for i, (l, v, u, col) in enumerate([
            ("Point de réappro.", fmtInt(PR),     "unités",              C_GREEN),
            ("Jours restants",    fmt(dL, 1),     "jours",               C_MUTED),
            ("Délai critique",    fmt(dL - lt_rp, 1), "jours avant urgence", C_GOLD),
        ]):
            with c[i]: st.markdown(mcard(l, v, u, color=col), unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
def alertes():
    st.markdown("## 🔔 Alertes & Surveillance des stocks")
    st.markdown("<p style='color:#7a9a7a'>Deux types d'alertes : <strong>dépassement du stock de sécurité</strong> et <strong>déclenchement de commande</strong></p>", unsafe_allow_html=True)

    def clean_num(v):
        if v is None: return 0.0
        s = str(v).strip().replace(" ", "").replace(",", ".")
        if s in ("#DIV/0!", "#N/A", "#VALEUR!", "#REF!", "", "-", "—"): return 0.0
        try: return float(s)
        except: return 0.0

    df_alerte_all = {}

    if "df_resultats" in st.session_state and st.session_state["df_resultats"]:
        src_choice = st.radio("Source des données", ["📥 Depuis calcul automatique", "📂 Importer fichier XLSX"], horizontal=True)
    else:
        src_choice = "📂 Importer fichier XLSX"

    if "calcul" in src_choice and "df_resultats" in st.session_state:
        for sheet_name, df_calc in st.session_state["df_resultats"].items():
            df_alerte_all[sheet_name] = df_calc.copy()
            df_alerte_all[sheet_name]["Stock sécurité calculé"] = df_calc["Stock sécurité"]
        st.success(f"✅ Données chargées depuis le calcul automatique — {sum(len(v) for v in df_alerte_all.values())} articles")
    else:
        uploaded = st.file_uploader(
            "📂 Importez le fichier XLSX exporté depuis le calcul automatique",
            type=["xlsx", "xls"], key="alertes_upload")

        if uploaded:
            try:
                all_raw = pd.read_excel(uploaded, sheet_name=None)
                sheet_names = list(all_raw.keys())

                if len(sheet_names) > 1:
                    st.info(f"📋 **{len(sheet_names)} feuilles détectées :** {', '.join(sheet_names)}")
                    feuille_alerte = st.radio("Feuille à analyser", sheet_names, horizontal=True, key="alerte_sheet_sel")
                    sheets_proc = [feuille_alerte]
                else:
                    sheets_proc = sheet_names

                for sn in sheets_proc:
                    df_raw = all_raw[sn]
                    df_raw.columns = [str(c).strip() for c in df_raw.columns]
                    df_raw = df_raw.dropna(how="all")

                    cmap = {}
                    for col in df_raw.columns:
                        cl = col.lower().strip()
                        if "code" in cl:                                         cmap["Code article"] = col
                        elif "description" in cl or "désignation" in cl:        cmap["Description Article"] = col
                        elif "source" in cl:                                     cmap["Source"] = col
                        elif "stock sécurité" in cl or "stock secu" in cl or "ss (u)" in cl: cmap["Stock sécurité calculé"] = col
                        elif "stock actuel" in cl or "stock actuel (u)" in cl:  cmap["Stock actuel"] = col
                        elif "point de commande" in cl:                          cmap["Point de commande"] = col
                        elif "commande en cours" in cl or "cmd en cours" in cl: cmap["Commande en cours"] = col
                        elif "consommation" in cl or "conso" in cl:             cmap["Conso mois moy"] = col

                    df_m = df_raw.rename(columns={v: k for k, v in cmap.items()})
                    for c in ["Stock sécurité calculé", "Stock actuel", "Point de commande",
                               "Commande en cours", "Conso mois moy"]:
                        if c in df_m.columns:
                            df_m[c] = df_m[c].apply(clean_num)

                    if "Code article" in df_m.columns:
                        df_m = df_m[df_m["Code article"].notna()]
                    df_alerte_all[sn] = df_m
                    st.success(f"✅ Feuille **{sn}** : {len(df_m)} articles chargés")
            except Exception as e:
                st.error(f"Erreur : {e}")

    if not df_alerte_all:
        st.info("⬆️ Lancez d'abord un **Calcul automatique** ou importez un fichier XLSX.")
        return

    for sheet_name, df in df_alerte_all.items():
        if len(df_alerte_all) > 1:
            st.markdown(f"---\n### 📄 {sheet_name}")

        if "Stock actuel" not in df.columns or df["Stock actuel"].sum() == 0:
            st.warning("⚠️ Saisissez le stock actuel pour chaque article :")
            stk_vals = {}
            cols_stk = st.columns(min(4, len(df)))
            for i, (_, row) in enumerate(df.iterrows()):
                with cols_stk[i % len(cols_stk)]:
                    stk_vals[str(row.get("Code article", i))] = st.number_input(
                        f"Stock {row.get('Code article', 'Art'+str(i))}",
                        value=0.0, min_value=0.0, step=1.0, key=f"stk_{sheet_name}_{i}")
            df["Stock actuel"] = df["Code article"].astype(str).map(stk_vals).fillna(0)

        df = df.copy()
        ss_col  = "Stock sécurité calculé" if "Stock sécurité calculé" in df.columns else "Stock sécurité"
        stk_col = "Stock actuel"
        pr_col  = "Point de commande"

        if ss_col  not in df.columns: df[ss_col]  = 0
        if stk_col not in df.columns: df[stk_col] = 0
        if pr_col  not in df.columns: df[pr_col]  = 0

        def get_alerte(row):
            stk = clean_num(row.get(stk_col, 0))
            ss  = clean_num(row.get(ss_col, 0))
            pr  = clean_num(row.get(pr_col, 0))
            cmd = clean_num(row.get("Commande en cours", 0))
            if stk <= 0 or stk < ss * 0.5:
                return "🔴 Rupture critique"
            elif stk < ss:
                return "🟠 SS dépassé"
            elif stk <= pr and not (stk + cmd > pr):
                return "🟡 Commander maintenant"
            else:
                return "🟢 Normal"

        df["Alerte"] = df.apply(get_alerte, axis=1)

        rupt    = df[df["Alerte"] == "🔴 Rupture critique"]
        ss_dep  = df[df["Alerte"] == "🟠 SS dépassé"]
        cmd_now = df[df["Alerte"] == "🟡 Commander maintenant"]
        ok      = df[df["Alerte"] == "🟢 Normal"]

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(mcard("🔴 Rupture critique",    str(len(rupt)),    "articles",  color=C_RED),    unsafe_allow_html=True)
        with c2: st.markdown(mcard("🟠 SS dépassé",           str(len(ss_dep)),  "articles",  color="#e65100"), unsafe_allow_html=True)
        with c3: st.markdown(mcard("🟡 Commander maintenant", str(len(cmd_now)), "articles",  color=C_GOLD),   unsafe_allow_html=True)
        with c4: st.markdown(mcard("🟢 Normal",               str(len(ok)),      "articles",  color=C_GREEN2), unsafe_allow_html=True)

        st.divider()

        filtre = st.multiselect(
            "Filtrer par alerte",
            options=["🔴 Rupture critique", "🟠 SS dépassé", "🟡 Commander maintenant", "🟢 Normal"],
            default=["🔴 Rupture critique", "🟠 SS dépassé", "🟡 Commander maintenant"],
            key=f"filtre_{sheet_name}")

        df_show = df[df["Alerte"].isin(filtre)] if filtre else df

        disp_cols = ["Code article", "Description Article", "Source"]
        for c in [ss_col, stk_col, pr_col, "Commande en cours", "Alerte"]:
            if c in df_show.columns:
                disp_cols.append(c)

        st.dataframe(df_show[[c for c in disp_cols if c in df_show.columns]],
                     use_container_width=True, hide_index=True)

        if stk_col in df.columns and ss_col in df.columns:
            st.markdown("---")
            st.markdown("#### 📈 Stock actuel vs Stock de sécurité vs Point de commande")
            labels = df["Code article"].astype(str).tolist()
            color_map = {"🔴 Rupture critique": C_RED, "🟠 SS dépassé": "#e65100",
                          "🟡 Commander maintenant": C_GOLD, "🟢 Normal": C_GREEN2}
            bar_colors = [color_map.get(a, C_GREEN2) for a in df["Alerte"]]
            fig = go.Figure()
            fig.add_trace(go.Bar(name="Stock actuel", x=labels, y=df[stk_col], marker_color=bar_colors))
            fig.add_trace(go.Scatter(name="Stock sécurité", x=labels, y=df[ss_col],
                mode="markers+lines", marker=dict(color=C_GOLD, size=9, symbol="diamond"),
                line=dict(color=C_GOLD, width=2, dash="dash")))
            if pr_col in df.columns:
                fig.add_trace(go.Scatter(name="Point de commande", x=labels, y=df[pr_col],
                    mode="markers+lines", marker=dict(color=C_RED, size=7, symbol="x"),
                    line=dict(color=C_RED, width=1.5, dash="dot")))
            fig.update_layout(**light_layout(), height=360,
                xaxis=dict(tickangle=-45, gridcolor="rgba(0,0,0,.06)"),
                yaxis=dict(title="Unités", gridcolor="rgba(0,0,0,.06)"))
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.markdown("#### 🚨 Actions prioritaires")
        if len(rupt) == 0 and len(ss_dep) == 0 and len(cmd_now) == 0:
            st.markdown("<div class='alert-ok'>✅ Aucune alerte — tous les stocks sont suffisants.</div>", unsafe_allow_html=True)

        for _, r in rupt.iterrows():
            stk = clean_num(r.get(stk_col, 0)); ss = clean_num(r.get(ss_col, 0))
            st.markdown(
                f"<div class='alert-critical'>🔴 <strong>{r.get('Code article','')} — {r.get('Description Article','')}</strong><br/>"
                f"<span style='font-size:.85rem'>Stock actuel : <strong>{stk:,.0f}</strong> u | SS : {ss:,.0f} u | "
                f"Écart : <strong>{stk-ss:,.0f}</strong> u — Commander immédiatement !</span></div>",
                unsafe_allow_html=True)

        for _, r in ss_dep.iterrows():
            stk = clean_num(r.get(stk_col, 0)); ss = clean_num(r.get(ss_col, 0))
            pr  = clean_num(r.get(pr_col, 0))
            st.markdown(
                f"<div class='alert-warning'>🟠 <strong>{r.get('Code article','')} — {r.get('Description Article','')}</strong><br/>"
                f"<span style='font-size:.85rem'>Stock : {stk:,.0f} u | SS : {ss:,.0f} u | PR : {pr:,.0f} u — Stock de sécurité entamé</span></div>",
                unsafe_allow_html=True)

        for _, r in cmd_now.iterrows():
            stk = clean_num(r.get(stk_col, 0)); pr = clean_num(r.get(pr_col, 0))
            src = r.get("Source", "")
            lt  = get_lt_jours(src) if src else "?"
            st.markdown(
                f"<div class='alert-info'>🟡 <strong>{r.get('Code article','')} — {r.get('Description Article','')}</strong><br/>"
                f"<span style='font-size:.85rem'>Stock : {stk:,.0f} u ≤ PR : {pr:,.0f} u | Délai : {lt}j — Lancer la commande maintenant</span></div>",
                unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# Catalogue complet des articles MedOil (codes fixes)
# ─────────────────────────────────────────────────────────────────────────────
CATALOGUE_ARTICLES = [
    {"code": "0050004", "description": "POUDRE DE LACTOSERUM",                "source": "EX"},
    {"code": "0040024", "description": "HUILE OMEGA 3",                        "source": "EX"},
    {"code": "0040030", "description": "MIXED TOCOPHEROL 95",                  "source": "EX"},
    {"code": "0040085", "description": "PALSGAARD PG 1155",                    "source": "EX"},
    {"code": "0040026", "description": "AROME 502916 BHI - Savoureuse 2.5kg", "source": "EX"},
    {"code": "0040088", "description": "PALSGAARD DMG 0291",                   "source": "EX"},
    {"code": "0140233", "description": "AROME HUILES D'OLIVE",                 "source": "EX"},
    {"code": "0040032", "description": "AROME T00262",                         "source": "EX"},
    {"code": "0040048", "description": "AROME 9261",                           "source": "EX"},
    {"code": "0040004", "description": "B.H.T",                               "source": "EX"},
    {"code": "0040100", "description": "TERRE RAFINOL",                        "source": "EX"},
    {"code": "0040007", "description": "NYSOSEL (NICKEL CATALYSEUR)",          "source": "EX"},
    {"code": "0040092", "description": "ACIDE LACTIQUE",                       "source": "BM"},
    {"code": "0040010", "description": "TERRE TONSIL",                         "source": "BM"},
    {"code": "0040055", "description": "TRISYL",                               "source": "BM"},
    {"code": "0050003", "description": "SORBATE DE POTASSIUM",                 "source": "BM"},
    {"code": "0040101", "description": "ACIDE CITRIQUE",                       "source": "BM"},
    {"code": "0200467", "description": "Tube 100 gr",                          "source": "EX"},
    {"code": "0200424", "description": "Tube 200gr Petit Format",               "source": "EX"},
    {"code": "0200456", "description": "Tube 200gr Grand Format admission",     "source": "EX"},
    {"code": "0200402", "description": "Tube 200gr Grand Format",               "source": "EX"},
    {"code": "0200465", "description": "Tube 200 Huile d'olive",                "source": "EX"},
    {"code": "0200711", "description": "Tube 200 Omega",                        "source": "EX"},
    {"code": "0200428", "description": "Tube 200gr Allégé",                     "source": "EX"},
    {"code": "0200401", "description": "Tube 250gr Local",                      "source": "Local"},
    {"code": "0200400", "description": "Tube 250gr Export",                     "source": "EX"},
    {"code": "0200434", "description": "Tube 400gr",                            "source": "EX"},
    {"code": "0200488", "description": "Barquette 200gr 4+",                    "source": "EX"},
    {"code": "0200652", "description": "Barquette Omega 3 400",                 "source": "EX"},
    {"code": "0200472", "description": "1KG Export",                            "source": "EX"},
    {"code": "0200471", "description": "1KG Local",                             "source": "Local"},
    {"code": "0200279", "description": "Opercule Aluminium 100g Turquie",       "source": "EX"},
    {"code": "0200374", "description": "Opercule Aluminium 200/250/400g Turquie","source": "EX"},
    {"code": "0210158", "description": "Papier sulfirisé sav avec PE imprimé",  "source": "EX"},
    {"code": "0210153", "description": "Papier Complexe MP JADIDA 2.5 KG",     "source": "Local"},
    {"code": "0210161", "description": "Papier Complexe MP JADIDA 2.5 KG (bis)","source": "Local"},
    {"code": "0210114", "description": "Film Savoureuse 200gr Gros",            "source": "EX"},
]


def evolution_gains():
    st.markdown("## 📈 Évolution SS & Évaluation des gains")
    st.markdown(
        "<p style='color:#7a9a7a'>"
        "Saisissez un code article — la description se remplit automatiquement. "
        "Les valeurs SS et coûts sont pré-remplis si un calcul a déjà été lancé.</p>",
        unsafe_allow_html=True)

    # ── Construire catalogue enrichi (fixe + session) ─────────────────────────
    catalogue_map = {a["code"]: dict(a) for a in CATALOGUE_ARTICLES}

    # Index rapide code -> données
    code_to_desc = {c: d.get("description","") for c, d in catalogue_map.items()}
    code_to_cu   = {c: float(d.get("cu",  0.0)  or 0.0)  for c, d in catalogue_map.items()}
    code_to_tx   = {c: float(d.get("tx",  0.20) or 0.20) for c, d in catalogue_map.items()}
    code_to_ss   = {c: d.get("ss_calcule", None)           for c, d in catalogue_map.items()}
    all_codes_list = sorted(catalogue_map.keys())

    n_session = sum(1 for d in catalogue_map.values() if d.get("ss_calcule"))
    st.markdown(
        f"<div style='background:#eef4ee;border:1px solid #d4e0d4;border-radius:10px;"
        f"padding:.7rem 1.1rem;margin-bottom:1rem;font-size:.85rem;color:#3d5c3d'>"
        f"📋 <strong>{len(all_codes_list)} articles dans le catalogue</strong> · "
        f"<strong style='color:#2d6a2d'>{n_session} avec SS calculé</strong> (depuis calcul automatique)</div>",
        unsafe_allow_html=True)

    # Mini-catalogue de référence
    with st.expander("📋 Catalogue complet — cliquez pour voir tous les codes", expanded=False):
        rows_cat = [{"Code": c, "Description": catalogue_map[c].get("description",""),
                     "Source": catalogue_map[c].get("source",""),
                     "SS calculé": f"{int(catalogue_map[c]['ss_calcule'])} u"
                                   if catalogue_map[c].get("ss_calcule") else "—"}
                    for c in all_codes_list]
        st.dataframe(pd.DataFrame(rows_cat), use_container_width=True, hide_index=True,
            column_config={
                "Code":        st.column_config.TextColumn("Code article"),
                "Description": st.column_config.TextColumn("Description"),
                "Source":      st.column_config.TextColumn("Source"),
                "SS calculé":  st.column_config.TextColumn("SS calculé"),
            })
        st.caption("💡 Copiez-collez le code dans le tableau ci-dessous")

    tab1, tab2 = st.tabs(["💰 Évaluation des gains", "📊 Évolution temporelle SS"])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 — GAINS
    # ══════════════════════════════════════════════════════════════════════════
    with tab1:
        st.markdown(
            fbox("Formule du gain annuel",
                 "Gain = (SS_avant − SS_calculé) × Coût_unitaire × Taux_stockage",
                 "Capital libéré = (SS_avant − SS_calculé) × Coût_unitaire · "
                 "Gain annuel = Capital libéré × Taux_stockage"),
            unsafe_allow_html=True)

        st.markdown("#### ✏️ Tableau de saisie — codes articles")
        st.caption("Tapez le code → Description, Coût unitaire, SS calculé et Taux stockage se remplissent automatiquement.")

        if "gain_table" not in st.session_state:
            st.session_state["gain_table"] = pd.DataFrame([
                {"Code article": "", "Description": "", "SS avant (u)": 500,
                 "SS calculé (u)": 350, "Coût unit. (TND)": 0.0, "Taux stockage": 0.20}
                for _ in range(8)
            ])

        df_gain_input = st.data_editor(
            st.session_state["gain_table"],
            use_container_width=True, hide_index=True, num_rows="dynamic",
            key="gain_editor",
            column_config={
                "Code article":     st.column_config.TextColumn("Code article",
                    help="Tapez le code — les autres champs se remplissent automatiquement", width="small"),
                "Description":      st.column_config.TextColumn("Description (auto)", disabled=True, width="large"),
                "SS avant (u)":     st.column_config.NumberColumn("SS avant (u)",    min_value=0, step=10,
                    help="Valeur du SS avant optimisation"),
                "SS calculé (u)":   st.column_config.NumberColumn("SS calculé (u)", min_value=0, step=10,
                    help="Pré-rempli depuis le calcul automatique"),
                "Coût unit. (TND)": st.column_config.NumberColumn("Coût unit. (TND)", min_value=0.0, step=0.01, format="%.4f"),
                "Taux stockage":    st.column_config.NumberColumn("Taux stockage (0.20=20%)",
                    min_value=0.0, max_value=1.0, step=0.01, format="%.2f"),
            })

        # Auto-remplissage
        df_gain_updated = df_gain_input.copy()
        changed_g = False
        for i, row in df_gain_updated.iterrows():
            code_s = str(row.get("Code article","")).strip()
            if code_s and code_s in catalogue_map:
                desc_a = code_to_desc.get(code_s,"")
                cu_a   = float(code_to_cu.get(code_s, 0.0) or 0.0)
                tx_a   = float(code_to_tx.get(code_s, 0.20) or 0.20)
                ss_c   = code_to_ss.get(code_s)
                if df_gain_updated.at[i,"Description"] != desc_a:
                    df_gain_updated.at[i,"Description"] = desc_a; changed_g = True
                if cu_a > 0 and float(df_gain_updated.at[i,"Coût unit. (TND)"]) == 0.0:
                    df_gain_updated.at[i,"Coût unit. (TND)"] = cu_a; changed_g = True
                if tx_a and float(df_gain_updated.at[i,"Taux stockage"]) == 0.20 and tx_a != 0.20:
                    df_gain_updated.at[i,"Taux stockage"] = tx_a; changed_g = True
                if ss_c:
                    if int(df_gain_updated.at[i,"SS calculé (u)"]) == 350:
                        df_gain_updated.at[i,"SS calculé (u)"] = int(ss_c); changed_g = True
                    if int(df_gain_updated.at[i,"SS avant (u)"]) == 500:
                        df_gain_updated.at[i,"SS avant (u)"] = int(ss_c * 1.3); changed_g = True
            elif code_s and code_s not in catalogue_map:
                msg = "⚠️ Code non trouvé"
                if df_gain_updated.at[i,"Description"] != msg:
                    df_gain_updated.at[i,"Description"] = msg; changed_g = True

        if changed_g:
            st.session_state["gain_table"] = df_gain_updated
            st.rerun()

        articles_gain = df_gain_updated[
            (df_gain_updated["Code article"].astype(str).str.strip() != "") &
            (~df_gain_updated["Description"].astype(str).str.startswith("⚠️"))
        ].copy()

        if articles_gain.empty:
            st.info("Saisissez des codes articles dans le tableau ci-dessus pour calculer les gains.")
        else:
            n_ss_ok = (articles_gain["SS calculé (u)"] != 350).sum()
            c1, c2 = st.columns(2)
            with c1: st.markdown(mcard("Articles reconnus",
                f"{len(articles_gain)}", "du catalogue", color=C_GREEN), unsafe_allow_html=True)
            with c2: st.markdown(mcard("SS pré-remplis",
                str(n_ss_ok), "depuis calcul auto", color=C_GREEN2), unsafe_allow_html=True)

            st.divider()
            if st.button("💰 Calculer les gains", key="btn_gain_calc", type="primary"):
                rows_result = []
                for _, r in articles_gain.iterrows():
                    delta   = int(r["SS avant (u)"]) - int(r["SS calculé (u)"])
                    cap_lib = delta * float(r["Coût unit. (TND)"])
                    gain_an = cap_lib * float(r["Taux stockage"])
                    rows_result.append({
                        "Code":                  r["Code article"],
                        "Description":           r["Description"],
                        "SS avant (u)":          int(r["SS avant (u)"]),
                        "SS calculé (u)":        int(r["SS calculé (u)"]),
                        "Réduction SS (u)":      delta,
                        "Coût unit. (TND)":      round(float(r["Coût unit. (TND)"]),4),
                        "Taux stockage":         f"{float(r['Taux stockage'])*100:.0f}%",
                        "Capital libéré (TND)":  round(cap_lib,2),
                        "Gain annuel (TND)":     round(gain_an,2),
                        "Statut": ("🟢 Économie" if delta>0 else "🔴 Surstock" if delta<0 else "⚪ Inchangé"),
                    })
                df_result = pd.DataFrame(rows_result)
                st.dataframe(df_result, use_container_width=True, hide_index=True,
                    column_config={
                        "Capital libéré (TND)": st.column_config.NumberColumn(format="%.2f"),
                        "Gain annuel (TND)":    st.column_config.NumberColumn(format="%.2f"),
                    })
                total_cap  = df_result["Capital libéré (TND)"].sum()
                total_gain = df_result["Gain annuel (TND)"].sum()
                n_eco      = (df_result["Statut"]=="🟢 Économie").sum()
                n_surs     = (df_result["Statut"]=="🔴 Surstock").sum()
                st.divider()
                c1,c2,c3,c4 = st.columns(4)
                with c1: st.markdown(mcard("Capital libéré",    f"{total_cap:,.0f}",  "TND",    color=C_GREEN2), unsafe_allow_html=True)
                with c2: st.markdown(mcard("Gain annuel total", f"{total_gain:,.0f}", "TND/an", color=C_GOLD),   unsafe_allow_html=True)
                with c3: st.markdown(mcard("Optimisés", str(n_eco), f"/{len(df_result)}", color=C_GREEN), unsafe_allow_html=True)
                with c4: st.markdown(mcard("En surstock", str(n_surs), "à vérifier", color=C_RED), unsafe_allow_html=True)

                fig_g = go.Figure()
                fig_g.add_trace(go.Bar(name="SS avant",   x=df_result["Code"], y=df_result["SS avant (u)"],   marker_color=C_RED,   opacity=0.75))
                fig_g.add_trace(go.Bar(name="SS calculé", x=df_result["Code"], y=df_result["SS calculé (u)"], marker_color=C_GREEN))
                fig_g.update_layout(**light_layout(), barmode="group", height=300,
                    xaxis=dict(title="Article", gridcolor="rgba(0,0,0,.06)"),
                    yaxis=dict(title="SS (u)",   gridcolor="rgba(0,0,0,.06)"), title="SS avant vs SS calculé")
                st.plotly_chart(fig_g, use_container_width=True)

                fig_gain = go.Figure(go.Bar(
                    x=df_result["Code"], y=df_result["Gain annuel (TND)"],
                    marker_color=[C_GREEN if v>=0 else C_RED for v in df_result["Gain annuel (TND)"]],
                    text=df_result["Gain annuel (TND)"].apply(lambda v: f"{v:,.0f} TND"),
                    textposition="outside"))
                fig_gain.update_layout(**light_layout(), height=280,
                    xaxis=dict(title="Article", gridcolor="rgba(0,0,0,.06)"),
                    yaxis=dict(title="Gain/an (TND)", gridcolor="rgba(0,0,0,.06)"), title="Gain annuel par article")
                st.plotly_chart(fig_gain, use_container_width=True)

                buf_gain = io.BytesIO()
                with pd.ExcelWriter(buf_gain, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="Gains SS")
                buf_gain.seek(0)
                st.download_button(label="⬇️ Exporter le rapport des gains", data=buf_gain,
                    file_name=f"gains_ss_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_gain")

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 — ÉVOLUTION SS
    # ══════════════════════════════════════════════════════════════════════════
    with tab2:
        c1, c2 = st.columns(2)
        with c1:
            debut_ev = st.date_input("Début période", value=date(date.today().year,1,1), key="ev_debut")
        with c2:
            nb_mois = st.number_input("Nombre de mois", value=6, min_value=2, max_value=24, step=1, key="ev_nb_mois")

        mois_ev = []
        for i in range(int(nb_mois)):
            m = (debut_ev.month - 1 + i) % 12 + 1
            y = debut_ev.year + ((debut_ev.month - 1 + i) // 12)
            mois_ev.append(f"{MOIS_ABREV[m-1]} {y}")

        st.divider()
        st.markdown("#### ✏️ Articles à suivre")
        st.caption("Tapez le code → les valeurs SS se pré-remplissent. Ajustez les valeurs mensuelles.")

        if "ev_table" not in st.session_state:
            st.session_state["ev_table"] = pd.DataFrame([
                {"Code article":"","Description":"","SS avant (u)":500,"SS cible (u)":350,"Coût unit. (TND)":0.0}
                for _ in range(5)
            ])

        df_ev_input = st.data_editor(
            st.session_state["ev_table"],
            use_container_width=True, hide_index=True, num_rows="dynamic",
            key="ev_editor",
            column_config={
                "Code article":    st.column_config.TextColumn("Code article", width="medium"),
                "Description":     st.column_config.TextColumn("Description (auto)", disabled=True, width="large"),
                "SS avant (u)":    st.column_config.NumberColumn("SS avant (u)",  min_value=0, step=10),
                "SS cible (u)":    st.column_config.NumberColumn("SS cible (u)",  min_value=0, step=10),
                "Coût unit. (TND)":st.column_config.NumberColumn("Coût unit. (TND)", min_value=0.0, step=0.01, format="%.4f"),
            })

        df_ev_updated = df_ev_input.copy()
        changed = False
        for i, row in df_ev_updated.iterrows():
            code_s = str(row.get("Code article","")).strip()
            if code_s and code_s in catalogue_map:
                desc_a = code_to_desc.get(code_s,"")
                cu_a   = float(code_to_cu.get(code_s, 0.0) or 0.0)
                ss_c   = code_to_ss.get(code_s)
                if df_ev_updated.at[i,"Description"] != desc_a:
                    df_ev_updated.at[i,"Description"] = desc_a; changed = True
                if cu_a > 0 and float(df_ev_updated.at[i,"Coût unit. (TND)"]) == 0.0:
                    df_ev_updated.at[i,"Coût unit. (TND)"] = cu_a; changed = True
                if ss_c:
                    if int(df_ev_updated.at[i,"SS cible (u)"]) == 350:
                        df_ev_updated.at[i,"SS cible (u)"] = int(ss_c)
                        df_ev_updated.at[i,"SS avant (u)"] = int(ss_c * 1.3); changed = True
            elif code_s and code_s not in catalogue_map:
                if df_ev_updated.at[i,"Description"] != "⚠️ Code non trouvé":
                    df_ev_updated.at[i,"Description"] = "⚠️ Code non trouvé"; changed = True
        if changed:
            st.session_state["ev_table"] = df_ev_updated
            st.rerun()

        articles_ev = df_ev_updated[df_ev_updated["Code article"].astype(str).str.strip() != ""].copy()
        if articles_ev.empty:
            st.info("Saisissez des codes articles ci-dessus.")
        else:
            st.divider()
            st.markdown("#### 📅 Valeurs SS mensuelles")
            mois_cols_cfg = {m: st.column_config.NumberColumn(m, min_value=0, step=10) for m in mois_ev}
            ev_key = f"ev_monthly_{len(articles_ev)}_{nb_mois}"
            if ev_key not in st.session_state:
                rows_m = []
                for _, art_row in articles_ev.iterrows():
                    code  = str(art_row["Code article"]).strip()
                    ss_av = int(art_row.get("SS avant (u)",500))
                    ss_nv = int(art_row.get("SS cible (u)",350))
                    row_m = {"Code":code,"Description":str(art_row.get("Description",""))[:25]}
                    for j, mois in enumerate(mois_ev):
                        row_m[mois] = int(ss_av-(ss_av-ss_nv)*j/max(1,int(nb_mois)-1))
                    rows_m.append(row_m)
                st.session_state[ev_key] = pd.DataFrame(rows_m)

            df_monthly = st.data_editor(st.session_state[ev_key],
                use_container_width=True, hide_index=True,
                key=f"ev_monthly_editor_{nb_mois}",
                column_config={
                    "Code":        st.column_config.TextColumn("Code", disabled=True, width="small"),
                    "Description": st.column_config.TextColumn("Description", disabled=True, width="medium"),
                    **mois_cols_cfg})

            if st.button("📊 Générer le graphique", key="ev_btn"):
                fig = go.Figure()
                colors_c = [C_GREEN,C_GOLD,C_GREEN2,C_PURPLE,C_RED,C_MUTED]
                for i, (_, art_row) in enumerate(articles_ev.iterrows()):
                    code  = str(art_row["Code article"]).strip()
                    ss_av = int(art_row.get("SS avant (u)",500))
                    ss_nv = int(art_row.get("SS cible (u)",350))
                    col   = colors_c[i % len(colors_c)]
                    match = df_monthly[df_monthly["Code"]==code]
                    vals  = [int(match.iloc[0].get(m,ss_av)) for m in mois_ev] if not match.empty else                             [int(ss_av-(ss_av-ss_nv)*j/max(1,int(nb_mois)-1)) for j in range(int(nb_mois))]
                    fig.add_trace(go.Scatter(x=mois_ev, y=vals,
                        name=f"{code} — {str(art_row.get('Description',''))[:20]}",
                        mode="lines+markers", line=dict(color=col,width=2.5), marker=dict(size=8,color=col)))
                    fig.add_hline(y=ss_av, line_dash="dot",  line_color=C_RED,   line_width=1,
                                   annotation_text=f"{code} avant={ss_av}u", annotation_font_size=9)
                    fig.add_hline(y=ss_nv, line_dash="dash", line_color=C_GREEN, line_width=1,
                                   annotation_text=f"{code} cible={ss_nv}u", annotation_font_size=9)
                fig.update_layout(**light_layout(), height=440,
                    xaxis=dict(title="Mois", gridcolor="rgba(0,0,0,.06)"),
                    yaxis=dict(title="SS (u)", gridcolor="rgba(0,0,0,.06)"),
                    title="Évolution SS — avant / après optimisation")
                st.plotly_chart(fig, use_container_width=True)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    URL_LOGO = "https://raw.githubusercontent.com/Ghofrane13/medoil-supply-app/main/logo.png"
    st.markdown(
        f"""
    <div style="display: flex; align-items: center; padding: .3rem 0 .5rem">
        <img src={URL_LOGO} style="width: 44px; margin-right: 10px; border-radius:6px">
        <div>
            <div style="color:#ffffff;font-weight:700;font-size:1.15rem;line-height:1.1">Med oil</div>
            <div style="color:#f5c518;font-size:.65rem;letter-spacing:.1em;text-transform:uppercase">Supply Chain · v4.0</div>
        </div>
    </div>
    """,
        unsafe_allow_html=True
    )
    st.divider()

    pg = st.navigation([
        st.Page(accueil,         title="Accueil",               icon="🏠"),
        st.Page(import_calcul,   title="Calcul automatique",    icon="📥"),
        st.Page(calculateurs,    title="Calculateurs SC",        icon="⚙️"),
        st.Page(alertes,         title="Alertes & Stocks",       icon="🔔"),
        st.Page(evolution_gains, title="Évolution & Gains",      icon="📈"),
    ])

with st.sidebar:
    st.divider()
    st.markdown("""
    <div style='font-size:.72rem;color:#3a5a3a'>
        <div style='color:#f5c518;margin-bottom:4px'>⏱ Délais fixes</div>
        <div>Export : 4 mois (120j)</div>
        <div>Local : 2 sem. (14j)</div>
        <div>BM : 3 sem. (21j)</div>
        <div style='margin-top:6px;color:#f5c518'>⚙️ Niveau service</div>
        <div>A → 99% (Z=2.33)</div>
        <div>B → 95% (Z=1.65)</div>
        <div>C → 90% (Z=1.28)</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    st.markdown("<p style='color:#4a6a4a;font-size:.72rem;'>© 2025 Medoil — Tous droits réservés</p>",
                unsafe_allow_html=True)

pg.run()
