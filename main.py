import streamlit as st
import math
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="medoil ", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700&family=DM+Mono:wght@400;500&family=DM+Serif+Display:ital@0;1&display=swap');

/* ── PALETTE FIGMA MED OIL ──────────────────────────────────────────────────
   Vert foncé  : #1a3d1a  /  #2d6a2d  /  #4a8c3f
   Or / Jaune  : #f5c518  /  #e8b400  /  #c9a000
   Fond sidebar: #152415
   Fond page   : #f5f5f0  (blanc cassé naturel)
   Texte       : #1a2e1a  (vert très foncé)
   Gris clair  : #e8e8e0
   Blanc cards : #ffffff
   ──────────────────────────────────────────────────────────────────────── */

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

/* ── SIDEBAR ─────────────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background-color: var(--bg-sidebar) !important;
    border-right: 1px solid #2d4a2d !important;
}
[data-testid="stSidebar"] * { color: #c8e6c9 !important; }
[data-testid="stSidebar"] h1 { color: #ffffff !important; font-size: 1.2rem !important; }
[data-testid="stSidebar"] hr { border-color: #2d4a2d !important; }

/* Nav links */
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

/* ── MAIN CONTENT AREA ───────────────────────────────────────────────────── */
.main .block-container {
    background-color: var(--bg-page) !important;
    padding-top: 1.5rem !important;
}

/* ── HEADINGS ────────────────────────────────────────────────────────────── */
h1, h2, h3, h4 { color: var(--text-dark) !important; }

/* ── FORMULA BOX ─────────────────────────────────────────────────────────── */
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

/* ── ALERT BOXES ─────────────────────────────────────────────────────────── */
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

/* ── STREAMLIT NATIVE WIDGETS (light mode override) ──────────────────────── */
.stSelectbox > div > div,
.stNumberInput > div > div > input,
.stTextInput > div > div > input {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    color: var(--text-dark) !important;
    border-radius: 8px !important;
}
.stSelectbox > div > div:hover,
.stNumberInput > div > div > input:focus,
.stTextInput > div > div > input:focus {
    border-color: var(--green-light) !important;
    box-shadow: 0 0 0 2px rgba(74,140,63,.15) !important;
}

/* Labels */
[data-testid="stWidgetLabel"] p { color: var(--text-mid) !important; font-size: .85rem !important; }

/* ── BUTTONS ─────────────────────────────────────────────────────────────── */
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

/* Primary download button */
.stDownloadButton > button[kind="primary"] {
    background: var(--gold) !important;
    font-size: 1rem !important;
    padding: .7rem 2rem !important;
}

/* ── TABS ────────────────────────────────────────────────────────────────── */
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
    transition: color .15s !important;
}
[data-testid="stTabs"] [role="tab"]:hover { color: var(--green-mid) !important; }
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: var(--green-dark) !important;
    border-bottom: 3px solid var(--gold) !important;
    font-weight: 700 !important;
}

/* ── DATAFRAME ───────────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
}

/* ── EXPANDER ────────────────────────────────────────────────────────────── */
[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
}
[data-testid="stExpander"] summary {
    color: var(--text-dark) !important;
    font-weight: 600 !important;
}

/* ── INFO / SUCCESS / WARNING banners ────────────────────────────────────── */
.stInfo    { background: rgba(45,106,45,.07) !important; border-color: var(--green-light) !important; color: var(--green-mid) !important; }
.stSuccess { background: rgba(74,140,63,.08) !important; border-color: var(--green-light) !important; }
.stWarning { background: rgba(245,197,24,.1) !important; border-color: var(--gold-dark) !important; }
.stError   { background: rgba(220,53,53,.07) !important; }

/* ── DIVIDER ─────────────────────────────────────────────────────────────── */
hr { border-color: var(--border) !important; }

/* ── MULTISELECT ─────────────────────────────────────────────────────────── */
[data-testid="stMultiSelect"] span {
    background: var(--green-pale) !important;
    color: var(--green-dark) !important;
    border-radius: 4px !important;
}

/* ── FILE UPLOADER ───────────────────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    background: #ffffff !important;
    border: 2px dashed var(--border) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--green-light) !important;
}

/* ── CHECKBOX ────────────────────────────────────────────────────────────── */
[data-testid="stCheckbox"] input:checked + span {
    background: var(--green-light) !important;
    border-color: var(--green-light) !important;
}
</style>""", unsafe_allow_html=True)

# ── Plotly theme (light) ──────────────────────────────────────────────────────
def light_layout():
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(248,250,248,0.6)",
        font=dict(color="#3d5c3d", family="Syne"),
        legend=dict(bgcolor="rgba(255,255,255,.9)", bordercolor="#d4e0d4", borderwidth=1),
        margin=dict(t=20, b=10, l=0, r=0)
    )

# ── Brand colors for charts ────────────────────────────────────────────────────
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
 
def sc(s): return C_GREEN2 if s>=85 else "#14b8a6" if s>=70 else C_GOLD if s>=55 else C_RED
def sl(s): return "Excellent" if s>=85 else "Bon" if s>=70 else "Moyen" if s>=55 else "Insuffisant"

# ── Calculs SC ────────────────────────────────────────────────────────────────
def calc_ss(conso_mois, z=1.65, lt_mois=1.0, variab=0.15):
    sigma_d = conso_mois * variab
    lt_j    = lt_mois * 30
    d_j     = conso_mois / 30
    return round(z * math.sqrt(lt_j * sigma_d**2 + d_j**2 * (lt_j * 0.05)**2))
 
def calc_eoq(conso_an, cout_unit, sc_cost=150, taux=0.20):
    if cout_unit <= 0 or conso_an <= 0: return 0, 0, 0
    h   = taux * cout_unit
    eoq = math.sqrt(2 * conso_an * sc_cost / h)
    nb  = conso_an / eoq
    ct  = nb * sc_cost + (eoq / 2) * h
    return round(eoq), round(nb, 1), round(ct, 2)
 
def calc_pr(conso_mois, lt_mois, ss):
    return round((conso_mois / 30) * (lt_mois * 30) + ss)
 
# ── Export Excel ──────────────────────────────────────────────────────────────
def build_excel(df_result):
    wb  = Workbook()
    ws  = wb.active
    ws.title = "Calculs Supply Chain"
    thin = Side(style="thin", color="BFBFBF")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
 
    def hdr(cell, val, bg):
        cell.value = val
        cell.font  = Font(bold=True, size=9, name="Arial")
        cell.fill  = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd
 
    def dat(cell, val, nf=None):
        cell.value = val
        cell.font  = Font(size=9, name="Arial")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = brd
        if nf: cell.number_format = nf
 
    # Updated colors to match Figma palette
    groups = [(1,3,"Produit","C8E6C9"),(4,7,"Fournisseur","FFF9C4"),
              (8,9,"Demand","C8E6C9"),(10,12,"EOQ","DCEDC8"),(13,15,"SS","FFE0B2")]
    for c1,c2,label,bg in groups:
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        hdr(ws.cell(1, c1), label, bg)
 
    hdrs = [("Code article","C8E6C9"),("Description Article","C8E6C9"),("Classe","C8E6C9"),
            ("Fournisseur","FFF9C4"),("Délai livraison (mois)","FFF9C4"),
            ("Incertitude fournisseur","FFF9C4"),("Délai de Sécurité (j)","FFF9C4"),
            ("ABC","C8E6C9"),("XYZ","C8E6C9"),
            ("EOQ","DCEDC8"),("MOQ","DCEDC8"),("Diff MOQ et EOQ","DCEDC8"),
            ("Stock de sécurité","FFE0B2"),("Coût stock de sécurité","FFE0B2"),("Point de commande","FFE0B2")]
    for col,(label,bg) in enumerate(hdrs, 1):
        hdr(ws.cell(2, col), label, bg)
 
    fmts = [None,None,None,None,"0.0","0%","0",None,None,"#,##0","#,##0","#,##0","#,##0","#,##0.00","#,##0"]
    for i, row in df_result.iterrows():
        r    = i + 3
        lt_v = row.get("Délai livraison (mois)", 1)
        vals = [row.get("Code article",""), row.get("Description Article",""), row.get("Classe",""),
                row.get("Fournisseur",""), lt_v, row.get("Incertitude", 0.15),
                round(lt_v * 30) if lt_v else "",
                row.get("ABC",""), row.get("XYZ",""),
                row.get("EOQ",0), row.get("MOQ",""), row.get("Diff MOQ EOQ",""),
                row.get("Stock sécurité",0), row.get("Coût SS",0), row.get("Point de commande",0)]
        for col,(val,nf) in enumerate(zip(vals, fmts), 1):
            dat(ws.cell(r, col), val, nf)
 
    for col,w in enumerate([12,32,8,20,14,14,12,6,6,10,10,14,14,18,16], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 42
    ws.freeze_panes = "A3"
 
    ws2 = wb.create_sheet("Données sources")
    src_hdrs = ["Code article","Description Article","Classe","SS fichier",
                "Consommation/mois","Coût de revient","Consommation/an","Coût total annuel",
                "SS calculé","EOQ calculé","Point de commande"]
    for col,h in enumerate(src_hdrs, 1):
        c = ws2.cell(1, col)
        c.value = h; c.font = Font(bold=True, size=9, name="Arial")
        c.fill  = PatternFill("solid", start_color="C8E6C9")
        c.alignment = Alignment(horizontal="center", wrap_text=True); c.border = brd
        ws2.column_dimensions[get_column_letter(col)].width = 18
    for i, row in df_result.iterrows():
        r = i + 2
        vals2 = [row.get("Code article"), row.get("Description Article"), row.get("Classe"),
                 row.get("SS existant",""), row.get("Conso mois",""), row.get("Coût revient",""),
                 row.get("Conso an",""), row.get("Coût total an",""),
                 row.get("Stock sécurité",""), row.get("EOQ",""), row.get("Point de commande","")]
        for col,val in enumerate(vals2, 1):
            c = ws2.cell(r, col); c.value = val
            c.font = Font(size=9, name="Arial"); c.alignment = Alignment(horizontal="center"); c.border = brd
 
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf
 
 
# ════════════════════════════════════════════════════════════════════════════════
# PAGES
# ════════════════════════════════════════════════════════════════════════════════
 
def accueil():
    # Hero banner — dark green + gold accent (matches Figma)
    st.markdown("""
    <div style='background:linear-gradient(135deg,#1a3d1a,#2d6a2d);padding:2.2rem 2.5rem;border-radius:16px;
         border:1px solid rgba(245,197,24,.2);margin-bottom:1.5rem;position:relative;overflow:hidden'>
        <div style='position:absolute;top:-30px;right:-30px;width:200px;height:200px;
             background:radial-gradient(circle,rgba(245,197,24,.12),transparent 70%);border-radius:50%'></div>
        <div style='font-size:.7rem;color:#f5c518;letter-spacing:.14em;text-transform:uppercase;margin-bottom:12px;
             background:rgba(245,197,24,.15);display:inline-block;padding:.3rem .8rem;border-radius:20px;
             border:1px solid rgba(245,197,24,.3)'>
            ✦ Outil supply chain manager · v2.0</div>
        <h1 style='font-family:"DM Serif Display",serif;font-size:2.3rem;color:#ffffff;margin:.4rem 0;line-height:1.25'>
            Pilotez votre <em style='color:#f5c518'>supply chain</em><br/>avec précision</h1>
        <p style='color:#a5c8a5;font-size:.95rem;margin:.8rem 0 1.4rem'>
            Importez votre Excel, calculez automatiquement EOQ, SS et points de commande, exportez le tableau rempli.</p>
        <div style='display:inline-block;background:#f5c518;color:#1a3d1a;font-weight:700;font-size:.9rem;
             padding:.6rem 1.6rem;border-radius:8px;cursor:pointer'>Commencer l'analyse →</div>
    </div>""", unsafe_allow_html=True)

    # KPI row (matches Figma stats strip)
    st.markdown("""
    <div style='display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;margin-bottom:1.5rem'>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;
             box-shadow:0 2px 8px rgba(26,61,26,.06)'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;letter-spacing:.08em;margin-bottom:.5rem'>
                📦 Articles gérés</div>
            <div style='font-size:1.9rem;font-weight:700;color:#1a2e1a;font-family:"DM Mono",monospace'>1 240+</div>
        </div>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;
             box-shadow:0 2px 8px rgba(26,61,26,.06)'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;letter-spacing:.08em;margin-bottom:.5rem'>
                📈 Taux de service moyen</div>
            <div style='font-size:1.9rem;font-weight:700;color:#2d6a2d;font-family:"DM Mono",monospace'>94.7%</div>
        </div>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;
             box-shadow:0 2px 8px rgba(26,61,26,.06)'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;letter-spacing:.08em;margin-bottom:.5rem'>
                💰 Économies SS estimées</div>
            <div style='font-size:1.9rem;font-weight:700;color:#c9a000;font-family:"DM Mono",monospace'>12.4%</div>
        </div>
        <div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.2rem 1.4rem;
             box-shadow:0 2px 8px rgba(26,61,26,.06)'>
            <div style='font-size:.72rem;color:#7a9a7a;text-transform:uppercase;letter-spacing:.08em;margin-bottom:.5rem'>
                🚚 Fournisseurs évalués</div>
            <div style='font-size:1.9rem;font-weight:700;color:#1a2e1a;font-family:"DM Mono",monospace'>38</div>
        </div>
    </div>""", unsafe_allow_html=True)

    # Modules disponibles
    st.markdown("<div style='font-size:1rem;font-weight:700;color:#1a2e1a;margin-bottom:1rem;padding-left:4px;border-left:3px solid #f5c518;padding-left:10px'>Modules disponibles</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    mods = [
        ("📥", "#2d6a2d", "Import & Calcul auto",   "Importez votre fichier Excel et générez le tableau rempli à télécharger"),
        ("📐", "#c9a000", "Calculateurs SC",          "EOQ · SS · KPIs · Point de réappro. avec courbes interactives"),
        ("🚚", "#4a8c3f", "Éval. Fournisseurs",       "Scorecard fournisseur — OTD, fill rate, qualité, réactivité"),
        ("🔔", "#dc3535", "Alertes & Stocks",          "Surveillance des niveaux et alertes rupture par article"),
    ]
    for col,(ic,accent,ti,de) in zip(cols, mods):
        with col:
            st.markdown(
                f"<div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:12px;padding:1.3rem;"
                f"border-top:3px solid {accent};box-shadow:0 2px 8px rgba(26,61,26,.05)'>"
                f"<div style='font-size:1.6rem;margin-bottom:10px'>{ic}</div>"
                f"<div style='font-size:.9rem;font-weight:700;color:#1a2e1a;margin-bottom:6px'>{ti}</div>"
                f"<div style='font-size:.78rem;color:#7a9a7a;line-height:1.55'>{de}</div></div>",
                unsafe_allow_html=True)
    st.divider()
    st.info("Commencez par **Import & Calcul automatique** dans la barre latérale pour charger votre fichier Excel.")
 

# ─────────────────────────────────────────────────────────────────────────────
def import_calcul():
    st.markdown("## 📥 Import & Calcul automatique")
    st.markdown("<p style='color:#7a9a7a'>Importez votre fichier — EOQ, SS et point de commande calculés pour chaque article</p>",
                unsafe_allow_html=True)
 
    with st.expander("📋 Colonnes attendues dans votre fichier Excel"):
        st.markdown("""
| Colonne | Obligatoire | Description |
|---|---|---|
| `Code article` | ✅ | Identifiant unique |
| `Description Article` | ✅ | Désignation |
| `Classe` | — | MP, SF, PF… |
| `Consommation / mois` | ✅ | Consommation mensuelle moyenne |
| `Coût de revient` | ✅ | Coût unitaire (TND) |
| `Consommation/ an` | — | Calculée si absente |
| `Stock Sécurité` | — | SS existant (pour comparaison) |
| `Cout total annuelle` | — | Affiché dans le rapport |
| `Fournisseur` | — | Repris dans l'export |
| `Délai de livraison` | — | En mois (sinon valeur par défaut) |
| `MOQ` / `Taille de lot` | — | Minimum order quantity |
""")
 
    st.markdown("#### ⚙️ Paramètres de calcul globaux")
    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: p_z_opt = st.selectbox("Niveau de service", ["90% (Z=1.28)","95% (Z=1.65)","97.5% (Z=1.96)","99% (Z=2.33)"], index=1)
    with c2: p_lt    = st.number_input("Délai défaut (mois)", value=1.0, min_value=0.1, step=0.5)
    with c3: p_sc    = st.number_input("Coût passation (TND)", value=150.0, min_value=1.0, step=10.0)
    with c4: p_taux  = st.number_input("Taux stockage (%)", value=20.0, min_value=1.0, step=1.0)
    with c5: p_var   = st.number_input("Variabilité demande (%)", value=15.0, min_value=1.0, step=1.0)
    Z = {"90% (Z=1.28)":1.28,"95% (Z=1.65)":1.65,"97.5% (Z=1.96)":1.96,"99% (Z=2.33)":2.33}[p_z_opt]
 
    st.divider()
    st.markdown("#### 📂 Chargement du fichier")
    uploaded = st.file_uploader("Déposez votre fichier Excel (.xlsx / .xls)", type=["xlsx","xls"])
    use_demo = st.checkbox("Utiliser les données de démo (votre format)", value=not bool(uploaded))
 
    df_source = None
 
    if uploaded:
        try:
            all_sheets = pd.read_excel(uploaded, sheet_name=None, header=None)
            names = list(all_sheets.keys())
            sel   = names[0] if len(names)==1 else st.selectbox("Choisir la feuille", names)
            raw   = all_sheets[sel]
 
            hrow = 0
            for i, row in raw.iterrows():
                row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                if any(k in row_str for k in ["code","description","article","consommation"]):
                    hrow = i; break
 
            df_raw = pd.read_excel(uploaded, sheet_name=sel, header=hrow)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            df_raw = df_raw.dropna(how="all")
 
            cmap = {}
            for col in df_raw.columns:
                cl = col.lower().strip()
                if "code" in cl and "article" in cl:                       cmap["Code article"] = col
                elif "description" in cl or "désignation" in cl:           cmap["Description Article"] = col
                elif cl == "classe":                                        cmap["Classe"] = col
                elif "consommation" in cl and "mois" in cl:                cmap["Conso mois"] = col
                elif "coût de revient" in cl or "cout de revient" in cl:   cmap["Coût revient"] = col
                elif "consommation" in cl and "an" in cl:                  cmap["Conso an"] = col
                elif "stock" in cl and ("sécu" in cl or "sécurité" in cl): cmap["SS existant"] = col
                elif "cout total" in cl or "coût total" in cl:             cmap["Coût total an"] = col
                elif "fournisseur" in cl:                                   cmap["Fournisseur"] = col
                elif "délai" in cl and "livraison" in cl:                   cmap["Délai livraison (mois)"] = col
                elif "taille" in cl and "lot" in cl:                        cmap["MOQ"] = col
                elif cl == "moq":                                            cmap["MOQ"] = col
 
            df_source = df_raw.rename(columns={v:k for k,v in cmap.items()})
            keep = [k for k in ["Code article","Description Article","Classe","Fournisseur",
                                 "Délai livraison (mois)","MOQ","Conso mois","Coût revient",
                                 "Conso an","SS existant","Coût total an"] if k in df_source.columns]
            df_source = df_source[keep].copy()
            for c in ["Conso mois","Coût revient","Conso an","SS existant","Coût total an","Délai livraison (mois)","MOQ"]:
                if c in df_source.columns:
                    df_source[c] = pd.to_numeric(df_source[c], errors="coerce").fillna(0)
            if "Code article" in df_source.columns:
                df_source = df_source[df_source["Code article"].notna()]
                df_source = df_source[df_source["Code article"].astype(str).str.strip() != ""]
            st.success(f"✅ {len(df_source)} articles chargés depuis « {sel} »")
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")
 
    if use_demo and df_source is None:
        df_source = pd.DataFrame({
            "Code article":       [40007,40010,40011,40014,40015,40036,40042,40055,40062,40063],
            "Description Article":["NYSOSEL (NICKEL CATALYSEUR)","TERRE TONSIL","PAPIER FILTRE",
                                   "TOILE FILTRANTE TERRE","PERLITE FILTRATION PAF1","SEL MARIN FIN",
                                   "MYVEROL 18-04 MYVEROL","TRISYL","VITAMINE A","VITAMINE E"],
            "Classe":             ["MP"]*10,
            "Conso mois":         [324,5886,509,32460,959,3747,3147,4432,16,22],
            "Coût revient":       [59.23,1.709,3.823,49.787,1.700,0,0,8.189,0,0],
            "Conso an":           [3888,70632,6108,389520,11508,44964,37764,53184,192,264],
            "SS existant":        [2673,6353,2859,68,1171,1560,20000,5000,74,100],
            "Coût total an":      [230286,120710,23351,19393032,19564,0,0,435524,0,0],
        })
        st.info("📊 Données de démo chargées — basées sur votre fichier Excel d'origine")
 
    if df_source is not None and len(df_source) > 0:
        with st.expander("👁 Aperçu des données sources", expanded=False):
            st.dataframe(df_source, use_container_width=True, hide_index=True)
 
        results = []
        for _, row in df_source.iterrows():
            cm  = float(row.get("Conso mois", 0) or 0)
            ca  = float(row.get("Conso an", 0) or cm*12) or cm*12
            cu  = float(row.get("Coût revient", 0) or 0)
            lt  = float(row.get("Délai livraison (mois)", p_lt) or p_lt)
            moq = float(row.get("MOQ", 0) or 0)
 
            ss   = calc_ss(cm, z=Z, lt_mois=lt, variab=p_var/100)
            eoq, _, _ = calc_eoq(ca, cu, sc_cost=p_sc, taux=p_taux/100)
            pr   = calc_pr(cm, lt, ss)
            css  = round(ss * cu, 2) if cu > 0 else 0
            diff = round(eoq - moq) if moq > 0 else ""
 
            results.append({
                "Code article":          row.get("Code article",""),
                "Description Article":   row.get("Description Article",""),
                "Classe":                row.get("Classe",""),
                "Fournisseur":           row.get("Fournisseur",""),
                "Délai livraison (mois)":lt,
                "Incertitude":           p_var/100,
                "EOQ":                   eoq,
                "MOQ":                   int(moq) if moq>0 else "",
                "Diff MOQ EOQ":          diff,
                "Stock sécurité":        ss,
                "SS existant":           float(row.get("SS existant", 0) or 0),
                "Coût SS":               css,
                "Point de commande":     pr,
                "Conso mois":            cm,
                "Coût revient":          cu,
                "Conso an":              ca,
                "Coût total an":         float(row.get("Coût total an", 0) or 0),
            })
 
        df_res = pd.DataFrame(results)
 
        st.divider()
        st.markdown("#### 🧮 Tableau des calculs")
        disp = df_res[["Code article","Description Article","Classe",
                       "EOQ","MOQ","Diff MOQ EOQ","Stock sécurité","Coût SS","Point de commande"]].copy()
        st.dataframe(disp, use_container_width=True, hide_index=True,
            column_config={
                "EOQ":               st.column_config.NumberColumn("EOQ (u)",           format="%d"),
                "MOQ":               st.column_config.NumberColumn("MOQ (u)"),
                "Diff MOQ EOQ":      st.column_config.NumberColumn("Diff MOQ/EOQ"),
                "Stock sécurité":    st.column_config.NumberColumn("Stock sécu. (u)",   format="%d"),
                "Coût SS":           st.column_config.NumberColumn("Coût SS (TND)",      format="%.2f"),
                "Point de commande": st.column_config.NumberColumn("Point commande (u)", format="%d"),
            })
 
        st.markdown("---")
        st.markdown("#### 📊 Synthèse")
        c1,c2,c3,c4 = st.columns(4)
        eoq_mean = df_res[df_res["EOQ"]>0]["EOQ"].mean()
        ecarts   = (df_res["Stock sécurité"] > df_res["SS existant"] * 1.3).sum()
        with c1: st.markdown(mcard("Articles traités",    str(len(df_res)),                   "articles",                          color=C_GREEN),  unsafe_allow_html=True)
        with c2: st.markdown(mcard("Coût total SS",       f"{df_res['Coût SS'].sum():,.0f}",  "TND immobilisés",  color=C_GOLD),   unsafe_allow_html=True)
        with c3: st.markdown(mcard("EOQ moyen",           fmtInt(eoq_mean),                    "u / commande",     color=C_GREEN2), unsafe_allow_html=True)
        with c4: st.markdown(mcard("SS à réviser",        str(ecarts),                         "SS calculé > 130% SS fichier",     color=C_RED),    unsafe_allow_html=True)
 
        st.markdown("---")
        st.markdown("#### 📈 EOQ · Stock de sécurité · Point de commande")
        labels = df_res["Code article"].astype(str).tolist()
        fig = go.Figure()
        fig.add_trace(go.Bar(name="EOQ", x=labels, y=df_res["EOQ"], marker_color=C_GREEN,
                             text=df_res["EOQ"].apply(lambda v: f"{int(v):,}" if v>0 else ""), textposition="outside"))
        fig.add_trace(go.Bar(name="Stock sécurité", x=labels, y=df_res["Stock sécurité"], marker_color=C_GOLD,
                             text=df_res["Stock sécurité"].apply(lambda v: f"{int(v):,}"), textposition="outside"))
        fig.add_trace(go.Scatter(name="Point de commande", x=labels, y=df_res["Point de commande"],
                                 mode="markers+lines", marker=dict(color=C_RED,size=9,symbol="diamond"),
                                 line=dict(color=C_RED,width=2,dash="dash")))
        if df_res["SS existant"].sum() > 0:
            fig.add_trace(go.Scatter(name="SS existant (fichier)", x=labels, y=df_res["SS existant"],
                                     mode="markers+lines", marker=dict(color=C_PURPLE,size=7,symbol="circle"),
                                     line=dict(color=C_PURPLE,width=1.5,dash="dot")))
        fig.update_layout(**light_layout(), barmode="group", height=420,
                          xaxis=dict(title="Article", gridcolor="rgba(0,0,0,.06)"),
                          yaxis=dict(title="Unités",  gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)
 
        if df_res["SS existant"].sum() > 0:
            st.markdown("---")
            st.markdown("#### 🔍 SS calculé vs SS dans votre fichier")
            comp = df_res[["Code article","Description Article","SS existant","Stock sécurité"]].copy()
            comp["Écart (u)"] = comp["Stock sécurité"] - comp["SS existant"]
            comp["Écart %"]   = np.where(comp["SS existant"]>0,
                                         (comp["Écart (u)"] / comp["SS existant"] * 100).round(1), None)
            st.dataframe(comp, use_container_width=True, hide_index=True,
                column_config={
                    "SS existant":    st.column_config.NumberColumn("SS fichier",  format="%d"),
                    "Stock sécurité": st.column_config.NumberColumn("SS calculé",  format="%d"),
                    "Écart (u)":      st.column_config.NumberColumn("Écart (u)",   format="%d"),
                    "Écart %":        st.column_config.NumberColumn("Écart %",     format="%.1f%%"),
                })
 
        st.markdown("---")
        st.markdown("#### 💾 Export — tableau rempli")
        st.info("Le fichier exporté reprend exactement le format de votre tableau d'origine avec toutes les colonnes calculées.")
        st.download_button(
            label="⬇️ Télécharger le tableau Excel rempli",
            data=build_excel(df_res),
            file_name="supply_chain_calculs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
 
 
# ─────────────────────────────────────────────────────────────────────────────
def calculateurs():
    st.markdown("## ⚙️ Calculateurs Supply Chain")
    tab_ss, tab_eoq, tab_kpi, tab_rp = st.tabs(
        ["📦 Stock de sécurité", "📐 EOQ (Wilson)", "📊 KPIs stock", "🔄 Point de réappro."])
 
    with tab_ss:
        st.markdown(fbox("Formule","SS = Z × √(LT×σD² + D²×σLT²)",
                         "Z = facteur service · σD = écart-type demande · LT = délai (j)"), unsafe_allow_html=True)
 
        st.markdown("##### 📂 Option : importer vos données de consommation")
        st.caption("Importez un fichier Excel avec une colonne de consommations historiques — σD sera calculé automatiquement.")
 
        uploaded_ss = st.file_uploader("Fichier Excel (.xlsx / .xls) — colonne de consommations", type=["xlsx","xls"], key="ss_upload")
 
        sd_auto, ss2_auto, ssl_auto = None, None, None
 
        if uploaded_ss:
            try:
                df_up = pd.read_excel(uploaded_ss)
                num_cols = df_up.select_dtypes(include=[np.number]).columns.tolist()
                if num_cols:
                    col_chosen = st.selectbox("Colonne de consommation à utiliser", num_cols, key="ss_col")
                    series = df_up[col_chosen].dropna()
                    sd_auto  = round(series.mean() / 30, 2)
                    ss2_auto = round(float(series.std()), 2)
                    ss2_auto_daily = round(ss2_auto / math.sqrt(30), 2)
                    st.success(f"✅ Données importées — {len(series)} observations · Demande moy./j : **{sd_auto}** · σD journalier : **{ss2_auto_daily}**")
                else:
                    st.warning("Aucune colonne numérique détectée dans le fichier.")
                    ss2_auto_daily = None
            except Exception as e:
                st.error(f"Erreur de lecture : {e}")
                ss2_auto_daily = None
        else:
            ss2_auto_daily = None
 
        st.markdown("---")
        st.markdown("##### ✏️ Paramètres de calcul")
        c1, c2, c3 = st.columns(3)
        with c1:
            sd  = st.number_input("Demande moy. (u/j)",    value=float(sd_auto) if sd_auto else 50.0, step=1.0)
            ss2 = st.number_input("Écart-type demande σD (u/j)", value=float(ss2_auto_daily) if ss2_auto_daily else 8.0, step=0.5)
        with c2:
            slt = st.number_input("Délai fourni. (j)",     value=7.0,  step=1.0)
            ssl = st.number_input("Écart-type délai σLT (j)", value=1.5, step=0.1)
        with c3:
            Z_MAP_SS = {
                "95% — Produits stratégiques":  1.65,
                "96% — Produits stratégiques":  1.75,
                "97% — Produits stratégiques":  1.88,
                "98% — Produits stratégiques":  2.05,
                "99% — Produits stratégiques":  2.33,
                "90% — Intermédiaires":         1.28,
                "91% — Intermédiaires":         1.34,
                "92% — Intermédiaires":         1.41,
                "93% — Intermédiaires":         1.48,
                "94% — Intermédiaires":         1.55,
                "95% — Intermédiaires":         1.65,
                "80% — Faible valeur":          0.84,
                "81% — Faible valeur":          0.88,
                "82% — Faible valeur":          0.92,
                "83% — Faible valeur":          0.95,
                "84% — Faible valeur":          0.99,
                "85% — Faible valeur":          1.04,
            }
            szo = st.selectbox("Niveau de service", list(Z_MAP_SS.keys()), index=0)
            scu = st.number_input("Coût unitaire (TND)", value=25.0, step=1.0)
 
        Z2  = Z_MAP_SS[szo]
        SS1 = Z2 * ss2 * math.sqrt(slt)
        SS2 = Z2 * math.sqrt(slt * ss2**2 + sd**2 * ssl**2)
 
        st.markdown("---")
        c = st.columns(3)
        for i,(l,v,u,col) in enumerate([
            ("SS (σD seul)",          fmtInt(SS1),       "unités", C_GREEN),
            ("SS (formule complète)", fmtInt(SS2),       "unités", C_GREEN2),
            ("Coût immobilisé SS",    fmtInt(SS2 * scu), "TND",    C_GOLD),
        ]):
            with c[i]: st.markdown(mcard(l,v,u,color=col), unsafe_allow_html=True)
 
        st.markdown("<div style='margin-bottom:14px'></div>", unsafe_allow_html=True)
 
        z_vals  = [0.84, 1.04, 1.28, 1.65, 1.88, 2.05, 2.33]
        ns_vals = [80, 85, 90, 95, 97, 98, 99]
        ss_v    = [round(z * math.sqrt(slt * ss2**2 + sd**2 * ssl**2)) for z in z_vals]
        current_pct = int(szo.split("%")[0])
        fig = go.Figure(go.Bar(
            x=[f"{n}%" for n in ns_vals], y=ss_v,
            marker_color=[C_GOLD if n == current_pct else C_GREEN for n in ns_vals],
            text=ss_v, textposition="outside"))
        fig.update_layout(**light_layout(), height=260,
            xaxis=dict(title="Niveau de service", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="SS (unités)",        gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)
 
    with tab_eoq:
        st.markdown(fbox("Modèle de Wilson","EOQ = √(2×D×Sc / (Sh×Cu))",
                         "D=demande annuelle · Sc=coût passation · Sh=taux stockage · Cu=coût unitaire"), unsafe_allow_html=True)
        c1,c2,c3 = st.columns(3)
        with c1:
            eD  = st.number_input("Demande annuelle (u)",  value=12000, step=100)
            eSc = st.number_input("Coût passation (TND)",  value=150.0, step=5.0)
        with c2:
            eCu = st.number_input("Coût unitaire (TND)",   value=80.0,  step=1.0)
            eSh = st.number_input("Taux stockage (%)",     value=20.0,  step=1.0)
        with c3:
            eLT = st.number_input("Délai fourni. (j)",     value=10,    step=1)
            eDy = st.number_input("Jours ouvr./an",        value=250,   step=1)
        h   = (eSh/100) * eCu
        EOQ = math.sqrt(2 * eD * eSc / h) if h > 0 else 0
        nb  = eD / EOQ if EOQ > 0 else 0
        T   = eDy / nb if nb > 0 else 0
        Cp  = nb * eSc; Cs = (EOQ/2) * h
        st.markdown("---")
        c = st.columns(3)
        data = [("EOQ",            fmtInt(EOQ),    "u/cmd", C_GREEN),
                ("Nb cmd/an",      fmt(nb,1),      "cmds",  C_MUTED),
                ("Périodicité",    fmtInt(T),      "jours", C_MUTED),
                ("Coût passation/an", fmtInt(Cp),  "TND",   C_GOLD),
                ("Coût stockage/an",  fmtInt(Cs),  "TND",   C_GOLD),
                ("Coût total min.", fmtInt(Cp+Cs), "TND",   C_GREEN2)]
        for i,(l,v,u,col) in enumerate(data):
            with c[i%3]:
                st.markdown(mcard(l,v,u,color=col), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px'></div>", unsafe_allow_html=True)
        qty_r = np.linspace(max(10, EOQ*.2), EOQ*2.5, 200)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=qty_r, y=[(eD/q)*eSc for q in qty_r],            name="Passation", line=dict(color=C_GREEN,width=2)))
        fig.add_trace(go.Scatter(x=qty_r, y=[(q/2)*h    for q in qty_r],             name="Stockage",  line=dict(color=C_GOLD,width=2)))
        fig.add_trace(go.Scatter(x=qty_r, y=[(eD/q)*eSc+(q/2)*h for q in qty_r],    name="Total",     line=dict(color=C_GREEN2,width=2.5)))
        fig.add_vline(x=EOQ, line_dash="dash", line_color=C_RED,
                      annotation_text=f"EOQ={fmtInt(EOQ)}", annotation_font_color=C_RED)
        fig.update_layout(**light_layout(), height=300,
            xaxis=dict(title="Quantité (u)", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="Coût/an (TND)",gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)
 
    with tab_kpi:
        c1,c2 = st.columns(2)
        with c1:
            ks   = st.number_input("Stock moyen (u)",         value=3000,    step=100)
            kv   = st.number_input("Ventes annuelles (u)",    value=18000,   step=100)
            kval = st.number_input("Valeur stock (TND)",      value=240000,  step=1000)
            kca  = st.number_input("CA annuel (TND)",         value=2400000, step=10000)
        with c2:
            kot  = st.number_input("Livrées à temps",         value=465,     step=1)
            kto  = st.number_input("Total livrées",           value=490,     step=1)
            klc  = st.number_input("Coût logistique (TND)",   value=168000,  step=1000)
            ksh  = st.number_input("Taux stockage (%)",       value=22.0,    step=1.0)
        rot = kv/ks if ks>0 else 0
        cov = (ks/kv)*365 if kv>0 else 0
        srv = (kot/kto)*100 if kto>0 else 0
        crt = (klc/kca)*100 if kca>0 else 0
        st.markdown("---")
        c = st.columns(3)
        kdata = [
            ("Rotation",         fmt(rot,1)+"×",          "/an · >6×=excellent",  C_GREEN2 if rot>=6 else C_GOLD),
            ("Couverture",       fmtInt(cov)+"j",         "15–30j=optimal",        C_MUTED),
            ("Taux service OTD", fmt(srv,1)+"%",           ">95%=cible",           C_GREEN2 if srv>=95 else C_GOLD),
            ("Coût log/CA",      fmt(crt,1)+"%",           "objectif <8%",          C_GREEN2 if crt<8 else C_RED),
            ("Coût stockage/an", fmtInt(kval*(ksh/100)),  "TND",                   C_GOLD),
            ("DSI",              fmt(ks/(kv/365),0)+"j",  "jours de stock",        C_MUTED),
        ]
        for i,(l,v,u,col) in enumerate(kdata):
            with c[i%3]:
                st.markdown(mcard(l,v,u,color=col), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px'></div>", unsafe_allow_html=True)
 
    with tab_rp:
        st.markdown(fbox("Formule","PR = (Dmoy × LT) + SS",
                         "Dmoy = demande/jour · LT = délai (j) · SS = stock de sécurité"), unsafe_allow_html=True)
        c1,c2,c3 = st.columns(3)
        with c1:
            rd  = st.number_input("Demande moy. (u/j)", value=50.0,  step=1.0)
            rlt = st.number_input("Délai fourni. (j)",  value=7.0,   step=1.0)
        with c2:
            rss = st.number_input("Stock sécu. (u)",    value=132.0, step=10.0)
            rdm = st.number_input("Demande max (u/j)",  value=70.0,  step=1.0)
        with c3:
            rltm = st.number_input("Délai max (j)",     value=10.0,  step=1.0)
            rcu  = st.number_input("Stock actuel (u)",  value=500.0, step=10.0)
        PR = rd * rlt + rss
        dL = (rcu - rss) / rd if rd > 0 else 0
        if rcu <= rss:
            st.markdown("<div class='alert-critical'>🔴 RUPTURE IMMINENTE — Commander immédiatement !</div>", unsafe_allow_html=True)
        elif rcu <= PR:
            st.markdown("<div class='alert-warning'>🟡 COMMANDER MAINTENANT — Point de réappro. atteint</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='alert-ok'>🟢 Stock suffisant — {fmtInt(rcu)} u > PR ({fmtInt(PR)} u)</div>", unsafe_allow_html=True)
        c = st.columns(3)
        for i,(l,v,u,col) in enumerate([
            ("Point de réappro.", fmtInt(PR),      "unités",             C_GREEN),
            ("Jours restants",    fmt(dL,1),       "jours",              C_MUTED),
            ("Délai critique",    fmt(dL-rlt,1),   "jours avant urgence",C_GOLD),
        ]):
            with c[i]: st.markdown(mcard(l,v,u,color=col), unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────────────────────────────────────────
def processus():
    st.markdown("## 🏭 Évaluation processus fournisseurs")
    c1,c2,c3 = st.columns(3)
    with c1: pn     = st.text_input("Fournisseur", value="Würth Tunisie SA")
    with c2: pr_ref = st.text_input("Référence",   value="RB-7204")
    with c3: pc     = st.selectbox("Catégorie", ["Composants mécaniques","Matières premières",
                                                  "Fournitures industrielles","Sous-traitance"])
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        pt  = st.number_input("Commandes passées",  value=24, min_value=1)
        po  = st.number_input("Livrées à temps",    value=20, min_value=0)
    with c2:
        pk  = st.number_input("Livrées complètes",  value=22, min_value=0)
        pnc = st.number_input("Non-conformités",    value=2,  min_value=0)
    with c3:
        pla = st.number_input("Délai annoncé (j)",  value=7.0,  step=.5)
        plr = st.number_input("Délai réel (j)",     value=8.4,  step=.1)
    with c4:
        ppr = st.number_input("Traitement (h)",     value=18.0, step=1.0)
        pac = st.number_input("Accusé récep. (h)",  value=4.0,  step=.5)
 
    otd = (po/pt)*100; fr  = (pk/pt)*100; qu  = ((pt-pnc)/pt)*100
    ltp = max(0, 100-((plr-pla)/pla)*100) if pla>0 else 0
    prs = 100 if ppr<=8 else 85 if ppr<=24 else 60 if ppr<=48 else 30
    acs = 100 if pac<=2 else 85 if pac<=8  else 60 if pac<=24 else 30
    gs  = otd*.30 + fr*.20 + qu*.20 + ltp*.15 + prs*.10 + acs*.05
 
    col_c = sc(gs); lbl = sl(gs)
    adv   = {"Excellent":"Fournisseur stratégique. Envisager un accord-cadre.",
             "Bon":      "Performance satisfaisante. Optimiser le délai réel.",
             "Moyen":    "Plan d'amélioration sous 30 jours recommandé.",
             "Insuffisant":"Audit qualité et diversification fournisseur urgent."}[lbl]
 
    st.markdown("---")
    fg = go.Figure(go.Indicator(mode="gauge+number", value=round(gs,1),
        number={"suffix":"/100","font":{"color":col_c,"family":"DM Mono","size":28}},
        gauge={"axis":{"range":[0,100]}, "bar":{"color":col_c,"thickness":.25}, "bgcolor":"#f4f4ee",
               "steps":[{"range":[0,55],"color":"rgba(220,53,53,.15)"},{"range":[55,70],"color":"rgba(245,197,24,.15)"},
                        {"range":[70,85],"color":"rgba(74,140,63,.15)"},{"range":[85,100],"color":"rgba(45,106,45,.2)"}]}))
    fg.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#3d5c3d",family="Syne"),
                     height=220, margin=dict(t=10,b=10,l=20,r=20))
    cg,cd = st.columns([1,2])
    with cg: st.plotly_chart(fg, use_container_width=True)
    with cd:
        st.markdown(f"<div style='background:#ffffff;border-left:4px solid {col_c};border-radius:10px;"
                    f"padding:1.2rem 1.4rem;margin-top:1rem;border:1px solid #d4e0d4'>"
                    f"<div style='font-size:1.2rem;font-weight:600;color:{col_c};margin-bottom:8px'>{lbl} — {pn}</div>"
                    f"<div style='font-size:.85rem;color:#3d5c3d;line-height:1.6'>{adv}</div></div>",
                    unsafe_allow_html=True)
 
    st.markdown("---")
    crit = [("Livraison à temps (OTD)", otd, f"{po}/{pt} · Poids 30%"),
            ("Fill rate",               fr,  f"{pk}/{pt} · Poids 20%"),
            ("Qualité / conformité",    qu,  f"{pt-pnc}/{pt} · Poids 20%"),
            ("Fiabilité délai",         ltp, f"Annoncé:{pla}j Réel:{plr}j · Poids 15%"),
            ("Traitement commande",     prs, f"{ppr}h · Poids 10%"),
            ("Réactivité ACK",          acs, f"{pac}h · Poids 5%")]
    c1,c2 = st.columns(2)
    for i,(nm,scr,nt) in enumerate(crit):
        cc = sc(scr)
        with (c1 if i%2==0 else c2):
            st.markdown(
                f"<div style='background:#ffffff;border:1px solid #d4e0d4;border-radius:10px;"
                f"padding:1rem;margin-bottom:10px;box-shadow:0 1px 4px rgba(26,61,26,.05)'>"
                f"<div style='display:flex;justify-content:space-between;margin-bottom:8px'>"
                f"<span style='font-size:.85rem;font-weight:600;color:#1a2e1a'>{nm}</span>"
                f"<span style='font-family:\"DM Mono\",monospace;color:{cc};font-weight:600'>{scr:.1f}%</span></div>"
                f"<div style='background:#f4f4ee;border-radius:4px;height:6px;overflow:hidden'>"
                f"<div style='background:{cc};height:100%;width:{scr}%'></div></div>"
                f"<div style='font-size:.7rem;color:#7a9a7a;margin-top:6px'>{nt}</div></div>",
                unsafe_allow_html=True)
 
    if otd < 85: st.markdown(f"<div class='alert-warning'>⏱ OTD insuffisant ({otd:.1f}%) — Négocier des pénalités. Majorer le stock de sécurité.</div>", unsafe_allow_html=True)
    if qu  < 95: st.markdown(f"<div class='alert-warning'>⚠️ {pnc} non-conformités — Audit qualité recommandé.</div>", unsafe_allow_html=True)
    if plr > pla: st.markdown(f"<div class='alert-warning'>📅 Dérive délai +{plr-pla:.1f}j — Recalculer le point de réappro. avec le délai réel.</div>", unsafe_allow_html=True)
    if gs  >= 85: st.markdown("<div class='alert-ok'>✅ Performance excellente — Envisager un accord-cadre annuel.</div>", unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────────────────────────────────────────
def alertes():
    st.markdown("## 🔔 Alertes & Surveillance des stocks")
    st.markdown("<p style='color:#7a9a7a'>Importez votre tableau de suivi — les alertes sont générées automatiquement</p>",
                unsafe_allow_html=True)
 
    def clean_num(v):
        if v is None: return 0.0
        s = str(v).strip().replace(" ", "").replace(",", ".")
        if s in ("#DIV/0!", "#N/A", "#VALEUR!", "#REF!", "", "-", "—"): return 0.0
        try: return float(s)
        except: return 0.0
 
    uploaded = st.file_uploader(
        "📂 Déposez votre fichier Excel de suivi des stocks (.xlsx / .xls)",
        type=["xlsx", "xls"], key="alertes_upload")
 
    df_alertes = None
 
    if uploaded:
        try:
            all_sheets = pd.read_excel(uploaded, sheet_name=None, header=None)
            names = list(all_sheets.keys())
            sel = names[0] if len(names) == 1 else st.selectbox("Choisir la feuille", names, key="alert_sheet")
            raw = all_sheets[sel]
 
            hrow = 0
            for i, row in raw.iterrows():
                row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                if any(k in row_str for k in ["codearticle", "article", "stock sécurité", "cons moy", "stk du jour"]):
                    hrow = i
                    break
 
            df_raw = pd.read_excel(uploaded, sheet_name=sel, header=hrow)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            df_raw = df_raw.dropna(how="all")
 
            cmap = {}
            for col in df_raw.columns:
                cl = col.lower().strip()
                if "codearticle" in cl or (cl.startswith("code") and "article" in cl):
                    cmap["Code article"] = col
                elif cl == "article" or "désignation" in cl or "description" in cl:
                    cmap["Désignation"] = col
                elif "stock sécurité" in cl or "stock sécu" in cl or ("stock" in cl and "sécu" in cl):
                    cmap["Stock Sécurité"] = col
                elif "stk du jour tot" in cl or ("stk" in cl and "jour" in cl and "tot" in cl):
                    cmap["Stock actuel"] = col
                elif "stk du jour ss" in cl or ("stk" in cl and "jour" in cl and "ss" in cl):
                    cmap["Stk jour SS"] = col
                elif "cons moy" in cl or "consommation moy" in cl or cl == "cons moy":
                    cmap["Cons Moy"] = col
                elif "couverture" in cl:
                    cmap["Couverture"] = col
                elif "besoin m1" in cl or cl == "besoin m1":
                    cmap["Besoin M1"] = col
                elif "besoin m2" in cl or cl == "besoin m2":
                    cmap["Besoin M2"] = col
                elif "besoin m3" in cl or cl == "besoin m3":
                    cmap["Besoin M3"] = col
                elif "stock m+1" in cl or cl == "stock m+1":
                    cmap["Stock M+1"] = col
                elif "stock m+2" in cl or cl == "stock m+2":
                    cmap["Stock M+2"] = col
                elif "stock m+3" in cl or cl == "stock m+3":
                    cmap["Stock M+3"] = col
                elif "commande en cours" in cl or "commande" in cl:
                    cmap["Commande en cours"] = col
                elif "commentaire" in cl or "comment" in cl:
                    cmap["Commentaire"] = col
                elif "source" in cl:
                    cmap["Source"] = col
                elif cl in ("um", "unité"):
                    cmap["UM"] = col
 
            df_mapped = df_raw.rename(columns={v: k for k, v in cmap.items()})
 
            num_cols = ["Stock Sécurité", "Stock actuel", "Stk jour SS", "Cons Moy",
                        "Besoin M1", "Besoin M2", "Besoin M3",
                        "Stock M+1", "Stock M+2", "Stock M+3", "Commande en cours"]
            for c in num_cols:
                if c in df_mapped.columns:
                    df_mapped[c] = df_mapped[c].apply(clean_num)
 
            if "Code article" in df_mapped.columns:
                df_mapped = df_mapped[df_mapped["Code article"].notna()]
                df_mapped = df_mapped[df_mapped["Code article"].astype(str).str.strip() != ""]
 
            df_alertes = df_mapped
            st.success(f"✅ {len(df_alertes)} articles chargés")
 
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")
 
    if df_alertes is None or len(df_alertes) == 0:
        st.info("⬆️ Importez votre fichier Excel pour générer les alertes automatiquement.")
        st.caption("Format attendu : Codearticle · article · Stock Sécurité · stk du jour tot · Cons Moy · couverture · besoin M1/M2/M3 · Stock M+1/M+2/M+3 · Commentaire")
        return
 
    df = df_alertes.copy()
 
    if "Couverture" not in df.columns or df["Couverture"].sum() == 0:
        df["Couverture"] = np.where(
            df.get("Cons Moy", pd.Series([0]*len(df))) > 0,
            (df.get("Stock actuel", pd.Series([0]*len(df))) / df["Cons Moy"]).round(1),
            np.nan)
 
    def get_status(row):
        stk     = clean_num(row.get("Stock actuel", 0))
        ss      = clean_num(row.get("Stock Sécurité", 0))
        m1      = clean_num(row.get("Stock M+1", 0))
        m2      = clean_num(row.get("Stock M+2", 0))
        m3      = clean_num(row.get("Stock M+3", 0))
        cov     = clean_num(row.get("Couverture", 99))
        comment = str(row.get("Commentaire", "")).strip()
 
        if stk <= 0 or stk < ss:
            return "🔴 Rupture imminente"
        if m1 < 0 or m2 < 0 or cov < 1:
            return "🔴 Rupture imminente"
        if comment.lower().startswith("lancer"):
            return "🟡 DA à lancer"
        if m3 < 0 or cov < 1.5:
            return "🟠 Risque M+3"
        if cov < 3:
            return "🟡 DA à lancer"
        return "🟢 Normal"
 
    df["Statut"] = df.apply(get_status, axis=1)
 
    rupt  = df[df["Statut"] == "🔴 Rupture imminente"]
    da    = df[df["Statut"].isin(["🟡 DA à lancer", "🟠 Risque M+3"])]
    ok    = df[df["Statut"] == "🟢 Normal"]
 
    st.markdown("---")
    st.markdown("#### 📊 Synthèse")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(mcard("🔴 Rupture imminente",   str(len(rupt)), "articles à commander d'urgence", color=C_RED),    unsafe_allow_html=True)
    with c2: st.markdown(mcard("🟡 DA à lancer",         str(len(da)),   "articles nécessitant une DA",    color=C_GOLD),   unsafe_allow_html=True)
    with c3: st.markdown(mcard("🟢 Normal",              str(len(ok)),   "articles en stock suffisant",    color=C_GREEN2), unsafe_allow_html=True)
    with c4: st.markdown(mcard("📦 Total articles",       str(len(df)),   "articles dans le fichier",       color=C_GREEN),  unsafe_allow_html=True)
 
    st.markdown("---")
    st.markdown("#### 📋 Tableau de bord complet")
 
    cols_display = ["Code article", "Désignation"]
    for c in ["Source", "UM", "Stock Sécurité", "Stock actuel", "Cons Moy",
              "Couverture", "Besoin M1", "Besoin M2", "Besoin M3",
              "Stock M+1", "Stock M+2", "Stock M+3", "Commentaire", "Statut"]:
        if c in df.columns:
            cols_display.append(c)
 
    df_show = df[cols_display].copy()
 
    filtre = st.multiselect(
        "Filtrer par statut",
        options=["🔴 Rupture imminente", "🟡 DA à lancer", "🟠 Risque M+3", "🟢 Normal"],
        default=["🔴 Rupture imminente", "🟡 DA à lancer", "🟠 Risque M+3"],
        key="alert_filter")
    if filtre:
        df_show = df_show[df_show["Statut"].isin(filtre)]
 
    col_cfg = {}
    for c in ["Stock Sécurité", "Stock actuel", "Cons Moy", "Besoin M1",
              "Besoin M2", "Besoin M3", "Stock M+1", "Stock M+2", "Stock M+3"]:
        if c in df_show.columns:
            col_cfg[c] = st.column_config.NumberColumn(c, format="%.1f")
    if "Couverture" in df_show.columns:
        col_cfg["Couverture"] = st.column_config.NumberColumn("Couverture (mois)", format="%.1f")
 
    st.dataframe(df_show, use_container_width=True, hide_index=True, column_config=col_cfg)
 
    if "Stock actuel" in df.columns and "Stock Sécurité" in df.columns:
        st.markdown("---")
        st.markdown("#### 📈 Stock actuel vs Stock de sécurité")
 
        labels = df["Code article"].astype(str).tolist()
        colors = [C_RED if s == "🔴 Rupture imminente"
                  else C_GOLD if s in ("🟡 DA à lancer", "🟠 Risque M+3")
                  else C_GREEN2 for s in df["Statut"]]
 
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Stock actuel", x=labels, y=df["Stock actuel"], marker_color=colors))
        fig.add_trace(go.Scatter(name="Stock Sécurité", x=labels, y=df["Stock Sécurité"],
            mode="markers+lines", marker=dict(color=C_PURPLE, size=7, symbol="diamond"),
            line=dict(color=C_PURPLE, width=1.5, dash="dash")))
        if "Stock M+1" in df.columns:
            fig.add_trace(go.Scatter(name="Stock M+1", x=labels, y=df["Stock M+1"],
                mode="markers+lines", marker=dict(color=C_GREEN, size=6),
                line=dict(color=C_GREEN, width=1.2, dash="dot")))
        fig.update_layout(**light_layout(), height=360, barmode="group",
            xaxis=dict(tickangle=-45, gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="Unités", gridcolor="rgba(0,0,0,.06)"),
            shapes=[dict(type="line", x0=-0.5, x1=len(labels)-0.5, y0=0, y1=0,
                         line=dict(color=C_RED, width=1.5, dash="dash"))])
        st.plotly_chart(fig, use_container_width=True)
 
    st.markdown("---")
    st.markdown("#### 🚨 Actions prioritaires")
 
    if len(rupt) == 0 and len(da) == 0:
        st.markdown("<div class='alert-ok'>✅ Aucune alerte critique — tous les stocks sont suffisants.</div>", unsafe_allow_html=True)
 
    for _, r in rupt.iterrows():
        code    = str(r.get("Code article", ""))
        nom     = str(r.get("Désignation", ""))
        stk     = clean_num(r.get("Stock actuel", 0))
        ss      = clean_num(r.get("Stock Sécurité", 0))
        cov     = clean_num(r.get("Couverture", 0))
        comment = str(r.get("Commentaire", "")).strip()
        m1      = clean_num(r.get("Stock M+1", 0))
        detail  = f"Stock actuel : <strong>{stk:,.0f}</strong> | SS : {ss:,.0f} | Couverture : {cov:.1f} mois | Stock M+1 : {m1:,.0f}"
        da_txt  = f" → <em>{comment}</em>" if comment and comment != "nan" else ""
        st.markdown(
            f"<div class='alert-critical'>🔴 <strong>{code} — {nom}</strong>{da_txt}<br/>"
            f"<span style='font-size:.85rem'>{detail}</span></div>",
            unsafe_allow_html=True)
 
    for _, r in da.iterrows():
        code    = str(r.get("Code article", ""))
        nom     = str(r.get("Désignation", ""))
        stk     = clean_num(r.get("Stock actuel", 0))
        ss      = clean_num(r.get("Stock Sécurité", 0))
        cov     = clean_num(r.get("Couverture", 0))
        comment = str(r.get("Commentaire", "")).strip()
        m3      = clean_num(r.get("Stock M+3", 0))
        detail  = f"Stock actuel : <strong>{stk:,.0f}</strong> | SS : {ss:,.0f} | Couverture : {cov:.1f} mois | Stock M+3 : {m3:,.0f}"
        da_txt  = f" → <em>{comment}</em>" if comment and comment != "nan" else ""
        cls     = "alert-warning" if r["Statut"] == "🟡 DA à lancer" else "alert-info"
        st.markdown(
            f"<div class='{cls}'>{r['Statut']} <strong>{code} — {nom}</strong>{da_txt}<br/>"
            f"<span style='font-size:.85rem'>{detail}</span></div>",
            unsafe_allow_html=True)
 
 
# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    URL_LOGO = "https://raw.githubusercontent.com/Ghofrane13/medoil-supply-app/main/logo.png"
    st.markdown(
    f"""
    <div style="display: flex; align-items: center; padding: .3rem 0 .5rem">
        <img src={URL_LOGO} style="width: 44px; margin-right: 10px; border-radius:6px">
        <div>
            <div style="color:#ffffff;font-weight:700;font-size:1.15rem;line-height:1.1">Med oil</div>
            <div style="color:#f5c518;font-size:.65rem;letter-spacing:.1em;text-transform:uppercase">Supply Chain · v2.0</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
    )
    st.divider()
 
    pg = st.navigation([
        st.Page(accueil,       title="Accueil",                    icon="🏠"),
        st.Page(import_calcul, title="Import & Calcul automatique", icon="📥"),
        st.Page(calculateurs,  title="Calculateurs SC",             icon="⚙️"),
        st.Page(processus,     title="Éval. Fournisseurs",          icon="🏭"),
        st.Page(alertes,       title="Alertes & Stocks",            icon="🔔"),
    ])
 
with st.sidebar:
    st.divider()
    st.markdown("<p style='color:#4a6a4a;font-size:.72rem;'>© 2025 Medoil — Tous droits réservés</p>",
                unsafe_allow_html=True)
 
pg.run()
