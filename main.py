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

def sc(s): return C_GREEN2 if s >= 85 else "#14b8a6" if s >= 70 else C_GOLD if s >= 55 else C_RED
def sl(s): return "Excellent" if s >= 85 else "Bon" if s >= 70 else "Moyen" if s >= 55 else "Insuffisant"

# ── Délais par source (en jours) ──────────────────────────────────────────────
SOURCE_DELAIS = {
    "Export":  4 * 30,   # 4 mois
    "Local":   14,        # 2 semaines
    "BM":      21,        # 3 semaines
}

# ── Calculs SC ────────────────────────────────────────────────────────────────
def get_lt_jours(source):
    return SOURCE_DELAIS.get(source, 30)

def calc_ss(sigma_d, lt_j, z=1.65):
    """SS = Z × σD × √L   (σD = écart-type demande, L = délai en jours)."""
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

# ── CBN sur 3 mois ────────────────────────────────────────────────────────────
def calc_cbn(d_j, ss, stock_actuel, commande_cours, mois_labels):
    """Calcule le CBN sur 3 périodes (mois)."""
    rows = []
    stock = stock_actuel + commande_cours
    for i, mois in enumerate(mois_labels):
        besoin_brut = d_j * 30
        besoin_net  = max(0, besoin_brut - stock)
        stock_fin   = max(0, stock - besoin_brut)
        rows.append({
            "Période":        mois,
            "Besoin brut (u)": round(besoin_brut),
            "Stock début (u)": round(stock),
            "Besoin net (u)":  round(besoin_net),
            "Stock fin (u)":   round(stock_fin),
            "Statut":          "🔴 Rupture" if stock_fin <= ss else ("🟡 Sous SS" if stock_fin < ss * 1.5 else "🟢 OK"),
        })
        stock = stock_fin
    return rows

# ── Export Excel complet ──────────────────────────────────────────────────────
def build_excel_complet(df_res, periode_debut=None):
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

    # ── Feuille 1 : Indicateurs SC ────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Indicateurs SC"

    headers_row1 = [
        ("Code article",           "C8E6C9"), ("Description",             "C8E6C9"),
        ("Source approv.",          "FFF9C4"), ("Délai (j)",               "FFF9C4"),
        ("Consommation/mois (moy)","DCEDC8"), ("Sigma demande (u/j)",     "DCEDC8"),
        ("Niveau service",          "FFF9C4"), ("Z",                       "FFF9C4"),
        ("Stock sécurité (u)",     "FFE0B2"), ("Coût unitaire (TND)",     "FFE0B2"),
        ("Coût SS total (TND)",    "FFE0B2"), ("EOQ (u)",                 "B3E5FC"),
        ("Nb cmd/an",              "B3E5FC"), ("Point de commande (u)",   "FFE0B2"),
        ("Stock actuel (u)",       "E8F5E9"), ("Commande en cours (u)",   "E8F5E9"),
    ]

    for col, (label, bg) in enumerate(headers_row1, 1):
        hdr(ws1.cell(1, col), label, bg)

    fmts1 = [None, None, None, "0", "#,##0.0", "#,##0.00", "0%", "0.00",
              "#,##0", "#,##0.00", "#,##0.00", "#,##0", "0.0", "#,##0", "#,##0", "#,##0"]

    for i, row in df_res.iterrows():
        r = i + 2
        vals = [
            row.get("Code article", ""),
            row.get("Description Article", ""),
            row.get("Source", ""),
            row.get("LT jours", 0),
            round(row.get("Conso mois moy", 0), 1),
            round(row.get("Sigma D", 0), 3),
            row.get("Niveau service pct", 0.95),
            row.get("Z val", 1.65),
            row.get("Stock sécurité", 0),
            row.get("Coût unitaire", 0),
            row.get("Coût SS", 0),
            row.get("EOQ", 0),
            row.get("Nb cmd an", 0),
            row.get("Point de commande", 0),
            row.get("Stock actuel", 0),
            row.get("Commande en cours", 0),
        ]
        for col, (val, nf) in enumerate(zip(vals, fmts1), 1):
            dat(ws1.cell(r, col), val, nf)

    for col, w in enumerate([14, 30, 10, 10, 18, 16, 14, 8, 18, 18, 18, 12, 10, 18, 16, 18], 1):
        ws1.column_dimensions[get_column_letter(col)].width = w
    ws1.row_dimensions[1].height = 42
    ws1.freeze_panes = "A2"

    # ── Feuille 2 : CBN 3 mois ────────────────────────────────────────────────
    ws2 = wb.create_sheet("CBN 3 mois")

    if periode_debut is None:
        now = datetime.now()
        mois_labels = []
        for i in range(3):
            m = (now.month - 1 + i) % 12 + 1
            y = now.year + ((now.month - 1 + i) // 12)
            mois_labels.append(f"{calendar.month_abbr[m]} {y}")
    else:
        mois_labels = []
        for i in range(3):
            m = (periode_debut.month - 1 + i) % 12 + 1
            y = periode_debut.year + ((periode_debut.month - 1 + i) // 12)
            mois_labels.append(f"{calendar.month_abbr[m]} {y}")

    cbn_headers = ["Code article", "Description", "Source", "SS (u)", "Stock actuel (u)", "Cmd en cours (u)"]
    for mois in mois_labels:
        cbn_headers += [f"Besoin brut {mois}", f"Stock début {mois}", f"Besoin net {mois}", f"Stock fin {mois}", f"Statut {mois}"]

    cbn_colors = {
        "Code article": "C8E6C9", "Description": "C8E6C9", "Source": "FFF9C4",
        "SS (u)": "FFE0B2", "Stock actuel (u)": "E8F5E9", "Cmd en cours (u)": "E8F5E9",
    }
    mois_bgs = ["DCEDC8", "B3E5FC", "F8BBD9"]
    for i, mois in enumerate(mois_labels):
        bg = mois_bgs[i % 3]
        for suffix in [f"Besoin brut {mois}", f"Stock début {mois}", f"Besoin net {mois}", f"Stock fin {mois}", f"Statut {mois}"]:
            cbn_colors[suffix] = bg

    for col, label in enumerate(cbn_headers, 1):
        hdr(ws2.cell(1, col), label, cbn_colors.get(label, "F5F5F5"))

    for i, row in df_res.iterrows():
        r  = i + 2
        d_j = row.get("D jour", 0)
        ss  = row.get("Stock sécurité", 0)
        stk = row.get("Stock actuel", 0)
        cmd = row.get("Commande en cours", 0)
        cbn = calc_cbn(d_j, ss, stk, cmd, mois_labels)

        base_vals = [
            row.get("Code article", ""), row.get("Description Article", ""),
            row.get("Source", ""), ss, stk, cmd
        ]
        for col, val in enumerate(base_vals, 1):
            dat(ws2.cell(r, col), val)

        col_idx = len(base_vals) + 1
        for cbn_row in cbn:
            dat(ws2.cell(r, col_idx),     cbn_row["Besoin brut (u)"], "#,##0")
            dat(ws2.cell(r, col_idx + 1), cbn_row["Stock début (u)"], "#,##0")
            dat(ws2.cell(r, col_idx + 2), cbn_row["Besoin net (u)"],  "#,##0")
            dat(ws2.cell(r, col_idx + 3), cbn_row["Stock fin (u)"],   "#,##0")
            dat(ws2.cell(r, col_idx + 4), cbn_row["Statut"])
            col_idx += 5

    for col in range(1, len(cbn_headers) + 1):
        ws2.column_dimensions[get_column_letter(col)].width = 16
    ws2.column_dimensions["B"].width = 30
    ws2.row_dimensions[1].height = 42
    ws2.freeze_panes = "G2"

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
             border:1px solid rgba(245,197,24,.3)'>✦ Outil supply chain manager · v3.0</div>
        <h1 style='font-family:"DM Serif Display",serif;font-size:2.3rem;color:#ffffff;margin:.4rem 0;line-height:1.25'>
            Pilotez votre <em style='color:#f5c518'>supply chain</em><br/>avec précision</h1>
        <p style='color:#a5c8a5;font-size:.95rem;margin:.8rem 0 1.4rem'>
            Importez vos données de consommation — SS, EOQ, CBN et alertes calculés automatiquement selon la source d'approvisionnement.</p>
    </div>""", unsafe_allow_html=True)

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

    st.markdown("<div style='font-size:1rem;font-weight:700;color:#1a2e1a;margin-bottom:1rem;border-left:3px solid #f5c518;padding-left:10px'>Modules disponibles</div>", unsafe_allow_html=True)
    cols = st.columns(4)
    mods = [
        ("📥", "#2d6a2d", "Calcul automatique",    "Importez consommations → SS · EOQ · CBN → Export XLSX"),
        ("📐", "#c9a000", "Calculateurs SC",         "EOQ · SS · KPIs · Point de réappro. avec courbes"),
        ("🔔", "#dc3535", "Alertes & Stocks",         "Alertes rupture · Déclenchement commandes automatique"),
        ("📈", "#4a8c3f", "Évolution & Gains",        "Suivi SS sur période · évaluation économies réalisées"),
    ]
    for col, (ic, accent, ti, de) in zip(cols, mods):
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


# ─────────────────────────────────────────────────────────────────────────────
def import_calcul():
    st.markdown("## 📥 Calcul automatique — Stock de sécurité, EOQ & CBN")
    st.markdown("<p style='color:#7a9a7a'>Importez votre fichier avec les consommations historiques — délais calculés automatiquement selon la source</p>", unsafe_allow_html=True)

    with st.expander("📋 Format attendu du fichier Excel"):
        st.markdown("""
**Le fichier peut avoir 1 ou 2 feuilles — l'application traite chaque feuille séparément.**

| Colonne | Obligatoire | Description |
|---|---|---|
| `Code article` | ✅ | Identifiant unique |
| `Description Article` | ✅ | Désignation |
| `Source` | ✅ | **Export**, **Local** ou **BM** → délai calculé automatiquement |
| `Consommation` | ✅ | Une **seule colonne** avec la valeur mensuelle (la moyenne est calculée) |
| `Coût unitaire` | ✅ | Saisi manuellement (en TND) |
| `Stock actuel` | — | Stock physique du jour |
| `Commande en cours` | — | Quantité commandée non reçue |

> **Délais automatiques :** Export = 4 mois · Local = 2 semaines · BM = 3 semaines
""")

    st.markdown("---")
    st.markdown("#### ⚙️ Paramètres globaux")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        p_z_opt = st.selectbox("Niveau de service", ["90% (Z=1.28)", "95% (Z=1.65)", "97.5% (Z=1.96)", "99% (Z=2.33)"], index=1)
    with c2:
        p_sc = st.number_input("Coût passation (TND)", value=150.0, min_value=1.0, step=10.0)
    with c3:
        p_taux = st.number_input("Taux stockage (%)", value=20.0, min_value=1.0, step=1.0)
    with c4:
        periode_debut = st.date_input("Début période CBN", value=date.today())

    Z_MAP = {"90% (Z=1.28)": 1.28, "95% (Z=1.65)": 1.65, "97.5% (Z=1.96)": 1.96, "99% (Z=2.33)": 2.33}
    Z = Z_MAP[p_z_opt]
    NS_pct = float(p_z_opt.split("%")[0]) / 100

    st.divider()
    st.markdown("#### 📂 Chargement du fichier")
    uploaded = st.file_uploader("Déposez votre fichier Excel (.xlsx / .xls)", type=["xlsx", "xls"])
    use_demo = st.checkbox("Utiliser les données de démo", value=not bool(uploaded))

    # ── Gestion multi-feuilles ─────────────────────────────────────────────────
    all_sheets_data = {}

    if uploaded:
        try:
            all_raw = pd.read_excel(uploaded, sheet_name=None, header=None)
            sheet_names = list(all_raw.keys())

            if len(sheet_names) > 1:
                st.info(f"📋 **{len(sheet_names)} feuilles détectées :** {', '.join(sheet_names)}")
                feuille_active = st.radio("Feuille à analyser", sheet_names, horizontal=True)
                sheets_to_process = [feuille_active]
            else:
                sheets_to_process = sheet_names

            for sheet_sel in sheets_to_process:
                raw = all_raw[sheet_sel]
                hrow = 0
                for i, row in raw.iterrows():
                    row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                    if any(k in row_str for k in ["code", "article", "consommation", "source"]):
                        hrow = i
                        break

                df_raw = pd.read_excel(uploaded, sheet_name=sheet_sel, header=hrow)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                df_raw = df_raw.dropna(how="all")

                cmap = {}
                for col in df_raw.columns:
                    cl = col.lower().strip()
                    if "code" in cl and "article" in cl:                        cmap["Code article"] = col
                    elif "description" in cl or "désignation" in cl:            cmap["Description Article"] = col
                    elif "source" in cl:                                         cmap["Source"] = col
                    elif "consommation" in cl or cl in ("conso", "cons"):        cmap["Consommation"] = col
                    elif "coût unitaire" in cl or "cout unitaire" in cl or "prix" in cl: cmap["Coût unitaire"] = col
                    elif "stock actuel" in cl or "stk" in cl or ("stock" in cl and "actuel" in cl): cmap["Stock actuel"] = col
                    elif "commande en cours" in cl or "cmd en cours" in cl:      cmap["Commande en cours"] = col

                df_s = df_raw.rename(columns={v: k for k, v in cmap.items()})
                for c in ["Consommation", "Coût unitaire", "Stock actuel", "Commande en cours"]:
                    if c in df_s.columns:
                        df_s[c] = pd.to_numeric(df_s[c], errors="coerce").fillna(0)

                if "Code article" in df_s.columns:
                    df_s = df_s[df_s["Code article"].notna()]
                    df_s = df_s[df_s["Code article"].astype(str).str.strip() != ""]

                if "Source" not in df_s.columns:
                    df_s["Source"] = "Export"

                all_sheets_data[sheet_sel] = df_s
                st.success(f"✅ Feuille **{sheet_sel}** : {len(df_s)} articles chargés")

        except Exception as e:
            st.error(f"Erreur de lecture : {e}")

    if use_demo and not all_sheets_data:
        demo_df = pd.DataFrame({
            "Code article":       [40007, 40010, 40011, 40014, 40015, 40036, 40042, 40055, 40062, 40063],
            "Description Article":["NYSOSEL (NICKEL CATALYSEUR)", "TERRE TONSIL", "PAPIER FILTRE",
                                    "TOILE FILTRANTE TERRE", "PERLITE FILTRATION PAF1", "SEL MARIN FIN",
                                    "MYVEROL 18-04", "TRISYL", "VITAMINE A", "VITAMINE E"],
            "Source":             ["Export", "Export", "Local", "Export", "BM", "Local", "Export", "BM", "Export", "Local"],
            "Consommation":       [324, 5886, 509, 32460, 959, 3747, 3147, 4432, 16, 22],
            "Coût unitaire":      [59.23, 1.709, 3.823, 49.787, 1.700, 0.85, 2.10, 8.189, 12.5, 18.0],
            "Stock actuel":       [1200, 8000, 500, 50000, 2000, 5000, 8000, 9000, 30, 50],
            "Commande en cours":  [0, 5000, 0, 0, 1000, 0, 3000, 0, 0, 0],
        })
        all_sheets_data["Démo"] = demo_df
        st.info("📊 Données de démo chargées")

    if not all_sheets_data:
        return

    # ── Traitement par feuille ──────────────────────────────────────────────────
    for sheet_name, df_source in all_sheets_data.items():
        if len(all_sheets_data) > 1:
            st.markdown(f"---\n### 📄 Feuille : **{sheet_name}**")

        # Saisie coût unitaire si manquant
        if "Coût unitaire" not in df_source.columns or df_source["Coût unitaire"].sum() == 0:
            st.warning("⚠️ Colonne **Coût unitaire** non détectée — saisissez les valeurs manuellement :")
            cu_vals = {}
            cols_cu = st.columns(min(4, len(df_source)))
            for i, (_, row) in enumerate(df_source.iterrows()):
                with cols_cu[i % len(cols_cu)]:
                    cu_vals[row.get("Code article", i)] = st.number_input(
                        f"Coût {row.get('Code article','Art'+str(i))}",
                        value=0.0, min_value=0.0, step=0.1,
                        key=f"cu_{sheet_name}_{i}"
                    )
            df_source["Coût unitaire"] = df_source["Code article"].map(cu_vals).fillna(0)

        with st.expander(f"👁 Aperçu données sources — {sheet_name}", expanded=False):
            st.dataframe(df_source, use_container_width=True, hide_index=True)

        # ── Calculs ────────────────────────────────────────────────────────────
        results = []
        for _, row in df_source.iterrows():
            conso   = float(row.get("Consommation", 0) or 0)
            source  = str(row.get("Source", "Export")).strip()
            cu      = float(row.get("Coût unitaire", 0) or 0)
            stk     = float(row.get("Stock actuel", 0) or 0)
            cmd     = float(row.get("Commande en cours", 0) or 0)

            lt_j    = get_lt_jours(source)
            d_j     = conso / 30
            # σD = écart-type de la demande journalière (15% de variabilité)
            sig_d   = conso * 0.15 / 30

            # SS = Z × σD × √L
            ss      = calc_ss(sig_d, lt_j, Z)
            pr      = calc_pr(d_j, lt_j, ss)
            conso_an = conso * 12
            eoq, nb_cmd, _ = calc_eoq(conso_an, cu, sc_cost=p_sc, taux=p_taux / 100)
            css     = round(ss * cu, 2) if cu > 0 else 0

            results.append({
                "Code article":       row.get("Code article", ""),
                "Description Article": row.get("Description Article", ""),
                "Source":             source,
                "LT jours":           lt_j,
                "Conso mois moy":     conso,
                "Sigma D":            round(sig_d, 4),
                "D jour":             round(d_j, 4),
                "Niveau service pct": NS_pct,
                "Z val":              Z,
                "Stock sécurité":     ss,
                "Coût unitaire":      cu,
                "Coût SS":            css,
                "EOQ":                eoq,
                "Nb cmd an":          nb_cmd,
                "Point de commande":  pr,
                "Stock actuel":       stk,
                "Commande en cours":  cmd,
            })

        df_res = pd.DataFrame(results)

        # ── Tableau résultats ──────────────────────────────────────────────────
        st.divider()
        st.markdown("#### 🧮 Indicateurs calculés")
        disp = df_res[["Code article", "Description Article", "Source", "LT jours",
                        "Stock sécurité", "Coût SS", "EOQ", "Point de commande",
                        "Stock actuel", "Commande en cours"]].copy()
        st.dataframe(disp, use_container_width=True, hide_index=True,
            column_config={
                "LT jours":           st.column_config.NumberColumn("Délai (j)", format="%d"),
                "Stock sécurité":     st.column_config.NumberColumn("SS (u)", format="%d"),
                "Coût SS":            st.column_config.NumberColumn("Coût SS (TND)", format="%.2f"),
                "EOQ":                st.column_config.NumberColumn("EOQ (u)", format="%d"),
                "Point de commande":  st.column_config.NumberColumn("Point de commande (u)", format="%d"),
                "Stock actuel":       st.column_config.NumberColumn("Stock actuel (u)", format="%d"),
                "Commande en cours":  st.column_config.NumberColumn("Commande en cours (u)", format="%d"),
            })

        # ── CBN ────────────────────────────────────────────────────────────────
        st.divider()
        st.markdown("#### 📅 CBN — Calcul des Besoins Nets sur 3 mois")
        pd_debut = datetime(periode_debut.year, periode_debut.month, 1)
        mois_labels = []
        for i in range(3):
            m = (pd_debut.month - 1 + i) % 12 + 1
            y = pd_debut.year + ((pd_debut.month - 1 + i) // 12)
            mois_labels.append(f"{calendar.month_abbr[m]} {y}")

        all_cbn = []
        for _, row in df_res.iterrows():
            cbn_rows = calc_cbn(row["D jour"], row["Stock sécurité"],
                                row["Stock actuel"], row["Commande en cours"], mois_labels)
            for cbn_row in cbn_rows:
                cbn_row["Code article"] = row["Code article"]
                cbn_row["Description"]  = row["Description Article"]
                all_cbn.append(cbn_row)

        df_cbn = pd.DataFrame(all_cbn)
        st.dataframe(
            df_cbn[["Code article", "Description", "Période", "Besoin brut (u)",
                     "Stock début (u)", "Besoin net (u)", "Stock fin (u)", "Statut"]],
            use_container_width=True, hide_index=True,
            column_config={
                "Besoin brut (u)":  st.column_config.NumberColumn(format="%d"),
                "Stock début (u)":  st.column_config.NumberColumn(format="%d"),
                "Besoin net (u)":   st.column_config.NumberColumn(format="%d"),
                "Stock fin (u)":    st.column_config.NumberColumn(format="%d"),
            })

        # ── Synthèse ────────────────────────────────────────────────────────────
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        ruptures_m1 = (df_cbn[df_cbn["Période"] == mois_labels[0]]["Statut"] == "🔴 Rupture").sum()
        with c1: st.markdown(mcard("Articles traités", str(len(df_res)), "articles", color=C_GREEN), unsafe_allow_html=True)
        with c2: st.markdown(mcard("Coût total SS", f"{df_res['Coût SS'].sum():,.0f}", "TND immobilisés", color=C_GOLD), unsafe_allow_html=True)
        with c3: st.markdown(mcard("EOQ moyen", fmtInt(df_res[df_res['EOQ'] > 0]['EOQ'].mean()), "u / commande", color=C_GREEN2), unsafe_allow_html=True)
        with c4: st.markdown(mcard("🔴 Ruptures M+1", str(ruptures_m1), "articles en risque", color=C_RED), unsafe_allow_html=True)

        # ── Graphique ───────────────────────────────────────────────────────────
        st.markdown("---")
        labels = df_res["Code article"].astype(str).tolist()
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Stock actuel", x=labels, y=df_res["Stock actuel"], marker_color=C_GREEN2))
        fig.add_trace(go.Bar(name="EOQ", x=labels, y=df_res["EOQ"], marker_color=C_GREEN))
        fig.add_trace(go.Scatter(name="Stock sécurité", x=labels, y=df_res["Stock sécurité"],
                                  mode="markers+lines", marker=dict(color=C_GOLD, size=9, symbol="diamond"),
                                  line=dict(color=C_GOLD, width=2, dash="dash")))
        fig.add_trace(go.Scatter(name="Point de commande", x=labels, y=df_res["Point de commande"],
                                  mode="markers+lines", marker=dict(color=C_RED, size=8, symbol="x"),
                                  line=dict(color=C_RED, width=1.5, dash="dot")))
        fig.update_layout(**light_layout(), barmode="group", height=380,
                           xaxis=dict(title="Article", gridcolor="rgba(0,0,0,.06)"),
                           yaxis=dict(title="Unités", gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)

        # ── Export ─────────────────────────────────────────────────────────────
        st.divider()
        st.markdown("#### 💾 Export XLSX — Indicateurs + CBN")
        st.info("Le fichier exporté contient 2 feuilles : **Indicateurs SC** et **CBN 3 mois** — prêt à importer dans l'onglet Alertes.")
        buf = build_excel_complet(df_res, periode_debut=pd_debut)
        st.download_button(
            label="⬇️ Télécharger le fichier Excel complet",
            data=buf,
            file_name=f"medoil_sc_{sheet_name}_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # Sauvegarder en session pour les alertes
        if "df_resultats" not in st.session_state:
            st.session_state["df_resultats"] = {}
        st.session_state["df_resultats"][sheet_name] = df_res
        st.session_state["Z"] = Z
        st.session_state["periode_debut"] = pd_debut


# ─────────────────────────────────────────────────────────────────────────────
def calculateurs():
    st.markdown("## ⚙️ Calculateurs Supply Chain")
    tab_ss, tab_eoq, tab_kpi, tab_rp = st.tabs(
        ["📦 Stock de sécurité", "📐 EOQ (Wilson)", "📊 KPIs stock", "🔄 Point de réappro."])

    with tab_ss:
        st.markdown(fbox("Formule", "SS = Z × σD × √L",
                         "Z = facteur service · σD = écart-type demande (u/j) · L = délai (j) · Délais : Export=120j · Local=14j · BM=21j"), unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            source_ss = st.selectbox("Source approvisionnement", ["Export", "Local", "BM"], key="ss_source")
            lt_auto   = get_lt_jours(source_ss)
            st.info(f"Délai automatique : **{lt_auto} jours**")
            sd   = st.number_input("Demande moy. (u/j)",          value=50.0, step=1.0, key="ss_sd")
            ss2  = st.number_input("Écart-type demande σD (u/j)", value=8.0,  step=0.5, key="ss_sigma")
        with c2:
            scu  = st.number_input("Coût unitaire (TND)",          value=25.0, step=1.0, key="ss_cu")
        with c3:
            Z_MAP_SS = {
                "95%": 1.65, "96%": 1.75, "97%": 1.88, "98%": 2.05, "99%": 2.33,
                "90%": 1.28, "85%": 1.04, "80%": 0.84,
            }
            szo = st.selectbox("Niveau de service", list(Z_MAP_SS.keys()), index=0, key="ss_ns")

        Z2  = Z_MAP_SS[szo]
        # SS = Z × σD × √L
        SS_calc = Z2 * ss2 * math.sqrt(lt_auto)

        c = st.columns(3)
        for i, (l, v, u, col) in enumerate([
            ("SS calculé",         fmtInt(SS_calc),       "unités", C_GREEN),
            ("Z × σD × √L",        f"{Z2} × {ss2} × √{lt_auto}", "détail calcul", C_GREEN2),
            ("Coût immobilisé SS", fmtInt(SS_calc * scu), "TND",    C_GOLD),
        ]):
            with c[i]: st.markdown(mcard(l, v, u, color=col), unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom:14px'></div>", unsafe_allow_html=True)
        z_vals  = [0.84, 1.04, 1.28, 1.65, 1.88, 2.05, 2.33]
        ns_vals = [80,   85,   90,   95,   97,   98,   99]
        ss_v    = [round(z * ss2 * math.sqrt(lt_auto)) for z in z_vals]
        current_pct = int(szo.replace("%", ""))
        fig = go.Figure(go.Bar(
            x=[f"{n}%" for n in ns_vals], y=ss_v,
            marker_color=[C_GOLD if n == current_pct else C_GREEN for n in ns_vals],
            text=ss_v, textposition="outside"))
        fig.update_layout(**light_layout(), height=260,
            xaxis=dict(title="Niveau de service", gridcolor="rgba(0,0,0,.06)"),
            yaxis=dict(title="SS (unités)", gridcolor="rgba(0,0,0,.06)"))
        st.plotly_chart(fig, use_container_width=True)

    with tab_eoq:
        st.markdown(fbox("Modèle de Wilson", "EOQ = √(2×D×Sc / (Sh×Cu))",
                         "D=demande annuelle · Sc=coût passation · Sh=taux stockage · Cu=coût unitaire"), unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            eD  = st.number_input("Demande annuelle (u)", value=12000, step=100,  key="eoq_D")
            eSc = st.number_input("Coût passation (TND)", value=150.0, step=5.0,  key="eoq_Sc")
        with c2:
            eCu = st.number_input("Coût unitaire (TND)",  value=80.0,  step=1.0,  key="eoq_Cu")
            eSh = st.number_input("Taux stockage (%)",    value=20.0,  step=1.0,  key="eoq_Sh")
        with c3:
            eLT = st.number_input("Délai fourni. (j)",    value=10,    step=1,    key="eoq_LT")
            eDy = st.number_input("Jours ouvr./an",       value=250,   step=1,    key="eoq_Dy")

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
            rss = st.number_input("Stock sécu. (u)",    value=132.0, step=10.0, key="rp_rss")
            rcu = st.number_input("Stock actuel (u)",   value=500.0, step=10.0, key="rp_rcu")
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

    # Charger depuis session ou fichier
    df_from_session = None
    if "df_resultats" in st.session_state and st.session_state["df_resultats"]:
        sheets_calc = list(st.session_state["df_resultats"].keys())
        src_choice = st.radio("Source des données", ["📥 Depuis calcul automatique", "📂 Importer fichier XLSX"], horizontal=True)
    else:
        src_choice = "📂 Importer fichier XLSX"

    df_alerte_all = {}

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

                    # Map colonnes
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

        # ── Saisie stock actuel si manquant ──────────────────────────────────
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

        if ss_col not in df.columns:
            df[ss_col] = 0
        if stk_col not in df.columns:
            df[stk_col] = 0
        if pr_col not in df.columns:
            df[pr_col] = 0

        # ── Calcul alertes ──────────────────────────────────────────────────
        def get_alerte(row):
            stk = clean_num(row.get(stk_col, 0))
            ss  = clean_num(row.get(ss_col, 0))
            pr  = clean_num(row.get(pr_col, 0))
            cmd = clean_num(row.get("Commande en cours", 0))

            alerte_ss  = stk < ss         # Dépassement SS
            alerte_cmd = stk <= pr and not (stk + cmd > pr)  # Commander maintenant

            if stk <= 0 or stk < ss * 0.5:
                return "🔴 Rupture critique"
            elif alerte_ss:
                return "🟠 SS dépassé"
            elif alerte_cmd:
                return "🟡 Commander maintenant"
            else:
                return "🟢 Normal"

        df["Alerte"] = df.apply(get_alerte, axis=1)

        # Synthèse
        rupt  = df[df["Alerte"] == "🔴 Rupture critique"]
        ss_dep = df[df["Alerte"] == "🟠 SS dépassé"]
        cmd_now = df[df["Alerte"] == "🟡 Commander maintenant"]
        ok    = df[df["Alerte"] == "🟢 Normal"]

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(mcard("🔴 Rupture critique",   str(len(rupt)),    "articles",  color=C_RED),    unsafe_allow_html=True)
        with c2: st.markdown(mcard("🟠 SS dépassé",          str(len(ss_dep)),  "articles",  color="#e65100"), unsafe_allow_html=True)
        with c3: st.markdown(mcard("🟡 Commander maintenant", str(len(cmd_now)), "articles",  color=C_GOLD),   unsafe_allow_html=True)
        with c4: st.markdown(mcard("🟢 Normal",              str(len(ok)),      "articles",  color=C_GREEN2), unsafe_allow_html=True)

        st.divider()

        # Tableau
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

        # Graphique alertes
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

        # Détail alertes critiques
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
def evolution_gains():
    st.markdown("## 📈 Évolution SS & Évaluation des gains")
    st.markdown("<p style='color:#7a9a7a'>Suivez l'évolution du stock de sécurité sur une période et mesurez les économies générées par les nouveaux indicateurs</p>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📊 Évolution temporelle SS", "💰 Évaluation des gains"])

    with tab1:
        st.markdown("#### Saisir les valeurs SS observées sur la période")
        st.caption("Entrez le stock de sécurité réel (ou théorique) pour chaque mois de la période d'analyse.")

        c1, c2 = st.columns(2)
        with c1:
            debut_ev = st.date_input("Début période", value=date(date.today().year, 1, 1), key="ev_debut")
        with c2:
            nb_mois = st.number_input("Nombre de mois", value=6, min_value=2, max_value=24, step=1)

        mois_ev = []
        for i in range(nb_mois):
            m = (debut_ev.month - 1 + i) % 12 + 1
            y = debut_ev.year + ((debut_ev.month - 1 + i) // 12)
            mois_ev.append(f"{calendar.month_abbr[m]} {y}")

        # Depuis session ou saisie
        n_articles = st.number_input("Nombre d'articles à suivre", value=3, min_value=1, max_value=20, step=1)

        data_ev = []
        for art_i in range(int(n_articles)):
            with st.expander(f"Article {art_i + 1}", expanded=(art_i == 0)):
                cc1, cc2, cc3 = st.columns(3)
                with cc1:
                    art_code = st.text_input("Code article", value=f"ART-{art_i+1:03d}", key=f"ev_code_{art_i}")
                with cc2:
                    ss_ancien = st.number_input("SS avant (u)", value=500, step=10, key=f"ev_ss_avant_{art_i}")
                with cc3:
                    ss_nouveau = st.number_input("SS calculé (u)", value=350, step=10, key=f"ev_ss_apres_{art_i}")

                cu_ev = st.number_input("Coût unitaire (TND)", value=25.0, step=1.0, key=f"ev_cu_{art_i}")

                ss_vals = []
                col_mois = st.columns(min(6, len(mois_ev)))
                for i, mois in enumerate(mois_ev):
                    with col_mois[i % 6]:
                        v = st.number_input(mois, value=int(ss_ancien - (ss_ancien - ss_nouveau) * i / max(1, len(mois_ev)-1)),
                                            step=10, key=f"ev_{art_i}_{i}")
                        ss_vals.append(v)

                data_ev.append({
                    "code": art_code, "ss_avant": ss_ancien, "ss_nouveau": ss_nouveau,
                    "cu": cu_ev, "vals": ss_vals
                })

        if st.button("📊 Générer le graphique d'évolution"):
            fig = go.Figure()
            for art in data_ev:
                fig.add_trace(go.Scatter(
                    x=mois_ev, y=art["vals"], name=art["code"],
                    mode="lines+markers", line=dict(width=2.5),
                    marker=dict(size=8)))
                fig.add_hline(y=art["ss_avant"],   line_dash="dot",  line_color=C_RED,
                               annotation_text=f"{art['code']} SS avant")
                fig.add_hline(y=art["ss_nouveau"], line_dash="dash", line_color=C_GREEN,
                               annotation_text=f"{art['code']} SS cible")

            fig.update_layout(**light_layout(), height=420,
                xaxis=dict(title="Mois", gridcolor="rgba(0,0,0,.06)"),
                yaxis=dict(title="Stock de sécurité (u)", gridcolor="rgba(0,0,0,.06)"),
                title="Évolution du stock de sécurité par article")
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.markdown("#### Évaluation des économies générées")
        st.markdown(fbox("Formule gain",
                         "Gain = (SS_ancien − SS_nouveau) × Coût_unitaire × Taux_stockage",
                         "Gain annuel = réduction du capital immobilisé × taux d'opportunité"), unsafe_allow_html=True)

        taux_opp = st.number_input("Taux d'opportunité / stockage (%)", value=20.0, step=1.0, key="gain_taux") / 100

        n_gain = st.number_input("Nombre d'articles", value=5, min_value=1, max_value=50, step=1)

        gain_data = []
        cols_g = st.columns([2, 2, 1, 1, 1])
        with cols_g[0]: st.markdown("**Code article**")
        with cols_g[1]: st.markdown("**Description**")
        with cols_g[2]: st.markdown("**SS avant (u)**")
        with cols_g[3]: st.markdown("**SS calculé (u)**")
        with cols_g[4]: st.markdown("**Coût unit. (TND)**")

        for i in range(int(n_gain)):
            cols_g = st.columns([2, 2, 1, 1, 1])
            with cols_g[0]: code_g = st.text_input("", value=f"ART-{i+1:03d}", key=f"g_code_{i}", label_visibility="collapsed")
            with cols_g[1]: desc_g = st.text_input("", value=f"Article {i+1}", key=f"g_desc_{i}", label_visibility="collapsed")
            with cols_g[2]: ss_av  = st.number_input("", value=500, step=10, key=f"g_avant_{i}", label_visibility="collapsed")
            with cols_g[3]: ss_nv  = st.number_input("", value=350, step=10, key=f"g_apres_{i}", label_visibility="collapsed")
            with cols_g[4]: cu_g   = st.number_input("", value=25.0, step=1.0, key=f"g_cu_{i}", label_visibility="collapsed")
            gain_data.append({"Code": code_g, "Description": desc_g, "SS avant": ss_av,
                               "SS calculé": ss_nv, "Coût unit.": cu_g})

        if st.button("💰 Calculer les gains"):
            rows_gain = []
            for r in gain_data:
                delta_ss = r["SS avant"] - r["SS calculé"]
                capital_lib = delta_ss * r["Coût unit."]
                gain_an = capital_lib * taux_opp
                rows_gain.append({
                    "Code":            r["Code"],
                    "Description":     r["Description"],
                    "SS avant (u)":    r["SS avant"],
                    "SS calculé (u)":  r["SS calculé"],
                    "Réduction SS (u)": delta_ss,
                    "Capital libéré (TND)": round(capital_lib, 2),
                    "Gain annuel (TND)":    round(gain_an, 2),
                    "Statut":          "🟢 Économie" if delta_ss > 0 else ("🔴 Surstock" if delta_ss < 0 else "⚪ Inchangé"),
                })

            df_gain = pd.DataFrame(rows_gain)
            st.dataframe(df_gain, use_container_width=True, hide_index=True,
                column_config={
                    "Capital libéré (TND)": st.column_config.NumberColumn(format="%.2f"),
                    "Gain annuel (TND)":    st.column_config.NumberColumn(format="%.2f"),
                })

            total_cap = df_gain["Capital libéré (TND)"].sum()
            total_gain = df_gain["Gain annuel (TND)"].sum()
            n_eco = (df_gain["Statut"] == "🟢 Économie").sum()

            st.divider()
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(mcard("Capital libéré total", f"{total_cap:,.0f}", "TND", color=C_GREEN2), unsafe_allow_html=True)
            with c2: st.markdown(mcard("Gain annuel total",    f"{total_gain:,.0f}", "TND/an", color=C_GOLD), unsafe_allow_html=True)
            with c3: st.markdown(mcard("Articles optimisés",   str(n_eco), f"/ {len(df_gain)}", color=C_GREEN), unsafe_allow_html=True)

            # Graphique gains
            fig_g = go.Figure()
            fig_g.add_trace(go.Bar(
                name="SS avant", x=df_gain["Code"], y=df_gain["SS avant (u)"],
                marker_color=C_RED, opacity=0.7))
            fig_g.add_trace(go.Bar(
                name="SS calculé", x=df_gain["Code"], y=df_gain["SS calculé (u)"],
                marker_color=C_GREEN))
            fig_g.update_layout(**light_layout(), barmode="group", height=320,
                xaxis=dict(title="Article", gridcolor="rgba(0,0,0,.06)"),
                yaxis=dict(title="SS (u)", gridcolor="rgba(0,0,0,.06)"),
                title="Comparaison SS avant / après optimisation")
            st.plotly_chart(fig_g, use_container_width=True)

            # Export gains
            buf_gain = io.BytesIO()
            with pd.ExcelWriter(buf_gain, engine="openpyxl") as writer:
                df_gain.to_excel(writer, index=False, sheet_name="Gains SS")
            buf_gain.seek(0)
            st.download_button(
                label="⬇️ Exporter le rapport des gains",
                data=buf_gain,
                file_name=f"gains_ss_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    URL_LOGO = "https://raw.githubusercontent.com/Ghofrane13/medoil-supply-app/main/logo.png"
    st.markdown(
        f"""
    <div style="display: flex; align-items: center; padding: .3rem 0 .5rem">
        <img src={URL_LOGO} style="width: 44px; margin-right: 10px; border-radius:6px">
        <div>
            <div style="color:#ffffff;font-weight:700;font-size:1.15rem;line-height:1.1">Med oil</div>
            <div style="color:#f5c518;font-size:.65rem;letter-spacing:.1em;text-transform:uppercase">Supply Chain · v3.0</div>
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
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    st.markdown("<p style='color:#4a6a4a;font-size:.72rem;'>© 2025 Medoil — Tous droits réservés</p>",
                unsafe_allow_html=True)

pg.run()
