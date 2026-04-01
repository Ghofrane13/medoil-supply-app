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
html, body, [class*="css"] { font-family: 'Syne', sans-serif; }
.formula-box { background:#d5d7de; border-left:3px solid #47b828; border-radius:8px; padding:.9rem 1.1rem;
    font-family:'DM Mono',monospace; font-size:.82rem; color:#8892a4; margin-bottom:1rem; line-height:1.7; }
.f-eq { color:#60a5fa; font-weight:500; }
.alert-critical { background:rgba(239,68,68,.08); border:1px solid rgba(239,68,68,.3); border-radius:10px;
    padding:1rem 1.2rem; margin-bottom:.7rem; color:#fca5a5; }
.alert-warning  { background:rgba(245,158,11,.08); border:1px solid rgba(245,158,11,.3); border-radius:10px;
    padding:1rem 1.2rem; margin-bottom:.7rem; color:#fcd34d; }
.alert-ok       { background:rgba(34,197,94,.08);  border:1px solid rgba(34,197,94,.3);  border-radius:10px;
    padding:1rem 1.2rem; margin-bottom:.7rem; color:#86efac; }
.alert-info     { background:rgba(20,184,166,.08); border:1px solid rgba(20,184,166,.3); border-radius:10px;
    padding:1rem 1.2rem; margin-bottom:.7rem; color:#5eead4; }
</style>""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    URL_LOGO = "https://raw.githubusercontent.com/Ghofrane13/medoil-supply-app/main/logo.png"
    st.markdown(
    f"""
    <div style="display: flex; align-items: center;">
        <img src={URL_LOGO} style="width: 50px; margin-right: 10px;">
        <h1 style="margin: 0;">Med oil</h1>
    </div>
    """,
    unsafe_allow_html=True
)
    st.divider()
    page = st.navigation("Nav", ["🏠 Accueil","📥 Import & Calcul automatique",
                             "⚙️ Calculateurs","🏭 Processus fournisseurs","🔔 Alertes stock"],
                    label_visibility="collapsed")
    st.divider()
    st.markdown("<p style='color:#5a6478;font-size:.75rem;'>v2.0 · Supply Chain Manager</p>", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(n, d=1):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))): return "—"
    return f"{n:,.{d}f}"

def fmtInt(n):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))): return "—"
    return f"{round(n):,}"

def mcard(label, value, unit, delta=None, color="#3b82f6"):
    dh = f"<div style='font-size:.7rem;color:{color};margin-top:4px'>{delta}</div>" if delta else ""
    return f"""<div style='background:#1e2535;border:1px solid rgba(255,255,255,.1);border-radius:12px;
        padding:1rem 1.2rem;text-align:center;border-top:3px solid {color}'>
        <div style='font-size:.65rem;color:#5a6478;text-transform:uppercase;letter-spacing:.1em;margin-bottom:.4rem'>{label}</div>
        <div style='font-family:"DM Mono",monospace;font-size:1.6rem;font-weight:500;color:#e8eaf0'>{value}</div>
        <div style='font-size:.68rem;color:#5a6478;margin-top:3px'>{unit}</div>{dh}</div>"""

def fbox(title, formula, details):
    return f"<div class='formula-box'><div style='font-size:.65rem;color:#5a6478;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px'>{title}</div><span class='f-eq'>{formula}</span><br/><span style='color:#5a6478'>{details}</span></div>"

def sc(s): return "#22c55e" if s>=85 else "#14b8a6" if s>=70 else "#f59e0b" if s>=55 else "#ef4444"
def sl(s): return "Excellent" if s>=85 else "Bon" if s>=70 else "Moyen" if s>=55 else "Insuffisant"

def dark_layout():
    return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#8892a4", family="Syne"),
                legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor="rgba(255,255,255,.1)", borderwidth=1),
                margin=dict(t=20, b=10, l=0, r=0))

# ── Calculs SC ────────────────────────────────────────────────────────────────
def calc_ss(conso_mois, z=1.65, lt_mois=1.0, variab=0.15):
    d_j = conso_mois / 30
    lt_j = lt_mois * 30
    sigma_d = conso_mois * variab
    return round(z * math.sqrt(lt_j * sigma_d**2 + d_j**2 * (lt_mois * 30 * 0.05)**2))

def calc_eoq(conso_an, cout_unit, sc_cost=150, taux=0.20):
    if cout_unit <= 0 or conso_an <= 0: return 0, 0, 0
    h = taux * cout_unit
    eoq = math.sqrt(2 * conso_an * sc_cost / h)
    nb = conso_an / eoq
    ct = nb * sc_cost + (eoq / 2) * h
    return round(eoq), round(nb, 1), round(ct, 2)

def calc_pr(conso_mois, lt_mois, ss):
    return round((conso_mois / 30) * (lt_mois * 30) + ss)

# ── Export Excel ──────────────────────────────────────────────────────────────
def build_excel(df_result):
    wb = Workbook()
    ws = wb.active
    ws.title = "Calculs Supply Chain"
    thin = Side(style="thin", color="BFBFBF")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

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
        if nf: cell.number_format = nf

    # Row 1 — group headers
    groups = [(1,3,"Produit","D9E1F2"),(4,7,"Fournisseur","FFF2CC"),
              (8,9,"Demand","D9E1F2"),(10,12,"EOQ","E2EFDA"),(13,15,"SS","FCE4D6")]
    for c1,c2,label,bg in groups:
        ws.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        hdr(ws.cell(1,c1), label, bg)

    # Row 2 — column headers
    hdrs = [("Code article","D9E1F2"),("Description Article","D9E1F2"),("Classe","D9E1F2"),
            ("Fournisseur","FFF2CC"),("Délai livraison (mois)","FFF2CC"),
            ("Incertitude fournisseur","FFF2CC"),("Délai de Sécurité (j)","FFF2CC"),
            ("ABC","D9E1F2"),("XYZ","D9E1F2"),
            ("EOQ","E2EFDA"),("MOQ","E2EFDA"),("Diff MOQ et EOQ","E2EFDA"),
            ("Stock de sécurité","FCE4D6"),("Coût stock de sécurité","FCE4D6"),("Point de commande","FCE4D6")]
    for col,(label,bg) in enumerate(hdrs,1):
        hdr(ws.cell(2,col), label, bg)

    # Data rows
    fmts = [None,None,None,None,"0.0","0%","0",None,None,"#,##0","#,##0","#,##0","#,##0","#,##0.00","#,##0"]
    for i, row in df_result.iterrows():
        r = i + 3
        vals = [row.get("Code article",""), row.get("Description Article",""), row.get("Classe",""),
                row.get("Fournisseur",""), row.get("Délai livraison (mois)",""),
                row.get("Incertitude",0.15), round(row.get("Délai livraison (mois)",1)*30) if row.get("Délai livraison (mois)") else "",
                row.get("ABC",""), row.get("XYZ",""),
                row.get("EOQ",0), row.get("MOQ",""), row.get("Diff MOQ EOQ",""),
                row.get("Stock sécurité",0), row.get("Coût SS",0), row.get("Point de commande",0)]
        for col,(val,nf) in enumerate(zip(vals,fmts),1):
            dat(ws.cell(r,col), val, nf)

    # Column widths
    for col,w in enumerate([12,32,8,20,14,14,12,6,6,10,10,14,14,18,16],1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 42
    ws.freeze_panes = "A3"

    # Sheet 2 — source data
    ws2 = wb.create_sheet("Données sources")
    src_hdrs = ["Code article","Description Article","Classe","Stock sécu. (fichier)",
                "Consommation/mois","Coût de revient","Consommation/an","Coût total annuel",
                "SS calculé","EOQ calculé","Point de commande"]
    for col,h in enumerate(src_hdrs,1):
        c = ws2.cell(1,col)
        c.value = h; c.font = Font(bold=True,size=9,name="Arial")
        c.fill = PatternFill("solid",start_color="D9E1F2")
        c.alignment = Alignment(horizontal="center",wrap_text=True); c.border = brd
        ws2.column_dimensions[get_column_letter(col)].width = 18
    for i,row in df_result.iterrows():
        r = i + 2
        vals2 = [row.get("Code article"), row.get("Description Article"), row.get("Classe"),
                 row.get("SS existant",""), row.get("Conso mois",""), row.get("Coût revient",""),
                 row.get("Conso an",""), row.get("Coût total an",""),
                 row.get("Stock sécurité",""), row.get("EOQ",""), row.get("Point de commande","")]
        for col,val in enumerate(vals2,1):
            c = ws2.cell(r,col); c.value = val
            c.font = Font(size=9,name="Arial"); c.alignment = Alignment(horizontal="center"); c.border = brd

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════════════════════════════
# PAGE ACCUEIL
# ════════════════════════════════════════════════════════════════════════════════
if page == "🏠 Accueil":
    st.markdown("""
    <div style='background:linear-gradient(135deg,#0f1117,#161b26);padding:2rem 2.5rem;border-radius:16px;
         border:1px solid rgba(255,255,255,.08);margin-bottom:1.5rem'>
        <div style='font-size:.7rem;color:#60a5fa;letter-spacing:.12em;text-transform:uppercase;margin-bottom:10px'>
            Outil supply chain manager · v2.0</div>
        <h1 style='font-family:"DM Serif Display",serif;font-size:2.2rem;color:#e8eaf0;margin:0 0 .4rem'>
            Pilotez votre <em style='color:#60a5fa'>supply chain</em><br/>avec précision</h1>
        <p style='color:#8892a4;font-size:.95rem;margin:0'>
            Importez votre Excel, calculez automatiquement EOQ, SS et points de commande, exportez le tableau rempli.</p>
    </div>""", unsafe_allow_html=True)

    cols = st.columns(4)
    mods = [("📥","Import & Calcul auto","Importez votre fichier Excel source et générez le tableau rempli à télécharger"),
            ("📐","EOQ / Wilson","Quantité économique — minimise le coût total logistique avec courbe interactive"),
            ("📦","Stock de sécurité","Formule complète avec variabilité de la demande et du délai fournisseur"),
            ("🔔","Alertes stock","Surveillance des niveaux et alertes rupture par article en temps réel")]
    for col,(ic,ti,de) in zip(cols,mods):
        with col:
            st.markdown(f"<div style='background:#161b26;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:1.2rem'>"
                        f"<div style='font-size:1.5rem;margin-bottom:8px'>{ic}</div>"
                        f"<div style='font-size:.9rem;font-weight:600;color:#e8eaf0;margin-bottom:6px'>{ti}</div>"
                        f"<div style='font-size:.78rem;color:#5a6478;line-height:1.5'>{de}</div></div>", unsafe_allow_html=True)
    st.divider()
    st.info("👈 Commencez par **📥 Import & Calcul automatique** pour charger votre fichier Excel.")


# ════════════════════════════════════════════════════════════════════════════════
# PAGE IMPORT & CALCUL AUTOMATIQUE
# ════════════════════════════════════════════════════════════════════════════════
elif page == "📥 Import & Calcul automatique":
    st.markdown("## 📥 Import & Calcul automatique")
    st.markdown("<p style='color:#8892a4'>Importez votre fichier — EOQ, SS et point de commande calculés automatiquement pour chaque article</p>", unsafe_allow_html=True)

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
            sel = names[0] if len(names)==1 else st.selectbox("Choisir la feuille", names)
            raw = all_sheets[sel]

            # Détecter la ligne d'en-tête
            hrow = 0
            for i, row in raw.iterrows():
                row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                if any(k in row_str for k in ["code","description","article","consommation"]):
                    hrow = i; break

            df_raw = pd.read_excel(uploaded, sheet_name=sel, header=hrow)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            df_raw = df_raw.dropna(how="all")

            # Mapping flexible
            cmap = {}
            for col in df_raw.columns:
                cl = col.lower().strip()
                if "code" in cl and "article" in cl:                      cmap["Code article"] = col
                elif "description" in cl or "désignation" in cl:          cmap["Description Article"] = col
                elif cl == "classe":                                       cmap["Classe"] = col
                elif "consommation" in cl and "mois" in cl:               cmap["Conso mois"] = col
                elif "coût de revient" in cl or "cout de revient" in cl:  cmap["Coût revient"] = col
                elif "consommation" in cl and "an" in cl:                 cmap["Conso an"] = col
                elif ("stock" in cl and ("sécu" in cl or "sécurité" in cl)):cmap["SS existant"] = col
                elif "cout total" in cl or "coût total" in cl:            cmap["Coût total an"] = col
                elif "fournisseur" in cl:                                  cmap["Fournisseur"] = col
                elif "délai" in cl and "livraison" in cl:                  cmap["Délai livraison (mois)"] = col
                elif "taille" in cl and "lot" in cl:                       cmap["MOQ"] = col
                elif cl == "moq":                                           cmap["MOQ"] = col

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
            "Code article":    [40007,40010,40011,40014,40015,40036,40042,40055,40062,40063],
            "Description Article":["NYSOSEL (NICKEL CATALYSEUR)","TERRE TONSIL","PAPIER FILTRE",
                                   "TOILE FILTRANTE TERRE","PERLITE FILTRATION PAF1","SEL MARIN FIN",
                                   "MYVEROL 18-04 MYVEROL","TRISYL","VITAMINE A","VITAMINE E"],
            "Classe":          ["MP"]*10,
            "Conso mois":      [324,5886,509,32460,959,3747,3147,4432,16,22],
            "Coût revient":    [59.23,1.709,3.823,49.787,1.700,0,0,8.189,0,0],
            "Conso an":        [3888,70632,6108,389520,11508,44964,37764,53184,192,264],
            "SS existant":     [2673,6353,2859,68,1171,1560,20000,5000,74,100],
            "Coût total an":   [230286,120710,23351,19393032,19564,0,0,435524,0,0],
        })
        st.info("📊 Données de démo chargées — basées sur votre fichier Excel d'origine")

    if df_source is not None and len(df_source) > 0:
        with st.expander("👁 Aperçu des données sources", expanded=False):
            st.dataframe(df_source, use_container_width=True, hide_index=True)

        # ── Calculs ──
        results = []
        for _, row in df_source.iterrows():
            cm  = float(row.get("Conso mois",0) or 0)
            ca  = float(row.get("Conso an",0) or cm*12) or cm*12
            cu  = float(row.get("Coût revient",0) or 0)
            lt  = float(row.get("Délai livraison (mois)",p_lt) or p_lt)
            moq = float(row.get("MOQ",0) or 0)

            ss  = calc_ss(cm, z=Z, lt_mois=lt, variab=p_var/100)
            eoq, nb_cmd, _ = calc_eoq(ca, cu, sc_cost=p_sc, taux=p_taux/100)
            pr  = calc_pr(cm, lt, ss)
            css = round(ss * cu, 2) if cu > 0 else 0
            diff = round(eoq - moq) if moq > 0 else ""

            results.append({
                "Code article":         row.get("Code article",""),
                "Description Article":  row.get("Description Article",""),
                "Classe":               row.get("Classe",""),
                "Fournisseur":          row.get("Fournisseur",""),
                "Délai livraison (mois)": lt,
                "Incertitude":          p_var/100,
                "EOQ":                  eoq,
                "MOQ":                  int(moq) if moq>0 else "",
                "Diff MOQ EOQ":         diff,
                "Stock sécurité":       ss,
                "SS existant":          float(row.get("SS existant",0) or 0),
                "Coût SS":              css,
                "Point de commande":    pr,
                "Conso mois":           cm,
                "Coût revient":         cu,
                "Conso an":             ca,
                "Coût total an":        float(row.get("Coût total an",0) or 0),
            })

        df_res = pd.DataFrame(results)

        st.divider()
        st.markdown("#### 🧮 Tableau des calculs")

        disp = df_res[["Code article","Description Article","Classe",
                        "EOQ","MOQ","Diff MOQ EOQ","Stock sécurité","Coût SS","Point de commande"]].copy()
        st.dataframe(disp, use_container_width=True, hide_index=True,
            column_config={
                "EOQ":              st.column_config.NumberColumn("EOQ (u)", format="%d"),
                "MOQ":              st.column_config.NumberColumn("MOQ (u)"),
                "Diff MOQ EOQ":     st.column_config.NumberColumn("Diff MOQ/EOQ"),
                "Stock sécurité":   st.column_config.NumberColumn("Stock sécu. (u)", format="%d"),
                "Coût SS":          st.column_config.NumberColumn("Coût SS (TND)", format="%.2f"),
                "Point de commande":st.column_config.NumberColumn("Point commande (u)", format="%d"),
            })

        # ── KPIs ──
        st.markdown("---")
        st.markdown("#### 📊 Synthèse")
        c1,c2,c3,c4 = st.columns(4)
        with c1: st.markdown(mcard("Articles traités", str(len(df_res)), "articles"), unsafe_allow_html=True)
        with c2: st.markdown(mcard("Coût total SS", f"{df_res['Coût SS'].sum():,.0f}", "TND immobilisés", color="#f59e0b"), unsafe_allow_html=True)
        eoq_mean = df_res[df_res["EOQ"]>0]["EOQ"].mean()
        with c3: st.markdown(mcard("EOQ moyen", fmtInt(eoq_mean), "u / commande", color="#22c55e"), unsafe_allow_html=True)
        ecarts_hauts = (df_res["Stock sécurité"] > df_res["SS existant"] * 1.3).sum()
        with c4: st.markdown(mcard("SS à réviser", str(ecarts_hauts), "SS calculé > 130% SS existant", color="#ef4444"), unsafe_allow_html=True)

        # ── Graphique EOQ vs SS ──
        st.markdown("---")
        st.markdown("#### 📈 EOQ · Stock de sécurité · Point de commande")
        labels = df_res["Code article"].astype(str).tolist()
        fig = go.Figure()
        fig.add_trace(go.Bar(name="EOQ", x=labels, y=df_res["EOQ"], marker_color="#3b82f6",
                             text=df_res["EOQ"].apply(lambda v: f"{int(v):,}" if v>0 else ""), textposition="outside"))
        fig.add_trace(go.Bar(name="Stock sécurité", x=labels, y=df_res["Stock sécurité"], marker_color="#f59e0b",
                             text=df_res["Stock sécurité"].apply(lambda v: f"{int(v):,}"), textposition="outside"))
        fig.add_trace(go.Scatter(name="Point de commande", x=labels, y=df_res["Point de commande"],
                                 mode="markers+lines", marker=dict(color="#ef4444",size=9,symbol="diamond"),
                                 line=dict(color="#ef4444",width=2,dash="dash")))
        if "SS existant" in df_res.columns and df_res["SS existant"].sum() > 0:
            fig.add_trace(go.Scatter(name="SS existant (fichier)", x=labels, y=df_res["SS existant"],
                                     mode="markers+lines", marker=dict(color="#a78bfa",size=7,symbol="circle"),
                                     line=dict(color="#a78bfa",width=1.5,dash="dot")))
        fig.update_layout(**dark_layout(), barmode="group", height=400,
                          xaxis=dict(title="Article", gridcolor="rgba(255,255,255,.05)"),
                          yaxis=dict(title="Unités", gridcolor="rgba(255,255,255,.05)"))
        st.plotly_chart(fig, use_container_width=True)

        # ── Comparaison SS calculé vs existant ──
        if df_res["SS existant"].sum() > 0:
            st.markdown("---")
            st.markdown("#### 🔍 SS calculé vs SS dans votre fichier")
            comp = df_res[["Code article","Description Article","SS existant","Stock sécurité"]].copy()
            comp["Écart (u)"] = comp["Stock sécurité"] - comp["SS existant"]
            comp["Écart %"] = np.where(comp["SS existant"]>0,
                                       (comp["Écart (u)"] / comp["SS existant"] * 100).round(1), None)
            st.dataframe(comp, use_container_width=True, hide_index=True,
                column_config={
                    "SS existant":     st.column_config.NumberColumn("SS fichier", format="%d"),
                    "Stock sécurité":  st.column_config.NumberColumn("SS calculé", format="%d"),
                    "Écart (u)":       st.column_config.NumberColumn("Écart (u)", format="%d"),
                    "Écart %":         st.column_config.NumberColumn("Écart %", format="%.1f%%"),
                })

        # ── Export ──
        st.markdown("---")
        st.markdown("#### 💾 Export — tableau rempli")
        st.info("Le fichier exporté reprend exactement le format de votre tableau d'origine avec toutes les colonnes calculées.")
        excel_buf = build_excel(df_res)
        st.download_button(
            label="⬇️ Télécharger le tableau Excel rempli",
            data=excel_buf,
            file_name="supply_chain_calculs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


# ════════════════════════════════════════════════════════════════════════════════
# PAGE CALCULATEURS
# ════════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Calculateurs":
    st.markdown("## ⚙️ Calculateurs Supply Chain")
    tab_ss,tab_eoq,tab_kpi,tab_rp = st.tabs(["📦 Stock de sécurité","📐 EOQ (Wilson)","📊 KPIs stock","🔄 Point de réappro."])

    with tab_ss:
        st.markdown(fbox("Formule","SS = Z × √(LT×σD² + D²×σLT²)","Z = facteur service · σD = écart-type demande · LT = délai (j)"), unsafe_allow_html=True)
        c1,c2,c3 = st.columns(3)
        with c1: sd=st.number_input("Demande moy. (u/j)",value=50.0,step=1.0); ss2=st.number_input("Écart-type demande σD",value=8.0,step=.5)
        with c2: slt=st.number_input("Délai fourni. (j)",value=7.0,step=1.0); ssl=st.number_input("Écart-type délai σLT",value=1.5,step=.1)
        with c3: szo=st.selectbox("Niveau de service",["90% (Z=1.28)","95% (Z=1.65)","97.5% (Z=1.96)","99% (Z=2.33)","99.5% (Z=2.58)"],index=1); scu=st.number_input("Coût unitaire (TND)",value=25.0,step=1.0)
        Z2={"90% (Z=1.28)":1.28,"95% (Z=1.65)":1.65,"97.5% (Z=1.96)":1.96,"99% (Z=2.33)":2.33,"99.5% (Z=2.58)":2.58}[szo]
        SS1=Z2*ss2*math.sqrt(slt); SS2=Z2*math.sqrt(slt*ss2**2+sd**2*ssl**2)
        st.markdown("---")
        c=st.columns(3)
        for i,(l,v,u,col) in enumerate([("SS (σD seul)",fmtInt(SS1),"unités","#3b82f6"),("SS (formule complète)",fmtInt(SS2),"unités","#22c55e"),("Coût immobilisé SS",fmtInt(SS2*scu),"TND","#f59e0b")]):
            with c[i]: st.markdown(mcard(l,v,u,color=col),unsafe_allow_html=True)
        ns_vals=[85,90,95,97.5,99,99.5]; z_vals=[1.04,1.28,1.65,1.96,2.33,2.58]
        ss_v=[round(z*math.sqrt(slt*ss2**2+sd**2*ssl**2)) for z in z_vals]
        fig=go.Figure(go.Bar(x=[f"{n}%" for n in ns_vals],y=ss_v,
            marker_color=["#22c55e" if abs(n-float(szo.split("%")[0]))<1 else "#3b82f6" for n in ns_vals],
            text=ss_v,textposition="outside"))
        fig.update_layout(**dark_layout(),height=260,
            xaxis=dict(title="Niveau de service",gridcolor="rgba(255,255,255,.05)"),
            yaxis=dict(title="SS (unités)",gridcolor="rgba(255,255,255,.05)"))
        st.plotly_chart(fig,use_container_width=True)

    with tab_eoq:
        st.markdown(fbox("Modèle de Wilson","EOQ = √(2×D×Sc / (Sh×Cu))","D=demande annuelle · Sc=coût passation · Sh=taux stockage · Cu=coût unitaire"), unsafe_allow_html=True)
        c1,c2,c3=st.columns(3)
        with c1: eD=st.number_input("Demande annuelle (u)",value=12000,step=100); eSc=st.number_input("Coût passation (TND)",value=150.0,step=5.0)
        with c2: eCu=st.number_input("Coût unitaire (TND)",value=80.0,step=1.0); eSh=st.number_input("Taux stockage (%)",value=20.0,step=1.0)
        with c3: eLT=st.number_input("Délai fourni. (j)",value=10,step=1); eDy=st.number_input("Jours ouvr./an",value=250,step=1)
        h=(eSh/100)*eCu; EOQ=math.sqrt(2*eD*eSc/h) if h>0 else 0
        nb=eD/EOQ if EOQ>0 else 0; T=eDy/nb if nb>0 else 0; Cp=nb*eSc; Cs=(EOQ/2)*h
        st.markdown("---")
        c=st.columns(3); data=[("EOQ",fmtInt(EOQ),"u/cmd","#3b82f6"),("Nb cmd/an",fmt(nb,1),"cmds","#8892a4"),
            ("Périodicité",fmtInt(T),"jours","#8892a4"),("Coût passation/an",fmtInt(Cp),"TND","#f59e0b"),
            ("Coût stockage/an",fmtInt(Cs),"TND","#f59e0b"),("Coût total min.",fmtInt(Cp+Cs),"TND","#22c55e")]
        for i,(l,v,u,col) in enumerate(data):
            with c[i%3]: st.markdown(mcard(l,v,u,color=col),unsafe_allow_html=True); st.markdown("<div style='margin-bottom:10px'></div>",unsafe_allow_html=True)
        qty_r=np.linspace(max(10,EOQ*.2),EOQ*2.5,200)
        fig=go.Figure()
        fig.add_trace(go.Scatter(x=qty_r,y=[(eD/q)*eSc for q in qty_r],name="Passation",line=dict(color="#60a5fa",width=2)))
        fig.add_trace(go.Scatter(x=qty_r,y=[(q/2)*h for q in qty_r],name="Stockage",line=dict(color="#f59e0b",width=2)))
        fig.add_trace(go.Scatter(x=qty_r,y=[(eD/q)*eSc+(q/2)*h for q in qty_r],name="Total",line=dict(color="#22c55e",width=2.5)))
        fig.add_vline(x=EOQ,line_dash="dash",line_color="#ef4444",annotation_text=f"EOQ={fmtInt(EOQ)}",annotation_font_color="#ef4444")
        fig.update_layout(**dark_layout(),height=300,
            xaxis=dict(title="Quantité (u)",gridcolor="rgba(255,255,255,.05)"),
            yaxis=dict(title="Coût/an (TND)",gridcolor="rgba(255,255,255,.05)"))
        st.plotly_chart(fig,use_container_width=True)

    with tab_kpi:
        c1,c2=st.columns(2)
        with c1: ks=st.number_input("Stock moyen (u)",value=3000,step=100); kv=st.number_input("Ventes annuelles (u)",value=18000,step=100); kval=st.number_input("Valeur stock (TND)",value=240000,step=1000); kca=st.number_input("CA annuel (TND)",value=2400000,step=10000)
        with c2: kot=st.number_input("Livrées à temps",value=465,step=1); kto=st.number_input("Total livrées",value=490,step=1); klc=st.number_input("Coût logistique (TND)",value=168000,step=1000); ksh=st.number_input("Taux stockage (%)",value=22.0,step=1.0)
        rot=kv/ks if ks>0 else 0; cov=(ks/kv)*365 if kv>0 else 0; srv=(kot/kto)*100 if kto>0 else 0; crt=(klc/kca)*100 if kca>0 else 0
        st.markdown("---")
        c=st.columns(3)
        kdata=[("Rotation",fmt(rot,1)+"×","/an · >6×=excellent","#22c55e" if rot>=6 else "#f59e0b"),
               ("Couverture",fmtInt(cov)+"j","15–30j=optimal","#8892a4"),
               ("Taux service OTD",fmt(srv,1)+"%",">95%=cible","#22c55e" if srv>=95 else "#f59e0b"),
               ("Coût log/CA",fmt(crt,1)+"%","objectif <8%","#22c55e" if crt<8 else "#ef4444"),
               ("Coût stockage/an",fmtInt(kval*(ksh/100)),"TND","#f59e0b"),
               ("DSI",fmt(ks/(kv/365),0)+"j","jours de stock","#8892a4")]
        for i,(l,v,u,col) in enumerate(kdata):
            with c[i%3]: st.markdown(mcard(l,v,u,color=col),unsafe_allow_html=True); st.markdown("<div style='margin-bottom:10px'></div>",unsafe_allow_html=True)

    with tab_rp:
        st.markdown(fbox("Formule","PR = (Dmoy × LT) + SS","Dmoy = demande/jour · LT = délai (j) · SS = stock de sécurité"), unsafe_allow_html=True)
        c1,c2,c3=st.columns(3)
        with c1: rd=st.number_input("Demande moy. (u/j)",value=50.0,step=1.0); rlt=st.number_input("Délai fourni. (j)",value=7.0,step=1.0)
        with c2: rss=st.number_input("Stock sécu. (u)",value=132.0,step=10.0); rdm=st.number_input("Demande max (u/j)",value=70.0,step=1.0)
        with c3: rltm=st.number_input("Délai max (j)",value=10.0,step=1.0); rcu=st.number_input("Stock actuel (u)",value=500.0,step=10.0)
        PR=rd*rlt+rss; dL=(rcu-rss)/rd if rd>0 else 0
        if rcu<=rss: st.markdown("<div class='alert-critical'>🔴 RUPTURE IMMINENTE — Commander immédiatement !</div>",unsafe_allow_html=True)
        elif rcu<=PR: st.markdown("<div class='alert-warning'>🟡 COMMANDER MAINTENANT — Point de réappro. atteint</div>",unsafe_allow_html=True)
        else: st.markdown(f"<div class='alert-ok'>🟢 Stock suffisant — {fmtInt(rcu)} u > PR ({fmtInt(PR)} u)</div>",unsafe_allow_html=True)
        c=st.columns(3)
        for i,(l,v,u,col) in enumerate([("Point de réappro.",fmtInt(PR),"unités","#3b82f6"),("Jours restants",fmt(dL,1),"jours","#8892a4"),("Délai critique",fmt(dL-rlt,1),"jours avant urgence","#f59e0b")]):
            with c[i]: st.markdown(mcard(l,v,u,color=col),unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════════
# PAGE PROCESSUS FOURNISSEURS
# ════════════════════════════════════════════════════════════════════════════════
elif page == "🏭 Processus fournisseurs":
    st.markdown("## 🏭 Évaluation processus fournisseurs")
    c1,c2,c3=st.columns(3)
    with c1: pn=st.text_input("Fournisseur",value="Würth Tunisie SA")
    with c2: pr_ref=st.text_input("Référence",value="RB-7204")
    with c3: pc=st.selectbox("Catégorie",["Composants mécaniques","Matières premières","Fournitures industrielles","Sous-traitance"])
    c1,c2,c3,c4=st.columns(4)
    with c1: pt=st.number_input("Commandes passées",value=24,min_value=1); po=st.number_input("Livrées à temps",value=20,min_value=0)
    with c2: pk=st.number_input("Livrées complètes",value=22,min_value=0); pnc=st.number_input("Non-conformités",value=2,min_value=0)
    with c3: pla=st.number_input("Délai annoncé (j)",value=7.0,step=.5); plr=st.number_input("Délai réel (j)",value=8.4,step=.1)
    with c4: ppr=st.number_input("Traitement (h)",value=18.0,step=1.0); pac=st.number_input("Accusé récep. (h)",value=4.0,step=.5)
    otd=(po/pt)*100; fr=(pk/pt)*100; qu=((pt-pnc)/pt)*100
    ltp=max(0,100-((plr-pla)/pla)*100) if pla>0 else 0
    prs=100 if ppr<=8 else 85 if ppr<=24 else 60 if ppr<=48 else 30
    acs=100 if pac<=2 else 85 if pac<=8 else 60 if pac<=24 else 30
    gs=otd*.30+fr*.20+qu*.20+ltp*.15+prs*.10+acs*.05
    col_c=sc(gs); lbl=sl(gs)
    adv={"Excellent":"Fournisseur stratégique. Envisager un accord-cadre.","Bon":"Performance satisfaisante. Optimiser le délai réel.",
         "Moyen":"Plan d'amélioration sous 30 jours recommandé.","Insuffisant":"Audit qualité et diversification fournisseur urgent."}[lbl]
    st.markdown("---")
    fg=go.Figure(go.Indicator(mode="gauge+number",value=round(gs,1),
        number={"suffix":"/100","font":{"color":col_c,"family":"DM Mono","size":28}},
        gauge={"axis":{"range":[0,100]},"bar":{"color":col_c,"thickness":.25},"bgcolor":"#1e2535",
               "steps":[{"range":[0,55],"color":"rgba(239,68,68,.2)"},{"range":[55,70],"color":"rgba(245,158,11,.2)"},
                        {"range":[70,85],"color":"rgba(20,184,166,.2)"},{"range":[85,100],"color":"rgba(34,197,94,.2)"}]}))
    fg.update_layout(paper_bgcolor="rgba(0,0,0,0)",font=dict(color="#8892a4",family="Syne"),height=220,margin=dict(t=10,b=10,l=20,r=20))
    cg,cd=st.columns([1,2])
    with cg: st.plotly_chart(fg,use_container_width=True)
    with cd: st.markdown(f"<div style='background:#1e2535;border-left:4px solid {col_c};border-radius:10px;padding:1.2rem 1.4rem;margin-top:1rem'>"
                         f"<div style='font-size:1.2rem;font-weight:600;color:{col_c};margin-bottom:8px'>{lbl} — {pn}</div>"
                         f"<div style='font-size:.85rem;color:#8892a4;line-height:1.6'>{adv}</div></div>",unsafe_allow_html=True)
    st.markdown("---")
    crit=[("Livraison à temps (OTD)",otd,f"{po}/{pt} · Poids 30%"),("Fill rate",fr,f"{pk}/{pt} · Poids 20%"),
          ("Qualité / conformité",qu,f"{pt-pnc}/{pt} · Poids 20%"),("Fiabilité délai",ltp,f"Annoncé:{pla}j Réel:{plr}j · Poids 15%"),
          ("Traitement commande",prs,f"{ppr}h · Poids 10%"),("Réactivité ACK",acs,f"{pac}h · Poids 5%")]
    c1,c2=st.columns(2)
    for i,(nm,scr,nt) in enumerate(crit):
        cc=sc(scr)
        with (c1 if i%2==0 else c2):
            st.markdown(f"<div style='background:#1e2535;border:1px solid rgba(255,255,255,.08);border-radius:10px;padding:1rem;margin-bottom:10px'>"
                        f"<div style='display:flex;justify-content:space-between;margin-bottom:8px'>"
                        f"<span style='font-size:.85rem;font-weight:600;color:#e8eaf0'>{nm}</span>"
                        f"<span style='font-family:\"DM Mono\",monospace;color:{cc}'>{scr:.1f}%</span></div>"
                        f"<div style='background:#0f1117;border-radius:4px;height:6px;overflow:hidden'>"
                        f"<div style='background:{cc};height:100%;width:{scr}%'></div></div>"
                        f"<div style='font-size:.7rem;color:#5a6478;margin-top:6px'>{nt}</div></div>",unsafe_allow_html=True)
    if otd<85: st.markdown(f"<div class='alert-warning'>⏱ OTD insuffisant ({otd:.1f}%) — Négocier des pénalités. Majorer le stock de sécurité.</div>",unsafe_allow_html=True)
    if qu<95:  st.markdown(f"<div class='alert-warning'>⚠️ {pnc} non-conformités — Audit qualité recommandé.</div>",unsafe_allow_html=True)
    if plr>pla: st.markdown(f"<div class='alert-warning'>📅 Dérive délai +{plr-pla:.1f}j — Recalculer le point de réappro. avec le délai réel.</div>",unsafe_allow_html=True)
    if gs>=85:  st.markdown("<div class='alert-ok'>✅ Performance excellente — Envisager un accord-cadre annuel.</div>",unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════════
# PAGE ALERTES STOCK
# ════════════════════════════════════════════════════════════════════════════════
elif page == "🔔 Alertes stock":
    st.markdown("## 🔔 Alertes & Surveillance des stocks")
    demo_data = {
        "Référence":      ["RB-7204","IS-316L","VD-0822","CA-1140","JT-2200","MO-1115"],
        "Désignation":    ["Roulements SKF","Acier inox 316L","Vannes pneumatiques","Câble 4×2.5mm²","Joints toriques","Moteur 0.75kW"],
        "Stock actuel":   [45,820,380,2400,1250,12],
        "Stock sécurité": [132,600,80,300,400,8],
        "Point réappro.": [482,1850,230,900,1200,35],
        "Conso. moy./j":  [50,120,15,80,95,3],
        "Lead time (j)":  [7,10,12,8,5,21],
    }
    if st.button("📂 Charger démo"): st.session_state.alerts_df = pd.DataFrame(demo_data)
    if "alerts_df" not in st.session_state: st.session_state.alerts_df = pd.DataFrame(demo_data)
    edited=st.data_editor(st.session_state.alerts_df,use_container_width=True,num_rows="dynamic",
        column_config={"Stock actuel":st.column_config.NumberColumn(format="%d"),"Stock sécurité":st.column_config.NumberColumn(format="%d"),
                       "Point réappro.":st.column_config.NumberColumn(format="%d"),"Conso. moy./j":st.column_config.NumberColumn(format="%d"),
                       "Lead time (j)":st.column_config.NumberColumn(format="%d")},hide_index=True)
    if st.button("🔍 Analyser les stocks",type="primary"):
        df=edited.copy()
        df["Jours restants"]=np.where(df["Conso. moy./j"]>0,((df["Stock actuel"]-df["Stock sécurité"])/df["Conso. moy./j"]).round(1),np.nan)
        def status(r):
            if r["Stock actuel"]<=r["Stock sécurité"]: return "🔴 Rupture"
            if r["Stock actuel"]<=r["Point réappro."]: return "🟡 Commander"
            return "🟢 Normal"
        df["Statut"]=df.apply(status,axis=1)
        cr=df[df["Statut"]=="🔴 Rupture"]; at=df[df["Statut"]=="🟡 Commander"]; ok=df[df["Statut"]=="🟢 Normal"]
        c1,c2,c3=st.columns(3)
        with c1: st.markdown(mcard("🔴 Rupture",str(len(cr)),"articles sous SS",color="#ef4444"),unsafe_allow_html=True)
        with c2: st.markdown(mcard("🟡 Commander",str(len(at)),"articles entre SS et PR",color="#f59e0b"),unsafe_allow_html=True)
        with c3: st.markdown(mcard("🟢 Normal",str(len(ok)),"articles OK",color="#22c55e"),unsafe_allow_html=True)
        fig=go.Figure()
        fig.add_trace(go.Bar(name="Stock actuel",x=df["Référence"],y=df["Stock actuel"],
            marker_color=["#ef4444" if s=="🔴 Rupture" else "#f59e0b" if s=="🟡 Commander" else "#22c55e" for s in df["Statut"]]))
        fig.add_trace(go.Scatter(name="Stock sécu.",x=df["Référence"],y=df["Stock sécurité"],mode="markers+lines",
            marker=dict(color="#a78bfa",size=8,symbol="diamond"),line=dict(color="#a78bfa",width=1.5,dash="dash")))
        fig.add_trace(go.Scatter(name="Point réappro.",x=df["Référence"],y=df["Point réappro."],mode="markers+lines",
            marker=dict(color="#f59e0b",size=8,symbol="triangle-up"),line=dict(color="#f59e0b",width=1.5,dash="dot")))
        fig.update_layout(**dark_layout(),height=320,
            xaxis=dict(gridcolor="rgba(255,255,255,.05)"),yaxis=dict(title="Unités",gridcolor="rgba(255,255,255,.05)"))
        st.plotly_chart(fig,use_container_width=True)
        st.dataframe(df[["Référence","Désignation","Stock actuel","Stock sécurité","Point réappro.","Jours restants","Statut"]],
            use_container_width=True,hide_index=True,
            column_config={"Jours restants":st.column_config.NumberColumn(format="%.1f j")})
        for _,r in cr.iterrows():
            st.markdown(f"<div class='alert-critical'>🔴 <strong>{r['Référence']} — {r['Désignation']}</strong> : {int(r['Stock actuel'])} u sous SS ({int(r['Stock sécurité'])} u). Commander immédiatement. LT : {int(r['Lead time (j)'])} j.</div>",unsafe_allow_html=True)
        for _,r in at.iterrows():
            st.markdown(f"<div class='alert-warning'>🟡 <strong>{r['Référence']} — {r['Désignation']}</strong> : Point de réappro. atteint. Lancer une commande. LT : {int(r['Lead time (j)'])} j.</div>",unsafe_allow_html=True)
