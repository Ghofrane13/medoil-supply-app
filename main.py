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
    return (f"<div class='formula-box'>"
            f"<div style='font-size:.65rem;color:#5a6478;text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px'>{title}</div>"
            f"<span class='f-eq'>{formula}</span><br/>"
            f"<span style='color:#5a6478'>{details}</span></div>")
 
def sc(s): return "#22c55e" if s>=85 else "#14b8a6" if s>=70 else "#f59e0b" if s>=55 else "#ef4444"
def sl(s): return "Excellent" if s>=85 else "Bon" if s>=70 else "Moyen" if s>=55 else "Insuffisant"
 
def dark_layout():
    return dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#8892a4", family="Syne"),
                legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor="rgba(255,255,255,.1)", borderwidth=1),
                margin=dict(t=20, b=10, l=0, r=0))
 
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
 
    groups = [(1,3,"Produit","D9E1F2"),(4,7,"Fournisseur","FFF2CC"),
              (8,9,"Demand","D9E1F2"),(10,12,"EOQ","E2EFDA"),(13,15,"SS","FCE4D6")]
    for c1,c2,label,bg in groups:
        ws.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
        hdr(ws.cell(1, c1), label, bg)
 
    hdrs = [("Code article","D9E1F2"),("Description Article","D9E1F2"),("Classe","D9E1F2"),
            ("Fournisseur","FFF2CC"),("Délai livraison (mois)","FFF2CC"),
            ("Incertitude fournisseur","FFF2CC"),("Délai de Sécurité (j)","FFF2CC"),
            ("ABC","D9E1F2"),("XYZ","D9E1F2"),
            ("EOQ","E2EFDA"),("MOQ","E2EFDA"),("Diff MOQ et EOQ","E2EFDA"),
            ("Stock de sécurité","FCE4D6"),("Coût stock de sécurité","FCE4D6"),("Point de commande","FCE4D6")]
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
        c.fill  = PatternFill("solid", start_color="D9E1F2")
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
# FONCTIONS DE PAGE  (chacune = une page st.navigation)
# ════════════════════════════════════════════════════════════════════════════════
 
def accueil():
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
    mods = [("📥","Import & Calcul auto","Importez votre fichier Excel et générez le tableau rempli à télécharger"),
            ("📐","EOQ / Wilson","Quantité économique — minimise le coût total logistique avec courbe interactive"),
            ("📦","Stock de sécurité","Formule complète avec variabilité de la demande et du délai fournisseur"),
            ("🔔","Alertes stock","Surveillance des niveaux et alertes rupture par article en temps réel")]
    for col,(ic,ti,de) in zip(cols, mods):
        with col:
            st.markdown(f"<div style='background:#161b26;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:1.2rem'>"
                        f"<div style='font-size:1.5rem;margin-bottom:8px'>{ic}</div>"
                        f"<div style='font-size:.9rem;font-weight:600;color:#e8eaf0;margin-bottom:6px'>{ti}</div>"
                        f"<div style='font-size:.78rem;color:#5a6478;line-height:1.5'>{de}</div></div>",
                        unsafe_allow_html=True)
    st.divider()
    st.info(" Commencez par **Import & Calcul automatique** dans la barre latérale pour charger votre fichier Excel.")
 
 
# ─────────────────────────────────────────────────────────────────────────────
def import_calcul():
    st.markdown(f"""
    <div class='medoil-card'>
        <h2>📥 Import & Calcul Automatique</h2>
        <p>Importez vos fichiers - Les calculs sont entièrement automatisés</p>
    </div>
    """, unsafe_allow_html=True)
    
    # ============================================================
    # CONFIGURATION DES DÉLAIS PAR SOURCE
    # ============================================================
    SOURCE_DELAYS = {
        "Export": {"mois": 4, "jours": 120, "label": "Export (120 jours)"},
        "Local": {"mois": 0.5, "jours": 15, "label": "Local (15 jours)"},
        "BM": {"mois": 0.066, "jours": 2, "label": "BM (2 jours)"}
    }
    
    # ============================================================
    # 1. TABLEAU DES CONSOMMATIONS (M1 à M12)
    # ============================================================
    st.markdown("### 📊 1. Tableau des consommations mensuelles")
    st.caption("Importez un fichier avec les consommations des 12 derniers mois pour calculer automatiquement la moyenne et l'écart-type")
    
    with st.expander("📋 Format attendu - Consommations mensuelles", expanded=False):
        st.markdown("""
| Colonne | Obligatoire | Description |
|---|---|---|
| `Code article` | ✅ | Identifiant unique du produit |
| `Article` | ✅ | Désignation du produit |
| `M1` | ✅ | Consommation mois 1 |
| `M2` | ✅ | Consommation mois 2 |
| `M3` | ✅ | Consommation mois 3 |
| `M4` | ✅ | Consommation mois 4 |
| `M5` | ✅ | Consommation mois 5 |
| `M6` | ✅ | Consommation mois 6 |
| `M7` | ✅ | Consommation mois 7 |
| `M8` | ✅ | Consommation mois 8 |
| `M9` | ✅ | Consommation mois 9 |
| `M10` | ✅ | Consommation mois 10 |
| `M11` | ✅ | Consommation mois 11 |
| `M12` | ✅ | Consommation mois 12 |
        """)
    
    uploaded_conso = st.file_uploader(
        "📂 Fichier des consommations mensuelles (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="conso_upload"
    )
    
    # ============================================================
    # 2. TABLEAU DES PRODUITS AVEC PARAMÈTRES
    # ============================================================
    st.markdown("### 📦 2. Tableau des produits et paramètres")
    st.caption("Importez le fichier avec les informations de base et les paramètres de calcul")
    
    with st.expander("📋 Format attendu - Produits", expanded=False):
        st.markdown("""
| Colonne | Obligatoire | Description |
|---|---|---|
| `Code article` | ✅ | Identifiant unique (doit correspondre au tableau consommations) |
| `Article` | ✅ | Désignation du produit |
| `Source` | ✅ | Export / Local / BM (définit le délai automatiquement) |
| `Unité` | — | Unité de mesure (kg, L, u, etc.) |
| `Coût unitaire` | ✅ | Coût d'achat unitaire (TND) |
| `Coût passation` | ✅ | Coût par commande (TND) |
| `Taux stockage (%)` | ✅ | Taux de détention du stock (%) |
        """)
    
    uploaded_produits = st.file_uploader(
        "📂 Fichier des produits et paramètres (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="produits_upload"
    )
    
    # ============================================================
    # BOUTON DE CALCUL
    # ============================================================
    st.markdown("---")
    
    calcul_disabled = (uploaded_conso is None) or (uploaded_produits is None)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 LANCER LE CALCUL AUTOMATIQUE", type="primary", disabled=calcul_disabled, use_container_width=True):
            process_calcul_auto(uploaded_conso, uploaded_produits, SOURCE_DELAYS)
    
    if calcul_disabled:
        st.info("📂 Veuillez importer les deux fichiers Excel pour lancer le calcul automatique")


def process_calcul_auto(uploaded_conso, uploaded_produits, SOURCE_DELAYS):
    """Fonction principale de traitement des calculs"""
    
    try:
        # ============================================================
        # CHARGEMENT ET TRAITEMENT DES CONSOMMATIONS
        # ============================================================
        df_conso = pd.read_excel(uploaded_conso)
        df_conso.columns = [str(c).strip() for c in df_conso.columns]
        
        # Identification des colonnes de consommation (M1 à M12 ou similaires)
        conso_cols = []
        for col in df_conso.columns:
            col_upper = col.upper()
            if col_upper.startswith('M') and col_upper[1:].isdigit():
                conso_cols.append(col)
            elif 'CONSO' in col_upper and any(str(i) in col_upper for i in range(1, 13)):
                conso_cols.append(col)
        
        if len(conso_cols) == 0:
            # Chercher des colonnes numériques qui pourraient être des consommations
            for col in df_conso.columns:
                if col.upper() not in ['CODE ARTICLE', 'CODE', 'ARTICLE', 'DESCRIPTION', 'PRODUIT']:
                    if df_conso[col].dtype in ['float64', 'int64']:
                        conso_cols.append(col)
        
        if len(conso_cols) < 3:
            st.error(f"⚠️ Colonnes de consommation insuffisantes. Trouvées: {conso_cols}. Attendu: M1 à M12 ou colonnes numériques")
            return
        
        # Mapping des colonnes
        code_col = None
        article_col = None
        
        for col in df_conso.columns:
            col_lower = col.lower().strip()
            if 'code' in col_lower and 'article' in col_lower:
                code_col = col
            elif col_lower == 'code':
                code_col = col
            elif 'article' in col_lower or 'designation' in col_lower or 'description' in col_lower:
                article_col = col
        
        if code_col is None:
            st.error("❌ Colonne 'Code article' introuvable dans le fichier des consommations")
            return
        
        # Calcul des statistiques par produit
        stats_produits = []
        
        for _, row in df_conso.iterrows():
            code = str(row[code_col]).strip()
            article = str(row[article_col]).strip() if article_col else code
            
            # Extraire les consommations
            consommations = []
            for col in conso_cols:
                val = clean_num(row[col])
                if val > 0:
                    consommations.append(val)
            
            if len(consommations) >= 2:
                moyenne_mensuelle = np.mean(consommations)
                ecart_type = np.std(consommations, ddof=1) if len(consommations) > 1 else moyenne_mensuelle * 0.15
                cv = (ecart_type / moyenne_mensuelle * 100) if moyenne_mensuelle > 0 else 15
            elif len(consommations) == 1:
                moyenne_mensuelle = consommations[0]
                ecart_type = moyenne_mensuelle * 0.15
                cv = 15
            else:
                moyenne_mensuelle = 0
                ecart_type = 0
                cv = 0
            
            stats_produits.append({
                "Code article": code,
                "Article": article,
                "Conso moyenne/mois": round(moyenne_mensuelle, 2),
                "Écart-type mensuel": round(ecart_type, 2),
                "CV (%)": round(cv, 1),
                "Nb mois": len(consommations),
                "Consommations": consommations
            })
        
        df_stats = pd.DataFrame(stats_produits)
        
        st.success(f"✅ {len(df_stats)} produits chargés depuis le fichier des consommations")
        
        # ============================================================
        # CHARGEMENT DES PRODUITS ET PARAMÈTRES
        # ============================================================
        df_produits = pd.read_excel(uploaded_produits)
        df_produits.columns = [str(c).strip() for c in df_produits.columns]
        
        # Mapping des colonnes produits
        code_col_prod = None
        article_col_prod = None
        source_col = None
        unite_col = None
        cout_unitaire_col = None
        cout_passation_col = None
        taux_stockage_col = None
        
        for col in df_produits.columns:
            col_lower = col.lower().strip()
            if 'code' in col_lower and 'article' in col_lower:
                code_col_prod = col
            elif col_lower == 'code':
                code_col_prod = col
            elif 'article' in col_lower or 'designation' in col_lower or 'description' in col_lower:
                article_col_prod = col
            elif col_lower == 'source':
                source_col = col
            elif 'unite' in col_lower or 'um' in col_lower:
                unite_col = col
            elif 'cout' in col_lower or 'coût' in col_lower:
                if 'passation' in col_lower or 'commande' in col_lower:
                    cout_passation_col = col
                elif 'unitaire' in col_lower or 'achat' in col_lower:
                    cout_unitaire_col = col
                else:
                    cout_unitaire_col = col
            elif 'taux' in col_lower and ('stockage' in col_lower or 'detention' in col_lower):
                taux_stockage_col = col
            elif 'stockage' in col_lower:
                taux_stockage_col = col
        
        # Vérification des colonnes obligatoires
        missing_cols = []
        if code_col_prod is None:
            missing_cols.append("Code article")
        if cout_unitaire_col is None:
            missing_cols.append("Coût unitaire")
        if cout_passation_col is None:
            missing_cols.append("Coût passation")
        if taux_stockage_col is None:
            missing_cols.append("Taux stockage (%)")
        if source_col is None:
            missing_cols.append("Source")
        
        if missing_cols:
            st.error(f"❌ Colonnes obligatoires manquantes dans le fichier produits: {', '.join(missing_cols)}")
            st.info("Veuillez ajouter ces colonnes avec les bonnes intitulés")
            return
        
        # Nettoyage des données
        df_produits[code_col_prod] = df_produits[code_col_prod].astype(str).str.strip()
        df_produits[cout_unitaire_col] = df_produits[cout_unitaire_col].apply(clean_num)
        df_produits[cout_passation_col] = df_produits[cout_passation_col].apply(clean_num)
        df_produits[taux_stockage_col] = df_produits[taux_stockage_col].apply(clean_num)
        
        # Valeur par défaut pour la source
        if source_col:
            df_produits[source_col] = df_produits[source_col].fillna("Local")
        
        st.success(f"✅ {len(df_produits)} produits chargés depuis le fichier des paramètres")
        
        # ============================================================
        # FUSION DES DONNÉES
        # ============================================================
        df_merged = df_stats.merge(
            df_produits[[code_col_prod, article_col_prod, source_col, unite_col, 
                         cout_unitaire_col, cout_passation_col, taux_stockage_col]],
            left_on="Code article",
            right_on=code_col_prod,
            how="inner"
        )
        
        if len(df_merged) == 0:
            st.error("❌ Aucun produit en commun entre les deux fichiers. Vérifiez les codes article.")
            return
        
        st.success(f"✅ {len(df_merged)} produits appariés avec succès")
        
        # Renommage pour faciliter l'utilisation
        df_merged = df_merged.rename(columns={
            cout_unitaire_col: "Coût unitaire",
            cout_passation_col: "Coût passation",
            taux_stockage_col: "Taux stockage (%)",
            source_col: "Source",
            unite_col: "Unité"
        })
        
        if article_col_prod:
            df_merged["Article"] = df_merged[article_col_prod]
        elif article_col:
            df_merged["Article"] = df_merged["Article"]
        else:
            df_merged["Article"] = df_merged["Code article"]
        
        # ============================================================
        # ANALYSE ABC (basée sur le CA)
        # ============================================================
        df_merged["CA mensuel"] = df_merged["Conso moyenne/mois"] * df_merged["Coût unitaire"]
        df_merged["CA annuel"] = df_merged["CA mensuel"] * 12
        
        df_abc = df_merged.sort_values("CA annuel", ascending=False).copy()
        df_abc["CA cumulé"] = df_abc["CA annuel"].cumsum()
        df_abc["CA cumulé %"] = df_abc["CA cumulé"] / df_abc["CA annuel"].sum() * 100
        
        def assign_abc(row):
            if row["CA cumulé %"] <= 70:
                return "A (70% CA)", 2.33  # Z = 99%
            elif row["CA cumulé %"] <= 90:
                return "B (20% CA)", 1.65  # Z = 95%
            else:
                return "C (10% CA)", 1.28  # Z = 90%
        
        df_abc[["Classe ABC", "Z"]] = df_abc.apply(lambda r: pd.Series(assign_abc(r)), axis=1)
        
        # ============================================================
        # CALCULS FINAUX
        # ============================================================
        results = []
        
        for _, row in df_abc.iterrows():
            code = row["Code article"]
            conso_mensuelle = row["Conso moyenne/mois"]
            conso_annuelle = conso_mensuelle * 12
            cout_unitaire = row["Coût unitaire"]
            cout_passation = row["Coût passation"]
            taux_stockage = row["Taux stockage (%)"]
            source = row["Source"]
            z = row["Z"]
            ecart_type = row["Écart-type mensuel"]
            
            # Délai selon source
            delay = SOURCE_DELAYS.get(source, SOURCE_DELAYS["Local"])
            lt_mois = delay["mois"]
            lt_jours = delay["jours"]
            
            if conso_mensuelle > 0 and cout_unitaire > 0:
                # Calcul de la variabilité
                variabilite = (ecart_type / conso_mensuelle) if conso_mensuelle > 0 else 0.15
                
                # Stock de sécurité
                ss = calc_ss(conso_mensuelle, z=z, lt_mois=lt_mois, variab=variabilite)
                
                # Quantité économique (EOQ)
                eoq, nb_cmd, ct_min = calc_eoq(conso_annuelle, cout_unitaire, cout_passation, taux_stockage/100)
                
                # Point de commande
                pr = calc_pr(conso_mensuelle, lt_mois, ss)
                
                # Coût annuel du stock de sécurité
                cout_ss = round(ss * cout_unitaire * (taux_stockage/100), 2)
                
                # Stock de sécurité en jours de couverture
                conso_journaliere = conso_mensuelle / 30
                ss_jours = round(ss / conso_journaliere, 1) if conso_journaliere > 0 else 0
            else:
                ss = 0
                eoq = 0
                pr = 0
                cout_ss = 0
                nb_cmd = 0
                ss_jours = 0
            
            results.append({
                "Code article": code,
                "Article": row.get("Article", ""),
                "Source": source,
                "Délai (jours)": lt_jours,
                "Unité": row.get("Unité", ""),
                "Conso moyenne/mois": conso_mensuelle,
                "Écart-type": round(ecart_type, 2),
                "CV (%)": row["CV (%)"],
                "Coût unitaire": cout_unitaire,
                "Coût passation": cout_passation,
                "Taux stockage (%)": taux_stockage,
                "Classe ABC": row["Classe ABC"],
                "Z (niveau service)": z,
                "CA annuel (TND)": round(row["CA annuel"], 0),
                "EOQ": eoq,
                "Stock sécurité (u)": ss,
                "SS (jours couverture)": ss_jours,
                "Point commande (u)": pr,
                "Coût SS annuel (TND)": cout_ss,
                "Nb commandes/an": round(nb_cmd, 1) if nb_cmd else 0
            })
        
        df_results = pd.DataFrame(results)
        
        # ============================================================
        # AFFICHAGE DES RÉSULTATS
        # ============================================================
        
        # Aperçu des consommations chargées
        with st.expander("👁️ Aperçu des consommations chargées", expanded=False):
            st.dataframe(df_stats[["Code article", "Article", "Conso moyenne/mois", "Écart-type mensuel", "CV (%)", "Nb mois"]], 
                        use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("#### 📊 Résultats des calculs")
        
        # Tableau principal
        display_cols = ["Code article", "Article", "Source", "Délai (jours)", "Unité",
                       "Conso moyenne/mois", "Classe ABC", "Z (niveau service)",
                       "EOQ", "Stock sécurité (u)", "SS (jours couverture)", 
                       "Point commande (u)", "Coût SS annuel (TND)", "Nb commandes/an"]
        
        st.dataframe(
            df_results[display_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Conso moyenne/mois": st.column_config.NumberColumn("Conso moyenne/mois", format="%.1f"),
                "Coût unitaire": st.column_config.NumberColumn("Coût unitaire (TND)", format="%.2f"),
                "Coût passation": st.column_config.NumberColumn("Coût passation (TND)", format="%.0f"),
                "Taux stockage (%)": st.column_config.NumberColumn("Taux stockage (%)", format="%.1f%%"),
                "CA annuel (TND)": st.column_config.NumberColumn("CA annuel (TND)", format="%.0f"),
                "EOQ": st.column_config.NumberColumn("EOQ", format="%d"),
                "Stock sécurité (u)": st.column_config.NumberColumn("Stock sécurité", format="%d"),
                "Point commande (u)": st.column_config.NumberColumn("Point commande", format="%d"),
                "Coût SS annuel (TND)": st.column_config.NumberColumn("Coût SS annuel", format="%.2f"),
            }
        )
        
        # ============================================================
        # INDICATEURS SYNTHÉTIQUES
        # ============================================================
        st.markdown("---")
        st.markdown("#### 📈 Synthèse des résultats")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Produits traités</div>
                <div class='metric-value'>{len(df_results)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            total_ss = df_results["Stock sécurité (u)"].sum()
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>SS total</div>
                <div class='metric-value'>{fmtInt(total_ss)} u</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            total_cout_ss = df_results["Coût SS annuel (TND)"].sum()
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>Coût SS annuel</div>
                <div class='metric-value'>{fmtInt(total_cout_ss)} TND</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            eoq_mean = df_results[df_results["EOQ"] > 0]["EOQ"].mean()
            st.markdown(f"""
            <div class='medoil-card' style='text-align:center'>
                <div class='metric-label'>EOQ moyen</div>
                <div class='metric-value'>{fmtInt(eoq_mean)} u</div>
            </div>
            """, unsafe_allow_html=True)
        
        # ============================================================
        # GRAPHIQUES
        # ============================================================
        st.markdown("---")
        st.markdown("#### 📊 Visualisation des indicateurs")
        
        tab1, tab2, tab3 = st.tabs(["📊 EOQ / SS / PR", "🏷️ Analyse ABC", "📈 Évolution CV"])
        
        with tab1:
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
                y=df_results["Stock sécurité (u)"],
                marker_color=COLORS["secondary"],
                text=df_results["Stock sécurité (u)"].apply(lambda v: f"{int(v):,}"),
                textposition="outside"
            ))
            
            fig.add_trace(go.Scatter(
                name="Point commande",
                x=df_results["Code article"].astype(str),
                y=df_results["Point commande (u)"],
                mode="markers+lines",
                marker=dict(color=COLORS["danger"], size=9, symbol="diamond"),
                line=dict(color=COLORS["danger"], width=2, dash="dash")
            ))
            
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                height=450,
                barmode="group",
                xaxis=dict(title="Article", tickangle=-45),
                yaxis=dict(title="Unités")
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            # Graphique Pareto ABC
            fig, ax = plt.subplots(figsize=(10, 5))
            
            colors_abc = {"A (70% CA)": COLORS["danger"], "B (20% CA)": COLORS["warning"], "C (10% CA)": COLORS["success"]}
            bar_colors = [colors_abc.get(classe, COLORS["gray"]) for classe in df_abc["Classe ABC"]]
            
            ax.bar(df_abc["Code article"].astype(str), df_abc["CA annuel"], color=bar_colors, alpha=0.7, label="CA annuel")
            ax2 = ax.twinx()
            ax2.plot(df_abc["Code article"].astype(str), df_abc["CA cumulé %"], color=COLORS["primary"], marker='o', linewidth=2, label="CA cumulé %")
            ax2.axhline(y=70, color='green', linestyle='--', alpha=0.7, label="Seuil A (70%)")
            ax2.axhline(y=90, color='orange', linestyle='--', alpha=0.7, label="Seuil B (90%)")
            
            ax.set_xlabel("Article")
            ax.set_ylabel("CA annuel (TND)")
            ax2.set_ylabel("CA cumulé (%)")
            ax.tick_params(axis='x', rotation=45)
            
            # Légende
            from matplotlib.patches import Patch
            legend_elements = [Patch(facecolor=COLORS["danger"], label='Classe A (70% CA)'),
                              Patch(facecolor=COLORS["warning"], label='Classe B (20% CA)'),
                              Patch(facecolor=COLORS["success"], label='Classe C (10% CA)')]
            ax.legend(handles=legend_elements, loc='upper left')
            
            st.pyplot(fig)
            plt.close()
        
        with tab3:
            # Graphique du coefficient de variation
            fig = go.Figure()
            
            cv_data = df_results.sort_values("CV (%)", ascending=False)
            colors_cv = ["#ef4444" if cv > 30 else "#f59e0b" if cv > 15 else "#22c55e" for cv in cv_data["CV (%)"]]
            
            fig.add_trace(go.Bar(
                x=cv_data["Code article"].astype(str),
                y=cv_data["CV (%)"],
                marker_color=colors_cv,
                text=cv_data["CV (%)"].apply(lambda v: f"{v:.1f}%"),
                textposition="outside"
            ))
            
            fig.add_hline(y=15, line_dash="dash", line_color="#f59e0b", annotation_text="Variabilité modérée (15%)")
            fig.add_hline(y=30, line_dash="dash", line_color="#ef4444", annotation_text="Variabilité élevée (30%)")
            
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                height=400,
                xaxis=dict(title="Article", tickangle=-45),
                yaxis=dict(title="Coefficient de variation (%)", range=[0, max(50, cv_data["CV (%)"].max() + 10)])
            )
            st.plotly_chart(fig, use_container_width=True)
        
        # ============================================================
        # EXPORT
        # ============================================================
        st.markdown("---")
        st.markdown("#### 💾 Export des résultats")
        
        excel_buf = build_excel_complete_calculs(df_results, df_stats, df_abc)
        
        st.download_button(
            label="⬇️ Télécharger le fichier Excel complet",
            data=excel_buf,
            file_name=f"medoil_supply_chain_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        # Stockage dans session_state
        st.session_state['df_calculs'] = df_results
        st.session_state['df_stats'] = df_stats
        st.session_state['df_abc'] = df_abc
        
        st.success("✅ Calculs terminés avec succès !")
        
    except Exception as e:
        st.error(f"❌ Erreur lors du calcul : {str(e)}")
        st.exception(e)


def build_excel_complete_calculs(df_results, df_stats, df_abc):
    """Crée un fichier Excel complet avec toutes les feuilles"""
    wb = Workbook()
    
    thin = Side(style="thin", color="BFBFBF")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    def hdr(cell, val, bg):
        cell.value = val
        cell.font = Font(bold=True, size=9, name="Arial")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd
    
    # Feuille 1: Résultats finaux
    ws1 = wb.active
    ws1.title = "Résultats calculs"
    
    hdrs = list(df_results.columns)
    for col, label in enumerate(hdrs, 1):
        hdr(ws1.cell(1, col), label, COLORS["primary"])
    
    for i, row in df_results.iterrows():
        r = i + 2
        for col, col_name in enumerate(hdrs, 1):
            val = row[col_name]
            cell = ws1.cell(r, col)
            cell.value = val
            cell.border = brd
            if isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
    
    for col, w in enumerate([15, 30, 10, 12, 8, 15, 12, 10, 12, 12, 12, 12, 15, 12, 12, 15, 12, 12, 15, 12], 1):
        ws1.column_dimensions[get_column_letter(col)].width = w
    
    # Feuille 2: Statistiques consommations
    ws2 = wb.create_sheet("Statistiques conso")
    
    stats_hdrs = ["Code article", "Article", "Conso moyenne/mois", "Écart-type mensuel", "CV (%)", "Nb mois"]
    for col, label in enumerate(stats_hdrs, 1):
        hdr(ws2.cell(1, col), label, COLORS["secondary"])
    
    for i, row in df_stats.iterrows():
        r = i + 2
        for col, col_name in enumerate(stats_hdrs, 1):
            ws2.cell(r, col, row.get(col_name, ""))
            ws2.cell(r, col).border = brd
    
    for col, w in enumerate([15, 30, 18, 15, 10, 10], 1):
        ws2.column_dimensions[get_column_letter(col)].width = w
    
    # Feuille 3: Analyse ABC
    ws3 = wb.create_sheet("Analyse ABC")
    
    abc_hdrs = ["Code article", "Article", "CA mensuel", "CA annuel", "CA cumulé %", "Classe ABC", "Z"]
    for col, label in enumerate(abc_hdrs, 1):
        hdr(ws3.cell(1, col), label, COLORS["accent"])
    
    for i, row in df_abc.iterrows():
        r = i + 2
        vals = [row.get("Code article", ""), row.get("Article", ""),
                row.get("CA mensuel", 0), row.get("CA annuel", 0),
                row.get("CA cumulé %", 0), row.get("Classe ABC", ""),
                row.get("Z", 0)]
        for col, val in enumerate(vals, 1):
            ws3.cell(r, col, val)
            ws3.cell(r, col).border = brd
    
    for col, w in enumerate([15, 30, 15, 15, 12, 15, 10], 1):
        ws3.column_dimensions[get_column_letter(col)].width = w
    
    # Feuille 4: Paramètres
    ws4 = wb.create_sheet("Paramètres")
    
    params_hdrs = ["Paramètre", "Valeur", "Description"]
    for col, label in enumerate(params_hdrs, 1):
        hdr(ws4.cell(1, col), label, COLORS["gray"])
    
    params = [
        ["Date calcul", datetime.now().strftime("%Y-%m-%d %H:%M"), ""],
        ["Source - Export", "Délai: 4 mois (120 jours)", ""],
        ["Source - Local", "Délai: 0.5 mois (15 jours)", ""],
        ["Source - BM", "Délai: 0.066 mois (2 jours)", ""],
        ["Classe A (70% CA)", "Z = 2.33 (99% service)", ""],
        ["Classe B (20% CA)", "Z = 1.65 (95% service)", ""],
        ["Classe C (10% CA)", "Z = 1.28 (90% service)", ""],
    ]
    
    for i, (param, val, desc) in enumerate(params, 2):
        ws4.cell(i, 1, param)
        ws4.cell(i, 2, val)
        ws4.cell(i, 3, desc)
        for col in range(1, 4):
            ws4.cell(i, col).border = brd
    
    ws4.column_dimensions['A'].width = 25
    ws4.column_dimensions['B'].width = 30
    ws4.column_dimensions['C'].width = 40
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
 
# ─────────────────────────────────────────────────────────────────────────────
def calculateurs():
    st.markdown("## ⚙️ Calculateurs Supply Chain")
    tab_ss, tab_eoq, tab_kpi, tab_rp = st.tabs(
        ["📦 Stock de sécurité", "📐 EOQ (Wilson)", "📊 KPIs stock", "🔄 Point de réappro."])
 
    with tab_ss:
        st.markdown(fbox("Formule","SS = Z × √(LT×σD² + D²×σLT²)",
                         "Z = facteur service · σD = écart-type demande · LT = délai (j)"), unsafe_allow_html=True)
 
        # ── Import de données pour calcul automatique de σD ──────────────────
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
                    sd_auto  = round(series.mean() / 30, 2)   # demande journalière
                    ss2_auto = round(float(series.std()), 2)   # écart-type mensuel → converti en journalier
                    # σD journalier = σ_mensuel / sqrt(30)
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
            ss2 = st.number_input("Écart-type demande σD (u/j)", value=float(ss2_auto_daily) if ss2_auto_daily else 8.0, step=0.5,
                                  help="Calculé automatiquement si fichier importé, sinon saisissez manuellement")
        with c2:
            slt = st.number_input("Délai fourni. (j)",     value=7.0,  step=1.0)
            ssl = st.number_input("Écart-type délai σLT (j)", value=1.5, step=0.1,
                                  help="Variabilité du délai fournisseur en jours")
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
            ("SS (σD seul)",          fmtInt(SS1),       "unités", "#3b82f6"),
            ("SS (formule complète)", fmtInt(SS2),       "unités", "#22c55e"),
            ("Coût immobilisé SS",    fmtInt(SS2 * scu), "TND",    "#f59e0b"),
        ]):
            with c[i]: st.markdown(mcard(l,v,u,color=col), unsafe_allow_html=True)
 
        st.markdown("<div style='margin-bottom:14px'></div>", unsafe_allow_html=True)
 
        # Graphique sensibilité
        z_vals  = [0.84, 1.04, 1.28, 1.65, 1.88, 2.05, 2.33]
        ns_vals = [80, 85, 90, 95, 97, 98, 99]
        ss_v    = [round(z * math.sqrt(slt * ss2**2 + sd**2 * ssl**2)) for z in z_vals]
        current_pct = int(szo.split("%")[0])
        fig = go.Figure(go.Bar(
            x=[f"{n}%" for n in ns_vals], y=ss_v,
            marker_color=["#22c55e" if n == current_pct else "#3b82f6" for n in ns_vals],
            text=ss_v, textposition="outside"))
        fig.update_layout(**dark_layout(), height=260,
            xaxis=dict(title="Niveau de service", gridcolor="rgba(255,255,255,.05)"),
            yaxis=dict(title="SS (unités)",        gridcolor="rgba(255,255,255,.05)"))
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
        data = [("EOQ",            fmtInt(EOQ),    "u/cmd", "#3b82f6"),
                ("Nb cmd/an",      fmt(nb,1),      "cmds",  "#8892a4"),
                ("Périodicité",    fmtInt(T),      "jours", "#8892a4"),
                ("Coût passation/an", fmtInt(Cp),  "TND",   "#f59e0b"),
                ("Coût stockage/an",  fmtInt(Cs),  "TND",   "#f59e0b"),
                ("Coût total min.", fmtInt(Cp+Cs), "TND",   "#22c55e")]
        for i,(l,v,u,col) in enumerate(data):
            with c[i%3]:
                st.markdown(mcard(l,v,u,color=col), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px'></div>", unsafe_allow_html=True)
        qty_r = np.linspace(max(10, EOQ*.2), EOQ*2.5, 200)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=qty_r, y=[(eD/q)*eSc for q in qty_r],            name="Passation", line=dict(color="#60a5fa",width=2)))
        fig.add_trace(go.Scatter(x=qty_r, y=[(q/2)*h    for q in qty_r],             name="Stockage",  line=dict(color="#f59e0b",width=2)))
        fig.add_trace(go.Scatter(x=qty_r, y=[(eD/q)*eSc+(q/2)*h for q in qty_r],    name="Total",     line=dict(color="#22c55e",width=2.5)))
        fig.add_vline(x=EOQ, line_dash="dash", line_color="#ef4444",
                      annotation_text=f"EOQ={fmtInt(EOQ)}", annotation_font_color="#ef4444")
        fig.update_layout(**dark_layout(), height=300,
            xaxis=dict(title="Quantité (u)", gridcolor="rgba(255,255,255,.05)"),
            yaxis=dict(title="Coût/an (TND)",gridcolor="rgba(255,255,255,.05)"))
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
            ("Rotation",         fmt(rot,1)+"×",          "/an · >6×=excellent",  "#22c55e" if rot>=6 else "#f59e0b"),
            ("Couverture",       fmtInt(cov)+"j",         "15–30j=optimal",        "#8892a4"),
            ("Taux service OTD", fmt(srv,1)+"%",           ">95%=cible",           "#22c55e" if srv>=95 else "#f59e0b"),
            ("Coût log/CA",      fmt(crt,1)+"%",           "objectif <8%",          "#22c55e" if crt<8 else "#ef4444"),
            ("Coût stockage/an", fmtInt(kval*(ksh/100)),  "TND",                   "#f59e0b"),
            ("DSI",              fmt(ks/(kv/365),0)+"j",  "jours de stock",        "#8892a4"),
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
            ("Point de réappro.", fmtInt(PR),      "unités",             "#3b82f6"),
            ("Jours restants",    fmt(dL,1),       "jours",              "#8892a4"),
            ("Délai critique",    fmt(dL-rlt,1),   "jours avant urgence","#f59e0b"),
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
        gauge={"axis":{"range":[0,100]}, "bar":{"color":col_c,"thickness":.25}, "bgcolor":"#1e2535",
               "steps":[{"range":[0,55],"color":"rgba(239,68,68,.2)"},{"range":[55,70],"color":"rgba(245,158,11,.2)"},
                        {"range":[70,85],"color":"rgba(20,184,166,.2)"},{"range":[85,100],"color":"rgba(34,197,94,.2)"}]}))
    fg.update_layout(paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#8892a4",family="Syne"),
                     height=220, margin=dict(t=10,b=10,l=20,r=20))
    cg,cd = st.columns([1,2])
    with cg: st.plotly_chart(fg, use_container_width=True)
    with cd:
        st.markdown(f"<div style='background:#1e2535;border-left:4px solid {col_c};border-radius:10px;"
                    f"padding:1.2rem 1.4rem;margin-top:1rem'>"
                    f"<div style='font-size:1.2rem;font-weight:600;color:{col_c};margin-bottom:8px'>{lbl} — {pn}</div>"
                    f"<div style='font-size:.85rem;color:#8892a4;line-height:1.6'>{adv}</div></div>",
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
                f"<div style='background:#1e2535;border:1px solid rgba(255,255,255,.08);border-radius:10px;"
                f"padding:1rem;margin-bottom:10px'>"
                f"<div style='display:flex;justify-content:space-between;margin-bottom:8px'>"
                f"<span style='font-size:.85rem;font-weight:600;color:#e8eaf0'>{nm}</span>"
                f"<span style='font-family:\"DM Mono\",monospace;color:{cc}'>{scr:.1f}%</span></div>"
                f"<div style='background:#0f1117;border-radius:4px;height:6px;overflow:hidden'>"
                f"<div style='background:{cc};height:100%;width:{scr}%'></div></div>"
                f"<div style='font-size:.7rem;color:#5a6478;margin-top:6px'>{nt}</div></div>",
                unsafe_allow_html=True)
 
    if otd < 85: st.markdown(f"<div class='alert-warning'>⏱ OTD insuffisant ({otd:.1f}%) — Négocier des pénalités. Majorer le stock de sécurité.</div>", unsafe_allow_html=True)
    if qu  < 95: st.markdown(f"<div class='alert-warning'>⚠️ {pnc} non-conformités — Audit qualité recommandé.</div>", unsafe_allow_html=True)
    if plr > pla: st.markdown(f"<div class='alert-warning'>📅 Dérive délai +{plr-pla:.1f}j — Recalculer le point de réappro. avec le délai réel.</div>", unsafe_allow_html=True)
    if gs  >= 85: st.markdown("<div class='alert-ok'>✅ Performance excellente — Envisager un accord-cadre annuel.</div>", unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────────────────────────────────────────
def alertes():
    st.markdown("## 🔔 Alertes & Surveillance des stocks")
    st.markdown("<p style='color:#8892a4'>Importez votre tableau de suivi — les alertes sont générées automatiquement</p>",
                unsafe_allow_html=True)
 
    # ── HELPER : nettoyer une valeur numérique (gère les virgules, espaces, #DIV/0!) ──
    def clean_num(v):
        if v is None: return 0.0
        s = str(v).strip().replace(" ", "").replace(",", ".")
        if s in ("#DIV/0!", "#N/A", "#VALEUR!", "#REF!", "", "-", "—"): return 0.0
        try: return float(s)
        except: return 0.0
 
    # ── UPLOAD ───────────────────────────────────────────────────────────────
    uploaded = st.file_uploader(
        "📂 Déposez votre fichier Excel de suivi des stocks (.xlsx / .xls)",
        type=["xlsx", "xls"], key="alertes_upload")
 
    df_alertes = None
 
    if uploaded:
        try:
            # 1. Lire tout le fichier sans charger les données pour avoir les noms des feuilles
            excel_file = pd.ExcelFile(uploaded)
            sheet_names = excel_file.sheet_names
            
            st.info(f"Fichier chargé : {len(sheet_names)} feuilles détectées.")
            
            # 2. Créer des onglets Streamlit dynamiquement
            tabs = st.tabs(sheet_names)
            
            for i, sheet_name in enumerate(sheet_names):
                with tabs[i]:
                    st.subheader(f"Analyse de la feuille : {sheet_name}")
                    
                    # --- ÉTAPE DE LECTURE SPÉCIFIQUE ---
                    # On relit la feuille pour trouver l'en-tête
                    raw = pd.read_excel(uploaded, sheet_name=sheet_name, header=None)
                    
                    hrow = 0
                    for idx, row in raw.iterrows():
                        row_str = " ".join([str(v).lower() for v in row if pd.notna(v)])
                        if any(k in row_str for k in ["codearticle", "article", "stock"]):
                            hrow = idx
                            break
                    
                    # Chargement final de la feuille en cours
                    df_current = pd.read_excel(uploaded, sheet_name=sheet_name, header=hrow)
                    
                    # --- APPEL DE LA LOGIQUE DE TRAITEMENT ---
                    # On passe le dataframe à une fonction de traitement
                    process_and_display_alerts(df_current, sheet_name)

        except Exception as e:
            st.error(f"Erreur lors du traitement : {e}")
 
            df_raw = pd.read_excel(uploaded, sheet_name=sel, header=hrow)
            df_raw.columns = [str(c).strip() for c in df_raw.columns]
            df_raw = df_raw.dropna(how="all")
 
            # ── Mapping flexible des colonnes ─────────────────────────────
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
 
            # Colonnes numériques à nettoyer
            num_cols = ["Stock Sécurité", "Stock actuel", "Stk jour SS", "Cons Moy",
                        "Besoin M1", "Besoin M2", "Besoin M3",
                        "Stock M+1", "Stock M+2", "Stock M+3", "Commande en cours"]
            for c in num_cols:
                if c in df_mapped.columns:
                    df_mapped[c] = df_mapped[c].apply(clean_num)
 
            # Filtrer les lignes sans code article
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
 
    # ── GÉNÉRER LES ALERTES ──────────────────────────────────────────────────
    df = df_alertes.copy()
 
    # Couverture en mois (depuis le fichier ou calculée)
    if "Couverture" not in df.columns or df["Couverture"].sum() == 0:
        df["Couverture"] = np.where(
            df.get("Cons Moy", pd.Series([0]*len(df))) > 0,
            (df.get("Stock actuel", pd.Series([0]*len(df))) / df["Cons Moy"]).round(1),
            np.nan)
 
    # Statut basé sur couverture et Stock M+1/M+2/M+3
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
 
    # ── KPIs synthèse ────────────────────────────────────────────────────────
    rupt  = df[df["Statut"] == "🔴 Rupture imminente"]
    da    = df[df["Statut"].isin(["🟡 DA à lancer", "🟠 Risque M+3"])]
    ok    = df[df["Statut"] == "🟢 Normal"]
 
    st.markdown("---")
    st.markdown("#### 📊 Synthèse")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(mcard("🔴 Rupture imminente",   str(len(rupt)), "articles à commander d'urgence", color="#ef4444"), unsafe_allow_html=True)
    with c2: st.markdown(mcard("🟡 DA à lancer",         str(len(da)),   "articles nécessitant une DA",    color="#f59e0b"), unsafe_allow_html=True)
    with c3: st.markdown(mcard("🟢 Normal",              str(len(ok)),   "articles en stock suffisant",    color="#22c55e"), unsafe_allow_html=True)
    with c4: st.markdown(mcard("📦 Total articles",       str(len(df)),   "articles dans le fichier",       color="#3b82f6"), unsafe_allow_html=True)
 
    # ── TABLEAU COMPLET ──────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### 📋 Tableau de bord complet")
 
    # Colonnes à afficher
    cols_display = ["Code article", "Désignation"]
    for c in ["Source", "UM", "Stock Sécurité", "Stock actuel", "Cons Moy",
              "Couverture", "Besoin M1", "Besoin M2", "Besoin M3",
              "Stock M+1", "Stock M+2", "Stock M+3", "Commentaire", "Statut"]:
        if c in df.columns:
            cols_display.append(c)
 
    df_show = df[cols_display].copy()
 
    # Filtre par statut
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
 
    # ── GRAPHIQUE Stock actuel vs Stock Sécurité ─────────────────────────────
    if "Stock actuel" in df.columns and "Stock Sécurité" in df.columns:
        st.markdown("---")
        st.markdown("#### 📈 Stock actuel vs Stock de sécurité")
 
        labels = df["Code article"].astype(str).tolist()
        colors = ["#ef4444" if s == "🔴 Rupture imminente"
                  else "#f59e0b" if s in ("🟡 DA à lancer", "🟠 Risque M+3")
                  else "#22c55e" for s in df["Statut"]]
 
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Stock actuel", x=labels, y=df["Stock actuel"], marker_color=colors))
        fig.add_trace(go.Scatter(name="Stock Sécurité", x=labels, y=df["Stock Sécurité"],
            mode="markers+lines", marker=dict(color="#a78bfa", size=7, symbol="diamond"),
            line=dict(color="#a78bfa", width=1.5, dash="dash")))
        if "Stock M+1" in df.columns:
            fig.add_trace(go.Scatter(name="Stock M+1", x=labels, y=df["Stock M+1"],
                mode="markers+lines", marker=dict(color="#60a5fa", size=6),
                line=dict(color="#60a5fa", width=1.2, dash="dot")))
        fig.update_layout(**dark_layout(), height=360, barmode="group",
            xaxis=dict(tickangle=-45, gridcolor="rgba(255,255,255,.05)"),
            yaxis=dict(title="Unités", gridcolor="rgba(255,255,255,.05)"),
            shapes=[dict(type="line", x0=-0.5, x1=len(labels)-0.5, y0=0, y1=0,
                         line=dict(color="#ef4444", width=1.5, dash="dash"))])
        st.plotly_chart(fig, use_container_width=True)
 
    # ── ALERTES TEXTUELLES ───────────────────────────────────────────────────
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
    <div style="display: flex; align-items: center;">
        <img src={URL_LOGO} style="width: 50px; margin-right: 10px;">
        <h1 style="margin: 0;">Med oil</h1>
    </div>
    """,
    unsafe_allow_html=True
)
    st.divider()
 
    pg = st.navigation([
    st.Page(accueil,       title="Accueil",                   icon="🏠"),
    st.Page(import_calcul, title="Import & Calcul automatique",icon="📥"),
    st.Page(calculateurs,  title="Calculateurs",               icon="⚙️"),
    st.Page(processus,     title="Processus fournisseurs",     icon="🏭"),
    st.Page(alertes,       title="Alertes stock",              icon="🔔"),
])
 
with st.sidebar:
    st.divider()
    st.markdown("<p style='color:#5a6478;font-size:.75rem;'>v2.0 · Supply Chain Manager</p>",
                unsafe_allow_html=True)
 
pg.run()
