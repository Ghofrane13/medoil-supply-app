import streamlit as st
import math
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="medoil-supply-app",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS CUSTOM ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700&family=DM+Mono:wght@400;500&family=DM+Serif+Display:ital@0;1&display=swap');

html, body, [class*="css"] { font-family: 'Syne', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #0f1117 0%, #161b26 100%);
    padding: 2rem 2.5rem;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.08);
    margin-bottom: 1.5rem;
}
.main-header h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 2.4rem;
    color: #e8eaf0;
    margin: 0 0 0.5rem 0;
}
.main-header h1 em { color: #60a5fa; font-style: italic; }
.main-header p { color: #8892a4; font-size: 1rem; margin: 0; }

.kpi-card {
    background: #161b26;
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    text-align: center;
}
.kpi-card .kpi-label { font-size: 0.7rem; color: #5a6478; text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 0.4rem; }
.kpi-card .kpi-value { font-family: 'DM Mono', monospace; font-size: 1.8rem; font-weight: 500; color: #e8eaf0; }
.kpi-card .kpi-unit { font-size: 0.7rem; color: #5a6478; margin-top: 0.2rem; }

.formula-box {
    background: #0f1117;
    border-left: 3px solid #3b82f6;
    border-radius: 8px;
    padding: 0.9rem 1.1rem;
    font-family: 'DM Mono', monospace;
    font-size: 0.82rem;
    color: #8892a4;
    margin-bottom: 1rem;
    line-height: 1.7;
}
.formula-box .f-eq { color: #60a5fa; font-weight: 500; }

.alert-critical {
    background: rgba(239,68,68,0.08);
    border: 1px solid rgba(239,68,68,0.3);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.7rem;
    color: #fca5a5;
}
.alert-warning {
    background: rgba(245,158,11,0.08);
    border: 1px solid rgba(245,158,11,0.3);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.7rem;
    color: #fcd34d;
}
.alert-ok {
    background: rgba(34,197,94,0.08);
    border: 1px solid rgba(34,197,94,0.3);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.7rem;
    color: #86efac;
}
.alert-info {
    background: rgba(20,184,166,0.08);
    border: 1px solid rgba(20,184,166,0.3);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.7rem;
    color: #5eead4;
}
.score-badge {
    display: inline-block;
    padding: 0.3rem 0.9rem;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
}
.section-card {
    background: #161b26;
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px;
    padding: 1.5rem;
    margin-bottom: 1.2rem;
}
.divider { border-top: 1px solid rgba(255,255,255,0.08); margin: 1.2rem 0; }
</style>
""", unsafe_allow_html=True)

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 SupplyIQ")
    st.markdown("<p style='color:#8892a4; font-size:0.82rem; margin-top:-10px;'>Outil Supply Chain Manager</p>", unsafe_allow_html=True)
    st.divider()
    page = st.radio(
        "Navigation",
        ["🏠 Accueil", "⚙️ Calculateurs", "🏭 Processus fournisseurs", "🔔 Alertes stock"],
        label_visibility="collapsed"
    )
    st.divider()
    st.markdown("<p style='color:#5a6478; font-size:0.75rem;'>v2.0 · Supply Chain Manager</p>", unsafe_allow_html=True)

# ─── HELPERS ─────────────────────────────────────────────────────────────────
def fmt(n, d=1):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))):
        return "—"
    return f"{n:,.{d}f}"

def fmtInt(n):
    if n is None or (isinstance(n, float) and (math.isnan(n) or math.isinf(n))):
        return "—"
    return f"{round(n):,}"

def metric_card(label, value, unit, delta=None, color="#3b82f6"):
    delta_html = f"<div style='font-size:0.7rem; color:{color}; margin-top:4px;'>{delta}</div>" if delta else ""
    return f"""
    <div style='background:#1e2535; border:1px solid rgba(255,255,255,0.1); border-radius:12px;
                padding:1rem 1.2rem; text-align:center; border-top: 3px solid {color};'>
        <div style='font-size:0.65rem; color:#5a6478; text-transform:uppercase; letter-spacing:0.1em; margin-bottom:0.4rem;'>{label}</div>
        <div style='font-family:"DM Mono",monospace; font-size:1.6rem; font-weight:500; color:#e8eaf0;'>{value}</div>
        <div style='font-size:0.68rem; color:#5a6478; margin-top:3px;'>{unit}</div>
        {delta_html}
    </div>"""

def formula_box(title, formula, details):
    return f"""
    <div class='formula-box'>
        <div style='font-size:0.65rem; color:#5a6478; text-transform:uppercase; letter-spacing:0.1em; margin-bottom:6px;'>{title}</div>
        <span class='f-eq'>{formula}</span><br/>
        <span style='color:#5a6478;'>{details}</span>
    </div>"""

def score_color(score):
    if score >= 85: return "#22c55e"
    if score >= 70: return "#14b8a6"
    if score >= 55: return "#f59e0b"
    return "#ef4444"

def score_label(score):
    if score >= 85: return "Excellent"
    if score >= 70: return "Bon"
    if score >= 55: return "Moyen"
    return "Insuffisant"

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : ACCUEIL
# ═══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Accueil":
    st.markdown("""
    <div class='main-header'>
        <div style='font-size:0.7rem; color:#60a5fa; letter-spacing:0.12em; text-transform:uppercase; margin-bottom:10px;'>
            Outil supply chain manager · v2.0
        </div>
        <h1>Pilotez votre <em>supply chain</em><br/>avec précision</h1>
        <p>Calculez vos indicateurs clés, évaluez vos processus d'achat et anticipez les ruptures de stock.</p>
    </div>
    """, unsafe_allow_html=True)

    cols = st.columns(4)
    modules = [
        ("📦", "Stock de sécurité", "Calcul selon la méthode statistique avec variabilité de la demande et du délai"),
        ("📐", "Quantité économique", "Modèle de Wilson (EOQ) — minimise le coût total logistique"),
        ("📊", "Indicateurs clés", "Taux de rotation, couverture, taux de service, coût logistique"),
        ("🔄", "Point de réappro.", "Seuil de déclenchement des commandes selon le lead time"),
    ]
    for col, (icon, title, desc) in zip(cols, modules):
        with col:
            st.markdown(f"""
            <div style='background:#161b26; border:1px solid rgba(255,255,255,0.08); border-radius:12px;
                        padding:1.2rem; cursor:pointer; transition:all 0.2s;'>
                <div style='font-size:1.5rem; margin-bottom:8px;'>{icon}</div>
                <div style='font-size:0.9rem; font-weight:600; color:#e8eaf0; margin-bottom:6px;'>{title}</div>
                <div style='font-size:0.78rem; color:#5a6478; line-height:1.5;'>{desc}</div>
            </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown("### 📌 Démarrage rapide")
    st.info("👈 Utilisez la **barre latérale** pour naviguer entre les modules. Commencez par **⚙️ Calculateurs** pour saisir vos données.")

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : CALCULATEURS
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Calculateurs":
    st.markdown("## ⚙️ Calculateurs Supply Chain")
    st.markdown("<p style='color:#8892a4;'>Saisissez vos données — les résultats se calculent automatiquement</p>", unsafe_allow_html=True)

    tab_ss, tab_eoq, tab_kpi, tab_rp = st.tabs([
        "📦 Stock de sécurité",
        "📐 Quantité économique (EOQ)",
        "📊 KPIs stock",
        "🔄 Point de réappro."
    ])

    # ── STOCK DE SÉCURITÉ ────────────────────────────────────────────────────
    with tab_ss:
        st.markdown(formula_box(
            "Formule",
            "SS = Z × σD × √LT",
            "Z = facteur de service · σD = écart-type demande · LT = délai fournisseur (jours)"
        ), unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            ss_demand = st.number_input("Demande moyenne (u/jour)", value=50.0, min_value=0.0, step=1.0, key="ss_d")
            ss_sigma  = st.number_input("Écart-type demande (σD)", value=8.0, min_value=0.0, step=0.5, key="ss_s")
        with col2:
            ss_lt     = st.number_input("Délai fournisseur moyen (jours)", value=7.0, min_value=0.0, step=1.0, key="ss_lt")
            ss_slt    = st.number_input("Écart-type du délai (σLT)", value=1.5, min_value=0.0, step=0.1, key="ss_slt")
        with col3:
            ss_z_opt  = st.selectbox("Niveau de service cible", ["90% (Z=1.28)", "95% (Z=1.65)", "97.5% (Z=1.96)", "99% (Z=2.33)", "99.5% (Z=2.58)"], index=1)
            ss_cost   = st.number_input("Coût unitaire (TND)", value=25.0, min_value=0.0, step=1.0, key="ss_cu")

        z_map = {"90% (Z=1.28)": 1.28, "95% (Z=1.65)": 1.65, "97.5% (Z=1.96)": 1.96, "99% (Z=2.33)": 2.33, "99.5% (Z=2.58)": 2.58}
        Z = z_map[ss_z_opt]

        SS_basic = Z * ss_sigma * math.sqrt(ss_lt) if ss_lt > 0 else 0
        SS_full  = Z * math.sqrt(ss_lt * ss_sigma**2 + ss_demand**2 * ss_slt**2) if ss_lt > 0 else 0
        stockMax = ss_demand * ss_lt + SS_full
        costSS   = SS_full * ss_cost
        cover    = SS_full / ss_demand if ss_demand > 0 else 0

        st.markdown("---")
        st.markdown("#### 📊 Résultats")
        cols = st.columns(3)
        results = [
            ("Stock de sécurité (σD seul)", fmtInt(SS_basic), "unités", f"= {Z} × {ss_sigma} × √{ss_lt}", "#3b82f6"),
            ("Stock de sécurité (formule complète)", fmtInt(SS_full), "unités", "avec variabilité du délai", "#22c55e"),
            ("Stock maximum théorique", fmtInt(stockMax), "unités", f"D×LT + SS = {fmtInt(ss_demand*ss_lt)} + {fmtInt(SS_full)}", "#8892a4"),
            ("Couverture du SS", fmt(cover, 1), "jours de conso.", "SS ÷ demande/jour", "#8892a4"),
            ("Coût immobilisé (SS)", fmtInt(costSS), "TND", f"{fmtInt(SS_full)} × {ss_cost} TND/u", "#f59e0b"),
            ("Niveau de service", ss_z_opt.split(" ")[0], f"Z = {Z}", "", "#a78bfa"),
        ]
        for i, (label, val, unit, delta, color) in enumerate(results):
            with cols[i % 3]:
                st.markdown(metric_card(label, val, unit, delta, color), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

        # Graphique : SS en fonction du niveau de service
        st.markdown("---")
        st.markdown("#### 📈 Sensibilité : SS selon le niveau de service")
        z_vals = [1.04, 1.28, 1.65, 1.96, 2.33, 2.58]
        ns_vals = [85, 90, 95, 97.5, 99, 99.5]
        ss_vals = [round(z * math.sqrt(ss_lt * ss_sigma**2 + ss_demand**2 * ss_slt**2)) for z in z_vals]
        fig = go.Figure(go.Bar(
            x=[f"{n}%" for n in ns_vals], y=ss_vals,
            marker_color=["#3b82f6" if n != float(ss_z_opt.split("%")[0]) else "#22c55e" for n in ns_vals],
            text=ss_vals, textposition='outside',
        ))
        fig.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#8892a4', family='Syne'),
            xaxis=dict(title="Niveau de service", gridcolor='rgba(255,255,255,0.05)'),
            yaxis=dict(title="Stock de sécurité (unités)", gridcolor='rgba(255,255,255,0.05)'),
            height=280, margin=dict(t=20, b=10, l=0, r=0),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── EOQ ──────────────────────────────────────────────────────────────────
    with tab_eoq:
        st.markdown(formula_box(
            "Modèle de Wilson",
            "EOQ = √(2 × D × Sc / (Sh × Cu))",
            "D = demande annuelle · Sc = coût passation · Sh = taux stockage · Cu = coût unitaire"
        ), unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            eoq_D    = st.number_input("Demande annuelle (unités)", value=12000, min_value=1, step=100)
            eoq_Sc   = st.number_input("Coût passation commande (TND)", value=150.0, min_value=0.0, step=5.0)
        with col2:
            eoq_Cu   = st.number_input("Coût unitaire d'achat (TND)", value=80.0, min_value=0.01, step=1.0)
            eoq_Sh   = st.number_input("Taux de stockage annuel (%)", value=20.0, min_value=0.1, max_value=100.0, step=1.0)
        with col3:
            eoq_LT   = st.number_input("Délai fournisseur (jours)", value=10, min_value=1, step=1)
            eoq_days = st.number_input("Jours ouvrables/an", value=250, min_value=1, step=1)

        h      = (eoq_Sh / 100) * eoq_Cu
        EOQ    = math.sqrt(2 * eoq_D * eoq_Sc / h) if h > 0 else 0
        nbCmd  = eoq_D / EOQ if EOQ > 0 else 0
        T      = eoq_days / nbCmd if nbCmd > 0 else 0
        Cpass  = nbCmd * eoq_Sc
        Cstock = (EOQ / 2) * h
        Ctotal = Cpass + Cstock
        RP_eoq = (eoq_D / eoq_days) * eoq_LT

        st.markdown("---")
        st.markdown("#### 📊 Résultats EOQ")
        cols = st.columns(3)
        results = [
            ("Quantité économique (EOQ)", fmtInt(EOQ), "unités / commande", None, "#3b82f6"),
            ("Nombre de commandes/an", fmt(nbCmd, 1), "commandes", None, "#8892a4"),
            ("Périodicité de commande", fmtInt(T), "jours ouvrables", None, "#8892a4"),
            ("Coût passation annuel", fmtInt(Cpass), "TND", None, "#f59e0b"),
            ("Coût stockage annuel", fmtInt(Cstock), "TND", None, "#f59e0b"),
            ("Coût total minimum", fmtInt(Ctotal), "TND/an", "Point optimal de Wilson", "#22c55e"),
        ]
        for i, (label, val, unit, delta, color) in enumerate(results):
            with cols[i % 3]:
                st.markdown(metric_card(label, val, unit, delta, color), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("#### 📈 Courbe des coûts — Modèle de Wilson")
        qty_range = np.linspace(max(10, EOQ * 0.2), EOQ * 2.5, 200)
        c_pass_arr = [(eoq_D / q) * eoq_Sc for q in qty_range]
        c_stock_arr = [(q / 2) * h for q in qty_range]
        c_total_arr = [cp + cs for cp, cs in zip(c_pass_arr, c_stock_arr)]

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=qty_range, y=c_pass_arr, name="Coût passation", line=dict(color="#60a5fa", width=2)))
        fig2.add_trace(go.Scatter(x=qty_range, y=c_stock_arr, name="Coût stockage", line=dict(color="#f59e0b", width=2)))
        fig2.add_trace(go.Scatter(x=qty_range, y=c_total_arr, name="Coût total", line=dict(color="#22c55e", width=2.5)))
        fig2.add_vline(x=EOQ, line_dash="dash", line_color="#ef4444", annotation_text=f"EOQ = {fmtInt(EOQ)} u", annotation_font_color="#ef4444")
        fig2.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#8892a4', family='Syne'),
            xaxis=dict(title="Quantité commandée (u)", gridcolor='rgba(255,255,255,0.05)'),
            yaxis=dict(title="Coût annuel (TND)", gridcolor='rgba(255,255,255,0.05)'),
            legend=dict(bgcolor='rgba(0,0,0,0)', bordercolor='rgba(255,255,255,0.1)', borderwidth=1),
            height=320, margin=dict(t=20, b=10, l=0, r=0),
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── KPIs STOCK ───────────────────────────────────────────────────────────
    with tab_kpi:
        col1, col2 = st.columns(2)
        with col1:
            kpi_avgstock = st.number_input("Stock moyen (unités)", value=3000, min_value=0, step=100)
            kpi_sales    = st.number_input("Ventes annuelles (unités)", value=18000, min_value=1, step=100)
            kpi_val      = st.number_input("Valeur du stock (TND)", value=240000, min_value=0, step=1000)
            kpi_ca       = st.number_input("CA annuel (TND)", value=2400000, min_value=1, step=10000)
        with col2:
            kpi_ontime   = st.number_input("Commandes livrées à temps", value=465, min_value=0, step=1)
            kpi_total    = st.number_input("Total commandes livrées", value=490, min_value=1, step=1)
            kpi_logcost  = st.number_input("Coût logistique total (TND)", value=168000, min_value=0, step=1000)
            kpi_shrate   = st.number_input("Taux de stockage (%)", value=22.0, min_value=0.1, max_value=100.0, step=1.0)

        rotation    = kpi_sales / kpi_avgstock if kpi_avgstock > 0 else 0
        cover       = (kpi_avgstock / kpi_sales) * 365 if kpi_sales > 0 else 0
        service     = (kpi_ontime / kpi_total) * 100 if kpi_total > 0 else 0
        cost_rate   = (kpi_logcost / kpi_ca) * 100 if kpi_ca > 0 else 0
        stock_cost  = kpi_val * (kpi_shrate / 100)
        dsi         = kpi_avgstock / (kpi_sales / 365) if kpi_sales > 0 else 0

        def kpi_color(val, good, ok):
            return "#22c55e" if val >= good else "#f59e0b" if val >= ok else "#ef4444"

        st.markdown("---")
        st.markdown("#### 📊 Résultats KPIs")
        cols = st.columns(3)
        kpis = [
            ("Taux de rotation", fmt(rotation, 1) + "×", "fois/an · >6× = excellent", kpi_color(rotation, 6, 4)),
            ("Couverture stocks", fmtInt(cover) + "j", "jours · 15–30j = optimal", kpi_color(60 - abs(cover - 22.5), 30, 10)),
            ("Taux de service OTD", fmt(service, 1) + "%", "livraisons à temps · >95% cible", kpi_color(service, 95, 85)),
            ("Coût logistique / CA", fmt(cost_rate, 1) + "%", "objectif < 8%", kpi_color(8 - cost_rate + 8, 8, 4)),
            ("Coût de stockage/an", fmtInt(stock_cost), "TND", "#f59e0b"),
            ("Jours de stock (DSI)", fmt(dsi, 0), "jours équivalents", "#8892a4"),
        ]
        for i, (label, val, unit, color) in enumerate(kpis):
            with cols[i % 3]:
                st.markdown(metric_card(label, val, unit, color=color), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

        # Radar chart
        st.markdown("---")
        st.markdown("#### 🕸️ Radar de performance")
        rot_score    = min(100, rotation / 10 * 100)
        serv_score   = service
        cost_score   = max(0, 100 - cost_rate * 5)
        cover_score  = max(0, 100 - abs(cover - 22.5) * 2)
        dsi_score    = max(0, 100 - abs(dsi - 22.5) * 2)

        categories = ["Rotation stocks", "Taux de service", "Coût logistique", "Couverture optimale", "DSI optimal"]
        values = [rot_score, serv_score, cost_score, cover_score, dsi_score]
        values += [values[0]]
        categories += [categories[0]]

        fig3 = go.Figure(go.Scatterpolar(
            r=values, theta=categories, fill='toself',
            fillcolor='rgba(59,130,246,0.15)', line=dict(color='#3b82f6', width=2),
        ))
        fig3.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 100], gridcolor='rgba(255,255,255,0.1)', color='#5a6478'),
                angularaxis=dict(gridcolor='rgba(255,255,255,0.1)', color='#8892a4'),
                bgcolor='rgba(0,0,0,0)',
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#8892a4', family='Syne'),
            height=350, margin=dict(t=20, b=20, l=40, r=40),
            showlegend=False,
        )
        st.plotly_chart(fig3, use_container_width=True)

    # ── POINT DE RÉAPPRO ─────────────────────────────────────────────────────
    with tab_rp:
        st.markdown(formula_box(
            "Formule",
            "PR = (Dmoy × LT) + SS",
            "Dmoy = demande moyenne journalière · LT = délai fournisseur · SS = stock de sécurité"
        ), unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            rp_demand  = st.number_input("Demande moyenne (u/jour)", value=50.0, min_value=0.0, step=1.0, key="rp_d")
            rp_lt      = st.number_input("Délai fournisseur (jours)", value=7.0, min_value=0.0, step=1.0, key="rp_lt")
        with col2:
            rp_ss      = st.number_input("Stock de sécurité (unités)", value=132.0, min_value=0.0, step=10.0, key="rp_ss")
            rp_dmax    = st.number_input("Demande max (u/jour)", value=70.0, min_value=0.0, step=1.0, key="rp_dm")
        with col3:
            rp_ltmax   = st.number_input("Délai fournisseur max (jours)", value=10.0, min_value=0.0, step=1.0, key="rp_ltm")
            rp_current = st.number_input("Stock actuel (unités)", value=500.0, min_value=0.0, step=10.0, key="rp_cur")

        PR       = rp_demand * rp_lt + rp_ss
        PR_max   = rp_dmax * rp_ltmax
        days_left = (rp_current - rp_ss) / rp_demand if rp_demand > 0 else 0
        days_crit = days_left - rp_lt

        st.markdown("---")
        st.markdown("#### 📊 Résultats")

        # Statut
        if rp_current <= rp_ss:
            st.markdown("<div class='alert-critical'>🔴 <strong>RUPTURE IMMINENTE</strong> — Stock sous le stock de sécurité ! Commander immédiatement.</div>", unsafe_allow_html=True)
        elif rp_current <= PR:
            st.markdown("<div class='alert-warning'>🟡 <strong>COMMANDER MAINTENANT</strong> — Le point de réapprovisionnement est atteint.</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='alert-ok'>🟢 <strong>Stock suffisant</strong> — Stock actuel ({fmtInt(rp_current)} u) au-dessus du PR ({fmtInt(PR)} u).</div>", unsafe_allow_html=True)

        cols = st.columns(3)
        rp_results = [
            ("Point de réappro. (PR)", fmtInt(PR), "unités", "#3b82f6"),
            ("Stock max théorique", fmtInt(PR_max), "unités", "#8892a4"),
            ("Stock actuel", fmtInt(rp_current), "unités", "#22c55e" if rp_current > PR else "#ef4444"),
            ("Jours de stock restants", fmt(days_left, 1), "jours (hors SS)", "#8892a4"),
            ("Délai critique restant", fmt(days_crit, 1), "jours avant urgence", "#f59e0b" if days_crit > 0 else "#ef4444"),
            ("Stock de sécurité", fmtInt(rp_ss), "unités", "#a78bfa"),
        ]
        for i, (label, val, unit, color) in enumerate(rp_results):
            with cols[i % 3]:
                st.markdown(metric_card(label, val, unit, color=color), unsafe_allow_html=True)
                st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

        # Jauge stock
        st.markdown("---")
        st.markdown("#### 📉 Niveau de stock actuel")
        max_display = max(PR_max * 1.2, rp_current * 1.2)
        fig4 = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=rp_current,
            delta={'reference': PR, 'valueformat': ',.0f'},
            number={'valueformat': ',.0f', 'font': {'color': '#e8eaf0', 'family': 'DM Mono'}},
            gauge={
                'axis': {'range': [0, max_display], 'tickcolor': '#5a6478', 'tickfont': {'color': '#5a6478'}},
                'bar': {'color': '#3b82f6', 'thickness': 0.25},
                'bgcolor': '#1e2535',
                'steps': [
                    {'range': [0, rp_ss], 'color': 'rgba(239,68,68,0.25)'},
                    {'range': [rp_ss, PR], 'color': 'rgba(245,158,11,0.2)'},
                    {'range': [PR, max_display], 'color': 'rgba(34,197,94,0.15)'},
                ],
                'threshold': {'line': {'color': '#ef4444', 'width': 3}, 'thickness': 0.8, 'value': PR},
            }
        ))
        fig4.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#8892a4', family='Syne'),
            height=260, margin=dict(t=10, b=10, l=30, r=30),
        )
        st.plotly_chart(fig4, use_container_width=True)
        st.caption(f"🔴 Rouge : sous stock sécu. ({fmtInt(rp_ss)} u)  |  🟡 Ambre : entre SS et PR ({fmtInt(PR)} u)  |  🟢 Vert : stock normal")

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : PROCESSUS FOURNISSEURS
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🏭 Processus fournisseurs":
    st.markdown("## 🏭 Évaluation des processus fournisseurs")
    st.markdown("<p style='color:#8892a4;'>Analysez la performance du processus de passation des commandes</p>", unsafe_allow_html=True)

    with st.expander("📋 Données du fournisseur", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            pf_name = st.text_input("Nom du fournisseur", value="Würth Tunisie SA")
        with col2:
            pf_ref  = st.text_input("Référence article", value="RB-7204")
        with col3:
            pf_cat  = st.selectbox("Catégorie", ["Composants mécaniques", "Matières premières", "Fournitures industrielles", "Sous-traitance"])

    st.markdown("#### Métriques de performance · 6 derniers mois")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        pf_total    = st.number_input("Commandes passées", value=24, min_value=1)
        pf_ontime   = st.number_input("Livrées à temps", value=20, min_value=0)
    with col2:
        pf_complete = st.number_input("Livrées complètes", value=22, min_value=0)
        pf_nc       = st.number_input("Non-conformités", value=2, min_value=0)
    with col3:
        pf_lt_a     = st.number_input("Délai annoncé (jours)", value=7.0, min_value=0.0, step=0.5)
        pf_lt_r     = st.number_input("Délai réel moyen (jours)", value=8.4, min_value=0.0, step=0.1)
    with col4:
        pf_proc     = st.number_input("Traitement commande (h)", value=18.0, min_value=0.0, step=1.0)
        pf_ack      = st.number_input("Réponse accusé réception (h)", value=4.0, min_value=0.0, step=0.5)

    # Calculs
    otd       = (pf_ontime / pf_total) * 100
    fill_rate = (pf_complete / pf_total) * 100
    quality   = ((pf_total - pf_nc) / pf_total) * 100
    lt_perf   = max(0, 100 - ((pf_lt_r - pf_lt_a) / pf_lt_a) * 100) if pf_lt_a > 0 else 0
    proc_s    = 100 if pf_proc <= 8 else 85 if pf_proc <= 24 else 60 if pf_proc <= 48 else 30
    ack_s     = 100 if pf_ack <= 2 else 85 if pf_ack <= 8 else 60 if pf_ack <= 24 else 30

    global_score = otd*0.30 + fill_rate*0.20 + quality*0.20 + lt_perf*0.15 + proc_s*0.10 + ack_s*0.05
    color  = score_color(global_score)
    label  = score_label(global_score)

    advice = {
        "Excellent": "Fournisseur stratégique. Maintenir le partenariat et envisager des accords-cadres.",
        "Bon": "Performance satisfaisante. Des améliorations mineures sur le délai réel peuvent optimiser la chaîne.",
        "Moyen": "Performance insuffisante. Plan d'amélioration fournisseur recommandé sous 30 jours.",
        "Insuffisant": "Performance critique. Envisager une diversification fournisseur ou un audit qualité immédiat.",
    }[label]

    st.markdown("---")
    st.markdown("#### 🎯 Score global")

    col_gauge, col_detail = st.columns([1, 2])
    with col_gauge:
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=round(global_score, 1),
            number={'suffix': '/100', 'font': {'color': color, 'family': 'DM Mono', 'size': 28}},
            gauge={
                'axis': {'range': [0, 100], 'tickcolor': '#5a6478', 'tickfont': {'color': '#5a6478'}},
                'bar': {'color': color, 'thickness': 0.25},
                'bgcolor': '#1e2535',
                'steps': [
                    {'range': [0, 55], 'color': 'rgba(239,68,68,0.2)'},
                    {'range': [55, 70], 'color': 'rgba(245,158,11,0.2)'},
                    {'range': [70, 85], 'color': 'rgba(20,184,166,0.2)'},
                    {'range': [85, 100], 'color': 'rgba(34,197,94,0.2)'},
                ],
            }
        ))
        fig_gauge.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', font=dict(color='#8892a4', family='Syne'),
            height=220, margin=dict(t=10, b=10, l=20, r=20),
        )
        st.plotly_chart(fig_gauge, use_container_width=True)

    with col_detail:
        st.markdown(f"""
        <div style='background:#1e2535; border-left:4px solid {color}; border-radius:10px; padding:1.2rem 1.4rem; margin-top:1rem;'>
            <div style='font-size:1.3rem; font-weight:600; color:{color}; margin-bottom:8px;'>{label} — {pf_name}</div>
            <div style='font-size:0.85rem; color:#8892a4; line-height:1.6;'>{advice}</div>
            <div style='margin-top:10px; font-size:0.75rem; color:#5a6478;'>
                Évaluation basée sur {pf_total} commandes · Article {pf_ref} · {pf_cat}
            </div>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 📋 Critères détaillés")

    criteria = [
        ("Livraison à temps (OTD)", otd, f"{pf_ontime}/{pf_total} commandes · Poids 30%"),
        ("Taux de service (fill rate)", fill_rate, f"{pf_complete}/{pf_total} complètes · Poids 20%"),
        ("Qualité / conformité", quality, f"{pf_total - pf_nc}/{pf_total} sans NC · Poids 20%"),
        ("Fiabilité délai annoncé", lt_perf, f"Annoncé: {pf_lt_a}j · Réel: {pf_lt_r}j · Poids 15%"),
        ("Délai traitement commande", proc_s, f"{pf_proc}h de traitement · Poids 10%"),
        ("Réactivité accusé réception", ack_s, f"{pf_ack}h pour ACK · Poids 5%"),
    ]

    col1, col2 = st.columns(2)
    for i, (name, score, note) in enumerate(criteria):
        c = score_color(score)
        with (col1 if i % 2 == 0 else col2):
            st.markdown(f"""
            <div style='background:#1e2535; border:1px solid rgba(255,255,255,0.08); border-radius:10px; padding:1rem; margin-bottom:10px;'>
                <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;'>
                    <span style='font-size:0.85rem; font-weight:600; color:#e8eaf0;'>{name}</span>
                    <span style='font-family:"DM Mono",monospace; font-size:0.9rem; color:{c};'>{score:.1f}%</span>
                </div>
                <div style='background:#0f1117; border-radius:4px; height:6px; overflow:hidden;'>
                    <div style='background:{c}; height:100%; width:{score}%; border-radius:4px;'></div>
                </div>
                <div style='font-size:0.7rem; color:#5a6478; margin-top:6px;'>{note}</div>
            </div>""", unsafe_allow_html=True)

    # Recommandations
    st.markdown("---")
    st.markdown("#### 💡 Recommandations")
    if otd < 85:
        st.markdown(f"<div class='alert-warning'>⏱ <strong>OTD insuffisant ({otd:.1f}%)</strong> — Négocier des pénalités de retard. Demander un plan de fiabilisation sur 90 jours. Majorer le stock de sécurité pour cet article.</div>", unsafe_allow_html=True)
    if quality < 95:
        st.markdown(f"<div class='alert-warning'>⚠️ <strong>Non-conformités détectées ({pf_nc} NC)</strong> — Déclencher un audit qualité fournisseur. Renforcer les contrôles à réception.</div>", unsafe_allow_html=True)
    if pf_lt_r > pf_lt_a:
        st.markdown(f"<div class='alert-warning'>📅 <strong>Dérive du délai réel (+{pf_lt_r - pf_lt_a:.1f} jours)</strong> — Recalculer le point de réappro. avec le délai réel. Demander une révision contractuelle.</div>", unsafe_allow_html=True)
    if pf_proc > 24:
        st.markdown(f"<div class='alert-info'>⚙️ <strong>Délai de traitement élevé ({pf_proc:.0f}h)</strong> — Proposer une commande électronique (EDI). Objectif : &lt; 8h.</div>", unsafe_allow_html=True)
    if global_score >= 85:
        st.markdown("<div class='alert-ok'>✅ <strong>Performance globale excellente</strong> — Envisager un accord-cadre annuel avec conditions préférentielles.</div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE : ALERTES STOCK
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🔔 Alertes stock":
    st.markdown("## 🔔 Alertes & Surveillance des stocks")
    st.markdown("<p style='color:#8892a4;'>Paramétrez vos articles et recevez des alertes automatiques sur vos niveaux de stock</p>", unsafe_allow_html=True)

    # Données de démo
    demo_data = {
        "Référence":        ["RB-7204", "IS-316L", "VD-0822", "CA-1140", "JT-2200", "MO-1115"],
        "Désignation":      ["Roulements à billes SKF", "Acier inox 316L (kg)", "Vannes pneumatiques", "Câble 4×2.5mm² (m)", "Joints toriques EPDM", "Moteur électrique 0.75kW"],
        "Stock actuel":     [45, 820, 380, 2400, 1250, 12],
        "Stock sécurité":   [132, 600, 80, 300, 400, 8],
        "Point réappro.":   [482, 1850, 230, 900, 1200, 35],
        "Conso. moy./j":    [50, 120, 15, 80, 95, 3],
        "Lead time (j)":    [7, 10, 12, 8, 5, 21],
    }

    col_btn1, col_btn2, _ = st.columns([1, 1, 3])
    with col_btn1:
        load_demo = st.button("📂 Charger données de démo", type="secondary")
    with col_btn2:
        st.markdown("")

    if "articles_df" not in st.session_state or load_demo:
        st.session_state.articles_df = pd.DataFrame(demo_data)

    st.markdown("#### ✏️ Tableau de bord des articles")
    st.caption("Modifiez directement les valeurs dans le tableau")

    edited_df = st.data_editor(
        st.session_state.articles_df,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Référence":      st.column_config.TextColumn("Référence", width="small"),
            "Désignation":    st.column_config.TextColumn("Désignation", width="medium"),
            "Stock actuel":   st.column_config.NumberColumn("Stock actuel", min_value=0, format="%d"),
            "Stock sécurité": st.column_config.NumberColumn("Stock sécu.", min_value=0, format="%d"),
            "Point réappro.": st.column_config.NumberColumn("Point réappro.", min_value=0, format="%d"),
            "Conso. moy./j":  st.column_config.NumberColumn("Conso./jour", min_value=0, format="%d"),
            "Lead time (j)":  st.column_config.NumberColumn("Lead time", min_value=0, format="%d"),
        },
        hide_index=True,
        key="articles_editor"
    )

    if st.button("🔍 Analyser les stocks", type="primary"):
        st.session_state.articles_df = edited_df
        df = edited_df.copy()

        # Calculs
        df["Jours restants"] = np.where(
            df["Conso. moy./j"] > 0,
            ((df["Stock actuel"] - df["Stock sécurité"]) / df["Conso. moy./j"]).round(1),
            np.nan
        )

        def get_status(row):
            if row["Stock actuel"] <= row["Stock sécurité"]: return "🔴 Rupture imminente"
            if row["Stock actuel"] <= row["Point réappro."]: return "🟡 Commander"
            return "🟢 Normal"

        def get_pct(row):
            return min(100, round(row["Stock actuel"] / row["Point réappro."] * 100)) if row["Point réappro."] > 0 else 0

        df["Statut"] = df.apply(get_status, axis=1)
        df["Couverture %"] = df.apply(get_pct, axis=1)

        critiques = df[df["Statut"] == "🔴 Rupture imminente"]
        attention = df[df["Statut"] == "🟡 Commander"]
        ok_stock  = df[df["Statut"] == "🟢 Normal"]

        st.markdown("---")
        st.markdown("#### 📊 Synthèse")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(metric_card("🔴 Rupture imminente", str(len(critiques)), "article(s) sous le stock de sécurité", color="#ef4444"), unsafe_allow_html=True)
        with c2:
            st.markdown(metric_card("🟡 Commander maintenant", str(len(attention)), "article(s) entre SS et point réappro.", color="#f59e0b"), unsafe_allow_html=True)
        with c3:
            st.markdown(metric_card("🟢 Stock normal", str(len(ok_stock)), "article(s) au-dessus du point de réappro.", color="#22c55e"), unsafe_allow_html=True)

        # Graphique
        st.markdown("---")
        st.markdown("#### 📈 Niveaux de stock vs seuils")
        fig_stock = go.Figure()
        fig_stock.add_trace(go.Bar(
            name="Stock actuel", x=df["Référence"], y=df["Stock actuel"],
            marker_color=[
                "#ef4444" if s == "🔴 Rupture imminente" else "#f59e0b" if s == "🟡 Commander" else "#22c55e"
                for s in df["Statut"]
            ],
            text=df["Stock actuel"], textposition='outside',
        ))
        fig_stock.add_trace(go.Scatter(
            name="Stock de sécurité", x=df["Référence"], y=df["Stock sécurité"],
            mode='markers+lines', marker=dict(color='#a78bfa', size=8, symbol='diamond'),
            line=dict(color='#a78bfa', width=1.5, dash='dash'),
        ))
        fig_stock.add_trace(go.Scatter(
            name="Point de réappro.", x=df["Référence"], y=df["Point réappro."],
            mode='markers+lines', marker=dict(color='#f59e0b', size=8, symbol='triangle-up'),
            line=dict(color='#f59e0b', width=1.5, dash='dot'),
        ))
        fig_stock.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#8892a4', family='Syne'),
            xaxis=dict(gridcolor='rgba(255,255,255,0.05)'),
            yaxis=dict(title="Unités", gridcolor='rgba(255,255,255,0.05)'),
            legend=dict(bgcolor='rgba(0,0,0,0)', bordercolor='rgba(255,255,255,0.1)', borderwidth=1),
            height=350, margin=dict(t=20, b=10, l=0, r=0),
            barmode='group',
        )
        st.plotly_chart(fig_stock, use_container_width=True)

        # Tableau résultat
        st.markdown("---")
        st.markdown("#### 📋 Détail par article")
        display_df = df[["Référence", "Désignation", "Stock actuel", "Stock sécurité", "Point réappro.", "Jours restants", "Statut"]].copy()
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Jours restants": st.column_config.NumberColumn("Jours restants", format="%.1f j"),
                "Stock actuel":   st.column_config.ProgressColumn("Stock actuel", min_value=0, max_value=int(df["Point réappro."].max() * 1.2), format="%d"),
            }
        )

        # Actions recommandées
        if len(critiques) > 0 or len(attention) > 0:
            st.markdown("---")
            st.markdown("#### 🚨 Actions recommandées")
            for _, row in critiques.iterrows():
                jours = f"{row['Jours restants']:.1f}" if not pd.isna(row["Jours restants"]) else "?"
                st.markdown(f"""<div class='alert-critical'>
                    🔴 <strong>{row['Référence']} — {row['Désignation']} : RUPTURE IMMINENTE</strong><br/>
                    <span style='font-size:0.82rem;'>Stock actuel <strong>{int(row['Stock actuel'])} u</strong> sous le stock de sécurité ({int(row['Stock sécurité'])} u).
                    Environ <strong>{jours} jours</strong> avant arrêt. Passer commande urgente. Lead time : {int(row['Lead time (j)'])} jours.</span>
                </div>""", unsafe_allow_html=True)

            for _, row in attention.iterrows():
                st.markdown(f"""<div class='alert-warning'>
                    🟡 <strong>{row['Référence']} — {row['Désignation']} : Commander maintenant</strong><br/>
                    <span style='font-size:0.82rem;'>Stock actuel <strong>{int(row['Stock actuel'])} u</strong> a atteint le point de réappro. ({int(row['Point réappro.'])} u).
                    Lancer une commande. Délai : {int(row['Lead time (j)'])} jours.</span>
                </div>""", unsafe_allow_html=True)
