import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io, re

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Tablero de Control — Farmacia",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.main { background: #f0f4f8; }

.metric-card {
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    border-left: 4px solid #2563eb;
    box-shadow: 0 1px 3px rgba(0,0,0,.08);
    margin-bottom: 12px;
}
.metric-card.green  { border-left-color: #16a34a; }
.metric-card.orange { border-left-color: #ea580c; }
.metric-card.red    { border-left-color: #dc2626; }
.metric-card.purple { border-left-color: #7c3aed; }
.metric-card .value { font-size: 2rem; font-weight: 700; color: #1e293b; line-height:1.1; }
.metric-card .label { font-size: .78rem; font-weight: 500; color: #64748b; text-transform: uppercase; letter-spacing:.05em; margin-bottom:4px; }
.metric-card .delta { font-size: .82rem; color: #64748b; margin-top: 4px; }

.section-header {
    font-size: 1.05rem; font-weight: 600; color: #1e293b;
    padding: 10px 0 6px; border-bottom: 2px solid #e2e8f0; margin-bottom: 14px;
}

.badge {
    display: inline-block; padding: 2px 10px; border-radius: 999px;
    font-size: .75rem; font-weight: 600; margin: 2px;
}
.badge-green  { background:#dcfce7; color:#16a34a; }
.badge-yellow { background:#fef9c3; color:#b45309; }
.badge-red    { background:#fee2e2; color:#dc2626; }
.badge-blue   { background:#dbeafe; color:#1d4ed8; }

div[data-testid="stSidebar"] { background: #1e293b; }
div[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
div[data-testid="stSidebar"] .stSelectbox label,
div[data-testid="stSidebar"] .stFileUploader label { color: #94a3b8 !important; font-size:.8rem; }

.upload-prompt {
    background: white; border-radius: 16px; padding: 60px 40px;
    text-align: center; border: 2px dashed #cbd5e1;
    box-shadow: 0 4px 24px rgba(0,0,0,.06);
}
.upload-prompt h2 { color: #1e293b; font-size: 1.6rem; margin-bottom: 8px; }
.upload-prompt p  { color: #64748b; font-size: .95rem; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
RUBROS_CFG = [
    ("ESTERILIZACION",3,53),("UTI",3,53),("ANESTESICOS",3,21),
    ("COMPRIMIDOS",3,80),("ANTIBIOTICOS",3,79),("PSICOTROPICOS",3,79),
    ("AMPOLLAS",3,106),("SUEROS",3,27),("DROGAS Y LIQUIDOS",3,38),
    ("ALIMENTACION",3,31),("SIES SAT URS",3,45),("BAJO COSTO",3,213),
    ("CREMAS GOTAS AEROSOLES",3,68),("DIALISIS",3,27),("RAYOS",3,33),
    ("FORMULACION Y SEDRONAR",3,32),("VARIOS",3,9),
    ("DESIERTOS",3,178),("ESTERILIZACION DESIERTOS",3,13),
]

RUBRO_COL = {
    "VARIOS": {"cm":"Y","cq":"R","cj":"W","cp":"M"},
    "DEFAULT":{"cm":"AA","cq":"R","cj":"Y","cp":"M"},
}


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_solicitudes(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    rows = []
    for sheet_name, r1, r2 in RUBROS_CFG:
        if sheet_name not in xl.sheet_names:
            continue
        df = xl.parse(sheet_name, header=None)
        cfg = RUBRO_COL.get(sheet_name, RUBRO_COL["DEFAULT"])
        cm_i = {"AA":26,"Y":24}[cfg["cm"]]
        cq_i = ord(cfg["cq"])-65
        cj_i = {"Y":24,"W":22}.get(cfg["cj"], ord(cfg["cj"])-65)
        cp_i = ord(cfg["cp"])-65

        for pi in range(r1-1, r2):
            if pi >= len(df): continue
            row = df.iloc[pi]
            desc = row[1] if pd.notna(row[1]) else None
            if not desc: continue
            codigo = str(row[0]).strip() if pd.notna(row[0]) else ""
            if not codigo or codigo == "nan": continue

            def sf(v):
                try: return float(v) if pd.notna(v) else 0.0
                except: return 0.0

            m = sf(row[cm_i]) if cm_i < len(row) else 0
            q = sf(row[cq_i]) if cq_i < len(row) else 0
            j = sf(row[cj_i]) if cj_i < len(row) else 0
            if m == 0 and q > 0 and j > 0: m = q * j
            prov = str(row[cp_i]).strip() if cp_i < len(row) and pd.notna(row[cp_i]) else ""
            oc   = str(row[15]).strip() if len(row)>15 and pd.notna(row[15]) else ""
            cons = sf(row[5]) if len(row)>5 else 0

            # Historical consumption for XYZ
            consums = [sf(row[c]) for c in [2,3,4,5] if c < len(row) and sf(row[c]) > 0]

            is_des  = prov.lower() in ["desierto","desierta"]
            has_oc  = bool(oc) and oc.lower() not in ["nan",""]
            rows.append({
                "Rubro":sheet_name, "Codigo":codigo, "Descripcion":str(desc)[:80],
                "Proveedor":prov, "Cantidad":q, "Justiprecio":j, "Monto":m,
                "Consumo_Mensual":cons, "Es_Desierto":is_des, "Tiene_OC":has_oc,
                "Consums":consums,
            })
    return pd.DataFrame(rows)


def compute_abc(df_in, value_col, prefix):
    df = df_in[df_in[value_col]>0].copy().sort_values(value_col, ascending=False)
    total = df[value_col].sum()
    if total == 0:
        df[f"{prefix}_pct_acum"] = 0
        df[f"{prefix}_abc"] = "C"
        return df
    df[f"{prefix}_pct_acum"] = df[value_col].cumsum() / total * 100
    df[f"{prefix}_abc"] = df[f"{prefix}_pct_acum"].apply(
        lambda x: "A" if x<=80 else ("B" if x<=95 else "C"))
    return df


def compute_xyz(consums):
    if len(consums) < 2: return "Z"
    arr = np.array(consums)
    m = arr.mean()
    if m == 0: return "Z"
    cv = arr.std() / m
    return "X" if cv<=0.5 else ("Y" if cv<=1.0 else "Z")


@st.cache_data(show_spinner=False)
def process(file_bytes):
    df = load_solicitudes(file_bytes)
    df = compute_abc(df, "Monto",    "M")
    df = compute_abc(df, "Cantidad", "Q")
    df["XYZ"] = df["Consums"].apply(compute_xyz)
    df["ABC_XYZ"] = df.get("M_abc","?") + df["XYZ"]
    prov = (df[~df["Proveedor"].str.lower().isin(["desierto","desierta","","nan"])]
            .groupby("Proveedor")
            .agg(Monto=("Monto","sum"), Items=("Descripcion","count"))
            .sort_values("Monto", ascending=False)
            .reset_index())
    total_prov = prov["Monto"].sum()
    prov["Pct"] = prov["Monto"]/total_prov*100
    prov["Pct_Acum"] = prov["Pct"].cumsum()
    return df, prov


# ─────────────────────────────────────────────────────────────────────────────
# CHART HELPERS
# ─────────────────────────────────────────────────────────────────────────────
PALETTE = {
    "A":"#16a34a","B":"#f59e0b","C":"#ef4444",
    "X":"#2563eb","Y":"#f59e0b","Z":"#94a3b8",
}
CHART_LAYOUT = dict(
    font_family="DM Sans",
    paper_bgcolor="white", plot_bgcolor="white",
    margin=dict(l=10,r=10,t=36,b=10),
)
LEGEND_H = dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)


def kpi(label, value, delta="", color="blue"):
    cls = {"blue":"","green":"green","orange":"orange","red":"red","purple":"purple"}.get(color,"")
    st.markdown(f"""
    <div class="metric-card {cls}">
        <div class="label">{label}</div>
        <div class="value">{value}</div>
        {'<div class="delta">'+delta+'</div>' if delta else ''}
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💊 Farmacia\n### Tablero de Control")
    st.markdown("---")

    uploaded = st.file_uploader(
        "Subir SOLICITUDES_2026.xlsx",
        type=["xlsx"],
        help="El mismo archivo de solicitudes que usaste antes.",
    )

    st.markdown("---")
    st.markdown("**Filtros globales**")

    rubros_all = [r[0] for r in RUBROS_CFG]

    if uploaded:
        df_raw, prov_df = process(uploaded.read())
        rubros_disp = sorted(df_raw["Rubro"].unique().tolist())
    else:
        rubros_disp = rubros_all

    rubro_sel = st.multiselect(
        "Rubros", rubros_disp,
        default=rubros_disp,
        help="Filtrá por rubro farmacéutico.",
    )
    show_desiertos = st.checkbox("Incluir desiertos", value=True)

    st.markdown("---")
    st.caption("Hospital Interzonal · 2026")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
if not uploaded:
    st.markdown("""
    <div class="upload-prompt">
        <h2>📊 Tablero de Control Farmacia</h2>
        <p>Subí el archivo <strong>SOLICITUDES_2026.xlsx</strong> desde la barra lateral para comenzar.</p>
        <br>
        <p style="color:#94a3b8;font-size:.85rem">
            Incluye: ABC por monto · ABC por cantidad · Análisis XYZ · Concentración de proveedores ·<br>
            Ejecución de presupuesto · Matriz ABC-XYZ
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Apply filters
df = df_raw[df_raw["Rubro"].isin(rubro_sel)].copy()
if not show_desiertos:
    df = df[~df["Es_Desierto"]]

total_monto  = df["Monto"].sum()
total_items  = len(df)
n_desiertos  = int(df["Es_Desierto"].sum())
n_con_oc     = int(df["Tiene_OC"].sum())
pct_ejec     = n_con_oc / total_items * 100 if total_items > 0 else 0

# ── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Resumen General",
    "🔢 ABC por Monto",
    "🔄 ABC por Cantidad + XYZ",
    "🏭 Proveedores",
    "🏥 Detalle por Rubro",
    "📋 Datos",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — RESUMEN GENERAL
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">Indicadores Globales</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi("Total Medicamentos",  f"{total_items:,}",    color="blue")
    with c2: kpi("Monto Total Justiprecio", f"${total_monto/1e9:.2f}MM", f"{len(rubro_sel)} rubros", color="green")
    with c3: kpi("Ítems Desiertos",     f"{n_desiertos:,}", f"{n_desiertos/total_items*100:.1f}% del total", color="orange")
    with c4: kpi("Con OC Confirmada",   f"{n_con_oc:,}",    f"{pct_ejec:.1f}% ejecución", color="purple")
    with c5: kpi("Sin OC (pendiente)",  f"{total_items-n_con_oc:,}", color="red")

    st.markdown('<div class="section-header" style="margin-top:20px">Monto por Rubro</div>', unsafe_allow_html=True)

    rubro_agg = (df.groupby("Rubro")
                 .agg(Monto=("Monto","sum"), Items=("Descripcion","count"),
                      Desiertos=("Es_Desierto","sum"))
                 .sort_values("Monto", ascending=True)
                 .reset_index())

    col_l, col_r = st.columns([3,2])

    with col_l:
        fig = px.bar(
            rubro_agg, x="Monto", y="Rubro", orientation="h",
            color="Monto",
            color_continuous_scale=[[0,"#bfdbfe"],[1,"#1d4ed8"]],
            text=rubro_agg["Monto"].apply(lambda v: f"${v/1e6:.1f}M"),
            labels={"Monto":"Monto ($)","Rubro":""},
            title="Monto Justiprecio por Rubro",
        )
        fig.update_traces(textposition="outside", textfont_size=10)
        fig.update_layout(**CHART_LAYOUT, height=460,
                          coloraxis_showscale=False,
                          legend=LEGEND_H,
                          xaxis=dict(showgrid=False,visible=False),
                          yaxis=dict(tickfont_size=11))
        st.plotly_chart(fig, use_container_width=True)

    with col_r:
        des_agg = (df.groupby("Rubro")["Es_Desierto"].agg(["sum","count"])
                   .rename(columns={"sum":"Desiertos","count":"Total"})
                   .reset_index()
                   .sort_values("Desiertos", ascending=False)
                   .head(12))
        des_agg["Adjudicados"] = des_agg["Total"] - des_agg["Desiertos"]

        fig2 = go.Figure()
        fig2.add_bar(x=des_agg["Rubro"], y=des_agg["Adjudicados"],
                     name="Adjudicados", marker_color="#4ade80")
        fig2.add_bar(x=des_agg["Rubro"], y=des_agg["Desiertos"],
                     name="Desiertos",   marker_color="#fb923c")
        fig2.update_layout(**CHART_LAYOUT, barmode="stack", height=460,
                           title="Ítems: Adjudicados vs Desiertos",
                           legend=LEGEND_H,
                           xaxis=dict(tickangle=-40, tickfont_size=10),
                           yaxis=dict(title="Cantidad"))
        st.plotly_chart(fig2, use_container_width=True)

    # Execution donut
    st.markdown('<div class="section-header">Ejecución del Presupuesto</div>', unsafe_allow_html=True)

    ejec_rubro = (df.groupby("Rubro")
                  .apply(lambda x: pd.Series({
                      "Monto_Total": x["Monto"].sum(),
                      "Monto_OC": x[x["Tiene_OC"]]["Monto"].sum(),
                  }))
                  .reset_index())
    ejec_rubro["Pct_Ejec"] = (ejec_rubro["Monto_OC"] / ejec_rubro["Monto_Total"] * 100).fillna(0)
    ejec_rubro = ejec_rubro.sort_values("Pct_Ejec", ascending=True)

    fig3 = px.bar(
        ejec_rubro, x="Pct_Ejec", y="Rubro", orientation="h",
        color="Pct_Ejec",
        color_continuous_scale=[[0,"#fca5a5"],[0.5,"#fde68a"],[1,"#86efac"]],
        range_color=[0,100],
        text=ejec_rubro["Pct_Ejec"].apply(lambda v: f"{v:.0f}%"),
        title="% Ejecución Presupuestal por Rubro (Monto con OC / Monto Total)",
    )
    fig3.update_traces(textposition="outside", textfont_size=10)
    fig3.update_layout(**CHART_LAYOUT, height=420, coloraxis_showscale=False,
                       xaxis=dict(range=[0,115], title="% Ejecución", showgrid=True,
                                  gridcolor="#f1f5f9"),
                       yaxis=dict(tickfont_size=11))
    fig3.add_vline(x=80, line_dash="dash", line_color="#dc2626",
                   annotation_text="80%", annotation_position="top right")
    st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ABC POR MONTO
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    df_abc = df[df["Monto"]>0].copy()
    if "M_abc" not in df_abc.columns:
        df_abc = compute_abc(df_abc, "Monto", "M")

    abc_sum = (df_abc.groupby("M_abc")
               .agg(Items=("Descripcion","count"), Monto=("Monto","sum"))
               .reindex(["A","B","C"]).fillna(0).reset_index())
    abc_sum["Pct_Items"] = abc_sum["Items"]/abc_sum["Items"].sum()*100
    abc_sum["Pct_Monto"] = abc_sum["Monto"]/abc_sum["Monto"].sum()*100

    st.markdown('<div class="section-header">Clasificación ABC por Monto Justiprecio</div>',
                unsafe_allow_html=True)

    # Summary cards
    for _, r in abc_sum.iterrows():
        cls = r["M_abc"]
        color_map = {"A":"green","B":"orange","C":"red"}
        desc_map  = {"A":"Gestión PRIORITARIA — stock alto, seguimiento semanal",
                     "B":"Control PERIÓDICO — revisión mensual",
                     "C":"Control SIMPLIFICADO — pedido por demanda"}
        kpi(f"Clase {cls} — {desc_map[cls]}",
            f"{int(r['Items'])} ítems ({r['Pct_Items']:.0f}%)",
            f"${r['Monto']/1e6:.1f}M · {r['Pct_Monto']:.0f}% del gasto",
            color=color_map[cls])

    col_l, col_r = st.columns([2,1])

    with col_l:
        # Pareto curve
        top80 = df_abc.sort_values("Monto", ascending=False).head(80).reset_index(drop=True)
        top80["idx"] = range(1, len(top80)+1)
        fig_p = go.Figure()
        # Bars colored by class
        colors = top80["M_abc"].map(PALETTE).tolist()
        fig_p.add_bar(x=top80["idx"], y=top80["Monto"],
                      marker_color=colors, name="Monto",
                      hovertemplate="<b>%{customdata}</b><br>$%{y:,.0f}<extra></extra>",
                      customdata=top80["Descripcion"].str[:40])
        # Pareto line
        pct_acum = top80["Monto"].cumsum() / df_abc["Monto"].sum() * 100
        fig_p.add_scatter(x=top80["idx"], y=pct_acum, mode="lines",
                          yaxis="y2", name="% Acumulado",
                          line=dict(color="#7c3aed", width=2))
        fig_p.add_hline(y=80, line_dash="dash", line_color="#dc2626",
                        annotation_text="80%", yref="y2")
        layout_p = {**CHART_LAYOUT,
            "height": 380,
            "title": "Curva de Pareto — Top 80 ítems por monto",
            "yaxis":  dict(title="Monto ($)", showgrid=False),
            "yaxis2": dict(title="% Acumulado", overlaying="y", side="right",
                           range=[0,105], ticksuffix="%"),
            "xaxis":  dict(title="Ítem (rank)", showgrid=False),
            "legend": dict(orientation="h"),
        }
        fig_p.update_layout(layout_p)
        st.plotly_chart(fig_p, use_container_width=True)

    with col_r:
        fig_pie = px.pie(
            abc_sum, names="M_abc", values="Monto",
            color="M_abc",
            color_discrete_map=PALETTE,
            hole=0.55,
            title="Distribución por Monto",
        )
        fig_pie.update_traces(textinfo="percent+label", textfont_size=13)
        fig_pie.update_layout(**CHART_LAYOUT, height=380,
                              showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)

    # Top 30 table
    st.markdown('<div class="section-header">Top 30 — Mayor Monto</div>', unsafe_allow_html=True)
    top30 = (df_abc.sort_values("Monto", ascending=False)
             .head(30)[["M_abc","Rubro","Codigo","Descripcion","Cantidad","Justiprecio","Monto","Proveedor"]]
             .rename(columns={"M_abc":"Clase","Monto":"Monto ($)"})
             .reset_index(drop=True))
    top30.index = top30.index + 1
    top30["Monto ($)"] = top30["Monto ($)"].apply(lambda v: f"${v:,.0f}")
    top30["Justiprecio"] = top30["Justiprecio"].apply(lambda v: f"${v:,.2f}")
    st.dataframe(top30, use_container_width=True, height=480)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — ABC CANTIDAD + XYZ
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    df_qty = df[df["Cantidad"]>0].copy()
    if "Q_abc" not in df_qty.columns:
        df_qty = compute_abc(df_qty, "Cantidad", "Q")

    col_l, col_r = st.columns(2)

    with col_l:
        st.markdown('<div class="section-header">ABC por Cantidad Solicitada</div>',
                    unsafe_allow_html=True)
        qty_sum = (df_qty.groupby("Q_abc")
                   .agg(Items=("Descripcion","count"), Unidades=("Cantidad","sum"))
                   .reindex(["A","B","C"]).fillna(0).reset_index())
        qty_sum["Pct_Items"]    = qty_sum["Items"]   /qty_sum["Items"].sum()*100
        qty_sum["Pct_Unidades"] = qty_sum["Unidades"]/qty_sum["Unidades"].sum()*100

        fig_qpie = px.pie(qty_sum, names="Q_abc", values="Unidades",
                          color="Q_abc", color_discrete_map=PALETTE,
                          hole=0.5, title="Unidades por Clase ABC")
        fig_qpie.update_traces(textinfo="percent+label")
        fig_qpie.update_layout(**CHART_LAYOUT, height=320, showlegend=False)
        st.plotly_chart(fig_qpie, use_container_width=True)

        st.dataframe(
            qty_sum.rename(columns={"Q_abc":"Clase","Unidades":"Total Unidades"})
                   .assign(**{"% Unidades":qty_sum["Pct_Unidades"].apply(lambda v:f"{v:.1f}%"),
                               "% Ítems":qty_sum["Pct_Items"].apply(lambda v:f"{v:.1f}%")})
                   [["Clase","Items","% Ítems","Total Unidades","% Unidades"]],
            use_container_width=True, hide_index=True,
        )

    with col_r:
        st.markdown('<div class="section-header">Análisis XYZ — Variabilidad de Demanda</div>',
                    unsafe_allow_html=True)
        xyz_sum = (df.groupby("XYZ")
                   .agg(Items=("Descripcion","count"), Monto=("Monto","sum"))
                   .reset_index())
        xyz_color = {"X":"#2563eb","Y":"#f59e0b","Z":"#94a3b8"}
        fig_xyz = px.bar(xyz_sum, x="XYZ", y="Items",
                         color="XYZ", color_discrete_map=xyz_color,
                         text="Items", title="Ítems por Clase XYZ")
        fig_xyz.update_traces(textposition="outside")
        fig_xyz.update_layout(**CHART_LAYOUT, height=200,
                              showlegend=False, xaxis_title="", yaxis_title="Ítems",
                              yaxis=dict(showgrid=False))
        st.plotly_chart(fig_xyz, use_container_width=True)

        xyz_desc = {
            "X":"Demanda ESTABLE (CV ≤ 0.5) — Stock fijo, reposición continua",
            "Y":"Demanda VARIABLE (CV 0.5-1.0) — Stock buffer, revisar mensual",
            "Z":"Demanda IRREGULAR o sin historial — Stock mínimo, pedido por demanda",
        }
        for k, desc in xyz_desc.items():
            n = int(xyz_sum[xyz_sum["XYZ"]==k]["Items"].sum()) if k in xyz_sum["XYZ"].values else 0
            c = {"X":"blue","Y":"orange","Z":"red"}[k]
            kpi(f"Clase {k}", f"{n} ítems", desc, color=c)

    # ABC-XYZ matrix
    st.markdown('<div class="section-header">Matriz ABC-XYZ</div>', unsafe_allow_html=True)

    df_m_full = df[df["Monto"]>0].copy()
    if "M_abc" not in df_m_full.columns:
        df_m_full = compute_abc(df_m_full, "Monto", "M")

    matrix_cnt = (df_m_full.groupby(["M_abc","XYZ"])
                  .agg(Items=("Descripcion","count"), Monto=("Monto","sum"))
                  .reset_index())

    strat = {
        "AX":"Stock alto\nProveedor fijo\nReposición continua",
        "AY":"Stock buffer\nRevisar mensual\nAnalizar causas",
        "AZ":"⚠ CRÍTICO\nAlto gasto + errática\nPlan contingencia",
        "BX":"Stock normal\nControl periódico",
        "BY":"Stock moderado\nRevisar trimestral",
        "BZ":"Stock mínimo\nAnalizar variabilidad",
        "CX":"Stock bajo\nPedido por demanda",
        "CY":"Sin stock fijo\nControl mínimo",
        "CZ":"Evaluar eliminar\nPedido esporádico",
    }

    bg_colors = {
        "AX":"#dcfce7","AY":"#fef9c3","AZ":"#fee2e2",
        "BX":"#dcfce7","BY":"#fef9c3","BZ":"#fef9c3",
        "CX":"#f0fdf4","CY":"#f0fdf4","CZ":"#f8fafc",
    }

    mcols = st.columns(4)
    mcols[0].markdown("**ABC \\ XYZ**")
    for j, xyz in enumerate(["X","Y","Z"]):
        mcols[j+1].markdown(f"**{xyz}**")

    for abc in ["A","B","C"]:
        row_cols = st.columns(4)
        row_cols[0].markdown(f"**Clase {abc}**")
        for j, xyz in enumerate(["X","Y","Z"]):
            key = abc+xyz
            sub = matrix_cnt[(matrix_cnt["M_abc"]==abc)&(matrix_cnt["XYZ"]==xyz)]
            cnt   = int(sub["Items"].sum()) if len(sub)>0 else 0
            monto = sub["Monto"].sum() if len(sub)>0 else 0
            bg = bg_colors.get(key,"#ffffff")
            row_cols[j+1].markdown(
                f'<div style="background:{bg};border-radius:8px;padding:10px;text-align:center;'
                f'border:1px solid #e2e8f0;font-size:.82rem">'
                f'<b style="font-size:1.1rem">{cnt}</b> ítems<br>'
                f'<span style="color:#64748b">${monto/1e6:.1f}M</span><br>'
                f'<span style="font-size:.72rem;color:#475569">{strat.get(key,"")}</span></div>',
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PROVEEDORES
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">Concentración de Proveedores</div>',
                unsafe_allow_html=True)

    prov_filt = (df[~df["Proveedor"].str.lower().isin(["desierto","desierta","","nan"])]
                 .groupby("Proveedor")
                 .agg(Monto=("Monto","sum"), Items=("Descripcion","count"))
                 .sort_values("Monto", ascending=False)
                 .reset_index())
    total_p = prov_filt["Monto"].sum()
    prov_filt["Pct"]     = prov_filt["Monto"]/total_p*100
    prov_filt["Pct_Acum"]= prov_filt["Pct"].cumsum()
    hhi = ((prov_filt["Pct"]/100)**2).sum()

    kp1, kp2, kp3, kp4 = st.columns(4)
    with kp1: kpi("Total Proveedores", f"{len(prov_filt)}", color="blue")
    with kp2: kpi("Top 5 % del Gasto", f"{prov_filt.head(5)['Pct'].sum():.1f}%",
                  "⚠ Riesgo alto" if prov_filt.head(5)['Pct'].sum()>70 else "Moderado",
                  color="red" if prov_filt.head(5)['Pct'].sum()>70 else "orange")
    with kp3: kpi("Top 10 % del Gasto",f"{prov_filt.head(10)['Pct'].sum():.1f}%", color="orange")
    with kp4: kpi("Índice HHI", f"{hhi:.4f}",
                  "Alta concentración" if hhi>0.25 else ("Moderada" if hhi>0.15 else "Baja"),
                  color="red" if hhi>0.25 else ("orange" if hhi>0.15 else "green"))

    col_l, col_r = st.columns([2,1])

    with col_l:
        top20 = prov_filt.head(20).copy()
        top20["Risk"] = top20["Pct"].apply(
            lambda v: "🔴 Crítico" if v>=20 else ("🟠 Alto" if v>=10 else ("🟡 Medio" if v>=5 else "🟢 Bajo")))

        fig_prov = go.Figure()
        fig_prov.add_bar(
            x=top20["Monto"], y=top20["Proveedor"], orientation="h",
            marker_color=top20["Pct"].apply(
                lambda v: "#dc2626" if v>=20 else ("#f97316" if v>=10 else ("#f59e0b" if v>=5 else "#4ade80"))),
            text=top20["Pct"].apply(lambda v: f"{v:.1f}%"),
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>$%{x:,.0f}<extra></extra>",
        )
        # Pareto line
        fig_prov.add_scatter(
            x=top20["Pct_Acum"]/100 * top20["Monto"].max(),
            y=top20["Proveedor"], mode="lines+markers",
            name="% Acum", line=dict(color="#7c3aed", width=2),
            xaxis="x2",
        )
        layout_prov = {**CHART_LAYOUT,
            "height": 500,
            "title": "Top 20 Proveedores por Monto",
            "xaxis":  dict(title="Monto ($)", showgrid=False),
            "xaxis2": dict(title="% Acumulado", overlaying="x", side="top",
                           range=[0,110], ticksuffix="%"),
            "yaxis":  dict(tickfont_size=10, autorange="reversed"),
            "showlegend": False,
        }
        fig_prov.update_layout(layout_prov)
        st.plotly_chart(fig_prov, use_container_width=True)

    with col_r:
        top5 = prov_filt.head(5)
        others_monto = prov_filt.iloc[5:]["Monto"].sum()
        pie_data = pd.concat([
            top5[["Proveedor","Monto"]],
            pd.DataFrame([{"Proveedor":"Otros","Monto":others_monto}])
        ])
        fig_pie2 = px.pie(pie_data, names="Proveedor", values="Monto",
                          hole=0.45, title="Top 5 vs Resto",
                          color_discrete_sequence=px.colors.qualitative.Bold)
        fig_pie2.update_traces(textinfo="percent+label", textfont_size=11)
        fig_pie2.update_layout(**CHART_LAYOUT, height=500, showlegend=False)
        st.plotly_chart(fig_pie2, use_container_width=True)

    st.markdown('<div class="section-header">Detalle de Proveedores</div>', unsafe_allow_html=True)
    prov_table = prov_filt.head(30).copy()
    prov_table["Monto"] = prov_table["Monto"].apply(lambda v: f"${v:,.0f}")
    prov_table["Pct"]   = prov_table["Pct"].apply(lambda v: f"{v:.2f}%")
    prov_table["Pct_Acum"]= prov_table["Pct_Acum"].apply(lambda v: f"{v:.1f}%")
    prov_table.index = range(1, len(prov_table)+1)
    st.dataframe(prov_table[["Proveedor","Monto","Pct","Pct_Acum","Items"]],
                 use_container_width=True, height=420)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — DETALLE POR RUBRO
# ══════════════════════════════════════════════════════════════════════════════
with tab5:

    # ── Selector de rubro ────────────────────────────────────────────────────
    rubros_disponibles = sorted(df["Rubro"].unique().tolist())
    rubro_elegido = st.selectbox(
        "Seleccioná un rubro para ver el análisis completo:",
        rubros_disponibles,
        key="rubro_sel_detalle",
    )

    df_r = df[df["Rubro"] == rubro_elegido].copy()
    df_r_m = df_r[df_r["Monto"] > 0].copy()
    if "M_abc" not in df_r_m.columns:
        df_r_m = compute_abc(df_r_m, "Monto", "M")
    if "Q_abc" not in df_r.columns:
        df_r_q = compute_abc(df_r[df_r["Cantidad"] > 0].copy(), "Cantidad", "Q")
    else:
        df_r_q = df_r[df_r["Cantidad"] > 0].copy()

    total_r       = len(df_r)
    monto_r       = df_r["Monto"].sum()
    des_r         = int(df_r["Es_Desierto"].sum())
    oc_r          = int(df_r["Tiene_OC"].sum())
    pct_ejec_r    = oc_r / total_r * 100 if total_r > 0 else 0
    monto_oc_r    = df_r[df_r["Tiene_OC"]]["Monto"].sum()
    monto_sin_oc  = monto_r - monto_oc_r

    # ── KPIs del rubro ────────────────────────────────────────────────────────
    st.markdown(f'<div class="section-header">📌 {rubro_elegido}</div>',
                unsafe_allow_html=True)

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1: kpi("Total ítems",      f"{total_r}",            color="blue")
    with k2: kpi("Monto justiprecio",f"${monto_r/1e6:.2f}M", color="green")
    with k3: kpi("Con OC",           f"{oc_r}",
                 f"{pct_ejec_r:.1f}% ejec.", color="purple")
    with k4: kpi("Sin OC",           f"{total_r - oc_r}",    color="orange")
    with k5: kpi("Desiertos",        f"{des_r}",
                 f"{des_r/total_r*100:.1f}%" if total_r > 0 else "",  color="red")
    with k6: kpi("Proveedores únicos",
                 str(df_r[~df_r["Proveedor"].str.lower().isin(
                     ["desierto","desierta","","nan"])]["Proveedor"].nunique()),
                 color="blue")

    st.markdown("---")

    # ── Fila 1: Monto por proveedor + Ejecución ───────────────────────────────
    c_l, c_r = st.columns(2)

    with c_l:
        st.markdown('<div class="section-header">Monto por proveedor</div>',
                    unsafe_allow_html=True)
        prov_r = (df_r[~df_r["Proveedor"].str.lower().isin(
                        ["desierto","desierta","","nan"])]
                  .groupby("Proveedor")
                  .agg(Monto=("Monto","sum"), Items=("Descripcion","count"))
                  .sort_values("Monto", ascending=True)
                  .reset_index())

        if len(prov_r) > 0:
            fig_pr = px.bar(
                prov_r, x="Monto", y="Proveedor", orientation="h",
                color="Monto",
                color_continuous_scale=[[0,"#bfdbfe"],[1,"#1d4ed8"]],
                text=prov_r["Monto"].apply(lambda v: f"${v/1e6:.2f}M" if v >= 1e6 else f"${v:,.0f}"),
                labels={"Monto":"","Proveedor":""},
                height=max(250, len(prov_r) * 32),
            )
            fig_pr.update_traces(textposition="outside", textfont_size=10)
            fig_pr.update_layout({**CHART_LAYOUT,
                "coloraxis_showscale": False,
                "xaxis": dict(showgrid=False, visible=False),
                "yaxis": dict(tickfont_size=10),
                "margin": dict(l=10, r=60, t=10, b=10),
            })
            st.plotly_chart(fig_pr, use_container_width=True)
        else:
            st.info("Sin proveedores adjudicados en este rubro.")

    with c_r:
        st.markdown('<div class="section-header">Ejecución presupuestal</div>',
                    unsafe_allow_html=True)

        # Donut ejecución
        fig_ej = go.Figure(go.Pie(
            labels=["Con OC", "Sin OC"],
            values=[monto_oc_r, max(0, monto_sin_oc)],
            hole=0.6,
            marker_colors=["#16a34a", "#e2e8f0"],
            textinfo="percent+label",
            textfont_size=12,
        ))
        fig_ej.update_layout({**CHART_LAYOUT,
            "height": 260,
            "showlegend": False,
            "annotations": [dict(
                text=f"{pct_ejec_r:.0f}%<br><span style='font-size:11px'>ejecución</span>",
                x=0.5, y=0.5, font_size=20, showarrow=False,
            )],
            "margin": dict(l=10, r=10, t=10, b=10),
        })
        st.plotly_chart(fig_ej, use_container_width=True)

        # Barra adjudicados vs desiertos
        fig_adj = go.Figure()
        fig_adj.add_bar(name="Adjudicados", x=["Ítems"], y=[total_r - des_r],
                        marker_color="#4ade80", text=[total_r - des_r],
                        textposition="inside")
        fig_adj.add_bar(name="Desiertos",   x=["Ítems"], y=[des_r],
                        marker_color="#fb923c", text=[des_r],
                        textposition="inside")
        fig_adj.update_layout({**CHART_LAYOUT,
            "barmode": "stack", "height": 110,
            "showlegend": True,
            "legend": dict(orientation="h", y=1.2),
            "xaxis": dict(visible=False),
            "yaxis": dict(visible=False),
            "margin": dict(l=10, r=10, t=30, b=5),
        })
        st.plotly_chart(fig_adj, use_container_width=True)

    st.markdown("---")

    # ── Fila 2: ABC por monto + ABC por cantidad ───────────────────────────────
    ca, cb = st.columns(2)

    with ca:
        st.markdown('<div class="section-header">ABC por monto — distribución</div>',
                    unsafe_allow_html=True)
        if len(df_r_m) > 0:
            abc_r = (df_r_m.groupby("M_abc")
                     .agg(Items=("Descripcion","count"), Monto=("Monto","sum"))
                     .reindex(["A","B","C"]).fillna(0).reset_index())
            fig_abc_pie = px.pie(
                abc_r, names="M_abc", values="Monto",
                color="M_abc", color_discrete_map=PALETTE,
                hole=0.5, height=280,
            )
            fig_abc_pie.update_traces(textinfo="percent+label", textfont_size=12)
            fig_abc_pie.update_layout({**CHART_LAYOUT,
                "showlegend": False,
                "margin": dict(l=10, r=10, t=10, b=10),
            })
            st.plotly_chart(fig_abc_pie, use_container_width=True)

            # Tabla resumen ABC
            abc_r["% Items"] = (abc_r["Items"] / abc_r["Items"].sum() * 100).apply(lambda v: f"{v:.1f}%")
            abc_r["% Monto"] = (abc_r["Monto"] / abc_r["Monto"].sum() * 100).apply(lambda v: f"{v:.1f}%")
            abc_r["Monto"]   = abc_r["Monto"].apply(lambda v: f"${v:,.0f}")
            st.dataframe(
                abc_r.rename(columns={"M_abc":"Clase","Items":"Ítems"}),
                use_container_width=True, hide_index=True, height=140,
            )

    with cb:
        st.markdown('<div class="section-header">ABC por cantidad — distribución</div>',
                    unsafe_allow_html=True)
        if len(df_r_q) > 0:
            abc_q_r = (df_r_q.groupby("Q_abc")
                       .agg(Items=("Descripcion","count"), Unidades=("Cantidad","sum"))
                       .reindex(["A","B","C"]).fillna(0).reset_index())
            fig_abcq_pie = px.pie(
                abc_q_r, names="Q_abc", values="Unidades",
                color="Q_abc", color_discrete_map=PALETTE,
                hole=0.5, height=280,
            )
            fig_abcq_pie.update_traces(textinfo="percent+label", textfont_size=12)
            fig_abcq_pie.update_layout({**CHART_LAYOUT,
                "showlegend": False,
                "margin": dict(l=10, r=10, t=10, b=10),
            })
            st.plotly_chart(fig_abcq_pie, use_container_width=True)

            abc_q_r["% Items"]    = (abc_q_r["Items"] / abc_q_r["Items"].sum() * 100).apply(lambda v: f"{v:.1f}%")
            abc_q_r["% Unidades"] = (abc_q_r["Unidades"] / abc_q_r["Unidades"].sum() * 100).apply(lambda v: f"{v:.1f}%")
            abc_q_r["Unidades"]   = abc_q_r["Unidades"].apply(lambda v: f"{v:,.0f}")
            st.dataframe(
                abc_q_r.rename(columns={"Q_abc":"Clase","Items":"Ítems"}),
                use_container_width=True, hide_index=True, height=140,
            )

    st.markdown("---")

    # ── Fila 3: Curva de Pareto del rubro ────────────────────────────────────
    st.markdown('<div class="section-header">Curva de Pareto — ítems del rubro ordenados por monto</div>',
                unsafe_allow_html=True)

    if len(df_r_m) > 0:
        top_r = df_r_m.sort_values("Monto", ascending=False).reset_index(drop=True)
        top_r["idx"] = range(1, len(top_r) + 1)
        pct_acum_r = top_r["Monto"].cumsum() / top_r["Monto"].sum() * 100

        fig_par = go.Figure()
        fig_par.add_bar(
            x=top_r["idx"], y=top_r["Monto"],
            marker_color=top_r["M_abc"].map(PALETTE),
            name="Monto",
            hovertemplate="<b>%{customdata}</b><br>$%{y:,.0f}<extra></extra>",
            customdata=top_r["Descripcion"].str[:50],
        )
        fig_par.add_scatter(
            x=top_r["idx"], y=pct_acum_r, mode="lines",
            yaxis="y2", name="% Acumulado",
            line=dict(color="#7c3aed", width=2),
        )
        fig_par.add_hline(y=80, line_dash="dash", line_color="#dc2626",
                          annotation_text="80%", yref="y2")
        layout_par = {**CHART_LAYOUT,
            "height": 340,
            "yaxis":  dict(title="Monto ($)", showgrid=False),
            "yaxis2": dict(overlaying="y", side="right", range=[0, 105],
                           ticksuffix="%", title="% Acumulado"),
            "xaxis":  dict(title="Ítem (rank)", showgrid=False),
            "legend": dict(orientation="h"),
        }
        fig_par.update_layout(layout_par)
        st.plotly_chart(fig_par, use_container_width=True)

    st.markdown("---")

    # ── Fila 4: Matriz ABC-XYZ del rubro ─────────────────────────────────────
    st.markdown('<div class="section-header">Matriz ABC-XYZ del rubro</div>',
                unsafe_allow_html=True)

    strat_r = {
        "AX": ("Stock alto · Proveedor fijo · Reposición continua",     "#dcfce7", "#16a34a"),
        "AY": ("Stock buffer · Revisar mensual · Analizar causas",       "#fef9c3", "#b45309"),
        "AZ": ("⚠ CRÍTICO: alto gasto + demanda errática",              "#fee2e2", "#dc2626"),
        "BX": ("Stock normal · Control periódico",                       "#dcfce7", "#16a34a"),
        "BY": ("Stock moderado · Revisar trimestral",                    "#fef9c3", "#b45309"),
        "BZ": ("Stock mínimo · Analizar variabilidad",                   "#fef9c3", "#b45309"),
        "CX": ("Stock bajo · Pedido por demanda",                        "#f0fdf4", "#15803d"),
        "CY": ("Sin stock fijo · Control mínimo",                        "#f0fdf4", "#15803d"),
        "CZ": ("Evaluar eliminar · Pedido esporádico",                   "#f8fafc", "#64748b"),
    }

    hdr_cols = st.columns(4)
    hdr_cols[0].markdown("**ABC \\ XYZ**")
    hdr_cols[1].markdown("**X — Estable**")
    hdr_cols[2].markdown("**Y — Variable**")
    hdr_cols[3].markdown("**Z — Irregular**")

    for abc_cls in ["A", "B", "C"]:
        row_c = st.columns(4)
        row_c[0].markdown(f"**Clase {abc_cls}**")
        for j, xyz_cls in enumerate(["X", "Y", "Z"]):
            key = abc_cls + xyz_cls
            sub_xyz = df_r_m[
                (df_r_m.get("M_abc", pd.Series(dtype=str)) == abc_cls) &
                (df_r_m["XYZ"] == xyz_cls)
            ] if "M_abc" in df_r_m.columns else pd.DataFrame()
            cnt_xyz   = len(sub_xyz)
            monto_xyz = sub_xyz["Monto"].sum() if len(sub_xyz) > 0 else 0
            txt, bg, fg = strat_r.get(key, ("", "#ffffff", "#000000"))
            row_c[j+1].markdown(
                f'<div style="background:{bg};border-radius:8px;padding:10px;'
                f'text-align:center;border:1px solid #e2e8f0;font-size:.82rem;min-height:80px">'
                f'<b style="font-size:1.1rem;color:{fg}">{cnt_xyz}</b>'
                f'<span style="color:{fg}"> ítems</span><br>'
                f'<span style="color:#64748b;font-size:.78rem">${monto_xyz/1e6:.2f}M</span><br>'
                f'<span style="font-size:.70rem;color:#475569">{txt}</span></div>',
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # ── Fila 5: Tabla completa del rubro con filtros ──────────────────────────
    st.markdown('<div class="section-header">Todos los ítems del rubro</div>',
                unsafe_allow_html=True)

    fc1, fc2, fc3 = st.columns(3)
    f_abc_r  = fc1.multiselect("Filtrar ABC", ["A","B","C"], default=["A","B","C"],
                               key="f_abc_r")
    f_xyz_r  = fc2.multiselect("Filtrar XYZ", ["X","Y","Z"], default=["X","Y","Z"],
                               key="f_xyz_r")
    f_des_r  = fc3.checkbox("Incluir desiertos", value=True, key="f_des_r")

    df_tabla = df_r.copy()
    if "M_abc" in df_tabla.columns:
        df_tabla = df_tabla[df_tabla["M_abc"].isin(f_abc_r)]
    df_tabla = df_tabla[df_tabla["XYZ"].isin(f_xyz_r)]
    if not f_des_r:
        df_tabla = df_tabla[~df_tabla["Es_Desierto"]]
    df_tabla = df_tabla.sort_values("Monto", ascending=False)

    st.caption(f"{len(df_tabla)} ítems · Monto total: ${df_tabla['Monto'].sum():,.0f}")

    cols_tabla = ["M_abc","Q_abc","XYZ","Codigo","Descripcion",
                  "Cantidad","Consumo_Mensual","Justiprecio","Monto",
                  "Proveedor","Es_Desierto","Tiene_OC"]
    cols_tabla = [c for c in cols_tabla if c in df_tabla.columns]

    st.dataframe(
        df_tabla[cols_tabla]
        .rename(columns={
            "M_abc":"ABC $","Q_abc":"ABC Ctd","Monto":"Monto ($)",
            "Consumo_Mensual":"Consumo M.","Tiene_OC":"OC","Es_Desierto":"Des.",
        })
        .style.format({
            "Monto ($)":    "${:,.0f}",
            "Justiprecio":  "${:,.2f}",
            "Cantidad":     "{:,.0f}",
            "Consumo M.":   "{:,.0f}",
        }),
        use_container_width=True,
        height=500,
    )

    # Descarga del rubro
    buf_r = io.BytesIO()
    df_tabla[cols_tabla].to_excel(buf_r, index=False)
    st.download_button(
        f"⬇ Descargar {rubro_elegido} como Excel",
        data=buf_r.getvalue(),
        file_name=f"rubro_{rubro_elegido.lower().replace(' ','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_rubro",
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — DATOS
# ══════════════════════════════════════════════════════════════════════════════
with tab6:
    st.markdown('<div class="section-header">Explorador de Datos</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    abc_filter  = c1.multiselect("Clase ABC (monto)", ["A","B","C"], default=["A","B","C"])
    xyz_filter  = c2.multiselect("Clase XYZ",          ["X","Y","Z"], default=["X","Y","Z"])
    search_term = c3.text_input("Buscar descripción")

    df_view = df.copy()
    if "M_abc" in df_view.columns:
        df_view = df_view[df_view["M_abc"].isin(abc_filter)]
    df_view = df_view[df_view["XYZ"].isin(xyz_filter)]
    if search_term:
        df_view = df_view[df_view["Descripcion"].str.contains(search_term, case=False, na=False)]

    st.caption(f"{len(df_view)} ítems encontrados")

    show_cols = ["M_abc","Q_abc","XYZ","Rubro","Codigo","Descripcion",
                 "Cantidad","Consumo_Mensual","Justiprecio","Monto","Proveedor","Es_Desierto","Tiene_OC"]
    show_cols = [c for c in show_cols if c in df_view.columns]

    st.dataframe(
        df_view[show_cols].rename(columns={
            "M_abc":"ABC_M","Q_abc":"ABC_Q","Monto":"Monto ($)",
            "Consumo_Mensual":"Consumo Mens.","Tiene_OC":"OC","Es_Desierto":"Desierto",
        }).style.format({"Monto ($)":"${:,.0f}","Justiprecio":"${:,.2f}",
                          "Cantidad":"{:,.0f}","Consumo Mens.":"{:,.0f}"}),
        use_container_width=True,
        height=600,
    )

    # Download
    buf = io.BytesIO()
    df_view[show_cols].to_excel(buf, index=False)
    st.download_button(
        "⬇ Descargar filtrado como Excel",
        data=buf.getvalue(),
        file_name="tablero_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
