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

/* ── Metric cards — adaptan a claro/oscuro ── */
.metric-card {
    background: rgba(255,255,255,0.06);
    border-radius: 12px;
    padding: 18px 22px;
    border-left: 4px solid #3b82f6;
    border: 1px solid rgba(255,255,255,0.12);
    border-left: 4px solid #3b82f6;
    margin-bottom: 12px;
}
.metric-card.green  { border-left-color: #22c55e; }
.metric-card.orange { border-left-color: #f97316; }
.metric-card.red    { border-left-color: #ef4444; }
.metric-card.purple { border-left-color: #a855f7; }

.metric-card .value {
    font-size: 1.9rem; font-weight: 700; line-height: 1.1;
    color: inherit;
}
.metric-card .label {
    font-size: .75rem; font-weight: 600;
    text-transform: uppercase; letter-spacing: .06em;
    margin-bottom: 4px; opacity: .65;
}
.metric-card .delta {
    font-size: .8rem; margin-top: 4px; opacity: .6;
}

/* ── Section headers ── */
.section-header {
    font-size: 1rem; font-weight: 600;
    padding: 8px 0 6px;
    border-bottom: 2px solid rgba(255,255,255,0.15);
    margin-bottom: 14px;
}

/* ── Upload prompt ── */
.upload-prompt {
    border-radius: 16px; padding: 60px 40px;
    text-align: center;
    border: 2px dashed rgba(255,255,255,0.2);
    background: rgba(255,255,255,0.04);
}
.upload-prompt h2 { font-size: 1.6rem; margin-bottom: 8px; }
.upload-prompt p  { font-size: .95rem; opacity: .7; }

/* ── Sidebar ── */
div[data-testid="stSidebar"] { background: #0f172a; }
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
    # VARIOS: columna desplazada (una col menos en historial)
    # consumo=idx4(E), qty=idx17(R), justi=idx22(W), monto=idx24(Y), prov=idx12(M)
    "VARIOS":  {"consumo":4, "cm":24, "cq":17, "cj":22, "cp":12},
    # ESTERILIZACION y ESTERILIZACION DESIERTOS: consumo en idx4(E) por columna corrida
    "ESTERILIZACION":          {"consumo":4, "cm":25, "cq":17, "cj":23, "cp":12},
    "ESTERILIZACION DESIERTOS":{"consumo":4, "cm":25, "cq":17, "cj":23, "cp":12},
    # Resto de hojas: consumo=idx5(F), justi=idx23(X), monto=idx25(Z)
    "DEFAULT": {"consumo":5, "cm":25, "cq":17, "cj":23, "cp":12},
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
        cfg    = RUBRO_COL.get(sheet_name, RUBRO_COL["DEFAULT"])
        cm_i   = cfg["cm"]        # Monto Total pedido s/Justiprecio
        cq_i   = cfg["cq"]        # Cantidad pedida 1er sem 2026
        cj_i   = cfg["cj"]        # Justiprecio 2026
        cp_i   = cfg["cp"]        # Proveedor
        cons_i = cfg["consumo"]   # Consumo mensual estimado

        for pi in range(r1-1, r2):
            if pi >= len(df): continue
            row = df.iloc[pi]
            desc = row[1] if pd.notna(row[1]) else None
            if not desc: continue
            codigo = str(row[0]).strip() if pd.notna(row[0]) else ""
            # VARIOS y similares no tienen código vademecum — usar descripcion como ID
            if not codigo or codigo == "nan":
                if sheet_name == "VARIOS":
                    codigo = f"VARIOS_{pi}"
                else:
                    continue

            def sf(v):
                try: return float(v) if pd.notna(v) else 0.0
                except: return 0.0

            m = sf(row[cm_i]) if cm_i < len(row) else 0
            q = sf(row[cq_i]) if cq_i < len(row) else 0
            j = sf(row[cj_i]) if cj_i < len(row) else 0
            if m == 0 and q > 0 and j > 0: m = q * j
            prov = str(row[cp_i]).strip() if cp_i < len(row) and pd.notna(row[cp_i]) else ""
            oc   = str(row[15]).strip() if len(row)>15 and pd.notna(row[15]) else ""
            cons = sf(row[cons_i]) if cons_i < len(row) else 0

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
    font_color="#e2e8f0",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=10, r=10, t=36, b=10),
)
LEGEND_H = dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                font=dict(color="#e2e8f0"))


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
            Ejecución de presupuesto · Matriz ABC-XYZ · Simuladores · Carga de OC
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Fallback vacío — st.stop() ya frenó la ejecución, esto es solo para evitar NameError
import pandas as _pd
df_raw = _pd.DataFrame() if "df_raw" not in dir() else df_raw

# Apply filters
if len(df_raw) == 0:
    df = _pd.DataFrame()
else:
    df = df_raw[df_raw["Rubro"].isin(rubro_sel)].copy()
    if not show_desiertos:
        df = df[~df["Es_Desierto"]]

total_monto  = df["Monto"].sum()
n_desiertos  = int(df["Es_Desierto"].sum())
n_con_oc     = int(df["Tiene_OC"].sum())

# Total medicamentos únicos: deduplicar por Codigo vademecum (cuando existe)
# y por Descripcion para los que no tienen código asignado
_df_con_cod = df[df["Codigo"] != ""].drop_duplicates(subset=["Codigo"])
_df_sin_cod = df[df["Codigo"] == ""].drop_duplicates(subset=["Descripcion"])
total_items  = len(_df_con_cod) + len(_df_sin_cod)
total_items_raw = len(df)  # con repetidos, para % internos

pct_ejec     = n_con_oc / total_items_raw * 100 if total_items_raw > 0 else 0

# ── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📊 Resumen General",
    "🔢 ABC por Monto",
    "🔄 ABC por Cantidad + XYZ",
    "🏭 Proveedores",
    "🏥 Detalle por Rubro",
    "📋 Datos",
    "📥 Carga de OC",
    "🧮 Simuladores",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — RESUMEN GENERAL
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">Indicadores Globales</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric(
            label="Total Medicamentos",
            value=f"{total_items:,}",
            delta=f"{total_items_raw - total_items:,} duplicados excluidos",
            help=(
                f"Medicamentos e insumos **únicos** en el archivo. "
                f"Se deduplicó por código vademecum (columna A): si el mismo código "
                f"aparece en más de un rubro, se cuenta una sola vez. "
                f"Para los {len(_df_sin_cod)} ítems sin código, se deduplicó por descripción. "
                f"Total de renglones originales (con repetidos): {total_items_raw:,}. "
                f"Duplicados eliminados: {total_items_raw - total_items:,}."
            ),
        )
    with c2:
        st.metric(
            label="Monto Total Justiprecio",
            value=f"${total_monto/1e9:.2f}MM" if total_monto >= 1e9 else f"${total_monto/1e6:.2f}M",
            delta=f"{len(rubro_sel)} rubros seleccionados",
            help=(
                "Suma del campo 'Monto Total Corregido s/Justiprecio' (columna AA) "
                "de todas las hojas. Cuando ese campo está vacío o en cero, se calcula "
                "como Cantidad solicitada × Justiprecio 2026 (columnas R × Y). "
                "Representa el valor estimado del total a comprar usando el precio "
                "de referencia (justiprecio), independientemente de si hay OC emitida."
            ),
        )
    with c3:
        st.metric(
            label="Ítems Desiertos",
            value=f"{n_desiertos:,}",
            delta=f"{n_desiertos/total_items_raw*100:.1f}% del total de renglones",
            delta_color="inverse",
            help=(
                "Renglones donde el proveedor adjudicado figura como 'Desierto' "
                "en la columna M (Proveedor). Incluye tanto los desiertos embebidos "
                "en cada hoja de rubro como los de las hojas 'DESIERTOS' y "
                "'ESTERILIZACION DESIERTOS'. Un ítem desierto significa que la "
                "licitación no tuvo oferentes válidos para ese renglón."
            ),
        )
    with c4:
        st.metric(
            label="Con OC Confirmada",
            value=f"{n_con_oc:,}",
            delta=f"{pct_ejec:.1f}% ejecución",
            help=(
                "Ítems que tienen un número de Orden de Compra registrado en "
                "la columna P del archivo (columna 'OC'). Indica que la compra "
                "fue formalizada y está en proceso de entrega o ya entregada. "
                "El porcentaje de ejecución es: Ítems con OC ÷ Total ítems × 100."
            ),
        )
    with c5:
        st.metric(
            label="Sin OC (pendiente)",
            value=f"{total_items_raw - n_con_oc:,}",
            delta=f"{(total_items_raw-n_con_oc)/total_items_raw*100:.1f}% pendiente",
            delta_color="inverse",
            help=(
                "Ítems adjudicados que aún no tienen Orden de Compra emitida. "
                "Son los renglones donde la columna P está vacía. Representa "
                "el presupuesto comprometido pero todavía no formalizado en una OC."
            ),
        )

    # Segunda fila: indicadores monetarios de ejecución
    monto_con_oc  = df[df["Tiene_OC"]]["Monto"].sum()
    monto_sin_oc  = total_monto - monto_con_oc
    pct_ejec_monto = monto_con_oc / total_monto * 100 if total_monto > 0 else 0
    n_des_total    = n_desiertos
    monto_desierto = df[df["Es_Desierto"]]["Monto"].sum()

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric(
            label="Monto con OC confirmada",
            value=f"${monto_con_oc/1e9:.2f}MM" if monto_con_oc >= 1e9 else f"${monto_con_oc/1e6:.1f}M",
            delta=f"{pct_ejec_monto:.1f}% del presupuesto total",
            help=(
                "Suma del monto justiprecio de todos los ítems que tienen "
                "Orden de Compra emitida (columna P no vacía). "
                "Indica el valor monetario ya formalizado en OC, "
                "es decir, el presupuesto efectivamente ejecutado."
            ),
        )
    with m2:
        st.metric(
            label="Monto sin OC (pendiente)",
            value=f"${monto_sin_oc/1e9:.2f}MM" if monto_sin_oc >= 1e9 else f"${monto_sin_oc/1e6:.1f}M",
            delta=f"{100-pct_ejec_monto:.1f}% aún no formalizado",
            delta_color="inverse",
            help=(
                "Monto justiprecio de ítems adjudicados pero sin OC emitida aún. "
                "Representa el presupuesto comprometido que todavía está pendiente "
                "de formalizar. Calculado como: Monto Total − Monto con OC."
            ),
        )
    with m3:
        st.metric(
            label="% Ejecución monetaria",
            value=f"{pct_ejec_monto:.1f}%",
            help=(
                "Porcentaje del presupuesto total que ya tiene OC emitida. "
                "Se calcula como: Monto con OC ÷ Monto Total Justiprecio × 100. "
                "Diferente al % de ejecución por ítems: aquí se pondera por "
                "el valor económico de cada renglón, no solo por cantidad."
            ),
        )
    with m4:
        st.metric(
            label="Monto estimado desiertos",
            value=f"${monto_desierto/1e9:.2f}MM" if monto_desierto >= 1e9 else f"${monto_desierto/1e6:.1f}M",
            delta=f"{monto_desierto/total_monto*100:.1f}% del presupuesto",
            delta_color="inverse",
            help=(
                "Valor estimado de los renglones declarados desiertos "
                "(calculado como Cantidad × Justiprecio cuando está disponible). "
                "Representa el presupuesto en riesgo por falta de oferentes. "
                "Este monto no tiene proveedor asignado y requiere re-licitación "
                "o compra directa para garantizar el abastecimiento."
            ),
        )
    st.markdown("---")

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
            color_continuous_scale=[[0,"#60a5fa"],[1,"#1d4ed8"]],
            text=rubro_agg["Monto"].apply(lambda v: f"${v/1e6:.1f}M"),
            labels={"Monto":"Monto ($)","Rubro":""},
            title="Monto Justiprecio por Rubro",
        )
        fig.update_traces(textposition="outside", textfont_size=10, textfont_color="#e2e8f0")
        fig.update_layout({**CHART_LAYOUT,
            "height": 460, "coloraxis_showscale": False, "legend": LEGEND_H,
            "xaxis": dict(showgrid=False, visible=False),
            "yaxis": dict(tickfont=dict(size=11, color="#e2e8f0")),
        })
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
        fig2.update_layout({**CHART_LAYOUT,
            "barmode":  "stack",
            "height":   480,
            "title":    dict(text="Ítems: Adjudicados vs Desiertos",
                             y=0.97, yanchor="top"),
            "legend":   dict(orientation="h", x=0.5, xanchor="center",
                             y=1.08, yanchor="bottom",
                             font=dict(color="#e2e8f0")),
            "margin":   dict(l=10, r=10, t=80, b=120),
            "xaxis":    dict(tickangle=-40,
                             tickfont=dict(size=10, color="#e2e8f0")),
            "yaxis":    dict(title="Cantidad",
                             tickfont=dict(color="#e2e8f0"),
                             gridcolor="rgba(255,255,255,0.08)"),
        })
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
    fig3.update_traces(textposition="outside", textfont_size=10, textfont_color="#e2e8f0")
    fig3.update_layout({**CHART_LAYOUT,
        "height": 420, "coloraxis_showscale": False,
        "xaxis": dict(range=[0,115], title="% Ejecución", showgrid=True,
                      gridcolor="rgba(255,255,255,0.12)",
                      tickfont=dict(color="#e2e8f0")),
        "yaxis": dict(tickfont=dict(size=11, color="#e2e8f0")),
    })
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
            "yaxis":  dict(title="Monto ($)", showgrid=False,
                           tickfont=dict(color="#e2e8f0")),
            "yaxis2": dict(title="% Acumulado", overlaying="y", side="right",
                           range=[0,105], ticksuffix="%",
                           tickfont=dict(color="#e2e8f0")),
            "xaxis":  dict(title="Ítem (rank)", showgrid=False,
                           tickfont=dict(color="#e2e8f0")),
            "legend": dict(orientation="h", font=dict(color="#e2e8f0")),
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
        fig_pie.update_layout({**CHART_LAYOUT, "height": 380, "showlegend": False})
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
        fig_qpie.update_layout({**CHART_LAYOUT, "height": 320, "showlegend": False})
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
        xyz_color = {"X":"#34d399","Y":"#fbbf24","Z":"#94a3b8"}
        fig_xyz = px.bar(xyz_sum, x="XYZ", y="Items",
                         color="XYZ", color_discrete_map=xyz_color,
                         text="Items", title="Ítems por Clase XYZ")
        fig_xyz.update_traces(textposition="outside", textfont_color="#e2e8f0")
        fig_xyz.update_layout({**CHART_LAYOUT,
            "height": 200, "showlegend": False,
            "xaxis_title": "", "yaxis_title": "Ítems",
            "xaxis": dict(tickfont=dict(color="#e2e8f0", size=13)),
            "yaxis": dict(showgrid=False, tickfont=dict(color="#e2e8f0")),
        })
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

    # ── Tablas de desglose XYZ ────────────────────────────────────────────────
    st.markdown('<div class="section-header">Detalle de ítems por clase XYZ</div>',
                unsafe_allow_html=True)

    xyz_tab_x, xyz_tab_y, xyz_tab_z = st.tabs([
        "✅ X — Demanda Estable",
        "⚡ Y — Demanda Variable",
        "❓ Z — Demanda Irregular / Sin datos",
    ])

    xyz_cols_show = ["M_abc","Q_abc","Rubro","Codigo","Descripcion",
                     "Cantidad","Consumo_Mensual","Justiprecio","Monto","Proveedor"]
    xyz_cols_show = [c for c in xyz_cols_show if c in df.columns]

    def xyz_detail_df(df_in, clase):
        sub = df_in[df_in["XYZ"] == clase].copy()
        if "M_abc" in sub.columns:
            sub = sub.sort_values(["M_abc","Monto"], ascending=[True, False])
        else:
            sub = sub.sort_values("Monto", ascending=False)
        return sub[xyz_cols_show].rename(columns={
            "M_abc":"ABC $","Q_abc":"ABC Ctd",
            "Monto":"Monto ($)","Consumo_Mensual":"Consumo M.",
        })

    for tab_xyz, clase_xyz, color_xyz in [
        (xyz_tab_x, "X", "#34d399"),
        (xyz_tab_y, "Y", "#fbbf24"),
        (xyz_tab_z, "Z", "#94a3b8"),
    ]:
        with tab_xyz:
            df_xyz_det = xyz_detail_df(df, clase_xyz)
            n_xyz = len(df_xyz_det)
            monto_xyz = df[df["XYZ"]==clase_xyz]["Monto"].sum()
            st.caption(f"{n_xyz} ítems · Monto total: ${monto_xyz:,.0f}")
            st.dataframe(
                df_xyz_det.style.format({
                    "Monto ($)":  "${:,.0f}",
                    "Justiprecio":"${:,.2f}",
                    "Cantidad":   "{:,.0f}",
                    "Consumo M.": "{:,.0f}",
                }),
                use_container_width=True,
                height=420,
            )
            buf_xyz = io.BytesIO()
            df_xyz_det.to_excel(buf_xyz, index=False)
            st.download_button(
                f"⬇ Descargar clase {clase_xyz}",
                data=buf_xyz.getvalue(),
                file_name=f"xyz_clase_{clase_xyz}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_xyz_{clase_xyz}",
            )

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
        "AX": ("#0f3d2e", "#4ade80", "#86efac"),
        "AY": ("#3d3000", "#fbbf24", "#fde68a"),
        "AZ": ("#3d0f0f", "#ef4444", "#fca5a5"),
        "BX": ("#0a2e1f", "#34d399", "#6ee7b7"),
        "BY": ("#2e2200", "#f59e0b", "#fcd34d"),
        "BZ": ("#2e2200", "#f59e0b", "#fcd34d"),
        "CX": ("#0a1f15", "#22c55e", "#86efac"),
        "CY": ("#1a1a0a", "#84cc16", "#bef264"),
        "CZ": ("#1a1a1a", "#94a3b8", "#cbd5e1"),
    }

    mcols = st.columns(4)
    mcols[0].markdown("**ABC \\ XYZ**")
    for j, xyz in enumerate(["X — Estable", "Y — Variable", "Z — Irregular"]):
        mcols[j+1].markdown(f"**{xyz}**")

    for abc in ["A","B","C"]:
        row_cols = st.columns(4)
        row_cols[0].markdown(f"**Clase {abc}**")
        for j, xyz in enumerate(["X","Y","Z"]):
            key = abc+xyz
            sub = matrix_cnt[(matrix_cnt["M_abc"]==abc)&(matrix_cnt["XYZ"]==xyz)]
            cnt   = int(sub["Items"].sum()) if len(sub)>0 else 0
            monto = sub["Monto"].sum() if len(sub)>0 else 0
            bg, border, fg = bg_colors.get(key, ("#1a1a1a","#475569","#cbd5e1"))
            row_cols[j+1].markdown(
                f'<div style="background:{bg};border-radius:10px;padding:14px 10px;'
                f'text-align:center;border:2px solid {border};font-size:.82rem;min-height:90px">'
                f'<b style="font-size:1.2rem;color:{border}">{cnt}</b>'
                f'<span style="color:{fg}"> ítems</span><br>'
                f'<span style="color:{fg};font-size:.85rem;font-weight:600">${monto/1e6:.1f}M</span><br>'
                f'<span style="font-size:.70rem;color:{fg};opacity:.85">'
                f'{strat.get(key,"").replace(chr(10)," · ")}</span></div>',
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
                lambda v: "#ef4444" if v>=20 else ("#f97316" if v>=10 else ("#fbbf24" if v>=5 else "#34d399"))),
            text=top20["Pct"].apply(lambda v: f"{v:.1f}%"),
            textposition="outside",
            textfont=dict(color="#e2e8f0", size=11),
            hovertemplate="<b>%{y}</b><br>$%{x:,.0f}<extra></extra>",
        )
        fig_prov.add_scatter(
            x=top20["Pct_Acum"]/100 * top20["Monto"].max(),
            y=top20["Proveedor"], mode="lines+markers",
            name="% Acum", line=dict(color="#c084fc", width=2),
            xaxis="x2",
        )
        layout_prov = {**CHART_LAYOUT,
            "height": 500,
            "title": "Top 20 Proveedores por Monto",
            "xaxis":  dict(title="Monto ($)", showgrid=False,
                           tickfont=dict(color="#cbd5e1")),
            "xaxis2": dict(title="% Acumulado", overlaying="x", side="top",
                           range=[0,110], ticksuffix="%",
                           tickfont=dict(color="#cbd5e1")),
            "yaxis":  dict(tickfont=dict(size=11, color="#e2e8f0"),
                           autorange="reversed"),
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
        fig_pie2.update_traces(textinfo="percent+label", textfont_size=11, textfont_color="#1e293b")
        fig_pie2.update_layout({**CHART_LAYOUT, "height": 500, "showlegend": False})
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
                color_continuous_scale=[[0,"#60a5fa"],[1,"#1d4ed8"]],
                text=prov_r["Monto"].apply(lambda v: f"${v/1e6:.2f}M" if v >= 1e6 else f"${v:,.0f}"),
                labels={"Monto":"","Proveedor":""},
                height=max(250, len(prov_r) * 32),
            )
            fig_pr.update_traces(textposition="outside", textfont_size=10, textfont_color="#e2e8f0")
            fig_pr.update_layout({**CHART_LAYOUT,
                "coloraxis_showscale": False,
                "xaxis": dict(showgrid=False, visible=False),
                "yaxis": dict(tickfont=dict(size=11, color="#e2e8f0")),
                "margin": dict(l=10, r=60, t=10, b=10),
            })
            st.plotly_chart(fig_pr, use_container_width=True)
        else:
            st.info("Sin proveedores adjudicados en este rubro.")

    with c_r:
        st.markdown('<div class="section-header">Ejecución presupuestal</div>',
                    unsafe_allow_html=True)

        # Donut ejecución — color gris oscuro para "Sin OC" visible en dark mode
        fig_ej = go.Figure(go.Pie(
            labels=["Con OC", "Sin OC"],
            values=[monto_oc_r, max(0, monto_sin_oc)],
            hole=0.6,
            marker_colors=["#22c55e", "#475569"],
            textinfo="percent+label",
            textfont=dict(size=12, color="#e2e8f0"),
        ))
        fig_ej.update_layout({**CHART_LAYOUT,
            "height": 260,
            "showlegend": False,
            "annotations": [dict(
                text=f"<b>{pct_ejec_r:.0f}%</b><br>ejecución",
                x=0.5, y=0.5, font=dict(size=18, color="#e2e8f0"),
                showarrow=False,
            )],
            "margin": dict(l=10, r=10, t=10, b=10),
        })
        st.plotly_chart(fig_ej, use_container_width=True)

        # Métricas simples en vez del bar chart (evita superposición)
        ma, mb = st.columns(2)
        ma.metric("Adjudicados", f"{total_r - des_r}",
                  f"{(total_r-des_r)/total_r*100:.0f}%" if total_r > 0 else "")
        mb.metric("Desiertos", f"{des_r}",
                  f"{des_r/total_r*100:.0f}%" if total_r > 0 else "",
                  delta_color="inverse")

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
        "AX": ("Stock alto · Proveedor fijo · Reposición continua",  "#0f3d2e", "#4ade80", "#86efac"),
        "AY": ("Stock buffer · Revisar mensual · Analizar causas",    "#3d3000", "#fbbf24", "#fde68a"),
        "AZ": ("⚠ CRÍTICO: alto gasto + demanda errática",           "#3d0f0f", "#ef4444", "#fca5a5"),
        "BX": ("Stock normal · Control periódico",                    "#0a2e1f", "#34d399", "#6ee7b7"),
        "BY": ("Stock moderado · Revisar trimestral",                 "#2e2200", "#f59e0b", "#fcd34d"),
        "BZ": ("Stock mínimo · Analizar variabilidad",                "#2e2200", "#f59e0b", "#fcd34d"),
        "CX": ("Stock bajo · Pedido por demanda",                     "#0a1f15", "#22c55e", "#86efac"),
        "CY": ("Sin stock fijo · Control mínimo",                     "#1a1a0a", "#84cc16", "#bef264"),
        "CZ": ("Evaluar eliminar · Pedido esporádico",                "#1a1a1a", "#94a3b8", "#cbd5e1"),
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
            txt, bg, border, fg = strat_r.get(key, ("", "#1a1a1a", "#475569", "#cbd5e1"))
            row_c[j+1].markdown(
                f'<div style="background:{bg};border-radius:10px;padding:14px 10px;'
                f'text-align:center;border:2px solid {border};font-size:.82rem;min-height:80px">'
                f'<b style="font-size:1.2rem;color:{border}">{cnt_xyz}</b>'
                f'<span style="color:{fg}"> ítems</span><br>'
                f'<span style="color:{fg};font-size:.85rem;font-weight:600">${monto_xyz/1e6:.2f}M</span><br>'
                f'<span style="font-size:.70rem;color:{fg};opacity:.85">{txt}</span></div>',
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


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — CARGA DE OC
# ══════════════════════════════════════════════════════════════════════════════
with tab7:

    # ── Importar módulos OC ───────────────────────────────────────────────────
    try:
        import pdfplumber
        from parsear_oc import parsear_oc
        from cargar_oc  import aplicar_oc
        PDF_OK = True
    except ImportError as _ie:
        PDF_OK = False
        _ie_msg = str(_ie)

    if not PDF_OK:
        st.error(f"Falta instalar: `pip install pdfplumber`  ({_ie_msg})")
        st.stop()

    st.markdown("### 📥 Carga de Orden de Compra")
    st.markdown(
        "Subí el PDF de la OC · Verificá los datos extraídos · "
        "Descargá el SOLICITUDES actualizado con N° OC, cantidad, precio y proveedor."
    )

    # ── Estado de sesión ──────────────────────────────────────────────────────
    for _k, _v in [("oc_data",None),("oc_df_edit",None),
                   ("oc_enc",{}),("oc_prov",{})]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

    # ── PASO 1: PDF ───────────────────────────────────────────────────────────
    st.markdown("---")
    col_pdf, col_info = st.columns([2, 1])

    with col_pdf:
        st.markdown("**Paso 1 — Subir PDF de la Orden de Compra**")
        pdf_up = st.file_uploader(
            "PDF de la OC (texto seleccionable)",
            type=["pdf"], key="oc_pdf_uploader",
            help="Solo se procesa la primera hoja. La hoja de remito se ignora.",
        )

    with col_info:
        st.markdown("**Columnas que se actualizan en el Excel:**")
        st.markdown("""
| Col | Campo |
|-----|-------|
| **P** | N° de OC |
| **M** | Proveedor *(si vacío)* |
| **U** | Cantidad con OC 2026 |
| **Y** | Precio Proveedor 2026 |
| **AA** | Monto Total adjudicado |
""")
        st.info(
            "**Prioridad de búsqueda:**\n"
            "1. N° Solicitud + N° Renglón\n"
            "2. N° Solicitud + Código\n"
            "3. Código único en el archivo",
            icon="🔍"
        )

    if pdf_up:
        with st.spinner("Leyendo el PDF..."):
            _pdf_bytes = pdf_up.read()
            _tmp_pdf   = f"/tmp/oc_{pdf_up.name}"
            with open(_tmp_pdf, "wb") as _f:
                _f.write(_pdf_bytes)
            _oc = parsear_oc(_tmp_pdf)
            st.session_state.oc_data    = _oc
            st.session_state.oc_enc     = dict(_oc["encabezado"])
            st.session_state.oc_prov    = dict(_oc["proveedor"])
            st.session_state.oc_df_edit = None

        for _e in _oc.get("errores", []):
            st.warning(f"⚠️ {_e}")

    # ── PASO 2: Encabezado ────────────────────────────────────────────────────
    if st.session_state.oc_data:
        enc  = st.session_state.oc_enc
        prov = st.session_state.oc_prov

        st.markdown("---")
        st.markdown("**Paso 2 — Verificar encabezado de la OC**")
        st.caption("El N° Solicitud es la clave principal de búsqueda. Corregilo si no coincide con el PDF.")

        c1, c2, c3 = st.columns(3)
        with c1:
            enc["oc_completo"]   = st.text_input("N° de OC",    value=enc.get("oc_completo",""),   key="oc_num")
            enc["fecha"]         = st.text_input("Fecha",        value=enc.get("fecha",""),         key="oc_fecha")
        with c2:
            enc["nro_solicitud"] = st.text_input(
                "N° Solicitud ⬅ clave de búsqueda",
                value=enc.get("nro_solicitud",""), key="oc_solic",
                help="Columna K del SOLICITUDES. Si hay múltiples OC para distintas solicitudes, cada PDF se carga por separado."
            )
            enc["licitacion"]    = st.text_input("Licitación",   value=enc.get("licitacion",""),    key="oc_lic")
        with c3:
            prov["razon_social"] = st.text_input("Proveedor",    value=prov.get("razon_social",""), key="oc_prov_rs")
            prov["cuit"]         = st.text_input("CUIT",         value=prov.get("cuit",""),         key="oc_cuit")

        if not enc.get("nro_solicitud"):
            st.error(
                "❌ No se detectó el N° de Solicitud en el PDF. "
                "Ingresalo manualmente arriba — es el número que figura en la columna K del Excel."
            )

        # ── PASO 3: Renglones ─────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("**Paso 3 — Revisar y editar renglones**")

        _renglones = st.session_state.oc_data.get("renglones", [])

        if not _renglones:
            st.error("❌ No se pudieron extraer renglones del PDF.")
            with st.expander("Ver texto extraído (diagnóstico)"):
                st.text(st.session_state.oc_data.get("texto_raw","")[:3000])
        else:
            st.success(f"✅ Se detectaron **{len(_renglones)} renglones**.")
            st.caption("Podés editar cualquier campo. Destildá **✓** para omitir un renglón.")

            if st.session_state.oc_df_edit is None:
                st.session_state.oc_df_edit = pd.DataFrame([{
                    "✓":               True,
                    "Renglón":         r.get("renglon",""),
                    "Código":          r.get("codigo",""),
                    "Descripción":     r.get("descripcion",""),
                    "Marca":           r.get("marca",""),
                    "Cantidad":        float(r.get("cantidad_num", 0)),
                    "Precio unitario": float(r.get("importe_unitario_num", 0)),
                    "Monto total":     float(r.get("importe_total_num", 0)),
                } for r in _renglones])

            _df_edit = st.data_editor(
                st.session_state.oc_df_edit,
                use_container_width=True,
                num_rows="fixed",
                hide_index=True,
                column_config={
                    "✓":               st.column_config.CheckboxColumn("✓", width="small"),
                    "Cantidad":        st.column_config.NumberColumn(format="%.0f"),
                    "Precio unitario": st.column_config.NumberColumn(format="$%.2f"),
                    "Monto total":     st.column_config.NumberColumn(format="$%.2f"),
                },
                key="oc_tabla_edit",
            )
            st.session_state.oc_df_edit = _df_edit

            _incl = _df_edit[_df_edit["✓"]]
            _m1, _m2, _m3 = st.columns(3)
            _m1.metric("Incluidos",  len(_incl))
            _m2.metric("Monto OC",   f"${_incl['Monto total'].sum():,.0f}")
            _m3.metric("Excluidos",  len(_df_edit) - len(_incl))

            # ── PASO 4: SOLICITUDES xlsx ──────────────────────────────────────
            st.markdown("---")
            st.markdown("**Paso 4 — Subir SOLICITUDES_2026.xlsx**")
            st.caption("El original no se modifica. Descargás una copia con los datos de la OC aplicados.")

            _xlsx_up = st.file_uploader(
                "SOLICITUDES_2026.xlsx",
                type=["xlsx"], key="oc_xlsx_uploader",
            )

            # ── PASO 5: Aplicar ───────────────────────────────────────────────
            if _xlsx_up:
                _xlsx_bytes = _xlsx_up.read()
                st.success(f"✅ {_xlsx_up.name} ({len(_xlsx_bytes)//1024} KB)")

                st.markdown("---")
                if st.button("🚀 Aplicar OC al archivo", type="primary",
                             use_container_width=True, key="oc_aplicar"):

                    _df_ok = st.session_state.oc_df_edit[
                        st.session_state.oc_df_edit["✓"]
                    ]
                    _rengls_conf = [{
                        "renglon":              str(r["Renglón"]),
                        "codigo":               str(r["Código"]).strip().upper(),
                        "descripcion":          str(r["Descripción"]),
                        "marca":                str(r.get("Marca","")),
                        "cantidad_num":         float(r["Cantidad"]),
                        "importe_unitario_num": float(r["Precio unitario"]),
                        "importe_total_num":    float(r["Monto total"]),
                    } for _, r in _df_ok.iterrows()]

                    _tmp_xlsx = "/tmp/solicitudes_oc_input.xlsx"
                    with open(_tmp_xlsx, "wb") as _f:
                        _f.write(_xlsx_bytes)

                    with st.spinner("Buscando renglones y actualizando Excel..."):
                        _wb_nuevo, _resultados = aplicar_oc(
                            _tmp_xlsx,
                            {"encabezado": enc, "proveedor": prov},
                            _rengls_conf,
                        )

                    _df_res  = pd.DataFrame(_resultados)
                    _ok_rows = _df_res[_df_res["estado"].str.startswith("✅")]
                    _nok_rows= _df_res[_df_res["estado"].str.startswith("⚠")]

                    _ra, _rb = st.columns(2)
                    _ra.metric("✅ Renglones cargados", len(_ok_rows))
                    _rb.metric("⚠️ No encontrados",     len(_nok_rows))

                    if len(_ok_rows):
                        st.success(f"Se actualizaron **{len(_ok_rows)} renglones** correctamente.")
                    if len(_nok_rows):
                        st.warning(
                            f"**{len(_nok_rows)} renglones no encontrados.** "
                            "Verificá que el N° de Solicitud del encabezado coincida "
                            "con la columna K del Excel, y que el N° de renglón "
                            "coincida con la columna O."
                        )

                    # Tabla de resultados con método de matching
                    st.dataframe(
                        _df_res[["renglon","codigo","descripcion","match","estado","hoja","fila"]].rename(columns={
                            "renglon":"Renglón","codigo":"Código",
                            "descripcion":"Descripción","match":"Encontrado por",
                            "estado":"Estado","hoja":"Hoja","fila":"Fila",
                        }),
                        use_container_width=True,
                        hide_index=True,
                    )

                    with st.expander("ℹ️ ¿Qué significa cada método de búsqueda?"):
                        st.markdown("""
| Método | Descripción |
|---|---|
| **Solicitud + Renglón** | Buscó por N° solicitud (col K) + N° renglón (col O). Método más confiable: identifica el renglón exacto aunque el código esté repetido en otras solicitudes. |
| **Solicitud + Código** | Buscó por N° solicitud (col K) + código vademecum (col A). Usado cuando el renglón no coincide exactamente. |
| **Código único** | El código vademecum aparece una sola vez en todo el archivo. Sin riesgo de confusión. |
| **Código ambiguo** | El código aparece en más de una fila y no se pudo determinar a cuál corresponde. Requiere N° solicitud correcto. |
| **⚠️ No encontrado** | Ningún método encontró coincidencia. Verificar manualmente. |
""")

                    # Descarga
                    _buf = io.BytesIO()
                    _wb_nuevo.save(_buf)
                    _buf.seek(0)
                    _oc_safe = enc.get("oc_completo","OC").replace("/","_")
                    st.download_button(
                        label=f"⬇ Descargar SOLICITUDES_2026_OC{_oc_safe}.xlsx",
                        data=_buf.getvalue(),
                        file_name=f"SOLICITUDES_2026_OC{_oc_safe}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary",
                        key="oc_descarga",
                    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — SIMULADORES
# ══════════════════════════════════════════════════════════════════════════════
with tab8:

    # Guard: verificar que haya datos cargados antes de ejecutar cualquier cálculo
    if not uploaded:
        st.info("📂 Subí el archivo **SOLICITUDES_2026.xlsx** desde la barra lateral para usar los simuladores.")
        st.stop()

    # Verificar que las columnas necesarias existan
    cols_requeridas = ["Clase_ABC_Monto", "Monto_Total", "Justiprecio", "Cantidad_Pedida", "Rubro"]
    cols_faltantes = [c for c in cols_requeridas if c not in df.columns]
    if cols_faltantes:
        st.error(f"Faltan columnas en el archivo: {cols_faltantes}")
        st.stop()

    st.markdown("### 🧮 Simuladores de licitación y presupuesto")
    st.markdown(
        "Herramientas para evaluar escenarios antes de tomar decisiones. "
        "Los cambios son **solo visuales** — no modifican el archivo original."
    )

    sim1, sim2 = st.tabs([
        "📈 Simulador de ampliaciones",
        "💰 Impacto presupuestario",
    ])

    # ─────────────────────────────────────────────────────────────────────────
    # SIMULADOR 1 — AMPLIACIONES DE LICITACIÓN
    # ─────────────────────────────────────────────────────────────────────────
    with sim1:
        st.markdown("#### Simulador de ampliaciones de licitación")
        st.markdown(
            "Calculá cuánto podés ampliar por rubro o por ítem, y el impacto "
            "en el monto total. Las licitaciones pueden ampliarse hasta un "
            "**20%** del monto original sin nueva licitación (art. 119 Ley 13.981)."
        )

        # ── Filtros ───────────────────────────────────────────────────────────
        sa1, sa2, sa3 = st.columns(3)
        with sa1:
            amp_rubro = st.multiselect(
                "Rubros a ampliar",
                options=sorted(df["Rubro"].unique().tolist()),
                default=[],
                key="amp_rubros",
                help="Dejá vacío para incluir todos los rubros",
            )
        with sa2:
            amp_abc = st.multiselect(
                "Clase ABC",
                options=["A", "B", "C"],
                default=["A", "B"],
                key="amp_abc",
                help="Clase A = top 80% del gasto. Lo más urgente.",
            )
        with sa3:
            solo_con_oc = st.checkbox(
                "Solo ítems con OC confirmada",
                value=True,
                key="amp_solo_oc",
                help="Las ampliaciones aplican sobre ítems ya adjudicados con OC.",
            )

        # Porcentaje de ampliación
        st.markdown("---")
        pct_amp = st.slider(
            "Porcentaje de ampliación (%)",
            min_value=1, max_value=30, value=20, step=1,
            key="amp_pct",
            help="La ley permite hasta 20% sin nueva licitación. Podés simular más para planificación.",
        )

        # ── Cálculo ───────────────────────────────────────────────────────────
        df_amp = df.copy()
        if amp_rubro:
            df_amp = df_amp[df_amp["Rubro"].isin(amp_rubro)]
        if amp_abc:
            df_amp = df_amp[df_amp["Clase_ABC_Monto"].isin(amp_abc)]
        if solo_con_oc:
            df_amp = df_amp[df_amp["Tiene_OC"] == "SI"]

        df_amp = df_amp[df_amp["Monto_Total"] > 0].copy()
        df_amp["Monto_Adicional"]   = df_amp["Monto_Total"] * (pct_amp / 100)
        df_amp["Monto_Ampliado"]    = df_amp["Monto_Total"] * (1 + pct_amp / 100)
        df_amp["Cantidad_Adicional"]= df_amp["Cantidad_Pedida"] * (pct_amp / 100)
        df_amp["Cantidad_Ampliada"] = df_amp["Cantidad_Pedida"] * (1 + pct_amp / 100)

        monto_orig  = df_amp["Monto_Total"].sum()
        monto_adic  = df_amp["Monto_Adicional"].sum()
        monto_total_amp = df_amp["Monto_Ampliado"].sum()

        # ── KPIs ─────────────────────────────────────────────────────────────
        k1, k2, k3, k4 = st.columns(4)
        k1.metric(
            "Ítems en el escenario",
            f"{len(df_amp):,}",
        )
        k2.metric(
            "Monto original seleccionado",
            f"${monto_orig/1e6:.1f}M",
        )
        k3.metric(
            f"Monto adicional (+{pct_amp}%)",
            f"${monto_adic/1e6:.1f}M",
            delta=f"+{pct_amp}% sobre lo adjudicado",
        )
        k4.metric(
            "Monto total ampliado",
            f"${monto_total_amp/1e6:.1f}M",
            delta=f"+${monto_adic/1e6:.1f}M vs original",
        )

        # ── Gráfico por rubro ─────────────────────────────────────────────────
        if len(df_amp) > 0:
            st.markdown("---")
            rubro_agg = (
                df_amp.groupby("Rubro")
                .agg(
                    Monto_Original=("Monto_Total", "sum"),
                    Monto_Adicional=("Monto_Adicional", "sum"),
                )
                .sort_values("Monto_Original", ascending=True)
                .reset_index()
            )

            fig_amp = go.Figure()
            fig_amp.add_bar(
                y=rubro_agg["Rubro"],
                x=rubro_agg["Monto_Original"],
                name="Monto original",
                orientation="h",
                marker_color="rgba(88,166,255,.75)",
            )
            fig_amp.add_bar(
                y=rubro_agg["Rubro"],
                x=rubro_agg["Monto_Adicional"],
                name=f"Ampliación +{pct_amp}%",
                orientation="h",
                marker_color="rgba(63,185,80,.8)",
            )
            fig_amp.update_layout({
                **CHART_LAYOUT,
                "barmode":  "stack",
                "height":   max(300, len(rubro_agg) * 38),
                "title":    f"Monto original vs ampliación +{pct_amp}% por rubro",
                "legend":   dict(orientation="h", y=1.05, font=dict(color="#e2e8f0")),
                "xaxis":    dict(title="Monto ($)", tickfont=dict(color="#e2e8f0"),
                                 tickformat="$,.0f", gridcolor="rgba(255,255,255,.08)"),
                "yaxis":    dict(tickfont=dict(color="#e2e8f0")),
            })
            st.plotly_chart(fig_amp, use_container_width=True)

            # ── Tabla detalle ─────────────────────────────────────────────────
            st.markdown("**Detalle por ítem**")
            df_tabla_amp = (
                df_amp[["Rubro","Descripcion","Proveedor","Clase_ABC_Monto",
                         "Cantidad_Pedida","Cantidad_Adicional","Cantidad_Ampliada",
                         "Justiprecio","Monto_Total","Monto_Adicional","Monto_Ampliado"]]
                .sort_values("Monto_Adicional", ascending=False)
                .rename(columns={
                    "Clase_ABC_Monto": "ABC",
                    "Cantidad_Pedida":    "Cant. original",
                    "Cantidad_Adicional": f"Cant. +{pct_amp}%",
                    "Cantidad_Ampliada":  "Cant. total",
                    "Monto_Total":        "Monto original",
                    "Monto_Adicional":    f"Monto +{pct_amp}%",
                    "Monto_Ampliado":     "Monto total",
                })
            )
            st.dataframe(
                df_tabla_amp.style.format({
                    "Cant. original":     "{:,.0f}",
                    f"Cant. +{pct_amp}%": "{:,.1f}",
                    "Cant. total":        "{:,.0f}",
                    "Justiprecio":        "${:,.2f}",
                    "Monto original":     "${:,.0f}",
                    f"Monto +{pct_amp}%": "${:,.0f}",
                    "Monto total":        "${:,.0f}",
                }),
                use_container_width=True,
                height=400,
                hide_index=True,
            )

            # ── Descarga ──────────────────────────────────────────────────────
            buf_amp = io.BytesIO()
            df_tabla_amp.to_excel(buf_amp, index=False)
            st.download_button(
                f"⬇ Descargar simulación de ampliación +{pct_amp}%",
                data=buf_amp.getvalue(),
                file_name=f"ampliacion_{pct_amp}pct.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_amp",
            )
        else:
            st.info("No hay ítems que cumplan los filtros seleccionados.")

    # ─────────────────────────────────────────────────────────────────────────
    # SIMULADOR 2 — IMPACTO PRESUPUESTARIO
    # ─────────────────────────────────────────────────────────────────────────
    with sim2:
        st.markdown("#### Simulador de impacto presupuestario")
        st.markdown(
            "Aplicá variaciones de precio, cantidad o incorporación de desiertos "
            "y visualizá el nuevo presupuesto requerido antes de presentar una solicitud."
        )

        # ── Escenarios ────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("**Configurar escenario**")

        esc1, esc2 = st.columns(2)

        with esc1:
            st.markdown("**Variación de precios**")
            var_precio_global = st.slider(
                "Variación global de justiprecio (%)",
                min_value=-50, max_value=100, value=0, step=5,
                key="esc_precio",
                help="Simula inflación de precios o renegociación. 0 = sin cambio.",
            )
            st.caption("Podés también definir variaciones por clase:")
            var_precio_A = st.slider("Variación clase A (%)", -50, 100, 0, 5, key="esc_pA")
            var_precio_B = st.slider("Variación clase B (%)", -50, 100, 0, 5, key="esc_pB")
            var_precio_C = st.slider("Variación clase C (%)", -50, 100, 0, 5, key="esc_pC")

        with esc2:
            st.markdown("**Incorporar desiertos**")
            inc_desiertos = st.checkbox(
                "Incluir ítems desiertos al presupuesto",
                value=False, key="esc_des",
                help="Simula el costo de re-licitar o comprar directamente los ítems sin oferentes.",
            )
            if inc_desiertos:
                pct_precio_des = st.slider(
                    "Precio estimado desiertos (% del justiprecio original)",
                    50, 200, 120, 5, key="esc_des_pct",
                    help="Los desiertos suelen comprarse por compra directa a mayor precio. 120% = 20% más caro.",
                )
            else:
                pct_precio_des = 100

            st.markdown("**Variación de cantidades**")
            var_cantidad = st.slider(
                "Variación global de cantidad (%)",
                -50, 100, 0, 5, key="esc_cant",
                help="Simula mayor o menor demanda proyectada.",
            )
            filtro_rubro_esc = st.multiselect(
                "Aplicar solo a estos rubros (vacío = todos)",
                options=sorted(df["Rubro"].unique().tolist()),
                default=[], key="esc_rubros",
            )

        # ── Cálculo del escenario ─────────────────────────────────────────────
        st.markdown("---")

        df_esc = df.copy()

        # Aplicar filtro de rubro si corresponde
        if filtro_rubro_esc:
            mask_rubro = df_esc["Rubro"].isin(filtro_rubro_esc)
        else:
            mask_rubro = pd.Series([True] * len(df_esc), index=df_esc.index)

        # Precio base (justiprecio original)
        df_esc["Precio_Sim"] = df_esc["Justiprecio"].copy()

        # Aplicar variación por clase ABC dentro del filtro de rubro
        for cls, var in [("A", var_precio_A), ("B", var_precio_B), ("C", var_precio_C)]:
            mask = mask_rubro & (df_esc["Clase_ABC_Monto"] == cls)
            df_esc.loc[mask, "Precio_Sim"] = (
                df_esc.loc[mask, "Justiprecio"] * (1 + var / 100)
            )

        # Variación global encima de la por-clase (acumulativa)
        if var_precio_global != 0:
            df_esc.loc[mask_rubro, "Precio_Sim"] = (
                df_esc.loc[mask_rubro, "Precio_Sim"] * (1 + var_precio_global / 100)
            )

        # Variación de cantidad
        df_esc["Cantidad_Sim"] = df_esc["Cantidad_Pedida"].copy()
        if var_cantidad != 0:
            df_esc.loc[mask_rubro, "Cantidad_Sim"] = (
                df_esc.loc[mask_rubro, "Cantidad_Pedida"] * (1 + var_cantidad / 100)
            )

        # Monto simulado para ítems adjudicados (no desiertos)
        mask_adj = df_esc["Es_Desierto"] == "NO"
        df_esc["Monto_Sim"] = 0.0
        df_esc.loc[mask_adj, "Monto_Sim"] = (
            df_esc.loc[mask_adj, "Precio_Sim"] *
            df_esc.loc[mask_adj, "Cantidad_Sim"]
        )

        # Incorporar desiertos si se pidió
        if inc_desiertos:
            mask_des = df_esc["Es_Desierto"] == "SI"
            df_esc.loc[mask_des, "Monto_Sim"] = (
                df_esc.loc[mask_des, "Justiprecio"] *
                df_esc.loc[mask_des, "Cantidad_Pedida"] *
                (pct_precio_des / 100)
            )

        # Resultados globales
        monto_base_esc = df_esc.loc[mask_adj, "Monto_Total"].sum()
        monto_sim_esc  = df_esc["Monto_Sim"].sum()
        delta_abs      = monto_sim_esc - monto_base_esc
        delta_pct      = delta_abs / monto_base_esc * 100 if monto_base_esc > 0 else 0

        # ── KPIs del escenario ────────────────────────────────────────────────
        e1, e2, e3, e4 = st.columns(4)
        e1.metric(
            "Presupuesto base",
            f"${monto_base_esc/1e6:.1f}M",
            help="Monto total del SOLICITUDES sin cambios (solo adjudicados).",
        )
        e2.metric(
            "Presupuesto simulado",
            f"${monto_sim_esc/1e6:.1f}M",
            delta=f"{delta_pct:+.1f}%",
            delta_color="inverse" if delta_abs > 0 else "normal",
        )
        e3.metric(
            "Diferencia absoluta",
            f"${abs(delta_abs)/1e6:.1f}M",
            delta="Aumento" if delta_abs > 0 else "Ahorro",
            delta_color="inverse" if delta_abs > 0 else "normal",
        )
        e4.metric(
            "Ítems desiertos incluidos",
            f"{(df_esc['Es_Desierto']=='SI').sum() if inc_desiertos else 0}",
            delta=f"${df_esc.loc[df_esc['Es_Desierto']=='SI','Monto_Sim'].sum()/1e6:.1f}M" if inc_desiertos else "No incluidos",
        )

        # ── Gráfico comparativo por rubro ─────────────────────────────────────
        rub_comp = (
            df_esc.groupby("Rubro")
            .agg(
                Base=("Monto_Total", lambda x: x[df_esc.loc[x.index, "Es_Desierto"] == "NO"].sum()),
                Simulado=("Monto_Sim", "sum"),
            )
            .reset_index()
        )
        rub_comp["Delta"] = rub_comp["Simulado"] - rub_comp["Base"]
        rub_comp = rub_comp.sort_values("Base", ascending=True)

        fig_esc = go.Figure()
        fig_esc.add_bar(
            y=rub_comp["Rubro"], x=rub_comp["Base"],
            name="Presupuesto base", orientation="h",
            marker_color="rgba(88,166,255,.7)",
        )
        fig_esc.add_bar(
            y=rub_comp["Rubro"],
            x=rub_comp["Delta"].clip(lower=0),
            name="Incremento simulado", orientation="h",
            marker_color="rgba(248,81,73,.75)",
        )
        fig_esc.add_bar(
            y=rub_comp["Rubro"],
            x=rub_comp["Delta"].clip(upper=0).abs(),
            name="Reducción simulada", orientation="h",
            marker_color="rgba(63,185,80,.75)",
        )
        fig_esc.update_layout({
            **CHART_LAYOUT,
            "barmode":  "stack",
            "height":   max(320, len(rub_comp) * 38),
            "title":    "Comparativa presupuesto base vs escenario simulado",
            "legend":   dict(orientation="h", y=1.05, font=dict(color="#e2e8f0")),
            "xaxis":    dict(title="Monto ($)", tickfont=dict(color="#e2e8f0"),
                             tickformat="$,.0f", gridcolor="rgba(255,255,255,.08)"),
            "yaxis":    dict(tickfont=dict(color="#e2e8f0")),
        })
        st.plotly_chart(fig_esc, use_container_width=True)

        # ── Tabla comparativa detallada ───────────────────────────────────────
        st.markdown("**Comparativa detallada por ítem**")
        with st.expander("Ver tabla completa de ítems con variación"):
            df_tabla_esc = df_esc[df_esc["Monto_Sim"] > 0].copy()
            df_tabla_esc["Variación $"] = df_tabla_esc["Monto_Sim"] - df_tabla_esc["Monto_Total"]
            df_tabla_esc["Variación %"] = np.where(
                df_tabla_esc["Monto_Total"] > 0,
                df_tabla_esc["Variación $"] / df_tabla_esc["Monto_Total"] * 100,
                0,
            )
            df_tabla_esc = df_tabla_esc.sort_values("Variación $", ascending=False)
            st.dataframe(
                df_tabla_esc[[
                    "Rubro","Descripcion","Clase_ABC_Monto","Es_Desierto",
                    "Justiprecio","Precio_Sim",
                    "Cantidad_Pedida","Cantidad_Sim",
                    "Monto_Total","Monto_Sim","Variación $","Variación %",
                ]].rename(columns={
                    "Clase_ABC_Monto": "ABC",
                    "Es_Desierto":     "Desierto",
                    "Justiprecio":     "Precio base",
                    "Precio_Sim":      "Precio sim.",
                    "Cantidad_Pedida": "Cant. base",
                    "Cantidad_Sim":    "Cant. sim.",
                    "Monto_Total":     "Monto base",
                    "Monto_Sim":       "Monto sim.",
                }).style.format({
                    "Precio base":  "${:,.2f}",
                    "Precio sim.":  "${:,.2f}",
                    "Cant. base":   "{:,.0f}",
                    "Cant. sim.":   "{:,.0f}",
                    "Monto base":   "${:,.0f}",
                    "Monto sim.":   "${:,.0f}",
                    "Variación $":  "${:+,.0f}",
                    "Variación %":  "{:+.1f}%",
                }),
                use_container_width=True,
                height=420,
                hide_index=True,
            )

        # ── Resumen por clase ABC ─────────────────────────────────────────────
        st.markdown("**Impacto por clase ABC**")
        abc_comp = (
            df_esc[df_esc["Es_Desierto"] == "NO"]
            .groupby("Clase_ABC_Monto")
            .agg(Base=("Monto_Total","sum"), Simulado=("Monto_Sim","sum"))
            .reindex(["A","B","C"]).fillna(0).reset_index()
        )
        abc_comp["Delta $"] = abc_comp["Simulado"] - abc_comp["Base"]
        abc_comp["Delta %"] = np.where(
            abc_comp["Base"] > 0,
            abc_comp["Delta $"] / abc_comp["Base"] * 100,
            0,
        )
        st.dataframe(
            abc_comp.rename(columns={"Clase_ABC_Monto":"Clase"})
            .style.format({
                "Base":     "${:,.0f}",
                "Simulado": "${:,.0f}",
                "Delta $":  "${:+,.0f}",
                "Delta %":  "{:+.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
            height=160,
        )

        # ── Descarga ──────────────────────────────────────────────────────────
        buf_esc = io.BytesIO()
        df_esc[df_esc["Monto_Sim"] > 0][[
            "Rubro","Descripcion","Clase_ABC_Monto","Es_Desierto",
            "Justiprecio","Precio_Sim","Cantidad_Pedida","Cantidad_Sim",
            "Monto_Total","Monto_Sim",
        ]].to_excel(buf_esc, index=False)
        st.download_button(
            "⬇ Descargar escenario simulado como Excel",
            data=buf_esc.getvalue(),
            file_name="escenario_presupuestario.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_esc",
        )
