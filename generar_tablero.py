"""
generar_tablero.py
==================
Script para GitHub Actions.
Lee SOLICITUDES_2026.xlsx y genera index.html con los datos embebidos.

Correcciones respecto a la versión anterior:
  - Usa index.html como plantilla (no tablero_template.html que no existe)
  - Actualiza el bloque STATS con los valores calculados del Excel nuevo
  - Actualiza la fecha en el footer
"""

import json, re, sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np

# ── Configuración de columnas ─────────────────────────────────────────────────
RUBROS_CFG = [
    ('ESTERILIZACION',3,53),('UTI',3,53),('ANESTESICOS',3,21),
    ('COMPRIMIDOS',3,80),('ANTIBIOTICOS',3,79),('PSICOTROPICOS',3,79),
    ('AMPOLLAS',3,106),('SUEROS',3,27),('DROGAS Y LIQUIDOS',3,38),
    ('ALIMENTACION',3,31),('SIES SAT URS',3,45),('BAJO COSTO',3,213),
    ('CREMAS GOTAS AEROSOLES',3,68),('DIALISIS',3,27),('RAYOS',3,33),
    ('FORMULACION Y SEDRONAR',3,32),('VARIOS',3,9),
    ('DESIERTOS',3,178),('ESTERILIZACION DESIERTOS',3,13),
]
RUBRO_COL = {
    "VARIOS":                  {"consumo":4,"cm":24,"cq":17,"cj":22,"cp":12,"caa":25,"cy":23},
    "ESTERILIZACION":          {"consumo":4,"cm":25,"cq":17,"cj":23,"cp":12,"caa":26,"cy":24},
    "ESTERILIZACION DESIERTOS":{"consumo":4,"cm":25,"cq":17,"cj":23,"cp":12,"caa":26,"cy":24},
    "DEFAULT":                 {"consumo":5,"cm":25,"cq":17,"cj":23,"cp":12,"caa":26,"cy":24},
}
HIST_COLS = {
    "ESTERILIZACION":          [2,3,4],
    "ESTERILIZACION DESIERTOS":[2,3,4],
    "VARIOS":                  [2,3,4],
    "DEFAULT":                 [2,3,4,5],
}

def sf(v):
    try: return float(v) if pd.notna(v) else 0.0
    except: return 0.0

def clasificar_xyz(hist_vals):
    vals = [v for v in hist_vals if v > 0]
    if len(vals) < 2: return "Z", 0.0, 0.0
    arr = np.array(vals, dtype=float)
    m = arr.mean()
    if m == 0: return "Z", 0.0, 0.0
    cv = arr.std() / m
    return ("X" if cv<=0.5 else "Y" if cv<=1.0 else "Z"), round(m,2), round(cv,4)

def clasificar_abc(serie, cortes=(80,95)):
    total = serie.sum()
    if total == 0: return pd.Series(["Sin datos"]*len(serie), index=serie.index)
    pct = serie.cumsum() / total * 100
    return pct.apply(lambda x: "A" if x<=cortes[0] else ("B" if x<=cortes[1] else "C"))

def procesar(ruta_xlsx):
    print(f"Leyendo: {ruta_xlsx}")
    all_sheets = pd.read_excel(ruta_xlsx, sheet_name=None, header=None)
    rows = []
    for sheet_name, r1, r2 in RUBROS_CFG:
        if sheet_name not in all_sheets: continue
        df = all_sheets[sheet_name]
        cfg  = RUBRO_COL.get(sheet_name, RUBRO_COL["DEFAULT"])
        hist = HIST_COLS.get(sheet_name, HIST_COLS["DEFAULT"])
        cm_i,cq_i,cj_i,cp_i,cons_i,caa_i,cy_i = cfg["cm"],cfg["cq"],cfg["cj"],cfg["cp"],cfg["consumo"],cfg["caa"],cfg["cy"]
        for pi in range(r1-1, r2):
            if pi >= len(df): continue
            row = df.iloc[pi]
            desc = row[1] if pd.notna(row[1]) else None
            if not desc: continue
            codigo = str(row[0]).strip() if pd.notna(row[0]) else ""
            if not codigo or codigo == "nan":
                if sheet_name == "VARIOS": codigo = f"VARIOS_{pi}"
                else: continue
            m    = sf(row[cm_i])   if cm_i   < len(row) else 0
            q    = sf(row[cq_i])   if cq_i   < len(row) else 0
            j    = sf(row[cj_i])   if cj_i   < len(row) else 0
            adj  = sf(row[caa_i])  if caa_i  < len(row) else 0
            prec = sf(row[cy_i])   if cy_i   < len(row) else 0
            if m==0 and q>0 and j>0: m = q*j
            prov = str(row[cp_i]).strip() if cp_i<len(row) and pd.notna(row[cp_i]) else ""
            oc   = str(row[15]).strip()   if len(row)>15  and pd.notna(row[15])    else ""
            cons = sf(row[cons_i]) if cons_i < len(row) else 0
            hist_vals = [sf(row[c]) for c in hist if c < len(row)]
            xyz, media, cv = clasificar_xyz(hist_vals)
            is_des = prov.lower() in ["desierto","desierta"]
            has_oc = bool(oc) and oc.lower() not in ["nan",""]
            # Limpiar saltos de línea y tabs en texto libre (pueden romper el JSON)
            desc_clean = re.sub(r'[\n\r\t]+', ' ', str(desc).strip())
            prov_clean = re.sub(r'[\n\r\t]+', ' ', prov)
            rows.append({
                "Rubro":str(sheet_name),"Codigo":codigo,
                "Descripcion":desc_clean,"Proveedor":prov_clean,
                "Consumo_Mensual":cons,"Cantidad_Pedida":q,
                "Justiprecio":j,"Monto_Total":m,"OC":oc if has_oc else "",
                "Monto_Adjudicado":adj,
                "Precio_Proveedor":prec,
                "Es_Desierto":"SI" if is_des else "NO",
                "Tiene_OC":"SI"  if has_oc else "NO",
                "Clase_XYZ":xyz,"Media_Historica":media,"CV":cv,
            })

    df_out = pd.DataFrame(rows)

    # ABC por monto
    df_m = df_out[df_out["Monto_Total"]>0].sort_values("Monto_Total",ascending=False).copy()
    df_m["Clase_ABC_Monto"] = clasificar_abc(df_m["Monto_Total"])
    df_out = df_out.merge(df_m[["Codigo","Rubro","Clase_ABC_Monto"]],on=["Codigo","Rubro"],how="left")
    df_out["Clase_ABC_Monto"] = df_out["Clase_ABC_Monto"].fillna("Sin monto")

    # ABC por cantidad
    df_q = df_out[df_out["Cantidad_Pedida"]>0].sort_values("Cantidad_Pedida",ascending=False).copy()
    df_q["Clase_ABC_Cantidad"] = clasificar_abc(df_q["Cantidad_Pedida"])
    df_out = df_out.merge(df_q[["Codigo","Rubro","Clase_ABC_Cantidad"]],on=["Codigo","Rubro"],how="left")
    df_out["Clase_ABC_Cantidad"] = df_out["Clase_ABC_Cantidad"].fillna("Sin cant.")

    df_out["Matriz_ABC_XYZ"] = df_out["Clase_ABC_Monto"] + df_out["Clase_XYZ"]
    # Deduplicación: si el mismo Código tiene descripciones distintas (ej. Sevoflurano común vs Quick Fill)
    # se trata cada combinación (Codigo, Descripcion_normalizada) como un ítem único independiente.
    # Para VARIOS_ nunca se marca como duplicado.
    df_out["_desc_norm"] = df_out["Descripcion"].str.strip().str.upper()
    df_out["Es_Duplicado"] = (
        df_out.duplicated(subset=["Codigo", "_desc_norm"], keep="first") &
        ~df_out["Codigo"].str.startswith("VARIOS_")
    ).map({True:"SI",False:"NO"})
    df_out.drop(columns=["_desc_norm"], inplace=True)

    print(f"  {len(df_out)} ítems procesados")
    print(f"  Monto total: ${df_out['Monto_Total'].sum():,.0f}")
    return df_out

def calcular_stats(df):
    """Calcula el bloque STATS que va hardcodeado en el HTML."""
    total      = len(df)
    monto      = float(df["Monto_Total"].sum())
    monto_oc   = float(df[df["Tiene_OC"]=="SI"]["Monto_Total"].sum())
    monto_des  = float(df[df["Es_Desierto"]=="SI"]["Monto_Total"].sum())
    monto_sin  = monto - monto_oc
    con_oc     = int((df["Tiene_OC"]=="SI").sum())
    des        = int((df["Es_Desierto"]=="SI").sum())
    pct_items  = round(con_oc / total * 100, 2) if total > 0 else 0
    pct_monto  = round(monto_oc / monto * 100, 2) if monto > 0 else 0

    # Monto real adjudicado: col AA, solo ítems con OC
    monto_adj  = float(df[df["Tiene_OC"]=="SI"]["Monto_Adjudicado"].sum())
    dif_adj    = round(monto_adj - monto_oc, 2)   # positivo = real > estimado
    pct_adj    = round((monto_adj / monto_oc - 1) * 100, 2) if monto_oc > 0 else 0

    # Insumos únicos
    con_cod = df[df["Codigo"] != ""].drop_duplicates(subset=["Codigo"])
    sin_cod = df[df["Codigo"] == ""].drop_duplicates(subset=["Descripcion"])
    uniq    = len(con_cod) + len(sin_cod)
    dups    = total - uniq

    return {
        "total":     total,
        "uniq":      uniq,
        "des":       des,
        "con_oc":    con_oc,
        "monto":     round(monto, 2),
        "monto_oc":  round(monto_oc, 2),
        "monto_sin": round(monto_sin, 2),
        "monto_des": round(monto_des, 2),
        "pct_items": pct_items,
        "pct_monto": pct_monto,
        "dups":      dups,
        "monto_adj": round(monto_adj, 2),
        "dif_adj":   dif_adj,
        "pct_adj":   pct_adj,
    }

def generar_html(df, plantilla_html, salida_html):
    records  = df.fillna("").to_dict(orient="records")
    json_data = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    stats    = calcular_stats(df)
    fecha    = datetime.now().strftime("%d/%m/%Y %H:%M")

    with open(plantilla_html, encoding="utf-8") as f:
        html = f.read()

    # 1. Reemplazar DATA
    html = re.sub(
        r'const DATA = \[.*?\];',
        f'const DATA = {json_data};',
        html, flags=re.DOTALL
    )

    # 2. Reemplazar STATS
    new_stats = (
        f"const STATS = {{\n"
        f"  total:      {stats['total']},\n"
        f"  uniq:       {stats['uniq']},\n"
        f"  des:        {stats['des']},\n"
        f"  con_oc:     {stats['con_oc']},\n"
        f"  monto:      {stats['monto']},\n"
        f"  monto_oc:   {stats['monto_oc']},\n"
        f"  monto_sin:  {stats['monto_sin']},\n"
        f"  monto_des:  {stats['monto_des']},\n"
        f"  pct_items:  {stats['pct_items']},\n"
        f"  pct_monto:  {stats['pct_monto']},\n"
        f"  dups:       {stats['dups']},\n"
        f"  monto_adj:  {stats['monto_adj']},\n"
        f"  dif_adj:    {stats['dif_adj']},\n"
        f"  pct_adj:    {stats['pct_adj']}\n"
        f"}};"
    )
    html = re.sub(
        r'const STATS = \{.*?\};',
        new_stats,
        html, flags=re.DOTALL
    )

    # 3. Actualizar fecha en el footer
    html = re.sub(
        r'(<span id="data-date">)[^<]*(</span>)',
        rf'\g<1>{fecha}\g<2>',
        html
    )

    with open(salida_html, "w", encoding="utf-8") as f:
        f.write(html)

    size = Path(salida_html).stat().st_size / 1024
    print(f"  HTML generado: {salida_html} ({size:.0f} KB)")
    print(f"  STATS actualizadas: {stats['total']} renglones, {stats['uniq']} únicos, ${stats['monto']:,.0f}")

if __name__ == "__main__":
    # Buscar el archivo Excel
    xlsx_files = sorted(Path(".").glob("SOLICITUDES*.xlsx"))
    if not xlsx_files:
        print("ERROR: No se encontró ningún archivo SOLICITUDES*.xlsx")
        sys.exit(1)

    xlsx_path = xlsx_files[0]
    print(f"Archivo encontrado: {xlsx_path}")

    # Verificar que existe el index.html como plantilla
    if not Path("index.html").exists():
        print("ERROR: No se encontró index.html (necesario como plantilla)")
        sys.exit(1)

    df = procesar(xlsx_path)
    # Usa index.html como plantilla Y como salida (se sobreescribe)
    generar_html(df, "index.html", "index.html")
    print("\nListo. index.html actualizado.")
