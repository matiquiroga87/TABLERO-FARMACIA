"""
generar_tablero.py
==================
Script simplificado para GitHub Actions.
Lee SOLICITUDES_2026.xlsx y genera index.html con los datos embebidos.

No requiere argumentos — GitHub Actions lo llama directamente.
"""

import json, csv, re, sys
from pathlib import Path
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
    "VARIOS":                  {"consumo":4,"cm":24,"cq":17,"cj":22,"cp":12},
    "ESTERILIZACION":          {"consumo":4,"cm":25,"cq":17,"cj":23,"cp":12},
    "ESTERILIZACION DESIERTOS":{"consumo":4,"cm":25,"cq":17,"cj":23,"cp":12},
    "DEFAULT":                 {"consumo":5,"cm":25,"cq":17,"cj":23,"cp":12},
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
        cfg = RUBRO_COL.get(sheet_name, RUBRO_COL["DEFAULT"])
        hist = HIST_COLS.get(sheet_name, HIST_COLS["DEFAULT"])
        cm_i,cq_i,cj_i,cp_i,cons_i = cfg["cm"],cfg["cq"],cfg["cj"],cfg["cp"],cfg["consumo"]
        for pi in range(r1-1, r2):
            if pi >= len(df): continue
            row = df.iloc[pi]
            desc = row[1] if pd.notna(row[1]) else None
            if not desc: continue
            codigo = str(row[0]).strip() if pd.notna(row[0]) else ""
            if not codigo or codigo == "nan":
                if sheet_name == "VARIOS": codigo = f"VARIOS_{pi}"
                else: continue
            m  = sf(row[cm_i])  if cm_i  < len(row) else 0
            q  = sf(row[cq_i])  if cq_i  < len(row) else 0
            j  = sf(row[cj_i])  if cj_i  < len(row) else 0
            if m==0 and q>0 and j>0: m = q*j
            prov = str(row[cp_i]).strip() if cp_i<len(row) and pd.notna(row[cp_i]) else ""
            oc   = str(row[15]).strip()   if len(row)>15  and pd.notna(row[15])    else ""
            cons = sf(row[cons_i]) if cons_i < len(row) else 0
            hist_vals = [sf(row[c]) for c in hist if c < len(row)]
            xyz, media, cv = clasificar_xyz(hist_vals)
            is_des = prov.lower() in ["desierto","desierta"]
            has_oc = bool(oc) and oc.lower() not in ["nan",""]
            rows.append({
                "Rubro":str(sheet_name),"Codigo":codigo,
                "Descripcion":str(desc).strip(),"Proveedor":prov,
                "Consumo_Mensual":cons,"Cantidad_Pedida":q,
                "Justiprecio":j,"Monto_Total":m,"OC":oc if has_oc else "",
                "Es_Desierto":"SI" if is_des else "NO",
                "Tiene_OC":"SI" if has_oc else "NO",
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
    df_out["Es_Duplicado"] = (
        df_out.duplicated(subset=["Codigo"], keep="first") &
        ~df_out["Codigo"].str.startswith("VARIOS_")
    ).map({True:"SI",False:"NO"})

    print(f"  {len(df_out)} ítems procesados")
    print(f"  Monto total: ${df_out['Monto_Total'].sum():,.0f}")
    return df_out

def generar_html(df, plantilla_html, salida_html):
    rows = df.fillna("").to_dict(orient="records")
    json_data = json.dumps(rows, ensure_ascii=False)
    with open(plantilla_html, encoding="utf-8") as f:
        template = f.read()
    # Replace embedded data
    result = re.sub(
        r'const DATA = \[.*?\];',
        f'const DATA = {json_data};',
        template, flags=re.DOTALL
    )
    with open(salida_html, "w", encoding="utf-8") as f:
        f.write(result)
    size = Path(salida_html).stat().st_size / 1024
    print(f"  HTML generado: {salida_html} ({size:.0f} KB)")

if __name__ == "__main__":
    # Buscar el archivo Excel en la carpeta actual
    xlsx_files = list(Path(".").glob("SOLICITUDES*.xlsx"))
    if not xlsx_files:
        print("ERROR: No se encontró ningún archivo SOLICITUDES*.xlsx")
        sys.exit(1)

    xlsx_path = xlsx_files[0]
    print(f"Archivo encontrado: {xlsx_path}")

    df = procesar(xlsx_path)
    generar_html(df, "tablero_template.html", "index.html")
    print("\nListo. index.html actualizado.")
