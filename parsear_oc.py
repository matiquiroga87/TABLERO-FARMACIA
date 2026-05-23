"""
parsear_oc.py
=============
Parser de PDFs de Órdenes de Compra — HIGA Oscar Alende
Formato: PDF texto seleccionable generado por SIGAF/sistema provincial

Devuelve un dict con:
  - encabezado: nro_oc, año_oc, fecha, licitacion, nro_solicitud, expediente
  - proveedor:  razon_social, cuit, domicilio
  - renglones:  lista de dicts con renglon, codigo, descripcion, marca,
                cantidad, importe_unitario, importe_total
  - pie:        subtotal, iva, total
"""

import re
import pdfplumber
from typing import Optional


# ── Helpers ───────────────────────────────────────────────────────────────────

def limpiar_monto(s: str) -> float:
    """Convierte '$375.863,00' → 375863.0 (formato argentino)"""
    if not s:
        return 0.0
    s = s.strip().replace("$", "").replace(" ", "")
    # Formato argentino: punto=miles, coma=decimal
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def buscar_campo(texto: str, patron: str, grupo: int = 1) -> str:
    """Busca un patrón regex y devuelve el grupo indicado, o '' si no encuentra."""
    m = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
    return m.group(grupo).strip() if m else ""


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear_oc(ruta_pdf: str) -> dict:
    """
    Lee el PDF y devuelve la estructura completa de la OC.
    Solo procesa la primera hoja (hoja 2 = remito, se ignora).
    """
    resultado = {
        "encabezado": {},
        "proveedor":  {},
        "renglones":  [],
        "pie":        {},
        "texto_raw":  "",
        "errores":    [],
    }

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            # Solo primera página
            pagina = pdf.pages[0]
            texto = pagina.extract_text(x_tolerance=3, y_tolerance=3) or ""
            resultado["texto_raw"] = texto
    except Exception as e:
        resultado["errores"].append(f"Error al abrir PDF: {e}")
        return resultado

    # ── ENCABEZADO ────────────────────────────────────────────────────────────
    enc = {}

    # Nro. OC: "222 / 2026" o "Nro. OC 222 / 2026" o "OC N° 222/2026"
    m_oc = re.search(
        r'(?:Nro\.?\s*OC|Orden\s+de\s+Compra\s+N[°º]?|OC\s*N[°º]?)[:\s]*(\d+)\s*/\s*(\d{4})',
        texto, re.IGNORECASE
    )
    if not m_oc:
        # Fallback: buscar patrón "222 / 2026" en las primeras líneas
        m_oc = re.search(r'\b(\d{1,4})\s*/\s*(20\d{2})\b', texto[:500])

    if m_oc:
        enc["nro_oc"]  = m_oc.group(1).strip()
        enc["año_oc"]  = m_oc.group(2).strip()
        enc["oc_completo"] = f"{enc['nro_oc']}/{enc['año_oc']}"
    else:
        enc["nro_oc"] = enc["año_oc"] = enc["oc_completo"] = ""
        resultado["errores"].append("No se encontró N° de OC")

    # Fecha
    enc["fecha"] = buscar_campo(texto,
        r'Fecha[:\s]+(\d{1,2}/\d{1,2}/\d{4})')
    if not enc["fecha"]:
        # Buscar cualquier fecha en el encabezado
        m_f = re.search(r'\b(\d{1,2}/\d{1,2}/20\d{2})\b', texto[:800])
        enc["fecha"] = m_f.group(1) if m_f else ""

    # Licitación
    enc["licitacion"] = buscar_campo(texto,
        r'Licitaci[oó]n[:\s]+(.+?)(?:\n|Solicitud|Expediente)', 1)

    # N° Solicitud
    enc["nro_solicitud"] = buscar_campo(texto,
        r'Solicitud\s+Nro\.?\s*[:\s]*(\d+)')
    if not enc["nro_solicitud"]:
        enc["nro_solicitud"] = buscar_campo(texto,
            r'N[°º]\s*Solicitud[:\s]*(\d+)')

    # Expediente
    enc["expediente"] = buscar_campo(texto,
        r'Expediente[:\s]+(EX-[\w\-]+)')
    if not enc["expediente"]:
        enc["expediente"] = buscar_campo(texto,
            r'(EX-\d{4}-\d+[\w\-]*)')

    resultado["encabezado"] = enc

    # ── PROVEEDOR ─────────────────────────────────────────────────────────────
    prov = {}

    # Razón Social — buscar después de ADJUDICATARIO
    m_adj = re.search(
        r'ADJUDICATARIO[:\s]*\n?\s*(.+?)(?:\n|CUIT|C\.U\.I\.T)',
        texto, re.IGNORECASE
    )
    if m_adj:
        prov["razon_social"] = m_adj.group(1).strip()
    else:
        # Buscar patrón "Razón Social: NOMBRE"
        prov["razon_social"] = buscar_campo(texto,
            r'Raz[oó]n\s+Social[:\s]+(.+?)(?:\n|CUIT)')

    # CUIT proveedor — evitar el CUIT del hospital (30-62698339-8)
    CUIT_HOSPITAL = "30626983398"
    cuits = re.findall(r'\b(\d{2}[-\s]?\d{8}[-\s]?\d{1})\b', texto)
    cuits_limpios = [re.sub(r'[-\s]', '', c) for c in cuits]
    prov_cuits = [c for c in cuits_limpios if c != CUIT_HOSPITAL]
    prov["cuit"] = prov_cuits[0] if prov_cuits else ""

    # Domicilio
    prov["domicilio"] = buscar_campo(texto,
        r'Domicilio[:\s]+(.+?)(?:\n|Localidad|CP|C\.P\.)')

    resultado["proveedor"] = prov

    # ── RENGLONES ─────────────────────────────────────────────────────────────
    # Estrategia: buscar bloque de tabla entre encabezado de columnas y pie
    # Encabezado típico: Renglón | Código | Descripción | Cantidad | Importe
    
    renglones = []

    # Intentar extraer tabla con pdfplumber primero
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            tablas = pagina.extract_tables()
            if tablas:
                # Buscar la tabla que tenga columnas de renglón
                for tabla in tablas:
                    if not tabla or len(tabla) < 2:
                        continue
                    # Detectar encabezado de la tabla
                    header = [str(c).lower().strip() if c else "" for c in tabla[0]]
                    tiene_renglon = any("renglon" in h or "renglón" in h or "renglon" in h.replace("ó","o") for h in header)
                    tiene_codigo  = any("c" in h and ("digo" in h or "ód" in h) for h in header)
                    tiene_cantidad = any("cantidad" in h or "cant" in h for h in header)

                    if tiene_renglon or (tiene_codigo and tiene_cantidad):
                        # Mapear columnas
                        ci = {h: i for i, h in enumerate(header)}
                        
                        def col(row, names):
                            for n in names:
                                for k, i in ci.items():
                                    if n in k:
                                        v = row[i] if i < len(row) else None
                                        return str(v).strip() if v else ""
                            return ""

                        for fila in tabla[1:]:
                            if not fila or not any(fila):
                                continue
                            reng_num = col(fila, ["renglon","renglón"])
                            if not reng_num or not reng_num.isdigit():
                                continue
                            reng = {
                                "renglon":         reng_num,
                                "codigo":          col(fila, ["código","codigo","cód"]),
                                "cod_sigaf":       col(fila, ["sigaf"]),
                                "descripcion":     col(fila, ["descripcion","descripción"]),
                                "marca":           col(fila, ["marca"]),
                                "cantidad":        col(fila, ["cantidad","cant"]),
                                "importe_unitario":col(fila, ["unitario","unit"]),
                                "importe_total":   col(fila, ["total","importe total"]),
                            }
                            reng["cantidad_num"]        = float(reng["cantidad"].replace(",",".")) if reng["cantidad"] else 0
                            reng["importe_unitario_num"]= limpiar_monto(reng["importe_unitario"])
                            reng["importe_total_num"]   = limpiar_monto(reng["importe_total"])
                            renglones.append(reng)

    except Exception as e:
        resultado["errores"].append(f"Error extrayendo tabla: {e}")

    # Fallback: parsear texto línea por línea si la tabla falló
    if not renglones:
        renglones = _parsear_renglones_texto(texto, resultado["errores"])

    resultado["renglones"] = renglones

    # ── PIE ───────────────────────────────────────────────────────────────────
    pie = {}
    pie["subtotal"] = limpiar_monto(buscar_campo(texto,
        r'Subtotal[:\s]+\$?([\d.,]+)'))
    pie["iva"] = limpiar_monto(buscar_campo(texto,
        r'IVA[:\s]+\$?([\d.,]+)'))
    pie["total"] = limpiar_monto(buscar_campo(texto,
        r'TOTAL[:\s]+\$?([\d.,]+)'))
    resultado["pie"] = pie

    return resultado


def _parsear_renglones_texto(texto: str, errores: list) -> list:
    """
    Fallback: parsea renglones línea por línea cuando extract_tables() falla.
    Busca líneas que empiecen con un número de renglón.
    """
    renglones = []
    lineas = texto.split("\n")
    i = 0
    while i < len(lineas):
        linea = lineas[i].strip()
        
        # Detectar inicio de renglón: línea que empiece con número solo
        m_reng = re.match(r'^(\d{1,3})\s+([A-Z0-9]+)\s+(.+)$', linea)
        if not m_reng:
            i += 1
            continue

        reng_num  = m_reng.group(1)
        codigo    = m_reng.group(2)
        resto     = m_reng.group(3)

        # Descripción puede continuar en la siguiente línea
        descripcion = resto.strip()
        marca = ""
        cantidad = ""
        imp_unit = ""
        imp_total = ""

        # Mirar siguientes líneas para más datos del mismo renglón
        j = i + 1
        while j < len(lineas) and j < i + 8:
            sig = lineas[j].strip()
            if re.match(r'^(\d{1,3})\s+[A-Z0-9]+', sig):
                break  # Siguiente renglón
            if sig.lower().startswith("marca"):
                marca = re.sub(r'^marca[:\s]*', '', sig, flags=re.IGNORECASE).strip()
            elif re.search(r'\$[\d.,]+', sig):
                # Línea con montos
                montos = re.findall(r'\$?([\d.]+,\d{2})', sig)
                nums   = re.findall(r'^\s*(\d+(?:[.,]\d+)?)\s', sig)
                if not cantidad and nums:
                    cantidad = nums[0]
                if len(montos) >= 2:
                    imp_unit  = montos[-2]
                    imp_total = montos[-1]
                elif len(montos) == 1:
                    imp_total = montos[0]
            elif not marca and not re.search(r'\$', sig) and sig:
                descripcion += " " + sig
            j += 1

        reng = {
            "renglon":          reng_num,
            "codigo":           codigo,
            "cod_sigaf":        "",
            "descripcion":      descripcion.strip(),
            "marca":            marca,
            "cantidad":         cantidad,
            "importe_unitario": imp_unit,
            "importe_total":    imp_total,
        }
        try:
            reng["cantidad_num"]         = float(cantidad.replace(",",".")) if cantidad else 0
            reng["importe_unitario_num"] = limpiar_monto(imp_unit)
            reng["importe_total_num"]    = limpiar_monto(imp_total)
        except:
            reng["cantidad_num"] = reng["importe_unitario_num"] = reng["importe_total_num"] = 0

        renglones.append(reng)
        i = j

    if not renglones:
        errores.append("No se pudieron extraer renglones del PDF")

    return renglones


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        print("Uso: python parsear_oc.py archivo.pdf")
        sys.exit(1)
    result = parsear_oc(sys.argv[1])
    print(json.dumps(result, ensure_ascii=False, indent=2))
