"""
parsear_oc.py
=============
Parser de PDFs de Órdenes de Compra — HIGA Oscar Alende

Estructura del PDF (SIGAF Buenos Aires):
  - La tabla de renglones es detectada por pdfplumber como UNA SOLA FILA
    donde cada columna contiene todos los valores separados por \\n.
  - La estrategia correcta es: detectar esa fila fusionada y hacer zip
    por columnas (Renglón, Código, Cantidad, Imp.Unitario, Imp.Total).
  - La descripción se extrae del texto crudo usando los números de renglón
    como anclas, ya que pdfplumber la fusiona en un bloque único.
  - Fallback: parseo línea por línea del texto crudo.
"""

import re
import pdfplumber
from typing import Optional


# ── Helpers ───────────────────────────────────────────────────────────────────

def limpiar_monto(s: str) -> float:
    """Convierte '$375.863,00' o '375.863,00' → 375863.0 (formato argentino)."""
    if not s:
        return 0.0
    s = str(s).strip().replace("$", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    # Quitar decimales extras (ej: 2.780,0000 → 2780.0)
    try:
        return float(s)
    except ValueError:
        return 0.0


def limpiar_cantidad(s: str) -> float:
    """Convierte '1.500' o '1500' o '1.500,00' → 1500.0"""
    if not s:
        return 0.0
    s = str(s).strip().replace("$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def buscar_campo(texto: str, patron: str, grupo: int = 1) -> str:
    m = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
    return m.group(grupo).strip() if m else ""


def limpiar_descripcion(texto: str) -> str:
    """Limpia la descripción de cantidades, importes y etiquetas pegadas."""
    texto = re.sub(r'\$?\s*\d{1,3}(?:\.\d{3})*,\d{2,4}', '', texto)
    texto = re.sub(r'\s+\d+(?:[.,]\d+)?\s*$', '', texto)
    texto = re.sub(r'^\s*\d+(?:[.,]\d+)?\s+', '', texto)
    texto = re.sub(r'(?:MARCA|BRAND|COD\.?|COD\.?\s*SIGAF)[:\s]\S+', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'\b\d{5}-\d{4}\b', '', texto)
    texto = re.sub(r'\s{2,}', ' ', texto).strip()
    return texto


def extraer_descripcion_desde_texto(texto: str, renglon: str, siguiente_renglon: str = None) -> str:
    """
    Extrae la descripción de un renglón usando el texto crudo.
    Busca el bloque entre el renglón actual y el siguiente.
    El formato de cada línea de renglón en el texto crudo es:
      3 M2022228 99994 DESCRIPCION... 2.5.2 1.500 $ X,XX $ X,XX
    """
    # Patrón: número de renglón + código + descripción hasta fin de la línea
    patron_reng = (
        r'^\s*' + re.escape(str(renglon)) +
        r'\s+[A-Z0-9]{4,}\s+\S+\s+'   # código + cód.sigaf
        r'(.+?)(?:\s+\d+\.\d+\.\d+)'  # descripción hasta imputación (X.X.X)
    )
    m = re.search(patron_reng, texto, re.MULTILINE)
    if m:
        desc = m.group(1).strip()
        # Limpiar cantidad e importes que pudieran haberse colado
        desc = re.sub(r'\s+[\d.]+(?:,\d+)?\s+\$.*$', '', desc)
        return desc.strip()

    # Fallback: buscar la línea completa y extraer la parte de descripción
    patron_simple = r'^\s*' + re.escape(str(renglon)) + r'\s+[A-Z0-9]{4,}\s+\S+\s+(.+)$'
    m2 = re.search(patron_simple, texto, re.MULTILINE)
    if m2:
        linea = m2.group(1)
        # Quitar desde la imputación en adelante (patrón X.X.X)
        linea = re.split(r'\s+\d+\.\d+\.\d+', linea)[0]
        return linea.strip()

    return ""


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear_oc(ruta_pdf: str) -> dict:
    """
    Lee el PDF y devuelve la estructura completa de la OC.
    Estrategia para tablas SIGAF (columnas fusionadas en una sola fila):
      zip de columnas Renglón / Código / Cantidad / Imp.Unitario / Imp.Total
    La descripción se extrae del texto crudo.
    """
    resultado = {
        "encabezado": {},
        "proveedor":  {},
        "renglones":  [],
        "pie":        {},
        "texto_raw":  "",
        "errores":    [],
    }

    # ── Extraer texto de página 1 ─────────────────────────────────────────────
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            texto  = pagina.extract_text(x_tolerance=3, y_tolerance=3) or ""
            resultado["texto_raw"] = texto
    except Exception as e:
        resultado["errores"].append(f"Error al abrir PDF: {e}")
        return resultado

    # ── ENCABEZADO ────────────────────────────────────────────────────────────
    enc = {}

    # N° OC: "Nro. de O.C.:283 / 2026" o "Nro. OC 222 / 2026"
    m_oc = re.search(
        r'(?:Nro\.?\s*(?:de\s*)?O\.?C\.?|Orden\s+de\s+Compra\s+N[°º]?|OC\s*N[°º]?)'
        r'[:\s]*(\d+)\s*/\s*(\d{4})',
        texto, re.IGNORECASE
    )
    if not m_oc:
        m_oc = re.search(r'\b(\d{1,4})\s*/\s*(20\d{2})\b', texto[:500])

    if m_oc:
        enc["nro_oc"]      = m_oc.group(1).strip()
        enc["año_oc"]      = m_oc.group(2).strip()
        enc["oc_completo"] = f"{enc['nro_oc']}/{enc['año_oc']}"
    else:
        enc["nro_oc"] = enc["año_oc"] = enc["oc_completo"] = ""
        resultado["errores"].append("No se encontró N° de OC")

    # Fecha
    enc["fecha"] = buscar_campo(texto, r'FECHA[:\s]+(\d{1,2}/\d{1,2}/\d{4})')
    if not enc["fecha"]:
        m_f = re.search(r'\b(\d{1,2}/\d{1,2}/20\d{2})\b', texto[:800])
        enc["fecha"] = m_f.group(1) if m_f else ""

    # Tipo de licitación
    m_lic = re.search(
        r'((?:Licitaci[oó]n\s+Privada|Procedimiento\s+Abreviado)'
        r'(?:\s+Nro\.?\s*\d+\s*/\s*\d{4})?)',
        texto, re.IGNORECASE
    )
    enc["licitacion"] = m_lic.group(1).strip() if m_lic else ""

    # N° Solicitud
    enc["nro_solicitud"] = buscar_campo(texto,
        r'SOLICITUD\s+Nro\.?\s*[:\s]*(\d+)')
    if not enc["nro_solicitud"]:
        enc["nro_solicitud"] = buscar_campo(texto, r'N[°º]\s*Solicitud[:\s]*(\d+)')

    # Expediente
    enc["expediente"] = buscar_campo(texto, r'(EX-\d{4}-\d+[\w\-]*)')

    resultado["encabezado"] = enc

    # ── PROVEEDOR ─────────────────────────────────────────────────────────────
    prov = {}

    # Razón Social: está en la misma línea que Destino, separados por espacios
    # Formato: "Razón Social: R.C. RADIOLOGIA CASTELAR S.R.L Destino: ..."
    # Se corta ante 2+ espacios (separador de columnas) o palabras clave
    m_rs = re.search(
        r'Raz[oó]n\s+Social:\s*(.+?)(?:\s{2,}|(?:\s+)(?:Destino:|Domicilio:|CUIT:|Registro\s+de|CBU:))',
        texto, re.IGNORECASE
    )
    if m_rs:
        prov["razon_social"] = m_rs.group(1).strip()
    else:
        # Fallback: capturar hasta fin de línea
        m_rs2 = re.search(r'Raz[oó]n\s+Social:\s*(.+?)$', texto, re.IGNORECASE | re.MULTILINE)
        razon = m_rs2.group(1).strip() if m_rs2 else ""
        # Quitar "Destino:" y todo lo que sigue si quedó pegado
        razon = re.split(r'\s+Destino:', razon, flags=re.IGNORECASE)[0].strip()
        prov["razon_social"] = razon

    # CUIT proveedor — excluir el CUIT del hospital
    CUIT_HOSPITAL = "30626983398"
    cuits = re.findall(r'\b(\d{2}[-\s]?\d{8}[-\s]?\d{1})\b', texto)
    cuits_limpios = [re.sub(r'[-\s]', '', c) for c in cuits]
    prov["cuit"] = next((c for c in cuits_limpios if c != CUIT_HOSPITAL), "")

    resultado["proveedor"] = prov

    # ── RENGLONES ─────────────────────────────────────────────────────────────
    renglones = []

    # Estrategia A: tabla SIGAF con columnas fusionadas en una sola fila
    # pdfplumber devuelve la tabla como 2 filas: header + 1 fila con todo junto
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            tablas = pagina.extract_tables()

            for tabla in (tablas or []):
                if not tabla or len(tabla) < 2:
                    continue

                header = [str(c).lower().strip() if c else "" for c in tabla[0]]
                header_norm = [
                    h.replace("ó","o").replace("é","e").replace("ú","u")
                     .replace("í","i").replace("á","a") for h in header
                ]

                # Verificar que tiene las columnas mínimas necesarias
                tiene_renglon  = any(kw in h for h in header_norm for kw in ["renglon","reng"])
                tiene_cantidad = any("cantidad" in h for h in header_norm)
                if not (tiene_renglon and tiene_cantidad):
                    continue

                def idx_col(keywords):
                    for kw in keywords:
                        for i, h in enumerate(header_norm):
                            if kw in h:
                                return i
                    return None

                i_reng  = idx_col(["renglon", "reng"])
                i_cod   = idx_col(["codigo", "cod"])
                i_cant  = idx_col(["cantidad"])
                i_uunit = idx_col(["unitario", "unit", "precio"])
                i_total = None
                # Importe Total: evitar confundir con Unitario
                for i, h in enumerate(header_norm):
                    if "total" in h and i != i_uunit:
                        i_total = i
                        break
                if i_total is None:
                    i_total = idx_col(["importe total", "total"])

                def celda(fila, idx):
                    if idx is None or idx >= len(fila):
                        return ""
                    v = fila[idx]
                    return str(v).strip() if v else ""

                for fila in tabla[1:]:
                    if not fila or not any(fila):
                        continue

                    col_reng = celda(fila, i_reng)
                    col_cod  = celda(fila, i_cod)
                    col_cant = celda(fila, i_cant)
                    col_uunt = celda(fila, i_uunit)
                    col_tot  = celda(fila, i_total)

                    # Detectar si es la fila "fusionada SIGAF":
                    # la celda de Renglón contiene múltiples valores separados por \n
                    rengs_lista  = [r.strip() for r in col_reng.split("\n") if r.strip() and re.match(r'^\d+$', r.strip())]
                    cods_lista   = [c.strip() for c in col_cod.split("\n")  if c.strip() and re.match(r'^[A-Z0-9]{4,}$', c.strip())]
                    cants_lista  = [c.strip() for c in col_cant.split("\n") if c.strip()]
                    uunts_lista  = [c.strip() for c in col_uunt.split("\n") if c.strip()]
                    tots_lista   = [c.strip() for c in col_tot.split("\n")  if c.strip()]

                    es_fusionada = len(rengs_lista) > 1

                    if es_fusionada:
                        # Zip por columnas — la descripción se extrae del texto crudo
                        n = len(rengs_lista)
                        # Asegurar que todas las listas tengan longitud n
                        def pad(lst, n):
                            return (lst + [""] * n)[:n]
                        cods_lista  = pad(cods_lista, n)
                        cants_lista = pad(cants_lista, n)
                        uunts_lista = pad(uunts_lista, n)
                        tots_lista  = pad(tots_lista, n)

                        siguientes = rengs_lista[1:] + [None]
                        for reng_num, cod, cant, uunt, tot, sig in zip(
                            rengs_lista, cods_lista, cants_lista,
                            uunts_lista, tots_lista, siguientes
                        ):
                            desc = extraer_descripcion_desde_texto(texto, reng_num, sig)
                            renglones.append({
                                "renglon":              reng_num,
                                "codigo":               cod,
                                "cod_sigaf":            "",
                                "descripcion":          limpiar_descripcion(desc),
                                "cantidad":             cant,
                                "importe_unitario":     uunt,
                                "importe_total":        tot,
                                "cantidad_num":         limpiar_cantidad(cant),
                                "importe_unitario_num": limpiar_monto(uunt),
                                "importe_total_num":    limpiar_monto(tot),
                            })
                        break  # Tabla procesada

                    else:
                        # Fila normal (1 renglón por fila)
                        reng_num = col_reng.strip()
                        if not reng_num or not re.match(r'^\d+$', reng_num):
                            continue
                        desc = extraer_descripcion_desde_texto(texto, reng_num)
                        renglones.append({
                            "renglon":              reng_num,
                            "codigo":               col_cod,
                            "cod_sigaf":            "",
                            "descripcion":          limpiar_descripcion(desc) or col_cod,
                            "cantidad":             col_cant,
                            "importe_unitario":     col_uunt,
                            "importe_total":        col_tot,
                            "cantidad_num":         limpiar_cantidad(col_cant),
                            "importe_unitario_num": limpiar_monto(col_uunt),
                            "importe_total_num":    limpiar_monto(col_tot),
                        })

    except Exception as e:
        resultado["errores"].append(f"Error extrayendo tabla: {e}")

    # Estrategia B: fallback texto línea por línea
    if not renglones:
        renglones = _parsear_renglones_texto(texto, resultado["errores"])

    resultado["renglones"] = renglones

    # ── PIE ───────────────────────────────────────────────────────────────────
    pie = {}
    pie["subtotal"] = limpiar_monto(buscar_campo(texto, r'SubTotal[:\s]+\$?([\d.,]+)'))
    pie["iva"]      = limpiar_monto(buscar_campo(texto, r'I\.V\.A\.[:\s]+\$?([\d.,]+)'))
    pie["total"]    = limpiar_monto(buscar_campo(texto, r'TOTAL[:\s]+\$?([\d.,]+)'))
    resultado["pie"] = pie

    return resultado


# ── Fallback: parseo línea por línea ─────────────────────────────────────────

def _parsear_renglones_texto(texto: str, errores: list) -> list:
    """
    Fallback cuando extract_tables() no produce resultados útiles.
    Formato por línea del texto crudo SIGAF:
      3 M2022228 99994 DESCRIPCION... 2.5.2 1.500 $ X,XX $ X,XX
    """
    renglones = []
    lineas    = texto.split("\n")
    i = 0

    while i < len(lineas):
        linea = lineas[i].strip()

        # Renglón: número + código + (cód.sigaf opcional) + descripción
        # El importe unitario puede tener 4 decimales: $2.780,0000
        m_reng = re.match(
            r'^(\d{1,3})\s+([A-Z][A-Z0-9]{3,})\s+(\S+)\s+(.+?)\s+'
            r'(\d+\.\d+\.\d+)\s+'           # imputación X.X.X
            r'([\d.]+(?:,\d+)?)\s+'         # cantidad
            r'\$\s*([\d.]+,\d+)\s+'         # importe unitario
            r'\$\s*([\d.]+,\d+)',            # importe total
            linea
        )
        if m_reng:
            renglones.append({
                "renglon":              m_reng.group(1),
                "codigo":               m_reng.group(2),
                "cod_sigaf":            m_reng.group(3),
                "descripcion":          limpiar_descripcion(m_reng.group(4)),
                "cantidad":             m_reng.group(6),
                "importe_unitario":     m_reng.group(7),
                "importe_total":        m_reng.group(8),
                "cantidad_num":         limpiar_cantidad(m_reng.group(6)),
                "importe_unitario_num": limpiar_monto(m_reng.group(7)),
                "importe_total_num":    limpiar_monto(m_reng.group(8)),
            })
            i += 1
            continue

        i += 1

    if not renglones:
        errores.append("No se pudieron extraer renglones del PDF (fallback también falló)")

    return renglones


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        print("Uso: python parsear_oc.py archivo.pdf")
        sys.exit(1)
    result = parsear_oc(sys.argv[1])
    print(json.dumps(result, ensure_ascii=False, indent=2))
