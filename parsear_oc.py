"""
parsear_oc.py
=============
Parser de PDFs de Órdenes de Compra — HIGA Oscar Alende

Tipos de licitación reconocidos:
  - Licitación Privada Nro. X / YYYY
  - Procedimiento Abreviado Nro. X / YYYY

Bloque ADJUDICATARIO:
  - Razón Social (campo "Razón Social:" o primer texto después de ADJUDICATARIO)
  - CUIT proveedor (distinto al del hospital: 30-62698339-8)

Renglones — campos que se extraen:
  renglon          → número de renglón
  codigo           → código vademecum (col A del SOLICITUDES)
  cod_sigaf        → código SIGAF (referencia, no se carga al Excel)
  descripcion      → SOLO la descripción del medicamento (limpiada)
  cantidad         → cantidad numérica
  importe_unitario → precio unitario (formato argentino $X.XXX,XX)
  importe_total    → importe total   (formato argentino $X.XXX,XX)

Cambios vs versión anterior:
  - Se eliminó el campo "marca" (no figura en la OC)
  - La descripción se limpia de cantidad, importe y otros datos pegados
  - Se reconocen ambos tipos de licitación
  - El proveedor se extrae del campo "Razón Social" dentro del bloque ADJUDICATARIO
  - La cantidad se extrae de la celda correcta (no de la descripción)
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
    try:
        return float(s)
    except ValueError:
        return 0.0


def limpiar_cantidad(s: str) -> float:
    """Convierte '1.500' o '1500' o '1.500,00' → 1500.0"""
    if not s:
        return 0.0
    s = str(s).strip().replace(" ", "")
    # Si tiene coma decimal: formato argentino
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        # Solo puntos → separador de miles
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def buscar_campo(texto: str, patron: str, grupo: int = 1) -> str:
    m = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
    return m.group(grupo).strip() if m else ""


def limpiar_descripcion(texto: str) -> str:
    """
    Elimina de la celda de descripción todo lo que no sea la descripción:
    - Números de cantidad pegados al inicio o al final
    - Importes con $ o formato numérico
    - Palabras clave como MARCA:, COD., SIGAF, etc.
    - Códigos alfanuméricos de tipo M1234567 o 01234-5678
    """
    # Eliminar importes ($X.XXX,XX o X.XXX,XX)
    texto = re.sub(r'\$?\s*\d{1,3}(?:\.\d{3})*,\d{2}', '', texto)
    # Eliminar cantidades puras al final (número seguido de fin o espacio+número)
    texto = re.sub(r'\s+\d+(?:[.,]\d+)?\s*$', '', texto)
    texto = re.sub(r'^\s*\d+(?:[.,]\d+)?\s+', '', texto)
    # Eliminar etiquetas de marca/código pegadas
    texto = re.sub(r'(?:MARCA|BRAND|COD\.?|COD\.?\s*SIGAF)[:\s]\S+', '', texto, flags=re.IGNORECASE)
    # Eliminar códigos tipo 01894-0074 que a veces se pegan a la descripción
    texto = re.sub(r'\b\d{5}-\d{4}\b', '', texto)
    # Limpiar espacios múltiples
    texto = re.sub(r'\s{2,}', ' ', texto).strip()
    return texto


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear_oc(ruta_pdf: str) -> dict:
    """
    Lee la primera página del PDF y devuelve la estructura completa de la OC.
    La segunda página (remito/firma) se ignora.
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
            pagina = pdf.pages[0]
            texto  = pagina.extract_text(x_tolerance=3, y_tolerance=3) or ""
            resultado["texto_raw"] = texto
    except Exception as e:
        resultado["errores"].append(f"Error al abrir PDF: {e}")
        return resultado

    # ── ENCABEZADO ────────────────────────────────────────────────────────────
    enc = {}

    # N° OC: "Nro. OC 222 / 2026"
    m_oc = re.search(
        r'(?:Nro\.?\s*OC|Orden\s+de\s+Compra\s+N[°º]?|OC\s*N[°º]?)'
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
    enc["fecha"] = buscar_campo(texto, r'Fecha[:\s]+(\d{1,2}/\d{1,2}/\d{4})')
    if not enc["fecha"]:
        m_f = re.search(r'\b(\d{1,2}/\d{1,2}/20\d{2})\b', texto[:800])
        enc["fecha"] = m_f.group(1) if m_f else ""

    # Tipo de licitación — reconoce ambos formatos:
    #   "Licitación Privada Nro. 5 / 2026"
    #   "Procedimiento Abreviado Nro. 12 / 2026"
    m_lic = re.search(
        r'((?:Licitaci[oó]n\s+Privada|Procedimiento\s+Abreviado)'
        r'(?:\s+Nro\.?\s*\d+\s*/\s*\d{4})?)',
        texto, re.IGNORECASE
    )
    enc["licitacion"] = m_lic.group(1).strip() if m_lic else buscar_campo(
        texto, r'Licitaci[oó]n[:\s]+(.+?)(?:\n|Solicitud|Expediente)')

    # N° Solicitud
    enc["nro_solicitud"] = buscar_campo(texto,
        r'Solicitud\s+Nro\.?\s*[:\s]*(\d+)')
    if not enc["nro_solicitud"]:
        enc["nro_solicitud"] = buscar_campo(texto,
            r'N[°º]\s*Solicitud[:\s]*(\d+)')

    # Expediente
    enc["expediente"] = buscar_campo(texto, r'(EX-\d{4}-\d+[\w\-]*)')

    resultado["encabezado"] = enc

    # ── PROVEEDOR ─────────────────────────────────────────────────────────────
    prov = {}

    # Razón Social: buscar dentro del bloque ADJUDICATARIO
    # El campo aparece como "Razón Social INNOVATE PHARMA S.A." o
    # "Razón Social: INNOVATE PHARMA S.A."
    m_rs = re.search(
        r'Raz[oó]n\s+Social[:\s]+(.+?)(?:\n|CUIT|C\.U\.I\.T\.)',
        texto, re.IGNORECASE
    )
    if m_rs:
        prov["razon_social"] = m_rs.group(1).strip()
    else:
        # Fallback: primer texto después de ADJUDICATARIO antes del CUIT
        m_adj = re.search(
            r'ADJUDICATARIO[:\s]*\n?\s*(.+?)(?:\n|CUIT|C\.U\.I\.T)',
            texto, re.IGNORECASE
        )
        prov["razon_social"] = m_adj.group(1).strip() if m_adj else ""

    # CUIT proveedor — excluir el CUIT del hospital
    CUIT_HOSPITAL = "30626983398"
    cuits = re.findall(r'\b(\d{2}[-\s]?\d{8}[-\s]?\d{1})\b', texto)
    cuits_limpios = [re.sub(r'[-\s]', '', c) for c in cuits]
    prov["cuit"] = next((c for c in cuits_limpios if c != CUIT_HOSPITAL), "")

    resultado["proveedor"] = prov

    # ── RENGLONES ─────────────────────────────────────────────────────────────
    renglones = []

    # Intento 1: extract_tables() de pdfplumber
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            pagina = pdf.pages[0]
            tablas = pagina.extract_tables()
            for tabla in (tablas or []):
                if not tabla or len(tabla) < 2:
                    continue
                header = [str(c).lower().strip() if c else "" for c in tabla[0]]
                # Normalizar tildes para comparar
                header_norm = [h.replace("ó","o").replace("é","e")
                                .replace("ú","u").replace("í","i") for h in header]

                tiene_renglon  = any("renglon" in h for h in header_norm)
                tiene_cantidad = any("cantidad" in h for h in header_norm)

                if not (tiene_renglon and tiene_cantidad):
                    continue

                # Mapeo de columnas: busca la primera que contenga el keyword
                def idx_col(keywords):
                    for kw in keywords:
                        for i, h in enumerate(header_norm):
                            if kw in h:
                                return i
                    return None

                i_reng  = idx_col(["renglon"])
                i_cod   = idx_col(["codigo", "cod"])
                i_desc  = idx_col(["descripcion"])
                i_cant  = idx_col(["cantidad"])
                i_uunit = idx_col(["unitario", "unit"])
                i_total = idx_col(["importe total", "total"])

                def celda(fila, idx):
                    if idx is None or idx >= len(fila):
                        return ""
                    v = fila[idx]
                    return str(v).strip() if v else ""

                for fila in tabla[1:]:
                    if not fila or not any(fila):
                        continue
                    reng_num = celda(fila, i_reng)
                    if not reng_num or not re.match(r'^\d+$', reng_num.strip()):
                        continue

                    desc_raw  = celda(fila, i_desc)
                    cant_raw  = celda(fila, i_cant)
                    imp_u_raw = celda(fila, i_uunit)
                    imp_t_raw = celda(fila, i_total)

                    # Limpiar descripción
                    descripcion = limpiar_descripcion(desc_raw)

                    renglones.append({
                        "renglon":              reng_num.strip(),
                        "codigo":               celda(fila, i_cod),
                        "cod_sigaf":            "",
                        "descripcion":          descripcion,
                        "cantidad":             cant_raw,
                        "importe_unitario":     imp_u_raw,
                        "importe_total":        imp_t_raw,
                        "cantidad_num":         limpiar_cantidad(cant_raw),
                        "importe_unitario_num": limpiar_monto(imp_u_raw),
                        "importe_total_num":    limpiar_monto(imp_t_raw),
                    })
                break  # Primera tabla válida encontrada

    except Exception as e:
        resultado["errores"].append(f"Error extrayendo tabla: {e}")

    # Intento 2: fallback texto línea por línea
    if not renglones:
        renglones = _parsear_renglones_texto(texto, resultado["errores"])

    resultado["renglones"] = renglones

    # ── PIE ───────────────────────────────────────────────────────────────────
    pie = {}
    pie["subtotal"] = limpiar_monto(buscar_campo(texto,
        r'Subtotal[:\s]+\$?([\d.,]+)'))
    pie["iva"]      = limpiar_monto(buscar_campo(texto,
        r'IVA[:\s]+\$?([\d.,]+)'))
    pie["total"]    = limpiar_monto(buscar_campo(texto,
        r'TOTAL[:\s]+\$?([\d.,]+)'))
    resultado["pie"] = pie

    return resultado


# ── Fallback: parseo línea por línea ─────────────────────────────────────────

def _parsear_renglones_texto(texto: str, errores: list) -> list:
    """
    Fallback cuando extract_tables() no encuentra tabla estructurada.
    Parsea el texto buscando líneas que empiecen con número de renglón + código.

    Formato esperado por línea:
      33  M3201130  RITUXIMAB FCO.AMPOLLA...
      (siguiente línea puede tener cantidad e importes)
    """
    renglones = []
    lineas    = texto.split("\n")
    i = 0

    while i < len(lineas):
        linea = lineas[i].strip()

        # Inicio de renglón: número + código alfanumérico + texto
        m_reng = re.match(r'^(\d{1,3})\s+([A-Z0-9]{5,})\s+(.+)$', linea)
        if not m_reng:
            i += 1
            continue

        reng_num    = m_reng.group(1)
        codigo      = m_reng.group(2)
        desc_inicio = m_reng.group(3).strip()

        descripcion  = desc_inicio
        cantidad_str = ""
        imp_unit_str = ""
        imp_total_str= ""

        j = i + 1
        while j < len(lineas) and j < i + 10:
            sig = lineas[j].strip()

            # Si empieza otro renglón, parar
            if re.match(r'^(\d{1,3})\s+[A-Z0-9]{5,}', sig):
                break

            # Línea con importes: contiene $X.XXX,XX
            if re.search(r'\$[\d.]+,\d{2}', sig):
                # Extraer cantidad (número entero al inicio de la línea)
                m_cant = re.match(r'^(\d[\d.]*)\s', sig)
                if m_cant and not cantidad_str:
                    cantidad_str = m_cant.group(1)

                # Extraer importes
                importes = re.findall(r'\$?([\d.]+,\d{2})', sig)
                if len(importes) >= 2:
                    imp_unit_str  = importes[-2]
                    imp_total_str = importes[-1]
                elif len(importes) == 1:
                    imp_total_str = importes[0]

            # Línea solo con número → puede ser la cantidad
            elif re.match(r'^[\d.]+$', sig) and not cantidad_str:
                cantidad_str = sig

            # Texto adicional de descripción (sin números, sin $)
            elif not re.search(r'[\$\d]', sig) and sig:
                descripcion += " " + sig

            j += 1

        renglones.append({
            "renglon":              reng_num,
            "codigo":               codigo,
            "cod_sigaf":            "",
            "descripcion":          limpiar_descripcion(descripcion),
            "cantidad":             cantidad_str,
            "importe_unitario":     imp_unit_str,
            "importe_total":        imp_total_str,
            "cantidad_num":         limpiar_cantidad(cantidad_str),
            "importe_unitario_num": limpiar_monto(imp_unit_str),
            "importe_total_num":    limpiar_monto(imp_total_str),
        })
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
