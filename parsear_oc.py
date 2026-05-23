"""
parsear_oc.py
=============
Parser de PDFs de Órdenes de Compra — HIGA Oscar Alende

Estructura del PDF (SIGAF Buenos Aires):
  - Los renglones pueden estar distribuidos en VARIAS páginas de contenido.
    Se procesan todas las páginas que contengan tabla de renglones
    (Renglón/Código/Cantidad/Importe), ignorando páginas de firma y carátula.
  - Cada tabla de renglones es detectada por pdfplumber como UNA SOLA FILA
    donde cada columna contiene todos los valores separados por \\n (columnas
    fusionadas). Se hace zip por columnas para reconstruir cada renglón.
  - La descripción se extrae del texto crudo acumulado de todas las páginas,
    usando el número de renglón como ancla.
  - Se filtran filas espurias (encabezado SAF, subtotales, afectación).
  - Se limpian marcas/nombres de empresa que pdfplumber pega a la descripción.
  - Fallback: parseo línea por línea del texto crudo acumulado.
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


def es_pagina_contenido(texto: str) -> bool:
    """
    Devuelve True si la página tiene contenido de OC (renglones/encabezado).
    Filtra páginas de firma (Hoja 3/3), carátula GEDO y hoja adicional.
    """
    if not texto:
        return False
    # Excluir hoja adicional de firmas GEDO
    if "Hoja Adicional de Firmas" in texto:
        return False
    # Excluir páginas que solo tienen firma/sello (Hoja 3/3 en este PDF)
    if re.search(r'JURIDISDICCION RESPONSABLE', texto, re.IGNORECASE):
        return False
    # Debe tener al menos referencia a OC o solicitud
    return bool(re.search(r'(?:Nro\.?\s*(?:de\s*)?O\.?C\.|SOLICITUD\s+Nro)', texto, re.IGNORECASE))


def es_fila_renglones(fila: list, header_norm: list) -> bool:
    """
    Verifica que la fila sea de datos de renglones y no una fila espuria
    (encabezado SAF repetido, subtotales, afectación presupuestaria).
    """
    if not fila or not any(fila):
        return False
    primera_celda = str(fila[0] or "").strip()
    # Fila de encabezado SAF (texto largo sin números de renglón)
    if "Saf:" in primera_celda or "Jurisdicción" in primera_celda:
        return False
    # Fila de subtotal/total
    if re.search(r'SubTotal|TOTAL AFECTADO|Remito:|Facturación:', primera_celda, re.IGNORECASE):
        return False
    # Fila de afectación presupuestaria
    if re.search(r'C\.\s*INSTITUCIONAL|PRG\s+\d+', primera_celda):
        return False
    # Debe tener al menos un número de renglón en la primera celda
    idx_reng = next((i for i, h in enumerate(header_norm) if "renglon" in h or "reng" in h), None)
    if idx_reng is not None and idx_reng < len(fila):
        col_reng = str(fila[idx_reng] or "").strip()
        return bool(re.search(r'\d', col_reng))
    return False


def limpiar_descripcion_celda(texto: str) -> str:
    """
    Limpia el texto de descripción de una celda:
    - Elimina nombres de empresa/marca que pdfplumber pega al inicio
      (palabras en mayúsculas sin dígitos que aparecen solas en la primera línea
       y no forman parte de la descripción del producto).
    - Toma la primera línea significativa como descripción principal.
    """
    if not texto:
        return ""
    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
    if not lineas:
        return ""

    # Si la primera línea parece un nombre de empresa/marca (solo mayúsculas,
    # sin números, sin paréntesis, corta), saltarla
    primera = lineas[0]
    es_marca = (
        primera.isupper() and
        len(primera.split()) <= 3 and
        not re.search(r'\d', primera) and
        not re.search(r'(?:ALCOHOL|IODO|DETER|VASEL|ORTO|RITU|AMOX|AMPIC|CEFTR)', primera, re.IGNORECASE)
    )
    desc_lineas = lineas[1:] if es_marca else lineas
    if not desc_lineas:
        return primera  # fallback: devolver la primera aunque sea marca

    # Tomar la primera línea significativa (la descripción principal)
    return desc_lineas[0].strip()


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

    # ── Extraer texto de TODAS las páginas de contenido ──────────────────────
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            textos_por_pagina = []
            for pagina in pdf.pages:
                t = pagina.extract_text(x_tolerance=3, y_tolerance=3) or ""
                if es_pagina_contenido(t):
                    textos_por_pagina.append(t)
            # Texto acumulado de todas las páginas (para búsqueda de descripciones)
            texto = "\n".join(textos_por_pagina)
            resultado["texto_raw"] = texto
    except Exception as e:
        resultado["errores"].append(f"Error al abrir PDF: {e}")
        return resultado

    if not texto:
        resultado["errores"].append("No se pudo extraer texto del PDF")
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

    # Estrategia A: tablas SIGAF con columnas fusionadas.
    # Los renglones pueden distribuirse en varias páginas. La página 1 tiene
    # el header (Renglón/Código/Cantidad/etc.); las páginas siguientes pueden
    # NO tener header — la tabla continúa directamente con datos.
    # Solución: guardar los índices de columna del primer header encontrado
    # y reutilizarlos en las páginas siguientes.

    # Índices de columna globales (se fijan en la primera tabla con header válido)
    _i_reng = _i_cod = _i_cant = _i_uunit = _i_total = None
    _header_fijado = False

    def _idx_col(header_norm, keywords):
        for kw in keywords:
            for i, h in enumerate(header_norm):
                if kw in h:
                    return i
        return None

    def _celda(fila, idx):
        if idx is None or idx >= len(fila):
            return ""
        v = fila[idx]
        return str(v).strip() if v else ""

    def _procesar_fila_fusionada(col_reng, col_cod, col_cant, col_uunt, col_tot):
        """Extrae renglones de una fila fusionada SIGAF y los agrega a la lista."""
        rengs = [r.strip() for r in col_reng.split("\n")
                 if r.strip() and re.match(r'^\d+$', r.strip())]
        cods  = [c.strip() for c in col_cod.split("\n")
                 if c.strip() and re.match(r'^[A-Z0-9]{4,}$', c.strip())]
        cants = [c.strip() for c in col_cant.split("\n") if c.strip()]
        uunts = [c.strip() for c in col_uunt.split("\n") if c.strip()]
        tots  = [c.strip() for c in col_tot.split("\n")  if c.strip()]

        if not rengs:
            return

        n = len(rengs)
        def pad(lst): return (lst + [""] * n)[:n]
        cods, cants, uunts, tots = pad(cods), pad(cants), pad(uunts), pad(tots)

        for reng_num, cod, cant, uunt, tot in zip(rengs, cods, cants, uunts, tots):
            desc = extraer_descripcion_desde_texto(texto, reng_num)
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

    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                t_pag = pagina.extract_text(x_tolerance=3, y_tolerance=3) or ""
                if not es_pagina_contenido(t_pag):
                    continue

                for tabla in (pagina.extract_tables() or []):
                    if not tabla or len(tabla) < 1:
                        continue

                    header_raw  = [str(c).lower().strip() if c else "" for c in tabla[0]]
                    header_norm = [
                        h.replace("ó","o").replace("é","e").replace("ú","u")
                         .replace("í","i").replace("á","a") for h in header_raw
                    ]

                    tiene_renglon  = any(kw in h for h in header_norm for kw in ["renglon","reng"])
                    tiene_cantidad = any("cantidad" in h for h in header_norm)

                    if tiene_renglon and tiene_cantidad:
                        # ── Tabla con header completo (página 1 o primera con renglones)
                        _i_reng  = _idx_col(header_norm, ["renglon", "reng"])
                        _i_cod   = _idx_col(header_norm, ["codigo", "cod"])
                        _i_cant  = _idx_col(header_norm, ["cantidad"])
                        _i_uunit = _idx_col(header_norm, ["unitario", "unit", "precio"])
                        _i_total = None
                        for i, h in enumerate(header_norm):
                            if "total" in h and i != _i_uunit:
                                _i_total = i
                                break
                        if _i_total is None:
                            _i_total = _idx_col(header_norm, ["importe total", "total"])
                        _header_fijado = True

                        for fila in tabla[1:]:
                            if not es_fila_renglones(fila, header_norm):
                                continue
                            _procesar_fila_fusionada(
                                _celda(fila, _i_reng), _celda(fila, _i_cod),
                                _celda(fila, _i_cant), _celda(fila, _i_uunit),
                                _celda(fila, _i_total)
                            )

                    elif _header_fijado:
                        # ── Página de continuación: no tiene header de renglones.
                        # La primera fila puede ser el encabezado SAF u otra cosa;
                        # buscamos directamente filas que tengan números de renglón
                        # en la columna _i_reng (ya fijada).
                        for fila in tabla:
                            if not fila or not any(fila):
                                continue
                            col_reng_raw = _celda(fila, _i_reng)
                            # Verificar que la celda de renglón tiene al menos un número
                            rengs_test = [r.strip() for r in col_reng_raw.split("\n")
                                          if r.strip() and re.match(r'^\d+$', r.strip())]
                            if not rengs_test:
                                continue
                            # Descartar filas de subtotal/afectación
                            primera = str(fila[0] or "").strip()
                            if re.search(r'SubTotal|TOTAL AFECTADO|Remito|Facturación|PRG\s+\d+|C\.\s*INSTITUCIONAL|Saf:', primera, re.IGNORECASE):
                                continue
                            _procesar_fila_fusionada(
                                col_reng_raw,
                                _celda(fila, _i_cod),
                                _celda(fila, _i_cant),
                                _celda(fila, _i_uunit),
                                _celda(fila, _i_total)
                            )

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
