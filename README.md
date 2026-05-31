# 💊 Tablero de Control — Farmacia HIGA Alende

Sistema integral de análisis y gestión de solicitudes de medicamentos del Hospital Interzonal General de Agudos (HIGA) Dr. Oscar Alende — Mar del Plata.

---

## 🗂️ Estructura del proyecto

```
TABLERO-FARMACIA/
├── app.py                    # App Streamlit — análisis interactivo (2070 líneas)
├── generar_tablero.py        # Script de generación del tablero estático HTML
├── parsear_oc.py             # Parser de órdenes de compra
├── cargar_oc.py              # Cargador de órdenes de compra
├── index.html                # Tablero estático generado (salida de generar_tablero.py)
├── tablero_template.html     # Plantilla base usada por generar_tablero.py
├── SOLICITUDES_2026.xlsx     # Datos de solicitudes anuales (⚠️ ver sección de privacidad)
├── requirements.txt          # Dependencias Python
├── .github/workflows/        # GitHub Actions (CI/CD)
└── .devcontainer/            # Entorno de desarrollo en contenedor
```

---

## 🚀 Cómo usar

### Opción 1 — GitHub Pages (tablero estático, sin servidor)

El `index.html` es el tablero **ya generado** y se puede servir directamente con GitHub Pages:

1. En el repo → **Settings → Pages**
2. Source: `Deploy from a branch` → `main` → `/ (root)`
3. GitHub publica automáticamente en `https://matiquiroga87.github.io/TABLERO-FARMACIA/`

> El tablero estático no requiere Python ni servidor. Es el modo de consulta recomendado para el equipo clínico.

---

### Opción 2 — Streamlit Cloud (app interactiva, recomendado para análisis)

1. Crear cuenta gratuita en [github.com](https://github.com) si no tenés
2. Ir a [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Seleccionar el repositorio y `app.py` como archivo principal
4. Subir `SOLICITUDES_2026.xlsx` cuando la app lo solicite
5. La app genera un link compartible con todo el equipo

---

### Opción 3 — Local

```bash
# Clonar el repositorio
git clone https://github.com/matiquiroga87/TABLERO-FARMACIA.git
cd TABLERO-FARMACIA

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar app interactiva
streamlit run app.py

# O regenerar el tablero estático
python generar_tablero.py
```

---

### Opción 4 — Google Colab

```python
# Ejecutar en Colab
!pip install streamlit pyngrok openpyxl pandas plotly
!wget -q https://raw.githubusercontent.com/matiquiroga87/TABLERO-FARMACIA/main/app.py

from pyngrok import ngrok
import subprocess, threading

def run():
    subprocess.run(["streamlit", "run", "app.py",
                    "--server.port", "8501",
                    "--server.headless", "true"])

t = threading.Thread(target=run)
t.start()

import time; time.sleep(4)
public_url = ngrok.connect(8501)
print(f"\n🚀 Tablero online en: {public_url}")
```

---

## 📊 Funcionalidades

| Pestaña | Contenido |
|---|---|
| 📊 **Resumen General** | KPIs globales, monto por rubro, ítems desiertos, % de ejecución presupuestaria |
| 🔢 **ABC por Monto** | Curva de Pareto, clasificación A/B/C, top 30 ítems por gasto |
| 🔄 **ABC Cantidad + XYZ** | Rotación de stock, variabilidad de demanda, matriz cruzada ABC-XYZ |
| 🏭 **Proveedores** | Concentración de compras, índice HHI, análisis de riesgo por proveedor |
| 📋 **Datos** | Explorador filtrable con descarga a Excel |

---

## 🔄 Flujo de actualización del tablero

```
SOLICITUDES_2026.xlsx  ──►  generar_tablero.py  ──►  index.html
        (datos)                  (script)              (tablero publicado)
```

Cuando el Excel se actualiza:

```bash
python generar_tablero.py   # regenera index.html desde tablero_template.html
git add index.html
git commit -m "Actualizar tablero - [fecha]"
git push
```

GitHub Pages publica automáticamente el nuevo `index.html`.

---

## 📁 Archivo de datos

El tablero requiere **`SOLICITUDES_2026.xlsx`** con una hoja por cada rubro (Medicamentos, Descartables, Reactivos, etc.).

> ⚠️ **Privacidad:** Este archivo contiene datos operativos del hospital. Se recomienda moverlo a un repositorio privado separado. Ver sección [Mover datos a repo privado](#-mover-solicitudes_2026xlsx-a-un-repositorio-privado).

---

## 🔒 Mover `SOLICITUDES_2026.xlsx` a un repositorio privado

### Por qué hacerlo

El Excel contiene información operativa institucional (montos, proveedores, ítems licitados). Mantenerlo en un repo público lo expone a cualquier persona. La separación también permite actualizar los datos sin tocar el código, y viceversa.

### Paso a paso

**1. Crear el repo privado de datos**

```
Nombre sugerido: FARMACIA-DATOS (privado)
Contenido inicial: SOLICITUDES_2026.xlsx
```

**2. Eliminar el Excel del repo público**

```bash
# Remover del tracking y del historial
git rm --cached SOLICITUDES_2026.xlsx
echo "SOLICITUDES_2026.xlsx" >> .gitignore
git commit -m "Mover datos a repo privado"
git push
```

> Si ya fue commiteado, usá `git filter-repo` o BFG Repo Cleaner para limpiarlo del historial.

**3. Ajustar `generar_tablero.py` y `app.py`**

Cambiá la ruta del archivo de:
```python
ARCHIVO_DATOS = "SOLICITUDES_2026.xlsx"
```
a una de estas estrategias:

**a) Variable de entorno** (recomendado para Streamlit Cloud):
```python
import os
ARCHIVO_DATOS = os.environ.get("RUTA_SOLICITUDES", "SOLICITUDES_2026.xlsx")
```
Configurar en Streamlit Cloud → **App settings → Secrets**:
```
RUTA_SOLICITUDES = "/ruta/al/archivo.xlsx"
```

**b) Carga manual desde la UI** (ya soportado en `app.py`):
```python
archivo = st.file_uploader("Subir SOLICITUDES_2026.xlsx", type=["xlsx"])
if archivo:
    df = pd.read_excel(archivo, sheet_name=None)
```

**c) GitHub Actions con acceso al repo privado** (para regeneración automática):
```yaml
# En .github/workflows/actualizar_tablero.yml
- uses: actions/checkout@v4
  with:
    repository: matiquiroga87/FARMACIA-DATOS
    token: ${{ secrets.PAT_DATOS }}
    path: datos
```
Requiere crear un **Personal Access Token** con permisos de lectura al repo privado y guardarlo como secret `PAT_DATOS`.

### ¿Qué hay que reconfigurar?

| Componente | Cambio necesario |
|---|---|
| `generar_tablero.py` | Actualizar ruta o leer desde env var |
| `app.py` | Actualizar ruta o habilitar `file_uploader` |
| GitHub Actions workflow | Agregar checkout del repo privado con PAT |
| Streamlit Cloud | Agregar secret con la ruta o usar uploader |
| `.gitignore` | Agregar `SOLICITUDES_2026.xlsx` |

---

## 🛠️ Requisitos

```
Python >= 3.9
streamlit
pandas
openpyxl
plotly
numpy
jinja2         # usado por generar_tablero.py con tablero_template.html
```

Ver `requirements.txt` para versiones exactas.

---

## 📄 Licencia

Proyecto interno — HIGA Dr. Oscar Alende, Mar del Plata.
Para uso institucional del servicio de Farmacia.
