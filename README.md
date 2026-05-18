# Tablero de Control — Farmacia 💊

App Streamlit para análisis de solicitudes de medicamentos 2026.

## Cómo usar

### Opción 1 — Streamlit Cloud (recomendado, sin instalación)

1. Creá una cuenta gratis en [github.com](https://github.com) si no tenés
2. Creá un repositorio nuevo (público o privado)
3. Subí los 3 archivos: `app.py`, `requirements.txt`, `README.md`
4. Andá a [share.streamlit.io](https://share.streamlit.io) → "New app"
5. Seleccioná tu repo y `app.py` como archivo principal
6. ¡Listo! Te da un link para compartir con todo el equipo

### Opción 2 — Google Colab

```python
# Ejecutar en Colab:
!pip install streamlit pyngrok
!wget -q https://raw.githubusercontent.com/TU_USUARIO/TU_REPO/main/app.py
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

### Opción 3 — Local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Funcionalidades

| Pestaña | Contenido |
|---|---|
| 📊 Resumen General | KPIs globales, monto por rubro, desiertos, ejecución |
| 🔢 ABC por Monto | Curva de Pareto, clasificación A/B/C, top 30 |
| 🔄 ABC Cantidad + XYZ | Rotación, variabilidad de demanda, matriz ABC-XYZ |
| 🏭 Proveedores | Concentración, índice HHI, riesgo por proveedor |
| 📋 Datos | Explorador filtrable + descarga Excel |

## Archivo necesario

Subir **SOLICITUDES_2026.xlsx** con las hojas de cada rubro.
