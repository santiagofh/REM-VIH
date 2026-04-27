# REM VIH 2025

Dashboard en Streamlit para consultar datos REM VIH 2025 de forma tabular.

## Contenido

- Explorador por serie REM y sección.
- Filtros por servicio de salud, comuna, establecimiento y prestación.
- Descarga de resultados filtrados en Excel y CSV.
- Descarga de secciones completas y archivos fuente consolidados.

## Fuente de datos

La aplicación usa los archivos consolidados ubicados en `2025/salida`:

- `A05_2025.xlsx`
- `P11_2025.xlsx`

## Ejecución local

Instala dependencias:

```powershell
pip install -r requirements.txt
```

Inicia la aplicación:

```powershell
streamlit run streamlit_dashboard.py
```

## Estructura principal

- `streamlit_dashboard.py`: entrada principal de la app.
- `dashboard_vih_pages.py`: páginas, filtros y exportaciones.
- `.streamlit/config.toml`: configuración de Streamlit.
- `assets/`: elementos visuales del dashboard.
