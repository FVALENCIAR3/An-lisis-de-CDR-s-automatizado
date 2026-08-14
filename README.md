# Análisis de CDRs Automatizado

Aplicación Streamlit convertida desde el notebook original.

## Funciones conservadas

- Carga y consolidación de CSV, XLS, XLSX y XLSB.
- Limpieza y normalización de fechas.
- Normalización de celdas.
- Búsqueda geográfica individual.
- Análisis de frecuencia por celda.
- Análisis de co-ubicación en ventanas de 15 minutos.
- Lectura de hechos y documentos TXT, PDF y DOCX.
- Extracción de números de 10 dígitos.
- Cruce documental con la base CDR.
- Búsqueda de posibles victimarios por co-desplazamiento.
- Perfilamiento de rutinas por franjas horarias.
- Mapas Folium.
- Exportación XLSX, PNG, HTML y ZIP.
- Exportación KML para Google Earth.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## GitHub / Streamlit Community Cloud

Sube `app.py` y `requirements.txt` al repositorio y configura el archivo principal como:

`app.py`
