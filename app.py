import streamlit as st
import pandas as pd
import io
import re
import folium
from folium.plugins import AntPath
from streamlit_folium import st_folium # Import the missing st_folium function

st.set_page_config(layout="wide", page_title="Análisis de Datos Geoespaciales")

st.title("🕵️‍♂️ Analisis-de-CDR-s-automatizado
Análisis de inteligencia artificial que permite limpiar, organizar y analizar las bases de datos en archivos consolidados! convierte celdas hexadecimales y excluye números plataformas, también representa gráficamente en el mapa, genera la sinopsis del comportamiento, zona de mayor confort, residencia o trabajo habitual")

# Funciones de limpieza de datos (replicadas de la Celda 1)
@st.cache_data # Cachear esta función para evitar recalcular fechas
def limpiar_fechas(s):
    s = s.astype(str).str.strip()
    mask_14 = s.str.match(r'^\d{14}$')
    fechas_14 = pd.to_datetime(s[mask_14], format='%Y%m%d%H%M%S', errors='coerce')
    fechas_otros = pd.to_datetime(s[~mask_14], format='mixed', dayfirst=True, errors='coerce')
    return pd.concat([fechas_14, fechas_otros])

@st.cache_data # Cachear esta función para evitar recalcular celdas
def normalizar_celda(val):
    val = str(val).strip().upper()
    if val.endswith('.0'):
        val = val[:-2]
    if re.match(r'^[0-9A-F]{4,6}-[0-9A-F]{2,4}$', val):
        try:
            parts = val.split('-')
            dec_parts = [str(int(p, 16)) for p in parts]
            return '-'.join(dec_parts)
        except:
            return val
    return val

@st.cache_data(show_spinner=False) # Cachear el proceso de carga y consolidación de datos
def load_and_process_files(uploaded_files_data):
    df_list = []
    for file_name, file_content in uploaded_files_data:
        try:
            if file_name.lower().endswith('.csv'):
                try:
                    df = pd.read_csv(io.BytesIO(file_content), sep=';', encoding='utf-8-sig', low_memory=False)
                    if len(df.columns) <= 1:
                        df = pd.read_csv(io.BytesIO(file_content), sep=',', encoding='utf-8-sig', low_memory=False)
                except:
                    df = pd.read_csv(io.BytesIO(file_content), sep=None, engine='python')
            else:
                df = pd.read_excel(io.BytesIO(file_content))

            df.columns = df.columns.str.strip().str.lower()

            col_tel = next((c for c in ['numero', 'numero_origen', 'numero_que_marca', 'numero_que_navega', 'originador'] if c in df.columns), None)
            col_fec = next((c for c in ['fecha_hora_inicio', 'fecha_trafico', 'fecha_hora_inicio_llamada', 'fecha_hora_inicio_sesion', 'fecha_hora'] if c in df.columns), None)
            col_lat = next((c for c in ['latitud', 'latitud_n'] if c in df.columns), None)
            col_lon = next((c for c in ['longitud', 'longitud_w'] if c in df.columns), None)
            col_cel = next((c for c in ['celda_decimal', 'cell_id_voz', 'celda', 'celda_inicio_llamada', 'bts_id', 'celda_hex'] if c in df.columns), None)

            df_temp = pd.DataFrame()

            if col_tel:
                df_temp['numero_limpio'] = df[col_tel].astype(str).str.replace(r'\D', '', regex=True).str[-10:]
            if col_fec:
                df_temp['fecha_limpia'] = limpiar_fechas(df[col_fec])
            if col_lat and col_lon:
                df_temp['latitud'] = pd.to_numeric(df[col_lat].astype(str).str.replace(',', '.'), errors='coerce')
                df_temp['longitud'] = pd.to_numeric(df[col_lon].astype(str).str.replace(',', '.'), errors='coerce')
            if col_cel:
                df_temp['celda'] = df[col_cel].apply(normalizar_celda)

            df_temp['registro_original'] = df.to_dict('records')
            df_list.append(df_temp)
            st.success(f"✅ Procesado: {file_name}")

        except Exception as e:
            st.error(f"❌ Error procesando {file_name}: {e}")

    if df_list:
        df_master = pd.concat(df_list, ignore_index=True)
        st.success(f"✅ ¡Archivos consolidados! Celdas Hexadecimales convertidas y números a 10 dígitos. Total registros: {len(df_master)}")
        return df_master
    else:
        st.warning("No se encontraron datos después de procesar los archivos.")
        return pd.DataFrame()


# --- Carga de Archivos (adaptado para Streamlit) ---
st.header("📂 Carga tus Archivos de Datos")
uploaded_files = st.file_uploader(
    "Sube tus archivos CSV o Excel (Soporta múltiples operadores al mismo tiempo):",
    type=['csv', 'xlsx'],
    accept_multiple_files=True
)

df_master = None

if uploaded_files:
    # Convertir uploaded_files a un formato cacheable
    uploaded_files_data = [(f.name, f.getvalue()) for f in uploaded_files]
    with st.spinner('Procesando archivos...'):
        df_master = load_and_process_files(uploaded_files_data)

# --- Búsqueda Geográfica Individual (adaptado para Streamlit) ---
st.header("📍 Búsqueda Geográfica Individual")

if df_master is not None and not df_master.empty:
    try:
        fecha_min = df_master['fecha_limpia'].min().date() if pd.notna(df_master['fecha_limpia'].min()) else None
        fecha_max = df_master['fecha_limpia'].max().date() if pd.notna(df_master['fecha_limpia'].max()) else None
    except:
        fecha_min = fecha_max = None

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        numero_buscado = st.text_input('Número:', placeholder='Ej: 3157658841')
    with col2:
        dropdown_top = st.selectbox('Mostrar:', ['Todos', 'Top 10', 'Top 5', 'Top 3', 'Top 1'])

    st.subheader("Rango de Fechas y Horas")
    col_dp_inicio, col_tm_inicio, col_dp_fin, col_tm_fin = st.columns(4)
    with col_dp_inicio:
        dp_inicio = st.date_input('Fecha Ini:', value=fecha_min if fecha_min else pd.Timestamp.now().date())
    with col_tm_inicio:
        tm_inicio = st.text_input('Hora Ini:', value='00:00')
    with col_dp_fin:
        dp_fin = st.date_input('Fecha Fin:', value=fecha_max if fecha_max else pd.Timestamp.now().date())
    with col_tm_fin:
        tm_fin = st.text_input('Hora Fin:', value='23:59')

    if st.button('Trazar Mapa 🗺️', type="primary"):
        if len(numero_buscado) != 10:
            st.warning("⚠️ Ingresa un número de exactamente 10 dígitos.")
        else:
            df_usuario = df_master[(df_master['numero_limpio'] == numero_buscado) & (df_master['latitud'].notna())].copy()

            try:
                dt_ini = pd.to_datetime(f"{dp_inicio} {tm_inicio}")
                dt_fin = pd.to_datetime(f"{dp_fin} {tm_fin}")
                df_usuario = df_usuario[(df_usuario['fecha_limpia'] >= dt_ini) & (df_usuario['fecha_limpia'] <= dt_fin)]
            except Exception as e:
                st.error(f"Error en el rango de fechas/horas: {e}")
                st.stop()

            if df_usuario.empty:
                st.info(f"❌ No hay historial de coordenadas para {numero_buscado} en este rango de tiempo.")
            else:
                df_usuario = df_usuario.sort_values(by='fecha_limpia')

                df_agrupado = df_usuario.groupby(['latitud', 'longitud']).agg(
                    visitas=('numero_limpio', 'count'),
                    primera_visita=('fecha_limpia', 'min'),
                    ultima_visita=('fecha_limpia', 'max'),
                    celdas=('celda', lambda x: ', '.join(x.dropna().unique().astype(str)))
                ).reset_index().sort_values(by='visitas', ascending=False).reset_index(drop=True)

                # --- SINOPSIS ---
                tot_conexiones = len(df_usuario)
                if not df_agrupado.empty:
                    top_lugar = df_agrupado.iloc[0]['celdas']
                    top_visitas = df_agrupado.iloc[0]['visitas']

                    st.markdown(f"""
                    <div style="background-color:#e8f4f8; padding:15px; border-left: 5px solid #5bc0de; border-radius:5px; margin-bottom: 15px;">
                        <h4 style="margin-top:0; color:#31708f;">💡 Sinopsis del Comportamiento (Patrón de Vida)</h4>
                        <p>El objetivo <b>{numero_buscado}</b> registra una actividad total de <b>{tot_conexiones}</b> conexiones geo-posicionadas dentro del periodo consultado. Su zona de mayor confort, residencia o trabajo habitual corresponde a la cobertura de la celda <b>{top_lugar}</b>, lugar donde acumuló el mayor número de impactos (<b>{top_visitas}</b> visitas registradas).</p>
                    </div>
                    """, unsafe_allow_html=True)

                # --- MAPA ---
                n = len(df_agrupado) if dropdown_top == 'Todos' else int(dropdown_top.replace('Top ', ''))
                df_mostrar = df_agrupado.head(n)

                if not df_mostrar.empty:
                    mapa = folium.Map(location=[df_mostrar.iloc[0]['latitud'], df_mostrar.iloc[0]['longitud']], zoom_start=14)

                    for i, row in df_mostrar.iterrows():
                        coord = [row['latitud'], row['longitud']]
                        ranking = i + 1
                        color = 'red' if ranking == 1 else 'orange' if ranking <=3 else 'blue'
                        popup_html = f"""<div style="min-width: 200px;"><b>Rank:</b> #{ranking}<br><b>Visitas:</b> {row['visitas']}<br><b>Celda:</b> {row['celdas']}<br><b>Inicio:</b> {row['primera_visita']}<br><b>Fin:</b> {row['ultima_visita']}</div>"""
                        folium.CircleMarker(location=coord, radius=min(20, 8 + row['visitas']), color=color, fill=True).add_to(mapa)
                        folium.Marker(location=coord, popup=folium.Popup(popup_html, max_width=300), tooltip=f"Rank #{ranking}", icon=folium.Icon(color=color)).add_to(mapa)

                    if dropdown_top == 'Todos' and len(df_usuario) > 1:
                        AntPath(locations=df_usuario[['latitud', 'longitud']].values.tolist(), delay=1000, color='purple', weight=4).add_to(mapa)

                    st.subheader("Mapa de Ubicaciones")
                    st_folium(mapa, width=1000, height=600)
                else:
                    st.info("No hay puntos para mostrar en el mapa con los filtros seleccionados.")

else:
    st.info("Por favor, sube archivos de datos para comenzar el análisis.")

# Para ejecutar esta aplicación, guarda el archivo como `streamlit_app.py` y luego en tu terminal ejecuta `streamlit run streamlit_app.py`
