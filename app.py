import streamlit as st
import pandas as pd
import io
import re
import folium
from folium.plugins import AntPath
from streamlit_folium import st_folium
import matplotlib.pyplot as plt
import seaborn as sns
import itertools
from collections import Counter
import zipfile
import PyPDF2
import docx
import simplekml
import datetime
import numpy as np

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA (ESTILO GAULA)
# ==========================================
st.set_page_config(page_title="Análisis de CDR-s automatizado", page_icon="🛡️", layout="wide")

st.markdown("""
    <style>
    .titulo-gaula { font-size: 2.2rem; font-weight: bold; color: #003366; margin-bottom: 0px; line-height: 1.2;}
    .desc-gaula { font-size: 1.05rem; color: #333333; text-align: justify; margin-bottom: 20px;}
    .caja-info { background-color: #f4f6f9; border-left: 5px solid #003366; padding: 15px; border-radius: 5px; margin-bottom: 15px;}
    .caja-alerta { background-color: #fcf8e3; border-left: 5px solid #d9534f; padding: 15px; border-radius: 5px; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

col_img, col_txt = st.columns([1, 8])
with col_img:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Escudo_de_la_Polic%C3%ADa_Nacional_de_Colombia.svg/200px-Escudo_de_la_Polic%C3%ADa_Nacional_de_Colombia.svg.png", width=90)
with col_txt:
    st.markdown('<div class="titulo-gaula">Análisis de CDR-s automatizado</div>', unsafe_allow_html=True)
    st.markdown('<div class="desc-gaula">Análisis de inteligencia artificial que permite limpiar, organizar y analizar las bases de datos en archivos consolidados! Convierte celdas hexadecimales y excluye números de plataformas, representa gráficamente en el mapa, genera la sinopsis del comportamiento, zona de mayor permanencia, residencia o trabajo habitual del objetivo.</div>', unsafe_allow_html=True)

st.divider()

if 'df_master' not in st.session_state: st.session_state.df_master = None
for key in ['run_t1', 'run_t2', 'run_t3', 'run_t4', 'run_t5']:
    if key not in st.session_state: st.session_state[key] = False

def activar_t1(): st.session_state.run_t1 = True
def activar_t2(): st.session_state.run_t2 = True
def activar_t3(): st.session_state.run_t3 = True
def activar_t4(): st.session_state.run_t4 = True
def activar_t5(): st.session_state.run_t5 = True

# ==========================================
# FUNCIONES DE LECTURA (TÉCNICA AVANZADA)
# ==========================================
@st.cache_data(show_spinner=False)
def procesar_archivos(uploaded_files):
    df_list = []
    
    def limpiar_fechas(s):
        if pd.api.types.is_datetime64_any_dtype(s): return s
        s_str = s.astype(str).str.strip()
        fechas = pd.to_datetime(s_str, format='mixed', dayfirst=False, errors='coerce')
        mask_nat = fechas.isna()
        if mask_nat.any():
            mask_14 = s_str.str.match(r'^\d{14}$')
            fechas.loc[mask_nat & mask_14] = pd.to_datetime(s_str[mask_nat & mask_14], format='%Y%m%d%H%M%S', errors='coerce')
        return fechas

    def limpiar_coord(c):
        return pd.to_numeric(c.astype(str).str.replace(',', '.').str.replace(r'[^\d.-]', '', regex=True), errors='coerce')

    def normalizar_celda(val):
        val = str(val).strip().upper()
        if val.endswith('.0'): val = val[:-2]
        # Convierte formatos hexadecimales cruzados (Movistar, etc)
        if re.match(r'^[0-9A-F]+-[0-9A-F]+$', val):
            try: return '-'.join([str(int(p, 16)) for p in val.split('-')])
            except: return val
        # Si es un hexadecimal puro de gran longitud, convertirlo a decimal
        if re.match(r'^[0-9A-F]{5,}$', val) and not val.isdigit():
            try: return str(int(val, 16))
            except: return val
        return val

    def leer_texto_plano(file):
        for enc in ['utf-8-sig', 'latin1', 'cp1252']:
            for sep in [';', ',', '\t', '|']:
                file.seek(0)
                try:
                    df_p = pd.read_csv(file, sep=sep, encoding=enc, low_memory=False, on_bad_lines='skip')
                    if len(df_p.columns) > 1: return df_p
                except: continue
        file.seek(0)
        try: return pd.read_csv(file, sep=None, engine='python', encoding='latin1', on_bad_lines='skip')
        except: return pd.DataFrame()

    for file in uploaded_files:
        try:
            file_ext = file.name.lower()
            df = pd.DataFrame()
            if file_ext.endswith(('.csv', '.txt')): df = leer_texto_plano(file)
            elif file_ext.endswith('.xlsb'): file.seek(0); df = pd.read_excel(file, engine='pyxlsb')
            elif file_ext.endswith(('.xls', '.xlsx')): file.seek(0); df = pd.read_excel(file)
            else: continue
            if df.empty: continue
            
            df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r'\W+', '_', regex=True)
            
            alias_tel = ['numero_origen', 'numero_a', 'numero_que_marca', 'numero_que_navega', 'originador', 'numero', 'msisdn', 'calling_number', 'min_origen', 'party_a']
            alias_fec = ['fecha_hora_inicio', 'fecha_trafico', 'fecha_hora_inicio_llamada', 'fecha_y_hora_origen', 'fecha_hora', 'fecha', 'start_time', 'timestamp', 'fechahora']
            alias_lat = ['latitud', 'latitud_n', 'latitude', 'lat', 'latitud_origen', 'lat_origen']
            alias_lon = ['longitud', 'longitud_w', 'longitude', 'lon', 'longitud_origen', 'long']
            alias_cel = ['celda_decimal', 'cell_id_voz', 'celda_inicio_llamada', 'celda_origen_truncada', 'celda_origen', 'celda_hex', 'cellid_nval', 'bts_id', 'site_id', 'cell_id', 'celda', 'cgi', 'codigo_celda']
            
            alias_lac = ['lac_decimal', 'lac', 'location_area_code', 'tac']
            alias_haz = ['sector', 'haz', 'beam', 'sector_id']
            alias_azi = ['azimut', 'azimuth', 'dir', 'direccion_antena']
            alias_nom = ['nombre_antena', 'descripcion', 'direccion', 'site_name', 'nombre_sitio', 'address', 'ubicacion', 'city_ds']
            
            col_tel = next((c for c in alias_tel if c in df.columns), None)
            col_fec = next((c for c in alias_fec if c in df.columns), None)
            col_lat = next((c for c in alias_lat if c in df.columns), None)
            col_lon = next((c for c in alias_lon if c in df.columns), None)
            col_cel = next((c for c in alias_cel if c in df.columns), None)
            
            col_lac = next((c for c in alias_lac if c in df.columns), None)
            col_haz = next((c for c in alias_haz if c in df.columns), None)
            col_azi = next((c for c in alias_azi if c in df.columns), None)
            col_nom = next((c for c in alias_nom if c in df.columns), None)
            
            df_temp = pd.DataFrame()
            if col_tel: df_temp['numero_limpio'] = df[col_tel].astype(str).str.replace(r'\D', '', regex=True).str[-10:]
            if col_fec: 
                df_temp['fecha_limpia'] = limpiar_fechas(df[col_fec])
                df_temp['fecha_limpia'] = df_temp['fecha_limpia'].dt.tz_localize(None)
            if col_lat and col_lon:
                df_temp['latitud'] = limpiar_coord(df[col_lat])
                df_temp['longitud'] = limpiar_coord(df[col_lon])
            
            if col_cel: 
                df_temp['celda'] = df[col_cel].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_temp['celda'] = df_temp['celda'].apply(normalizar_celda)
            else:
                df_temp['celda'] = 'Desconocida'
                
            df_temp['lac'] = df[col_lac].astype(str).str.replace(r'\.0$', '', regex=True) if col_lac else 'N/A'
            df_temp['haz'] = df[col_haz].astype(str).str.replace(r'\.0$', '', regex=True) if col_haz else 'N/A'
            df_temp['azimut'] = df[col_azi].astype(str).str.replace(r'\.0$', '', regex=True) if col_azi else 'N/A'
            df_temp['nombre_antena'] = df[col_nom].astype(str) if col_nom else 'N/A'
                
            if not df_temp.empty:
                df_temp['registro_original'] = df.to_dict('records')
                df_list.append(df_temp)
            
        except Exception as e:
            st.sidebar.error(f"Error en {file.name}: {e}")

    if df_list: return pd.concat(df_list, ignore_index=True)
    return None

# ==========================================
# BARRA LATERAL (CARGA)
# ==========================================
with st.sidebar.form("form_carga"):
    st.markdown("**📂 1. Cargar Archivos Consolidables**")
    uploaded_files = st.file_uploader("Soporta CDRs: CSV, TXT, XLS, XLSX, XLSB", accept_multiple_files=True, type=['csv', 'txt', 'xls', 'xlsx', 'xlsb'])
    submit_button = st.form_submit_button("⚙️ Procesar Archivos (I.A)")

if submit_button:
    if uploaded_files:
        with st.spinner("Limpiando, estandarizando y consolidando data (Hex, Fechas, Coordenadas)..."):
            st.session_state.df_master = procesar_archivos(uploaded_files)
            for key in ['run_t1', 'run_t2', 'run_t3', 'run_t4', 'run_t5']: st.session_state[key] = False
        if st.session_state.df_master is not None and not st.session_state.df_master.empty:
            st.sidebar.success(f"✅ {len(st.session_state.df_master)} registros consolidados exitosamente.")
        else:
            st.sidebar.error("❌ Archivos no contienen columnas telefónicas compatibles.")

# ==========================================
# CUERPO PRINCIPAL (PESTAÑAS)
# ==========================================
if st.session_state.df_master is not None:
    df_master = st.session_state.df_master
    try: f_min, f_max = df_master['fecha_limpia'].min().date(), df_master['fecha_limpia'].max().date()
    except: f_min = f_max = datetime.date.today()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📍 Búsqueda Geográfica", "⚙️ Inteligencia de Red", "📝 Informe PNL", "🕵️‍♂️ Búsqueda Victimarios", "👤 Patrón y KML"])

    # ------------------------------------------
    # TAB 1: BÚSQUEDA GEOGRÁFICA MÚLTIPLE
    # ------------------------------------------
    with tab1:
        st.header("📍 Análisis de Trayectorias y Permanencia")
        st.markdown("Busca hasta 5 objetivos simultáneos. Si varios números coinciden en el mismo lugar, el mapa los agrupará para mostrar la data de todos.")
        
        c_fd1, c_ft1, c_fd2, c_ft2, c_top = st.columns(5)
        d_ini = c_fd1.date_input("Fecha Inicio", f_min, key="t1_d1")
        t_ini = c_ft1.time_input("Hora Inicio", datetime.time(0, 0), key="t1_t1")
        d_fin = c_fd2.date_input("Fecha Fin", f_max, key="t1_d2")
        t_fin = c_ft2.time_input("Hora Fin", datetime.time(23, 59), key="t1_t2")
        top_mostrar = c_top.selectbox("Mostrar Top Lugares:", ["Todos", 10, 5, 3, 1], key="in_top1")

        cn1, cn2, cn3, cn4, cn5 = st.columns(5)
        n1 = cn1.text_input("Objetivo 1:", placeholder="300...", key="n1")
        n2 = cn2.text_input("Objetivo 2:", placeholder="310...", key="n2")
        n3 = cn3.text_input("Objetivo 3:", placeholder="320...", key="n3")
        n4 = cn4.text_input("Objetivo 4:", placeholder="315...", key="n4")
        n5 = cn5.text_input("Objetivo 5:", placeholder="318...", key="n5")

        st.button("🗺️ Trazar Mapa de Inteligencia", on_click=activar_t1, type="primary")

        if st.session_state.run_t1:
            lista_nums = [n.strip() for n in [n1, n2, n3, n4, n5] if len(n.strip()) >= 7]
            
            if lista_nums:
                df_u = df_master[(df_master['numero_limpio'].isin(lista_nums)) & (df_master['latitud'].notna())].copy()
                dt_ini, dt_fin = pd.to_datetime(f"{d_ini} {t_ini}"), pd.to_datetime(f"{d_fin} {t_fin}")
                df_u = df_u[(df_u['fecha_limpia'] >= dt_ini) & (df_u['fecha_limpia'] <= dt_fin)].sort_values('fecha_limpia')

                if not df_u.empty:
                    # Agrupar por Coordenadas para solucionar la sobreposición (Overlapping)
                    df_coords = df_u.groupby(['latitud', 'longitud']).apply(lambda x: pd.Series({
                        'total_visitas': len(x),
                        'objetivos_presentes': list(x['numero_limpio'].unique()),
                        'data_html': "".join([
                            f"<b>Línea:</b> <span style='color:blue'>{num}</span> | <b>Visitas:</b> {len(x[x['numero_limpio']==num])}<br>"
                            f"<b>Celda:</b> {', '.join([c for c in x[x['numero_limpio']==num]['celda'].astype(str).unique() if c not in ('nan', 'Desconocida')])}<br>"
                            f"<b>Fechas:</b> {x[x['numero_limpio']==num]['fecha_limpia'].min()} a {x[x['numero_limpio']==num]['fecha_limpia'].max()}<hr style='margin:5px 0px;'>"
                            for num in x['numero_limpio'].unique()
                        ])
                    })).reset_index()

                    df_coords = df_coords.sort_values('total_visitas', ascending=False)
                    if top_mostrar != "Todos": df_coords = df_coords.head(int(top_mostrar))

                    st.markdown('<div class="caja-info">💡 <b>Sinopsis Geográfica:</b> Mapa interactivo generado. Los puntos están agrupados; si varios objetivos estuvieron en el mismo sitio, el cursor te mostrará el desglose de cada uno.</div>', unsafe_allow_html=True)

                    m = folium.Map(location=[df_coords.iloc[0]['latitud'], df_coords.iloc[0]['longitud']], zoom_start=13)

                    for rank_idx, (_, row) in enumerate(df_coords.iterrows()):
                        rank = rank_idx + 1
                        # Color dependiendo si hay más de 1 objetivo en ese punto
                        color = 'red' if len(row['objetivos_presentes']) > 1 else 'blue'
                        
                        popup_html = f"""<div style="min-width:300px; font-size:12px;">
                        <h4 style="margin:0;">Lugar Top #{rank}</h4><hr style="margin:5px 0px;">
                        {row['data_html']}
                        </div>"""
                        
                        folium.CircleMarker([row['latitud'], row['longitud']], radius=min(25, 8+(row['total_visitas']/2)), color=color, fill=True, fill_opacity=0.5).add_to(m)
                        folium.Marker([row['latitud'], row['longitud']], popup=folium.Popup(popup_html, max_width=400), tooltip=f"Top #{rank} | {len(row['objetivos_presentes'])} Objetivo(s) aquí", icon=folium.Icon(color=color)).add_to(m)

                    st_folium(m, use_container_width=True, height=500, returned_objects=[])
                else:
                    st.warning("No hay coordenadas para los objetivos en este rango de fechas y horas.")
            else:
                st.error("Ingresa al menos un número válido.")

    # ------------------------------------------
    # TAB 2: INTELIGENCIA DE RED (Restaurada Full)
    # ------------------------------------------
    with tab2:
        st.header("⚙️ Inteligencia de Red (Top Celdas y Encuentros)")
        
        filtro_colombia = st.checkbox("Ignorar Plataformas (Solo Móviles Colombia)", value=True, key="t2_chk")
        c1, c2, c3, c4 = st.columns(4)
        d_ini_2 = c1.date_input("Fecha Inicio", f_min, key="t2_d1")
        t_ini_2 = c2.time_input("Hora Inicio", datetime.time(0, 0), key="t2_t1")
        d_fin_2 = c3.date_input("Fecha Fin", f_max, key="t2_d2")
        t_fin_2 = c4.time_input("Hora Fin", datetime.time(23, 59), key="t2_t2")
        excluir = st.text_input("Excluir Números (Separados por coma):")

        st.button("Generar Análisis General", on_click=activar_t2, type="primary")

        if st.session_state.run_t2:
            with st.spinner("Procesando redes y topologías..."):
                df_ana = df_master.dropna(subset=['numero_limpio', 'celda', 'fecha_limpia']).copy()
                if filtro_colombia: df_ana = df_ana[df_ana['numero_limpio'].str.match(r'^3\d{9}$')]
                
                dt_i, dt_f = pd.to_datetime(f"{d_ini_2} {t_ini_2}"), pd.to_datetime(f"{d_fin_2} {t_fin_2}")
                df_ana = df_ana[(df_ana['fecha_limpia'] >= dt_i) & (df_ana['fecha_limpia'] <= dt_f)]
                
                if excluir: df_ana = df_ana[~df_ana['numero_limpio'].isin([x.strip() for x in excluir.split(',')])]

                if not df_ana.empty:
                    top_c = df_ana.groupby(['celda', 'numero_limpio']).size().reset_index(name='conexiones').sort_values('conexiones', ascending=False).head(10)
                    
                    df_ana['vent'] = df_ana['fecha_limpia'].dt.floor('15min')
                    ct = Counter()
                    for nums in df_ana.groupby(['celda', 'vent'])['numero_limpio'].unique():
                        if 1 < len(nums) <= 150: ct.update(itertools.combinations(sorted(nums), 2))
                    top_e = pd.DataFrame([{'Número A': p[0], 'Número B': p[1], 'Match': c} for p, c in ct.most_common(10)])

                    s_txt = f"El tráfico más alto se ancló a la celda <b>{top_c.iloc[0]['celda'] if not top_c.empty else 'N/A'}</b>. "
                    if not top_e.empty: s_txt += f"Además, se constata un probable encuentro recurrente entre <b>{top_e.iloc[0]['Número A']}</b> y <b>{top_e.iloc[0]['Número B']}</b> ({top_e.iloc[0]['Match']} cruces confirmados en 15 min)."
                    st.markdown(f'<div class="caja-info">💡 <b>Inteligencia Activa:</b> {s_txt}</div>', unsafe_allow_html=True)

                    colA, colB = st.columns(2)
                    with colA:
                        st.subheader("🏆 Top Celdas Frecuentes")
                        st.dataframe(top_c, use_container_width=True, hide_index=True)
                    with colB:
                        st.subheader("📍 Encuentros Probables (15 min)")
                        st.dataframe(top_e, use_container_width=True, hide_index=True)

    # ------------------------------------------
    # TAB 3: INFORME NLP AVANZADO (CORRELACIÓN)
    # ------------------------------------------
    with tab3:
        st.header("📝 Informe PNL (Correlación de Relatos y Evidencia)")
        st.markdown("Extrae entidades (líneas) de documentos anexos y correlaciona Modo, Tiempo y Lugar.")
        
        c1, c2, c3, c4 = st.columns(4)
        d_ini_3 = c1.date_input("Fecha Inicio", f_min, key="t3_d1")
        t_ini_3 = c2.time_input("Hora Inicio", datetime.time(0, 0), key="t3_t1")
        d_fin_3 = c3.date_input("Fecha Fin", f_max, key="t3_d2")
        t_fin_3 = c4.time_input("Hora Fin", datetime.time(23, 59), key="t3_t2")

        texto_in = st.text_area("Texto de los hechos (opcional si subes docs):", height=100)
        docs_in = st.file_uploader("Documentos (PDF, DOCX, TXT):", accept_multiple_files=True, type=['pdf', 'docx', 'doc', 'txt'])

        st.button("🔍 Extraer y Correlacionar", on_click=activar_t3, type="primary")

        if st.session_state.run_t3:
            txt_full = texto_in + "\n"
            if docs_in:
                for f in docs_in:
                    try:
                        if f.name.endswith('.txt'): txt_full += f.read().decode('utf-8', errors='ignore') + "\n"
                        elif f.name.endswith('.pdf'): 
                            reader = PyPDF2.PdfReader(io.BytesIO(f.read()))
                            for p in reader.pages: txt_full += (p.extract_text() or "") + "\n"
                        elif f.name.endswith(('.docx', '.doc')):
                            doc = docx.Document(io.BytesIO(f.read()))
                            for p in doc.paragraphs: txt_full += p.text + "\n"
                    except: pass

            nums_implicados = list(set(re.findall(r'\b3\d{9}\b', txt_full)))
            
            if nums_implicados:
                df_c = df_master[df_master['numero_limpio'].isin(nums_implicados)].copy()
                dt_i, dt_f = pd.to_datetime(f"{d_ini_3} {t_ini_3}"), pd.to_datetime(f"{d_fin_3} {t_fin_3}")
                df_c = df_c[(df_c['fecha_limpia'] >= dt_i) & (df_c['fecha_limpia'] <= dt_f)].sort_values('fecha_limpia')

                if not df_c.empty:
                    n_activos = len(df_c['numero_limpio'].unique())
                    celdas_comunes = df_c['celda'].value_counts().head(3).index.tolist()
                    
                    st.markdown(f"""
                    <div class="caja-info">
                        <h4>📋 Sinopsis Criminal (Tiempo, Modo y Lugar)</h4>
                        <p><b>Números Correlacionados:</b> {', '.join(df_c['numero_limpio'].unique())} ({n_activos} de los hallados en el relato).</p>
                        <p><b>⏱️ Tiempo:</b> Operatividad confirmada entre el <b>{df_c['fecha_limpia'].min()}</b> y el <b>{df_c['fecha_limpia'].max()}</b>, generando <b>{len(df_c)}</b> trazas en la red.</p>
                        <p><b>📍 Lugar:</b> La interconexión geográfica sitúa su mayor convergencia en las celdas: <b>{', '.join([str(c) for c in celdas_comunes])}</b>.</p>
                        <p><b>⚙️ Modo:</b> El relato sugiere asociación entre las líneas perfiladas, confirmada por el volumen de tráfico solapado en la misma ventana espacial de los hechos.</p>
                    </div>
                    """, unsafe_allow_html=True)

                    df_m = df_c.dropna(subset=['latitud', 'longitud'])
                    if not df_m.empty:
                        df_agrup_m = df_m.groupby(['latitud', 'longitud']).apply(lambda x: pd.Series({
                            'visitas': len(x),
                            'data_html': "".join([f"<b>{num}</b>: {len(x[x['numero_limpio']==num])} impactos | Celda: {', '.join([c for c in x[x['numero_limpio']==num]['celda'].astype(str).unique() if c!='Desconocida'])} | Fechas: {x[x['numero_limpio']==num]['fecha_limpia'].min()} - {x[x['numero_limpio']==num]['fecha_limpia'].max()}<hr style='margin:2px;'>" for num in x['numero_limpio'].unique()])
                        })).reset_index()

                        m = folium.Map(location=[df_agrup_m.iloc[0]['latitud'], df_agrup_m.iloc[0]['longitud']], zoom_start=13)
                        for rank_idx, (_, r) in enumerate(df_agrup_m.sort_values('visitas', ascending=False).iterrows()):
                            pop = f"<div style='min-width:250px;'><b>Top #{rank_idx+1} (Visitas: {r['visitas']})</b><hr>{r['data_html']}</div>"
                            folium.CircleMarker([r['latitud'], r['longitud']], radius=min(20, 5+r['visitas']), color='darkred', fill=True).add_to(m)
                            folium.Marker([r['latitud'], r['longitud']], tooltip=f"Top #{rank_idx+1} | Visitas: {r['visitas']}", popup=folium.Popup(pop, max_width=400)).add_to(m)
                        st_folium(m, use_container_width=True, height=400)
                else:
                    st.warning("Los números existen, pero no registran actividad en las Fechas seleccionadas.")
            else:
                st.error("No se detectaron números móviles válidos (10 dígitos, inicio 3) en el relato.")

    # ------------------------------------------
    # TAB 4: CO-DESPLAZAMIENTO (CRUCE ESPACIAL AVANZADO)
    # ------------------------------------------
    with tab4:
        st.header("🕵️‍♂️ Búsqueda de Victimarios (Rolling Window Avanzado)")
        st.markdown("Identifica qué líneas ajenas clonaron la ruta de la víctima. Integra búsqueda espacial (redondeo geográfico) para unir operadores distintos.")
        
        c1, c2, c3, c4 = st.columns(4)
        vic_num = c1.text_input("Número Víctima (10 dígitos):", key="t4_vic")
        tol = c2.selectbox("Tolerancia Temporal:", ['15min', '30min', '1H', '2H'], index=1)
        d_i_4 = c3.date_input("Fecha Inicio", f_min, key="t4_d1")
        t_i_4 = c4.time_input("Hora Inicio", datetime.time(0, 0), key="t4_t1")
        
        c5, c6, c7, c8 = st.columns(4)
        f_col = c5.checkbox("Excluir Plataformas (Inicio 3)", value=True, key="t4_chk")
        d_f_4 = c7.date_input("Fecha Fin", f_max, key="t4_d2")
        t_f_4 = c8.time_input("Hora Fin", datetime.time(23, 59), key="t4_t2")

        st.button("🔍 Rastrear Cruces", on_click=activar_t4, type="primary")

        if st.session_state.run_t4:
            if len(vic_num) >= 7:
                with st.spinner("Correlacionando mallas de ubicación..."):
                    df_l = df_master.dropna(subset=['numero_limpio', 'fecha_limpia']).copy()
                    dt_i, dt_f = pd.to_datetime(f"{d_i_4} {t_i_4}"), pd.to_datetime(f"{d_f_4} {t_f_4}")
                    df_l = df_l[(df_l['fecha_limpia'] >= dt_i) & (df_l['fecha_limpia'] <= dt_f)]
                    
                    # Redondeo a 3 decimales (Aprox 100 metros) para atrapar antenas conjuntas entre operadores (Ej: Claro y Tigo en la misma torre)
                    df_l['lat_r'] = df_l['latitud'].round(3)
                    df_l['lon_r'] = df_l['longitud'].round(3)
                    
                    df_v = df_l[df_l['numero_limpio'] == vic_num].copy()
                    df_o = df_l[df_l['numero_limpio'] != vic_num].copy()
                    if f_col: df_o = df_o[df_o['numero_limpio'].str.match(r'^3\d{9}$')]

                    if not df_v.empty:
                        res = []
                        td = pd.Timedelta(tol)
                        for _, r in df_v.iterrows():
                            # 1. Match por Celda Exacta
                            if r['celda'] != 'Desconocida':
                                match = df_o[(df_o['celda'] == r['celda']) & (df_o['fecha_limpia'] >= r['fecha_limpia']-td) & (df_o['fecha_limpia'] <= r['fecha_limpia']+td)]
                            # 2. Si es Desconocida o falla celda, Match por Coordenada Cercana (Radio 100m)
                            else:
                                if pd.notna(r['lat_r']) and pd.notna(r['lon_r']):
                                    match = df_o[(df_o['lat_r'] == r['lat_r']) & (df_o['lon_r'] == r['lon_r']) & (df_o['fecha_limpia'] >= r['fecha_limpia']-td) & (df_o['fecha_limpia'] <= r['fecha_limpia']+td)]
                                else: match = pd.DataFrame()
                                
                            for _, sr in match.iterrows():
                                res.append({'numero_limpio': sr['numero_limpio'], 'celda': r['celda'] if r['celda']!='Desconocida' else 'Coords_Cercanas'})
                        
                        if res:
                            sosp = pd.DataFrame(res).groupby('numero_limpio').agg(ubicaciones_distintas=('celda','nunique'), impactos_totales=('celda','count')).reset_index().sort_values(['ubicaciones_distintas','impactos_totales'], ascending=False).head(10)
                            st.markdown(f'<div class="caja-alerta">💡 <b>Sospechosos Identificados:</b> El algoritmo señala a <b>{sosp.iloc[0]["numero_limpio"]}</b> como principal victimario, interceptando a la víctima en <b>{sosp.iloc[0]["ubicaciones_distintas"]}</b> ubicaciones/celdas distintas en el rango de {tol}.</div>', unsafe_allow_html=True)
                            st.dataframe(sosp, use_container_width=True, hide_index=True)
                        else:
                            st.success("Cruce Negativo: Ninguna línea siguió a la víctima en ese rango espacial-temporal.")
                    else:
                        st.warning("La víctima no registra actividad bajo estas fechas y horas.")

    # ------------------------------------------
    # TAB 5: PATRÓN Y DATOS TÉCNICOS
    # ------------------------------------------
    with tab5:
        st.header("👤 Perfilamiento de Rutinas y Ficha Técnica")
        st.markdown("Visualiza las zonas de pernocta, operaciones de día, análisis individual/grupal y extrae la **ficha técnica completa** para informes periciales.")
        
        c_1, c_2, c_3, c_4, c_5 = st.columns(5)
        o1 = c_1.text_input("Objetivo 1:", key="o1")
        o2 = c_2.text_input("Objetivo 2:", key="o2")
        o3 = c_3.text_input("Objetivo 3:", key="o3")
        o4 = c_4.text_input("Objetivo 4:", key="o4")
        o5 = c_5.text_input("Objetivo 5:", key="o5")

        cf1, cf2, cf3, cf4 = st.columns(4)
        f_col_5 = cf1.checkbox("Excluir Plataformas", value=True, key="t5_chk_p")
        d_i_5 = cf2.date_input("Fecha Inicio", f_min, key="t5_d1")
        d_f_5 = cf3.date_input("Fecha Fin", f_max, key="t5_d2")

        st.button("Generar Perfiles y Ficha Técnica", on_click=activar_t5, type="primary")
        
        if st.session_state.run_t5:
            list_obj = [o.strip() for o in [o1, o2, o3, o4, o5] if len(o.strip()) >= 7]
            if list_obj:
                df_p = df_master.dropna(subset=['numero_limpio', 'fecha_limpia']).copy()
                if f_col_5: df_p = df_p[df_p['numero_limpio'].str.match(r'^3\d{9}$')]
                
                dt_i, dt_f = pd.to_datetime(d_i_5), pd.to_datetime(d_f_5) + pd.Timedelta(days=1)
                df_p = df_p[(df_p['fecha_limpia'] >= dt_i) & (df_p['fecha_limpia'] <= dt_f)]
                
                df_obj = df_p[df_p['numero_limpio'].isin(list_obj)].copy()
                
                if not df_obj.empty:
                    # FICHA TÉCNICA
                    st.subheader("📡 Ficha Técnica Forense de Antenas")
                    ficha_tecnica = df_obj[['numero_limpio', 'fecha_limpia', 'nombre_antena', 'celda', 'lac', 'latitud', 'longitud', 'haz', 'azimut']].copy()
                    st.dataframe(ficha_tecnica.sort_values('fecha_limpia'), use_container_width=True, hide_index=True)
                    
                    def cl_h(h):
                        if h >= 22 or h <= 2: return '🌙 NOCHE (Pernocta: 22:00 - 02:59)'
                        elif 3 <= h <= 8: return '🌅 MADRUGADA (Amanecer: 03:00 - 08:59)'
                        else: return '☀️ DÍA (Operación: 09:00 - 21:59)'
                        
                    df_obj['franja'] = df_obj['fecha_limpia'].dt.hour.apply(cl_h)
                    rut = df_obj.groupby(['numero_limpio', 'franja', 'celda', 'nombre_antena']).size().reset_index(name='impactos').sort_values(['numero_limpio', 'franja','impactos'], ascending=[True, True, False])
                    
                    st.subheader("⏱️ Patrones Horarios de Ubicación")
                    
                    # Sinopsis individual y conjunta
                    txt_sinop = "<b>Análisis de Rutinas:</b><br>"
                    for obj in list_obj:
                        df_o_r = rut[rut['numero_limpio'] == obj]
                        if not df_o_r.empty:
                            noche = df_o_r[df_o_r['franja'].str.contains('NOCHE')]
                            noche_txt = noche.iloc[0]['celda'] if not noche.empty else "N/A"
                            dia = df_o_r[df_o_r['franja'].str.contains('DÍA')]
                            dia_txt = dia.iloc[0]['celda'] if not dia.empty else "N/A"
                            txt_sinop += f"• El objetivo <b>{obj}</b> pernocta principalmente en la celda {noche_txt} y opera de día en la celda {dia_txt}.<br>"
                    
                    st.markdown(f'<div class="caja-info">{txt_sinop}</div>', unsafe_allow_html=True)
                    st.dataframe(rut, use_container_width=True, hide_index=True)

                    df_kml = df_obj.dropna(subset=['latitud', 'longitud'])
                    if not df_kml.empty:
                        kml = simplekml.Kml()
                        for obj in list_obj:
                            df_kml_obj = df_kml[df_kml['numero_limpio'] == obj]
                            if not df_kml_obj.empty:
                                fol = kml.newfolder(name=f"Objetivo_{obj}")
                                for _, r in df_kml_obj.iterrows():
                                    pnt = fol.newpoint(name=str(r['celda']), coords=[(r['longitud'], r['latitud'])])
                                    pnt.description = f"<b>Número:</b> {r['numero_limpio']}<br><b>Fecha:</b> {r['fecha_limpia']}<br><b>Franja:</b> {r['franja']}<br><b>Antena:</b> {r['nombre_antena']}<br><b>LAC:</b> {r['lac']}<br><b>HAZ:</b> {r['haz']}<br><b>AZIMUT:</b> {r['azimut']}"
                        
                        st.download_button("🌍 Descargar KMZ/KML (Google Earth)", data=kml.kml(), file_name="Perfiles_Tecnicos.kml", mime="application/vnd.google-earth.kml+xml")
                else:
                    st.error("Sin registros validos bajo esos filtros y números.")
            else:
                st.warning("Ingresa al menos un número objetivo.")

else:
    st.info("👈 Por favor, carga las bases de datos CDR en el menú lateral para iniciar la Inteligencia Artificial.")