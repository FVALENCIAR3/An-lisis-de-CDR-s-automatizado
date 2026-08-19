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
import math

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
    st.markdown('<div class="desc-gaula">Análisis de inteligencia artificial que permite limpiar, organizar y analizar las bases de datos en archivos consolidados. Convierte celdas hexadecimales, excluye números de plataformas, representa gráficamente trayectorias, y genera sinopsis de comportamiento (permanencia, residencia, trabajo y co-desplazamiento) del objetivo.</div>', unsafe_allow_html=True)

st.divider()

if 'df_master' not in st.session_state: st.session_state.df_master = None
for key in ['run_t1', 'run_t2', 'run_t3', 'run_t4', 'run_t5', 'run_t6']:
    if key not in st.session_state: st.session_state[key] = False

def activar_t1(): st.session_state.run_t1 = True
def activar_t2(): st.session_state.run_t2 = True
def activar_t3(): st.session_state.run_t3 = True
def activar_t4(): st.session_state.run_t4 = True
def activar_t5(): st.session_state.run_t5 = True
def activar_t6(): st.session_state.run_t6 = True

MIN_DATE = datetime.date(1990, 1, 1)
MAX_DATE = datetime.date(2050, 12, 31)

# ==========================================
# FUNCIONES MATEMÁTICAS Y DE OPERADORES
# ==========================================
def obtener_operador_colombia(num_str):
    if len(num_str) != 10 or not num_str.startswith('3'): return "DESCONOCIDO / PLATAFORMA"
    try:
        pref = int(num_str[:3])
        num_int = int(num_str)
        if pref in [300, 301, 303, 304]: return "Tigo"
        elif pref == 302:
            if 3024700000 <= num_int <= 3028699999: return "Wom"
            else: return "Tigo"
        elif pref == 305: return "ETB"
        elif pref in [310, 311, 312, 313, 314, 320, 321, 322]: return "Claro"
        elif pref == 323:
            if 3236000000 <= num_int <= 3239999999: return "Wom"
            else: return "Claro"
        elif pref == 324:
            if 3241000000 <= num_int <= 3241999999 or 3247000000 <= num_int <= 3249999999: return "Wom"
            else: return "Tigo"
        elif pref in [315, 316, 317, 318]: return "Movistar"
        elif pref == 319: return "Virgin Mobile"
        elif pref in [350, 351]: return "Avantel"
        elif pref == 333: return "Suma Móvil"
        return "OTROS MÓVILES"
    except: return "DESCONOCIDO"

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
        if val in ('NAN', 'NONE', 'NULL', ''): return 'Desconocida'
        if re.match(r'^[0-9A-F]+-[0-9A-F]+$', val):
            try: return '-'.join([str(int(p, 16)) for p in val.split('-')])
            except: return val
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

    def buscar_columna(df_cols, alias_list):
        for a in alias_list:
            if a in df_cols: return a
        for a in alias_list:
            for c in df_cols:
                if a in c: return c
        return None

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
            alias_lat = ['latitud', 'latitud_n', 'latitude', 'lat', 'lat_origen']
            alias_lon = ['longitud', 'longitud_w', 'longitude', 'lon', 'longitud_origen', 'long']
            alias_cel = ['celda_origen_truncada', 'celda_decimal', 'cell_id_voz', 'celda_inicio', 'celda_origen', 'celda_hex', 'cellid_nval', 'bts_id', 'site_id', 'cell_id', 'cgi', 'codigo_celda', 'celda', 'cell']
            alias_lac = ['lac_decimal', 'lac', 'location_area_code', 'tac']
            alias_haz = ['horizontal_beamwidth', 'angulo_beam_horizontal', 'sector', 'haz', 'beam', 'sector_id']
            alias_azi = ['azimuth', 'azimut', 'dir', 'direccion_antena']
            alias_nom = ['nombre_antena', 'descripcion', 'direccion', 'site_name', 'nombre_sitio', 'address', 'ubicacion', 'city_ds']
            
            col_tel = buscar_columna(df.columns, alias_tel)
            col_fec = buscar_columna(df.columns, alias_fec)
            col_lat = buscar_columna(df.columns, alias_lat)
            col_lon = buscar_columna(df.columns, alias_lon)
            col_cel = buscar_columna(df.columns, alias_cel)
            col_lac = buscar_columna(df.columns, alias_lac)
            col_haz = buscar_columna(df.columns, alias_haz)
            col_azi = buscar_columna(df.columns, alias_azi)
            col_nom = buscar_columna(df.columns, alias_nom)
            
            df_temp = pd.DataFrame()
            if col_tel: 
                df_temp['numero_limpio'] = df[col_tel].astype(str).str.replace(r'\D', '', regex=True).str[-10:]
                df_temp['operador'] = df_temp['numero_limpio'].apply(obtener_operador_colombia)
            if col_fec: 
                df_temp['fecha_limpia'] = limpiar_fechas(df[col_fec])
                df_temp['fecha_limpia'] = df_temp['fecha_limpia'].dt.tz_localize(None)
            if col_lat and col_lon:
                df_temp['latitud'] = limpiar_coord(df[col_lat])
                df_temp['longitud'] = limpiar_coord(df[col_lon])
            
            if col_cel: 
                df_temp['celda'] = df[col_cel].astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
                df_temp['celda'] = df_temp['celda'].apply(normalizar_celda)
            else: df_temp['celda'] = 'Desconocida'
                
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
# BARRA LATERAL (SUBIDA DE ARCHIVOS)
# ==========================================
with st.sidebar.form("form_carga"):
    st.markdown("**📂 1. Cargar Archivos Consolidables**")
    uploaded_files = st.file_uploader("Soporta CDRs: CSV, TXT, XLS, XLSX, XLSB", accept_multiple_files=True, type=['csv', 'txt', 'xls', 'xlsx', 'xlsb'])
    submit_button = st.form_submit_button("⚙️ Procesar Archivos (I.A)")

if submit_button:
    if uploaded_files:
        with st.spinner("Procesando gigabytes de datos e indexando..."):
            st.session_state.df_master = procesar_archivos(uploaded_files)
            for key in ['run_t1', 'run_t2', 'run_t3', 'run_t4', 'run_t5', 'run_t6']: st.session_state[key] = False
        if st.session_state.df_master is not None and not st.session_state.df_master.empty:
            st.sidebar.success(f"✅ {len(st.session_state.df_master)} registros consolidados exitosamente.")
        else:
            st.sidebar.error("❌ Los archivos no contienen columnas compatibles.")

# ==========================================
# CUERPO PRINCIPAL (PESTAÑAS DE INTELIGENCIA)
# ==========================================
if st.session_state.df_master is not None:
    df_master = st.session_state.df_master
    try: f_min, f_max = df_master['fecha_limpia'].min().date(), df_master['fecha_limpia'].max().date()
    except: f_min = f_max = datetime.date.today()

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📍 Búsqueda Geográfica", "⚙️ Inteligencia de Red", "📝 Informe PNL", "🕵️‍♂️ Búsqueda Victimarios", "👤 Patrón y Ficha Técnica", "📡 Antenas 3D (G-NetTilt)"])

    # ------------------------------------------
    # TAB 1: BÚSQUEDA GEOGRÁFICA MÚLTIPLE
    # ------------------------------------------
    with tab1:
        st.header("📍 Análisis de Trayectorias y Permanencia")
        st.markdown("Busca hasta 5 objetivos simultáneos. Traza la ruta en sentido de movimientos y resalta en **fucsia** los puntos de convergencia.")
        
        c_fd1, c_ft1, c_fd2, c_ft2, c_top = st.columns(5)
        d_ini = c_fd1.date_input("Fecha Inicio", f_min, min_value=MIN_DATE, max_value=MAX_DATE, key="t1_d1")
        t_ini = c_ft1.time_input("Hora Inicio", datetime.time(0, 0), key="t1_t1")
        d_fin = c_fd2.date_input("Fecha Fin", f_max, min_value=MIN_DATE, max_value=MAX_DATE, key="t1_d2")
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
                    df_coords = df_u.groupby(['latitud', 'longitud']).apply(lambda x: pd.Series({
                        'total_visitas': len(x),
                        'objetivos_presentes': list(x['numero_limpio'].unique()),
                        'data_html': "".join([
                            f"<b>Línea:</b> <span style='color:#003366'>{num}</span> | <b>Visitas:</b> {len(x[x['numero_limpio']==num])}<br>"
                            f"<b>Celda:</b> {', '.join([str(c) for c in x[x['numero_limpio']==num]['celda'].unique() if str(c) not in ('nan', 'Desconocida')])}<br>"
                            f"<b>Fechas:</b> {x[x['numero_limpio']==num]['fecha_limpia'].min()} a {x[x['numero_limpio']==num]['fecha_limpia'].max()}<hr style='margin:5px 0px;'>"
                            for num in x['numero_limpio'].unique()
                        ])
                    })).reset_index().sort_values('total_visitas', ascending=False)

                    if top_mostrar != "Todos": df_coords = df_coords.head(int(top_mostrar))

                    sinop_txt = "<ul>"
                    for obj in lista_nums:
                        df_obj_sinop = df_u[df_u['numero_limpio'] == obj]
                        if not df_obj_sinop.empty:
                            top_lugar = df_obj_sinop.groupby('celda').size().idxmax()
                            n_vis = df_obj_sinop.groupby('celda').size().max()
                            sinop_txt += f"<li>El objetivo <b>{obj}</b> presenta su mayor zona de permanencia (residencia o trabajo) en la celda <b>{top_lugar}</b> con {n_vis} registros en las fechas seleccionadas.</li>"
                        else:
                            sinop_txt += f"<li>El objetivo <b>{obj}</b> no presenta tráfico georreferenciado en este periodo.</li>"
                    sinop_txt += "</ul>"

                    st.markdown(f'<div class="caja-info">💡 <b>Sinopsis Geográfica Analizada:</b><br>{sinop_txt}</div>', unsafe_allow_html=True)

                    m = folium.Map(location=[df_coords.iloc[0]['latitud'], df_coords.iloc[0]['longitud']], zoom_start=13)
                    colores_obj = ['red', 'blue', 'green', 'orange', 'black']
                    mapa_colores = {num: colores_obj[i % len(colores_obj)] for i, num in enumerate(lista_nums)}

                    for rank_idx, (_, row) in enumerate(df_coords.iterrows()):
                        rank = rank_idx + 1
                        color = 'fuchsia' if len(row['objetivos_presentes']) > 1 else mapa_colores.get(row['objetivos_presentes'][0], 'gray')
                        
                        popup_html = f"""<div style="min-width:300px; font-size:12px;">
                        <h4 style="margin:0;">Top Lugar #{rank}</h4><hr style="margin:5px 0px;">
                        {row['data_html']}
                        </div>"""
                        
                        folium.CircleMarker([row['latitud'], row['longitud']], radius=min(25, 8+(row['total_visitas']/2)), color=color, fill=True, fill_opacity=0.6).add_to(m)
                        folium.Marker([row['latitud'], row['longitud']], popup=folium.Popup(popup_html, max_width=400), tooltip=f"Top #{rank} | {len(row['objetivos_presentes'])} Objetivo(s)", icon=folium.Icon(color=color)).add_to(m)

                    for obj in lista_nums:
                        df_ruta = df_u[df_u['numero_limpio'] == obj].sort_values('fecha_limpia')
                        if len(df_ruta) > 1:
                            AntPath(locations=df_ruta[['latitud', 'longitud']].values.tolist(), color=mapa_colores[obj], weight=3, tooltip=f"Ruta: {obj}").add_to(m)

                    st_folium(m, use_container_width=True, height=500, returned_objects=[])
                else:
                    st.warning("No hay coordenadas para los objetivos en este rango de fechas y horas.")
            else:
                st.error("Ingresa al menos un número válido.")

    # ------------------------------------------
    # TAB 2: INTELIGENCIA DE RED
    # ------------------------------------------
    with tab2:
        st.header("⚙️ Inteligencia de Red (Top Celdas y Encuentros)")
        
        c1, c2, c3, c4 = st.columns(4)
        d_ini_2 = c1.date_input("Fecha Inicio", f_min, min_value=MIN_DATE, max_value=MAX_DATE, key="t2_d1")
        t_ini_2 = c2.time_input("Hora Inicio", datetime.time(0, 0), key="t2_t1")
        d_fin_2 = c3.date_input("Fecha Fin", f_max, min_value=MIN_DATE, max_value=MAX_DATE, key="t2_d2")
        t_fin_2 = c4.time_input("Hora Fin", datetime.time(23, 59), key="t2_t2")
        
        c5, c6 = st.columns(2)
        tol_enc = c5.selectbox("Tolerancia Temporal de Encuentros:", ['15min', '30min', '45min', '1H', '2H', '3H'])
        incluir = c6.text_input("Incluir Números (Separados por coma - Opcional):", placeholder="Ej: 3001234567, 3159876543")

        st.button("Generar Análisis General", on_click=activar_t2, type="primary")

        if st.session_state.run_t2:
            with st.spinner("Procesando redes y topologías..."):
                df_ana = df_master.dropna(subset=['numero_limpio', 'celda', 'fecha_limpia']).copy()
                
                dt_i, dt_f = pd.to_datetime(f"{d_ini_2} {t_ini_2}"), pd.to_datetime(f"{d_fin_2} {t_fin_2}")
                df_ana = df_ana[(df_ana['fecha_limpia'] >= dt_i) & (df_ana['fecha_limpia'] <= dt_f)]
                
                if incluir: 
                    nums_inc = [x.strip() for x in incluir.split(',') if x.strip()]
                    df_ana = df_ana[df_ana['numero_limpio'].isin(nums_inc)]

                if not df_ana.empty:
                    top_c = df_ana.groupby(['celda', 'numero_limpio']).size().reset_index(name='conexiones').sort_values('conexiones', ascending=False).head(10)
                    
                    tol_p = tol_enc.lower() 
                    df_ana['vent'] = df_ana['fecha_limpia'].dt.floor(tol_p)
                    
                    encuentros_list = []
                    for (celda, vent), df_g in df_ana.groupby(['celda', 'vent']):
                        nums = df_g['numero_limpio'].unique()
                        if 1 < len(nums) <= 150:
                            # CORRECCIÓN TYPEERROR: Convertir a str antes de hacer join
                            ops = ", ".join(df_g['operador'].dropna().astype(str).unique())
                            for pair in itertools.combinations(sorted(nums), 2):
                                encuentros_list.append({'Número A': pair[0], 'Número B': pair[1], 'Operadores': ops})
                                
                    if encuentros_list:
                        df_e = pd.DataFrame(encuentros_list)
                        top_e_all = df_e.groupby(['Número A', 'Número B']).size().reset_index(name='Coincidencias').sort_values('Coincidencias', ascending=False).head(10)
                        top_e_op = df_e.groupby(['Número A', 'Número B', 'Operadores']).size().reset_index(name='Coincidencias').sort_values('Coincidencias', ascending=False).head(10)
                    else:
                        top_e_all = pd.DataFrame()
                        top_e_op = pd.DataFrame()

                    s_txt = f"Se evaluó la ventana de datos del <b>{dt_i.date()}</b> al <b>{dt_f.date()}</b>. El tráfico máximo detectado se ancló a la celda <b>{top_c.iloc[0]['celda'] if not top_c.empty else 'N/A'}</b>. "
                    if not top_e_all.empty: s_txt += f"En el plano temporal (rango {tol_enc}), se constata un probable encuentro recurrente entre <b>{top_e_all.iloc[0]['Número A']}</b> y <b>{top_e_all.iloc[0]['Número B']}</b> ({top_e_all.iloc[0]['Coincidencias']} cruces confirmados)."
                    st.markdown(f'<div class="caja-info">💡 <b>Sinopsis de Red Analizada:</b><br>{s_txt}</div>', unsafe_allow_html=True)

                    st.subheader("🏆 Top Celdas Frecuentes")
                    st.dataframe(top_c, use_container_width=True, hide_index=True)
                    
                    colA, colB = st.columns(2)
                    with colA:
                        st.subheader(f"📍 Encuentros ({tol_enc} - Todos los Op.)")
                        if not top_e_all.empty: st.dataframe(top_e_all, use_container_width=True, hide_index=True)
                        else: st.warning("Sin encuentros.")
                    with colB:
                        st.subheader(f"📍 Encuentros ({tol_enc} - Detalle Operador)")
                        if not top_e_op.empty: st.dataframe(top_e_op, use_container_width=True, hide_index=True)
                        else: st.warning("Sin encuentros.")
                else:
                    st.warning("No hay registros bajo los parámetros solicitados.")

    # ------------------------------------------
    # TAB 3: INFORME NLP AVANZADO (CORRELACIÓN)
    # ------------------------------------------
    with tab3:
        st.header("📝 Informe PNL (Correlación de Relatos y Evidencia)")
        st.markdown("Extrae entidades (líneas) de documentos anexos y correlaciona Modo, Tiempo y Lugar exactos. Colorea los objetivos en el mapa y subraya coincidencias.")
        
        c1, c2, c3, c4 = st.columns(4)
        d_ini_3 = c1.date_input("Fecha Inicio", f_min, min_value=MIN_DATE, max_value=MAX_DATE, key="t3_d1")
        t_ini_3 = c2.time_input("Hora Inicio", datetime.time(0, 0), key="t3_t1")
        d_fin_3 = c3.date_input("Fecha Fin", f_max, min_value=MIN_DATE, max_value=MAX_DATE, key="t3_d2")
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
                    # Asignar colores fijos a los objetivos detectados
                    colores = ['red', 'blue', 'green', 'orange', 'black', 'cadetblue', 'darkred', 'purple']
                    c_map = {n: colores[i%len(colores)] for i, n in enumerate(df_c['numero_limpio'].unique())}
                    color_labels = [f"<span style='color:{c_map[n]}; font-weight:bold;'>{n}</span>" for n in df_c['numero_limpio'].unique()]

                    n_activos = len(df_c['numero_limpio'].unique())
                    celdas_comunes = df_c['celda'].value_counts().head(3).index.tolist()
                    
                    st.markdown(f"""
                    <div class="caja-info">
                        <h4>📋 Sinopsis Criminal (Tiempo, Modo y Lugar)</h4>
                        <p><b>Números Correlacionados:</b> {', '.join(color_labels)} ({n_activos} activos encontrados operando en este rango).</p>
                        <p><b>⏱️ Tiempo:</b> Operatividad confirmada entre el <b>{df_c['fecha_limpia'].min()}</b> y el <b>{df_c['fecha_limpia'].max()}</b>, con <b>{len(df_c)}</b> trazas en la red.</p>
                        <p><b>📍 Lugar:</b> Convergencia principal de estos objetivos detectada en las celdas: <b>{', '.join([str(c) for c in celdas_comunes if c != 'Desconocida'])}</b>.</p>
                        <p><b>⚙️ Modo:</b> Confirmada asociación espacial-temporal. Los puntos en <b>morado oscuro (fuchsia)</b> indican lugares donde la evidencia sitúa a varios sospechosos juntos.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.subheader("Tablas Explicativas (Actividad Real Detectada)")
                    st.dataframe(df_c.groupby(['numero_limpio', 'operador']).size().reset_index(name='Conexiones'), use_container_width=True, hide_index=True)

                    df_m = df_c.dropna(subset=['latitud', 'longitud'])
                    if not df_m.empty:
                        df_agrup_m = df_m.groupby(['latitud', 'longitud']).apply(lambda x: pd.Series({
                            'visitas': len(x),
                            'objetivos': list(x['numero_limpio'].unique()),
                            'data_html': "".join([f"<b><span style='color:{c_map.get(num,'black')}'>Línea {num}</span></b>: {len(x[x['numero_limpio']==num])} impactos | Celda: {', '.join([str(c) for c in x[x['numero_limpio']==num]['celda'].unique() if c!='Desconocida'])} | Fechas: {x[x['numero_limpio']==num]['fecha_limpia'].min()} a {x[x['numero_limpio']==num]['fecha_limpia'].max()}<hr style='margin:2px;'>" for num in x['numero_limpio'].unique()])
                        })).reset_index()

                        m = folium.Map(location=[df_agrup_m.iloc[0]['latitud'], df_agrup_m.iloc[0]['longitud']], zoom_start=13)
                        
                        for rank_idx, (_, r) in enumerate(df_agrup_m.sort_values('visitas', ascending=False).iterrows()):
                            pop = f"<div style='min-width:300px; font-size:12px;'><b>Top #{rank_idx+1} (Visitas Totales: {r['visitas']})</b><hr>{r['data_html']}</div>"
                            # MORADO (Fuchsia) si se encontraron juntos en ese punto
                            color_punto = 'fuchsia' if len(r['objetivos']) > 1 else c_map.get(r['objetivos'][0], 'gray')
                            
                            folium.CircleMarker([r['latitud'], r['longitud']], radius=min(20, 5+(r['visitas']/2)), color=color_punto, fill=True).add_to(m)
                            folium.Marker([r['latitud'], r['longitud']], tooltip=f"Top #{rank_idx+1} | Visitas: {r['visitas']}", popup=folium.Popup(pop, max_width=400), icon=folium.Icon(color=color_punto)).add_to(m)
                        
                        # Trazo de rutas por objetivo
                        for obj in df_c['numero_limpio'].unique():
                            df_ruta = df_c[df_c['numero_limpio'] == obj].sort_values('fecha_limpia')
                            if len(df_ruta) > 1:
                                AntPath(locations=df_ruta[['latitud', 'longitud']].values.tolist(), color=c_map[obj], weight=3, tooltip=f"Ruta: {obj}").add_to(m)

                        st_folium(m, use_container_width=True, height=400)
                        
                        # Generador KML con colores asimétricos y marcadores especiales
                        kml = simplekml.Kml()
                        for obj in df_m['numero_limpio'].unique():
                            df_kml_obj = df_m[df_m['numero_limpio'] == obj]
                            fol = kml.newfolder(name=f"Objetivo_{obj}")
                            for _, r in df_kml_obj.iterrows():
                                pnt = fol.newpoint(name=str(r['celda']), coords=[(r['longitud'], r['latitud'])])
                                pnt.description = f"Línea: {obj}<br>Fecha: {r['fecha_limpia']}<br>Antena: {r['nombre_antena']}"
                                
                                coincidencia = df_agrup_m[(df_agrup_m['latitud'] == r['latitud']) & (df_agrup_m['longitud'] == r['longitud'])]
                                if not coincidencia.empty and len(coincidencia.iloc[0]['objetivos']) > 1:
                                    pnt.style.iconstyle.color = simplekml.Color.fuchsia
                                else:
                                    c_name = c_map.get(obj, 'red')
                                    if c_name == 'blue': pnt.style.iconstyle.color = simplekml.Color.blue
                                    elif c_name == 'green': pnt.style.iconstyle.color = simplekml.Color.green
                                    elif c_name == 'orange': pnt.style.iconstyle.color = simplekml.Color.orange
                                    elif c_name == 'black': pnt.style.iconstyle.color = simplekml.Color.black
                                    else: pnt.style.iconstyle.color = simplekml.Color.red
                                pnt.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
                                
                        st.download_button("🌍 Descargar Ruta Georeferenciada (KML)", data=kml.kml(), file_name="Rutas_Correlacionadas.kml", mime="application/vnd.google-earth.kml+xml")

                else:
                    st.warning("Los números existen en BD, pero no registran actividad en las Fechas y Horas seleccionadas.")
            else:
                st.error("No se detectaron números móviles válidos (10 dígitos, inicio 3) en el relato.")

    # ------------------------------------------
    # TAB 4: CO-DESPLAZAMIENTO (CRUCE ESPACIAL AVANZADO)
    # ------------------------------------------
    with tab4:
        st.header("🕵️‍♂️ Búsqueda de Victimarios (Rolling Window Avanzado)")
        st.markdown("Identifica qué líneas ajenas clonaron la ruta de la víctima. Integra **búsqueda espacial (100 metros)** para unir operadores distintos.")
        
        c1, c2, c3, c4 = st.columns(4)
        vic_num = c1.text_input("Número Víctima (10 dígitos):", key="t4_vic")
        tol = c2.selectbox("Tolerancia Temporal:", ['15min', '30min', '1H', '2H'], index=1)
        d_i_4 = c3.date_input("Fecha Inicio", f_min, min_value=MIN_DATE, max_value=MAX_DATE, key="t4_d1")
        t_i_4 = c4.time_input("Hora Inicio", datetime.time(0, 0), key="t4_t1")
        
        c5, c6, c7, c8 = st.columns(4)
        f_col = c5.checkbox("Excluir Plataformas (Inicio 3)", value=True, key="t4_chk")
        d_f_4 = c7.date_input("Fecha Fin", f_max, min_value=MIN_DATE, max_value=MAX_DATE, key="t4_d2")
        t_f_4 = c8.time_input("Hora Fin", datetime.time(23, 59), key="t4_t2")

        st.button("🔍 Rastrear Cruces", on_click=activar_t4, type="primary")

        if st.session_state.run_t4:
            if len(vic_num) >= 7:
                with st.spinner("Correlacionando mallas de ubicación..."):
                    df_l = df_master.dropna(subset=['numero_limpio', 'fecha_limpia']).copy()
                    dt_i, dt_f = pd.to_datetime(f"{d_i_4} {t_i_4}"), pd.to_datetime(f"{d_f_4} {t_f_4}")
                    df_l = df_l[(df_l['fecha_limpia'] >= dt_i) & (df_l['fecha_limpia'] <= dt_f)]
                    
                    df_l['lat_r'] = df_l['latitud'].round(3)
                    df_l['lon_r'] = df_l['longitud'].round(3)
                    
                    df_v = df_l[df_l['numero_limpio'] == vic_num].copy()
                    df_o = df_l[df_l['numero_limpio'] != vic_num].copy()
                    if f_col: df_o = df_o[df_o['numero_limpio'].str.match(r'^3\d{9}$')]

                    if not df_v.empty:
                        res = []
                        td = pd.Timedelta(tol.lower())
                        for _, r in df_v.iterrows():
                            if r['celda'] != 'Desconocida':
                                match = df_o[(df_o['celda'] == r['celda']) & (df_o['fecha_limpia'] >= r['fecha_limpia']-td) & (df_o['fecha_limpia'] <= r['fecha_limpia']+td)]
                            else:
                                if pd.notna(r['lat_r']):
                                    match = df_o[(df_o['lat_r'] == r['lat_r']) & (df_o['lon_r'] == r['lon_r']) & (df_o['fecha_limpia'] >= r['fecha_limpia']-td) & (df_o['fecha_limpia'] <= r['fecha_limpia']+td)]
                                else: match = pd.DataFrame()
                                
                            for _, sr in match.iterrows():
                                res.append({'numero_limpio': sr['numero_limpio'], 'celda': r['celda'] if r['celda']!='Desconocida' else 'Radio_100m'})
                        
                        if res:
                            sosp = pd.DataFrame(res).groupby('numero_limpio').agg(ubicaciones_distintas=('celda','nunique'), impactos_totales=('celda','count')).reset_index().sort_values(['ubicaciones_distintas','impactos_totales'], ascending=False).head(10)
                            top_sospechoso = sosp.iloc[0]["numero_limpio"]
                            
                            st.markdown(f'<div class="caja-alerta">💡 <b>Sospechosos Identificados:</b> El algoritmo señala a <b>{top_sospechoso}</b> como principal victimario, interceptando a la víctima en <b>{sosp.iloc[0]["ubicaciones_distintas"]}</b> ubicaciones distintas dentro del margen de {tol}.</div>', unsafe_allow_html=True)
                            st.dataframe(sosp, use_container_width=True, hide_index=True)
                            
                            # MAPA DE RUTAS VÍCTIMA VS VICTIMARIO
                            df_v_map = df_v.dropna(subset=['latitud', 'longitud']).sort_values('fecha_limpia')
                            df_s_map = df_o[df_o['numero_limpio'] == top_sospechoso].dropna(subset=['latitud', 'longitud']).sort_values('fecha_limpia')
                            
                            if not df_v_map.empty:
                                m4 = folium.Map(location=[df_v_map.iloc[0]['latitud'], df_v_map.iloc[0]['longitud']], zoom_start=13)
                                AntPath(locations=df_v_map[['latitud', 'longitud']].values.tolist(), color='green', weight=5, tooltip=f"Ruta Víctima: {vic_num}").add_to(m4)
                                if not df_s_map.empty:
                                    AntPath(locations=df_s_map[['latitud', 'longitud']].values.tolist(), color='red', weight=5, tooltip=f"Ruta Victimario: {top_sospechoso}").add_to(m4)
                                st_folium(m4, use_container_width=True, height=500, returned_objects=[])
                        else:
                            st.success("Cruce Negativo: Ninguna línea siguió a la víctima en ese rango espacial-temporal.")
                    else:
                        st.warning("La víctima no registra actividad bajo estas fechas y horas exactas.")

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
        d_i_5 = cf2.date_input("Fecha Inicio", f_min, min_value=MIN_DATE, max_value=MAX_DATE, key="t5_d1")
        t_i_5 = cf3.time_input("Hora Inicio", datetime.time(0, 0), key="t5_t1")
        d_f_5 = cf4.date_input("Fecha Fin", f_max, min_value=MIN_DATE, max_value=MAX_DATE, key="t5_d2")
        t_f_5 = st.time_input("Hora Fin", datetime.time(23, 59), key="t5_t2")

        st.button("Generar Perfiles y Ficha Técnica", on_click=activar_t5, type="primary")
        
        if st.session_state.run_t5:
            list_obj = [o.strip() for o in [o1, o2, o3, o4, o5] if len(o.strip()) >= 7]
            if list_obj:
                df_p = df_master.dropna(subset=['numero_limpio', 'fecha_limpia']).copy()
                if f_col_5: df_p = df_p[df_p['numero_limpio'].str.match(r'^3\d{9}$')]
                
                dt_i, dt_f = pd.to_datetime(f"{d_i_5} {t_i_5}"), pd.to_datetime(f"{d_f_5} {t_f_5}")
                df_p = df_p[(df_p['fecha_limpia'] >= dt_i) & (df_p['fecha_limpia'] <= dt_f)]
                
                df_obj = df_p[df_p['numero_limpio'].isin(list_obj)].copy()
                
                if not df_obj.empty:
                    # FICHA TÉCNICA
                    st.subheader("📡 Ficha Técnica Forense de Antenas (Filtrada)")
                    ficha_tecnica = df_obj[['numero_limpio', 'fecha_limpia', 'nombre_antena', 'celda', 'lac', 'latitud', 'longitud', 'haz', 'azimut', 'operador']].copy()
                    st.dataframe(ficha_tecnica.sort_values('fecha_limpia'), use_container_width=True, hide_index=True)
                    
                    # PATRÓN DE VIDA EXACTO
                    def cl_h(h):
                        if 0 <= h < 6: return '🌙 MADRUGADA (00:00 - 05:59)'
                        elif 6 <= h < 12: return '🌅 MAÑANA (06:00 - 11:59)'
                        elif 12 <= h < 18: return '☀️ TARDE (12:00 - 17:59)'
                        else: return '🌃 NOCHE (18:00 - 23:59)'
                        
                    df_obj['franja'] = df_obj['fecha_limpia'].dt.hour.apply(cl_h)
                    rut = df_obj.groupby(['numero_limpio', 'franja', 'celda', 'nombre_antena']).size().reset_index(name='impactos').sort_values(['numero_limpio', 'franja','impactos'], ascending=[True, True, False])
                    
                    st.subheader("⏱️ Patrones Horarios de Ubicación")
                    
                    txt_sinop = "<b>Análisis de Rutinas (Filtrado por fechas/horas):</b><br>"
                    for obj in list_obj:
                        df_o_r = rut[rut['numero_limpio'] == obj]
                        if not df_o_r.empty:
                            noche = df_o_r[df_o_r['franja'].str.contains('NOCHE')]
                            n_txt = noche.iloc[0]['celda'] if not noche.empty else "N/A"
                            dia = df_o_r[df_o_r['franja'].str.contains('DÍA')]
                            d_txt = dia.iloc[0]['celda'] if not dia.empty else "N/A"
                            txt_sinop += f"• El objetivo <b>{obj}</b> pernocta principalmente en la celda {n_txt} y opera de día en la celda {d_txt}.<br>"
                    
                    st.markdown(f'<div class="caja-info">{txt_sinop}</div>', unsafe_allow_html=True)
                    st.dataframe(rut, use_container_width=True, hide_index=True)

                    df_kml = df_obj.dropna(subset=['latitud', 'longitud'])
                    if not df_kml.empty:
                        kml = simplekml.Kml()
                        kml_colors = [simplekml.Color.red, simplekml.Color.blue, simplekml.Color.green, simplekml.Color.purple, simplekml.Color.orange]
                        for idx, obj in enumerate(list_obj):
                            df_kml_obj = df_kml[df_kml['numero_limpio'] == obj]
                            if not df_kml_obj.empty:
                                fol = kml.newfolder(name=f"Objetivo_{obj}")
                                estilo = simplekml.Style()
                                estilo.iconstyle.color = kml_colors[idx % len(kml_colors)]
                                estilo.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
                                
                                for _, r in df_kml_obj.iterrows():
                                    pnt = fol.newpoint(name=str(r['celda']), coords=[(r['longitud'], r['latitud'])])
                                    pnt.style = estilo
                                    pnt.description = f"<b>Número:</b> {r['numero_limpio']}<br><b>Fecha:</b> {r['fecha_limpia']}<br><b>Franja:</b> {r['franja']}<br><b>Antena:</b> {r['nombre_antena']}<br><b>LAC:</b> {r['lac']}<br><b>HAZ:</b> {r['haz']}<br><b>AZIMUT:</b> {r['azimut']}<br><b>OPERADOR:</b> {r['operador']}"
                        
                        st.download_button("🌍 Descargar Ficha Técnica (Google Earth KML)", data=kml.kml(), file_name="Perfiles_Tecnicos.kml", mime="application/vnd.google-earth.kml+xml")
                else:
                    st.error("Sin registros validos bajo esos filtros de fecha, hora y números.")
            else:
                st.warning("Ingresa al menos un número objetivo.")

    # ------------------------------------------
    # TAB 6: ANTENAS 3D (G-NET TILT)
    # ------------------------------------------
    with tab6:
        st.header("📡 G-NetTilt 3D (Simulador de Antenas)")
        st.markdown("Genera un archivo KML con los **tres planos de propagación 3D** (Central, Superior e Inferior) de una antena, importable directamente a Google Earth para peritaje de cobertura espacial.")
        
        celdas_validas = df_master[(df_master['celda'] != 'Desconocida') & (df_master['latitud'].notna())]['celda'].unique()
        celda_auto = st.selectbox("Autocompletar datos desde Celda en Evidencia (Opcional):", ["Ninguna"] + list(celdas_validas))
        
        def_lat, def_lon, def_az, def_haz = 0.0, 0.0, 0.0, 65.0
        if celda_auto != "Ninguna":
            dat_c = df_master[df_master['celda'] == celda_auto].iloc[0]
            def_lat = float(dat_c['latitud']) if pd.notna(dat_c['latitud']) else 0.0
            def_lon = float(dat_c['longitud']) if pd.notna(dat_c['longitud']) else 0.0
            try: def_az = float(dat_c['azimut'])
            except: def_az = 0.0
            try: def_haz = float(dat_c['haz'])
            except: def_haz = 65.0

        with st.form("form_gnet"):
            c1, c2, c3 = st.columns(3)
            c_name = c1.text_input("CELLNAME:", value=celda_auto if celda_auto != "Ninguna" else "Antena_Objetivo")
            c_lat = c2.number_input("LATITUDE:", value=def_lat, format="%.6f")
            c_lon = c3.number_input("LONGITUDE:", value=def_lon, format="%.6f")
            
            c4, c5, c6 = st.columns(3)
            c_height = c4.number_input("HEIGHT (m):", value=30.0, help="Altura de la torre")
            c_azimuth = c5.number_input("AZIMUTH (°):", value=def_az, help="Orientación del haz")
            c_tilt = c6.number_input("TILT (°):", value=0.0, help="Inclinación mecánica/eléctrica")
            
            c7, c8, c9 = st.columns(3)
            c_hbeam = c7.number_input("HORIZONTAL BEAMWIDTH (°):", value=def_haz)
            c_vbeam = c8.number_input("VERTICAL BEAMWIDTH (°):", value=10.0)
            c_dist = c9.number_input("COVERAGE DISTANCE (m):", value=1000.0, help="Radio de cobertura simulado")
            
            btn_kml_3d = st.form_submit_button("⚙️ Generar KML 3D")
            
        if btn_kml_3d:
            def calc_dest(lat, lon, bearing, dist_m):
                R = 6378137.0
                lat_r, lon_r, br_r = math.radians(lat), math.radians(lon), math.radians(bearing)
                l2 = math.asin(math.sin(lat_r)*math.cos(dist_m/R) + math.cos(lat_r)*math.sin(dist_m/R)*math.cos(br_r))
                lo2 = lon_r + math.atan2(math.sin(br_r)*math.sin(dist_m/R)*math.cos(lat_r), math.cos(dist_m/R)-math.sin(lat_r)*math.sin(l2))
                return math.degrees(l2), math.degrees(lo2)

            kml = simplekml.Kml()
            fol = kml.newfolder(name=c_name)
            
            angles = {
                "Central (Max Power)": c_tilt,
                "Upper (-3dB)": c_tilt - (c_vbeam / 2),
                "Lower (-3dB)": c_tilt + (c_vbeam / 2)
            }
            
            az_l = (c_azimuth - c_hbeam / 2) % 360
            az_r = (c_azimuth + c_hbeam / 2) % 360
            
            for p_name, a_deg in angles.items():
                alt_drop = c_dist * math.tan(math.radians(a_deg))
                alt_e = max(0, c_height - alt_drop)
                
                points = [(c_lon, c_lat, c_height)]
                steps = 10
                for i in range(steps + 1):
                    az_step = az_l + (az_r - az_l) * (i / steps)
                    lat_s, lon_s = calc_dest(c_lat, c_lon, az_step, c_dist)
                    points.append((lon_s, lat_s, alt_e))
                points.append((c_lon, c_lat, c_height))
                
                pol = fol.newpolygon(name=p_name)
                pol.outerboundaryis = points
                pol.altitudemode = simplekml.AltitudeMode.relativetoground
                pol.extrude = 0
                
                # COLOR: Amarillo al 40% transparente (KML AABBGGRR)
                pol.style.polystyle.color = simplekml.Color.hexa('6600FFFF') 
                pol.style.polystyle.fill = 1
                pol.style.polystyle.outline = 1
                pol.style.linestyle.color = simplekml.Color.black
                pol.style.linestyle.width = 2
                
            st.success("✅ Modelo 3D de radiación calculado exitosamente.")
            st.download_button("🌍 Descargar Planos de Radiación (KML)", data=kml.kml(), file_name=f"GNetTilt_{c_name}.kml", mime="application/vnd.google-earth.kml+xml")

else:
    st.info("👈 Por favor, carga las bases de datos CDR en el menú lateral para iniciar el motor de Inteligencia Artificial.")