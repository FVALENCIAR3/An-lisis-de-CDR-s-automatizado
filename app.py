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

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Inteligencia Geoespacial & SIGINT", page_icon="📡", layout="wide")

# Inicializar Base de Datos en Memoria (Session State)
if 'df_master' not in st.session_state:
    st.session_state.df_master = None

# ==========================================
# FUNCIONES DE LIMPIEZA Y TRANSFORMACIÓN
# ==========================================
@st.cache_data(show_spinner=False)
def procesar_archivos(uploaded_files):
    df_list = []
    
    def limpiar_fechas(s):
        if pd.api.types.is_datetime64_any_dtype(s): return s
        s_str = s.astype(str).str.strip()
        fechas = pd.Series(pd.NaT, index=s.index)
        mask_14 = s_str.str.match(r'^\d{14}$')
        fechas.loc[mask_14] = pd.to_datetime(s_str[mask_14], format='%Y%m%d%H%M%S', errors='coerce')
        fechas.loc[~mask_14] = pd.to_datetime(s_str[~mask_14], format='mixed', dayfirst=True, errors='coerce')
        return fechas

    def normalizar_celda(val):
        val = str(val).strip().upper()
        if val.endswith('.0'): val = val[:-2]
        if re.match(r'^[0-9A-F]{4,6}-[0-9A-F]{2,4}$', val):
            try: return '-'.join([str(int(p, 16)) for p in val.split('-')])
            except: return val
        return val

    for file in uploaded_files:
        try:
            file_ext = file.name.lower()
            
            # 1. Leer CSV o TXT (Bases de datos en texto plano)
            if file_ext.endswith(('.csv', '.txt')):
                try: 
                    df = pd.read_csv(file, sep=';', encoding='utf-8-sig', low_memory=False)
                    if len(df.columns) <= 1:
                        df = pd.read_csv(file, sep=',', encoding='utf-8-sig', low_memory=False)
                except: 
                    # Si falla, intentamos que Pandas detecte el separador solo
                    file.seek(0)
                    df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8-sig')
            
            # 2. Leer Excel Binario
            elif file_ext.endswith('.xlsb'):
                df = pd.read_excel(file, engine='pyxlsb')
            
            # 3. Leer Excel Tradicional
            elif file_ext.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(file)
                
            # 4. Evitar que archivos no compatibles (PDF, DOCX) rompan el programa
            else:
                st.sidebar.warning(f"Ignorado '{file.name}': Este espacio es solo para bases de datos. Los PDF y Word súbelos en la pestaña 'Informe PNL'.")
                continue
                
            # Limpieza de cabeceras
            df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
            
            alias_tel = ['numero', 'numero_origen', 'numero_que_marca', 'numero_que_navega', 'originador', 'numero_a']
            alias_fec = ['fecha_hora_inicio', 'fecha_trafico', 'fecha_hora_inicio_llamada', 'fecha_hora_inicio_sesion', 'fecha_hora', 'fecha_y_hora_origen']
            alias_lat = ['latitud', 'latitud_n', 'latitude']
            alias_lon = ['longitud', 'longitud_w', 'longitude']
            alias_cel = ['celda_decimal', 'cell_id_voz', 'celda', 'celda_inicio_llamada', 'bts_id', 'celda_hex', 'celda_origen_truncada', 'cellid_nval']
            
            col_tel = next((c for c in alias_tel if c in df.columns), None)
            col_fec = next((c for c in alias_fec if c in df.columns), None)
            col_lat = next((c for c in alias_lat if c in df.columns), None)
            col_lon = next((c for c in alias_lon if c in df.columns), None)
            col_cel = next((c for c in alias_cel if c in df.columns), None)
            
            df_temp = pd.DataFrame()
            if col_tel: df_temp['numero_limpio'] = df[col_tel].astype(str).str.replace(r'\D', '', regex=True).str[-10:]
            if col_fec: df_temp['fecha_limpia'] = limpiar_fechas(df[col_fec])
            if col_lat and col_lon:
                df_temp['latitud'] = pd.to_numeric(df[col_lat].astype(str).str.replace(',', '.'), errors='coerce')
                df_temp['longitud'] = pd.to_numeric(df[col_lon].astype(str).str.replace(',', '.'), errors='coerce')
            if col_cel: df_temp['celda'] = df[col_cel].apply(normalizar_celda)
                
            df_temp['registro_original'] = df.to_dict('records')
            df_list.append(df_temp)
            
        except Exception as e:
            st.sidebar.error(f"Error procesando {file.name}: {e}")

    if df_list:
        return pd.concat(df_list, ignore_index=True)
    return None

# ==========================================
# BARRA LATERAL (CARGA DE DATOS)
# ==========================================
st.sidebar.title("📡 SIGINT App")
st.sidebar.markdown("---")

# Restricción de tipos de archivo en la barra lateral para evitar que suban PDFs aquí
uploaded_files = st.sidebar.file_uploader("1. Cargar Bases de Datos", accept_multiple_files=True, type=['csv', 'txt', 'xls', 'xlsx', 'xlsb'])

if st.sidebar.button("Procesar Archivos"):
    if uploaded_files:
        with st.spinner("Integrando y estandarizando datos..."):
            st.session_state.df_master = procesar_archivos(uploaded_files)
        if st.session_state.df_master is not None:
            st.sidebar.success(f"✅ {len(st.session_state.df_master)} registros cargados.")
    else:
        st.sidebar.warning("Sube al menos un archivo válido.")

# ==========================================
# CUERPO PRINCIPAL (PESTAÑAS)
# ==========================================
if st.session_state.df_master is not None:
    df_master = st.session_state.df_master
    try:
        f_min, f_max = df_master['fecha_limpia'].min().date(), df_master['fecha_limpia'].max().date()
    except:
        f_min = f_max = datetime.date.today()

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📍 Individual", "⚙️ Inteligencia de Red", "📝 Informe PNL", "🕵️‍♂️ Co-Desplazamiento", "👤 Patrón de Vida"
    ])

    # ------------------------------------------
    # TAB 1: BÚSQUEDA INDIVIDUAL
    # ------------------------------------------
    with tab1:
        st.header("📍 Búsqueda Geográfica Individual")
        col1, col2, col3 = st.columns(3)
        num_buscar = col1.text_input("Número (10 dígitos):", placeholder="3157658841")
        top_mostrar = col2.selectbox("Mostrar Top:", ["Todos", 10, 5, 3, 1])
        
        c1, c2, c3, c4 = st.columns(4)
        d_ini = c1.date_input("Fecha Inicio", f_min, key="t1_d1")
        t_ini = c2.time_input("Hora Inicio", datetime.time(0, 0), key="t1_t1")
        d_fin = c3.date_input("Fecha Fin", f_max, key="t1_d2")
        t_fin = c4.time_input("Hora Fin", datetime.time(23, 59), key="t1_t2")

        if st.button("Trazar Mapa", key="btn_tab1"):
            if len(num_buscar) == 10:
                df_u = df_master[(df_master['numero_limpio'] == num_buscar) & (df_master['latitud'].notna())].copy()
                dt_ini, dt_fin = pd.to_datetime(f"{d_ini} {t_ini}"), pd.to_datetime(f"{d_fin} {t_fin}")
                df_u = df_u[(df_u['fecha_limpia'] >= dt_ini) & (df_u['fecha_limpia'] <= dt_fin)].sort_values('fecha_limpia')

                if not df_u.empty:
                    df_agrup = df_u.groupby(['latitud', 'longitud']).agg(
                        visitas=('numero_limpio', 'count'), prim_vis=('fecha_limpia', 'min'), ult_vis=('fecha_limpia', 'max'),
                        celdas=('celda', lambda x: ', '.join(x.dropna().unique().astype(str)))
                    ).reset_index().sort_values('visitas', ascending=False)

                    st.info(f"💡 **Sinopsis:** El objetivo **{num_buscar}** registró **{len(df_u)}** conexiones. Su mayor zona de confort fue la celda **{df_agrup.iloc[0]['celdas']}** con **{df_agrup.iloc[0]['visitas']}** impactos.")

                    df_plot = df_agrup if top_mostrar == "Todos" else df_agrup.head(top_mostrar)
                    m = folium.Map(location=[df_plot.iloc[0]['latitud'], df_plot.iloc[0]['longitud']], zoom_start=13)

                    for i, row in df_plot.iterrows():
                        color = 'red' if i==0 else 'orange' if i<=2 else 'blue'
                        folium.CircleMarker([row['latitud'], row['longitud']], radius=min(20, 8+row['visitas']), color=color, fill=True).add_to(m)
                        folium.Marker([row['latitud'], row['longitud']], tooltip=f"Rank #{i+1} | Visitas: {row['visitas']}", icon=folium.Icon(color=color)).add_to(m)

                    if top_mostrar == "Todos" and len(df_u) > 1:
                        AntPath(locations=df_u[['latitud', 'longitud']].values.tolist(), color='purple').add_to(m)
                    
                    st_folium(m, use_container_width=True, height=500)
                else:
                    st.warning("No hay registros con coordenadas en este periodo.")
            else:
                st.error("Ingresa un número de 10 dígitos.")

    # ------------------------------------------
    # TAB 2: INTELIGENCIA DE RED
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

        if st.button("Generar Análisis General", key="btn_tab2"):
            with st.spinner("Cruzando registros..."):
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

                    st.info(f"💡 **Sinopsis:** El tráfico más alto se ancló a la celda **{top_c.iloc[0]['celda'] if not top_c.empty else 'N/A'}**. " + 
                            (f"Además, **{top_e.iloc[0]['Número A']}** y **{top_e.iloc[0]['Número B']}** coincidieron **{top_e.iloc[0]['Match']}** veces." if not top_e.empty else "No se detectaron encuentros simultáneos."))

                    colA, colB = st.columns(2)
                    with colA:
                        st.subheader("🏆 Top Celdas Frecuentes")
                        st.dataframe(top_c, use_container_width=True)
                    with colB:
                        st.subheader("📍 Encuentros Probables (15 min)")
                        st.dataframe(top_e, use_container_width=True)

    # ------------------------------------------
    # TAB 3: INFORME NLP
    # ------------------------------------------
    with tab3:
        st.header("📝 Informe PNL (Procesamiento Natural de Documentos)")
        st.info("Sube aquí tus documentos PDF o DOCX de fiscalía, policía, reportes, etc. El sistema extraerá los números de celular y los cruzará con la base de datos.")
        texto_in = st.text_area("O pega el texto del caso aquí:", height=150)
        docs_in = st.file_uploader("Sube PDF, DOCX, TXT:", accept_multiple_files=True, type=['pdf', 'docx', 'doc', 'txt'])

        if st.button("Generar Informe PNL", key="btn_tab3"):
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
                    except Exception as e:
                        st.error(f"No se pudo leer el archivo {f.name}: {e}")

            nums_implicados = list(set(re.findall(r'\b\d{10}\b', txt_full)))
            
            if nums_implicados:
                df_caso = df_master[df_master['numero_limpio'].isin(nums_implicados)].copy()
                if not df_caso.empty:
                    df_caso = df_caso.dropna(subset=['fecha_limpia']).sort_values('fecha_limpia')
                    st.success(f"Números encontrados y cruzados: {', '.join(nums_implicados)}")
                    
                    st.markdown(f"### 📋 Sinopsis de Tiempo, Modo y Lugar")
                    st.markdown(f"> **Tiempo:** Desde **{df_caso['fecha_limpia'].min()}** hasta **{df_caso['fecha_limpia'].max()}** con **{len(df_caso)}** impactos.\n"
                                f"> **Lugar:** Celdas predominantes: **{', '.join(df_caso['celda'].value_counts().head(3).index.astype(str))}**.\n"
                                f"> **Modo:** **{len(df_caso['numero_limpio'].unique())}** líneas activas del texto aportado.")

                    df_caso['dia'] = df_caso['fecha_limpia'].dt.date
                    conteo = df_caso.groupby(['dia', 'numero_limpio']).size().reset_index(name='eventos')
                    fig, ax = plt.subplots(figsize=(10, 3))
                    sns.lineplot(data=conteo, x='dia', y='eventos', hue='numero_limpio', marker='o', ax=ax)
                    st.pyplot(fig)
                else:
                    st.warning("Los números extraídos del texto/PDF no existen en la base de datos de llamadas.")
            else:
                st.warning("No se hallaron números de 10 dígitos en el texto o documentos.")

    # ------------------------------------------
    # TAB 4: CO-DESPLAZAMIENTO
    # ------------------------------------------
    with tab4:
        st.header("🕵️‍♂️ Búsqueda de Victimarios (Rolling Window)")
        col1, col2, col3 = st.columns(3)
        vic_num = col1.text_input("Número Víctima:")
        tol = col2.selectbox("Tolerancia:", ['15min', '30min', '1H', '2H'], index=1)
        f_col = col3.checkbox("Ignorar Plataformas", value=True, key="t4_chk")
        
        c1, c2, c3, c4 = st.columns(4)
        d_i_4 = c1.date_input("Fecha Inicio", f_min, key="t4_d1")
        t_i_4 = c2.time_input("Hora Inicio", datetime.time(0, 0), key="t4_t1")
        d_f_4 = c3.date_input("Fecha Fin", f_max, key="t4_d2")
        t_f_4 = c4.time_input("Hora Fin", datetime.time(23, 59), key="t4_t2")

        if st.button("Rastrear Ruta", key="btn_tab4"):
            if len(vic_num) == 10:
                with st.spinner("Escaneando intersecciones exactas..."):
                    df_l = df_master.dropna(subset=['numero_limpio', 'celda', 'fecha_limpia']).copy()
                    dt_i, dt_f = pd.to_datetime(f"{d_i_4} {t_i_4}"), pd.to_datetime(f"{d_f_4} {t_f_4}")
                    df_l = df_l[(df_l['fecha_limpia'] >= dt_i) & (df_l['fecha_limpia'] <= dt_f)]
                    
                    df_v = df_l[df_l['numero_limpio'] == vic_num].copy()
                    df_o = df_l[df_l['numero_limpio'] != vic_num].copy()
                    if f_col: df_o = df_o[df_o['numero_limpio'].str.match(r'^3\d{9}$')]

                    if not df_v.empty:
                        res = []
                        td = pd.Timedelta(tol)
                        for _, r in df_v.iterrows():
                            match = df_o[(df_o['celda'] == r['celda']) & (df_o['fecha_limpia'] >= r['fecha_limpia']-td) & (df_o['fecha_limpia'] <= r['fecha_limpia']+td)]
                            for _, sr in match.iterrows():
                                res.append({'numero_limpio': sr['numero_limpio'], 'celda': r['celda'], 'registro_original': sr['registro_original']})
                        
                        if res:
                            inter = pd.DataFrame(res)
                            sosp = inter.groupby('numero_limpio').agg(celdas=('celda','nunique'), imp=('celda','count')).reset_index().sort_values(['celdas','imp'], ascending=False).head(10)
                            
                            st.info(f"💡 **Sinopsis:** Se perfilaron posibles victimarios. El principal sospechoso es **{sosp.iloc[0]['numero_limpio']}**, compartiendo **{sosp.iloc[0]['celdas']}** celdas diferentes dentro del margen de {tol}.")
                            st.dataframe(sosp, use_container_width=True)
                        else:
                            st.success("Nadie siguió la ruta de la víctima.")
                    else:
                        st.error("La víctima no tiene registros en el rango.")

    # ------------------------------------------
    # TAB 5: PATRÓN DE VIDA Y KML
    # ------------------------------------------
    with tab5:
        st.header("👤 Perfilamiento de Rutinas y Google Earth (KML)")
        c1, c2 = st.columns(2)
        obj_num = c1.text_input("Número Victimario:")
        f_col_5 = c2.checkbox("Omitir Plataformas", value=True, key="t5_chk")

        if st.button("Perfilar Rutina", key="btn_tab5"):
            if len(obj_num) == 10:
                df_p = df_master.dropna(subset=['numero_limpio', 'celda', 'fecha_limpia']).copy()
                if f_col_5: df_p = df_p[df_p['numero_limpio'].str.match(r'^3\d{9}$')]
                df_obj = df_p[df_p['numero_limpio'] == obj_num].copy()
                
                if not df_obj.empty:
                    def cl_h(h):
                        if h >= 22 or h <= 2: return '🌙 NOCHE (Pernocta)'
                        elif 3 <= h <= 8: return '🌅 MADRUGADA (Amanecer)'
                        else: return '☀️ DÍA (Operación)'
                        
                    df_obj['franja'] = df_obj['fecha_limpia'].dt.hour.apply(cl_h)
                    rut = df_obj.groupby(['franja', 'celda']).size().reset_index(name='v').sort_values(['franja','v'], ascending=[True, False])
                    top_rut = rut.groupby('franja').head(3)
                    
                    st.dataframe(top_rut, use_container_width=True)

                    df_kml = df_obj.dropna(subset=['latitud', 'longitud'])
                    if not df_kml.empty:
                        kml = simplekml.Kml()
                        for _, r in df_kml.iterrows():
                            pnt = kml.newpoint(name=str(r['celda']), coords=[(r['longitud'], r['latitud'])])
                            pnt.description = f"{r['fecha_limpia']} | {r['franja']}"
                        
                        st.download_button("🌍 Descargar KML (Google Earth)", data=kml.kml(), file_name=f"Rutina_{obj_num}.kml", mime="application/vnd.google-earth.kml+xml")
                else:
                    st.error("Sin registros.")

else:
    st.info("👈 Por favor, carga los archivos de Excel/CSV/TXT en el menú lateral para iniciar.")