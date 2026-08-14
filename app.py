import io
import re
import gc
import itertools
import zipfile
from collections import Counter

import pandas as pd
import streamlit as st
import folium
from folium.plugins import AntPath
from streamlit_folium import st_folium
import matplotlib.pyplot as plt
import seaborn as sns
import PyPDF2
import docx

try:
    import simplekml
except ImportError:
    simplekml = None


# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Análisis de CDR´s  Automatizado",
    page_icon="🕵️‍♂️",
    layout="wide",
)

# ============================================================
# CONTROL DE ACCESO
# ============================================================
# Las credenciales deben configurarse en Streamlit Cloud > Settings
# > Secrets. NO las escribas directamente en este archivo.
#
# Ejemplo de secrets:
# [auth]
# username = "admin"
# password = "Valencia2026."

def verificar_acceso():
    if st.session_state.get("autenticado", False):
        return True

    st.title("🔐 Acceso restringido")
    st.caption("Análisis de CDRs Automatizado")

    try:
        credenciales = st.secrets["auth"]
        usuario_correcto = str(credenciales["username"])
        password_correcta = str(credenciales["password"])
    except Exception:
        st.error(
            "⚠️ La aplicación no tiene configuradas las credenciales. "
            "Configúralas en Streamlit Cloud → Settings → Secrets."
        )
        st.code(
            '[auth]\nusername = "admin"\npassword = "CAMBIA_ESTA_CONTRASENA"',
            language="toml",
        )
        st.stop()

    with st.form("login_form", clear_on_submit=False):
        usuario = st.text_input("👤 Usuario", autocomplete="username")
        password = st.text_input(
            "🔑 Contraseña",
            type="password",
            autocomplete="current-password",
        )
        ingresar = st.form_submit_button(
            "🚪 Ingresar",
            type="primary",
            use_container_width=True,
        )

    if ingresar:
        # Comparación directa de valores provenientes de st.secrets.
        if usuario == usuario_correcto and password == password_correcta:
            st.session_state["autenticado"] = True
            st.session_state["usuario_autenticado"] = usuario
            st.rerun()
        else:
            st.error("❌ Usuario o contraseña incorrectos.")

    return False


if not verificar_acceso():
    st.stop()


# ============================================================
# SESIÓN AUTENTICADA
# ============================================================
with st.sidebar:
    st.success(
        f"🔓 Sesión activa: "
        f"**{st.session_state.get('usuario_autenticado', 'usuario')}**"
    )

    if st.button("🔒 Cerrar sesión", use_container_width=True):
        for key in [
            "autenticado",
            "usuario_autenticado",
            "df_master",
            "loaded_names",
            "load_errors",
            "cdr_upload",
        ]:
            st.session_state.pop(key, None)
        st.rerun()


# ============================================================
# ESTADO PERSISTENTE PARA EVITAR QUE LOS RESULTADOS DESAPAREZCAN
# ============================================================
if "df_master" not in st.session_state:
    st.session_state.df_master = None
if "errores" not in st.session_state:
    st.session_state.errores = []
if "payload_cdr" not in st.session_state:
    st.session_state.payload_cdr = None

for _key in [
    "geo_ejecutado", "red_ejecutado", "informe_ejecutado",
    "co_ejecutado", "perfil_ejecutado"
]:
    if _key not in st.session_state:
        st.session_state[_key] = False


st.title("🕵️‍♂️ Análisis de CDR´s  Automatizado")
st.caption(
    " Análisis de inteligencia artificial que permite limpiar, organizar y analizar las bases de datos en archivos consolidados! "
    " convierte celdas hexadecimales y excluye números de plataformas, también representa gráficamente en el mapa,  "
    " genera la sinopsis del comportamiento, zona de mayor permanencia, residencia o trabajo habitual del objetivo. "
)


# ============================================================
# FUNCIONES ORIGINALES DE LIMPIEZA
# ============================================================
@st.cache_data(show_spinner=False)
def limpiar_fechas(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s

    s_str = s.astype(str).str.strip()
    fechas = pd.Series(pd.NaT, index=s.index)

    mask_14 = s_str.str.match(r"^\d{14}$")
    fechas.loc[mask_14] = pd.to_datetime(
        s_str[mask_14],
        format="%Y%m%d%H%M%S",
        errors="coerce",
    )

    fechas.loc[~mask_14] = pd.to_datetime(
        s_str[~mask_14],
        format="mixed",
        dayfirst=True,
        errors="coerce",
    )

    return fechas


@st.cache_data(show_spinner=False)
def normalizar_celda(val):
    val = str(val).strip().upper()

    if val.endswith(".0"):
        val = val[:-2]

    if re.match(r"^[0-9A-F]{4,6}-[0-9A-F]{2,4}$", val):
        try:
            parts = val.split("-")
            dec_parts = [str(int(p, 16)) for p in parts]
            return "-".join(dec_parts)
        except Exception:
            return val

    return val


# ============================================================
# CARGA Y CONSOLIDACIÓN
# ============================================================
@st.cache_data(show_spinner=False)
def load_and_process_files(uploaded_files_data):
    df_list = []
    errores = []

    alias_tel = [
        "numero",
        "numero_origen",
        "numero_que_marca",
        "numero_que_navega",
        "originador",
        "numero_a",
    ]
    alias_fec = [
        "fecha_hora_inicio",
        "fecha_trafico",
        "fecha_hora_inicio_llamada",
        "fecha_hora_inicio_sesion",
        "fecha_hora",
        "fecha_y_hora_origen",
    ]
    alias_lat = ["latitud", "latitud_n", "latitude"]
    alias_lon = ["longitud", "longitud_w", "longitude"]
    alias_cel = [
        "celda_decimal",
        "cell_id_voz",
        "celda",
        "celda_inicio_llamada",
        "bts_id",
        "celda_hex",
        "celda_origen_truncada",
        "cellid_nval",
    ]

    for file_name, file_content in uploaded_files_data:
        try:
            name = file_name.lower()

            if name.endswith(".csv"):
                try:
                    df = pd.read_csv(
                        io.BytesIO(file_content),
                        sep=";",
                        encoding="utf-8-sig",
                        low_memory=False,
                    )
                    if len(df.columns) <= 1:
                        df = pd.read_csv(
                            io.BytesIO(file_content),
                            sep=",",
                            encoding="utf-8-sig",
                            low_memory=False,
                        )
                except Exception:
                    df = pd.read_csv(
                        io.BytesIO(file_content),
                        sep=None,
                        engine="python",
                        encoding="utf-8-sig",
                    )

            elif name.endswith(".xlsb"):
                df = pd.read_excel(
                    io.BytesIO(file_content),
                    engine="pyxlsb",
                )

            else:
                df = pd.read_excel(io.BytesIO(file_content))

            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.lower()
                .str.replace(" ", "_", regex=False)
            )

            col_tel = next((c for c in alias_tel if c in df.columns), None)
            col_fec = next((c for c in alias_fec if c in df.columns), None)
            col_lat = next((c for c in alias_lat if c in df.columns), None)
            col_lon = next((c for c in alias_lon if c in df.columns), None)
            col_cel = next((c for c in alias_cel if c in df.columns), None)

            df_temp = pd.DataFrame(index=df.index)

            if col_tel:
                df_temp["numero_limpio"] = (
                    df[col_tel]
                    .astype(str)
                    .str.replace(r"\D", "", regex=True)
                    .str[-10:]
                )

            if col_fec:
                df_temp["fecha_limpia"] = limpiar_fechas(df[col_fec])

            if col_lat and col_lon:
                df_temp["latitud"] = pd.to_numeric(
                    df[col_lat].astype(str).str.replace(",", ".", regex=False),
                    errors="coerce",
                )
                df_temp["longitud"] = pd.to_numeric(
                    df[col_lon].astype(str).str.replace(",", ".", regex=False),
                    errors="coerce",
                )

            if col_cel:
                df_temp["celda"] = df[col_cel].apply(normalizar_celda)

            df_temp["registro_original"] = df.to_dict("records")
            df_list.append(df_temp)

        except Exception as e:
            errores.append(f"{file_name}: {e}")

    if not df_list:
        return pd.DataFrame(), errores

    return pd.concat(df_list, ignore_index=True), errores


def create_date_range(df):
    if "fecha_limpia" not in df.columns:
        return None, None

    fechas = pd.to_datetime(df["fecha_limpia"], errors="coerce").dropna()
    if fechas.empty:
        return None, None

    return fechas.min().date(), fechas.max().date()


def parse_datetime_range(fecha_ini, hora_ini, fecha_fin, hora_fin):
    try:
        dt_ini = pd.to_datetime(f"{fecha_ini} {hora_ini}")
        dt_fin = pd.to_datetime(f"{fecha_fin} {hora_fin}")
        return dt_ini, dt_fin
    except Exception:
        return None, None


def registros_a_excel_bytes(df_datos):
    if df_datos.empty:
        return b""

    df_excel = pd.DataFrame(df_datos["registro_original"].tolist()).astype(str)
    for col in df_excel.columns:
        df_excel[col] = (
            df_excel[col]
            .str.replace(r"\.0$", "", regex=True)
            .replace(["nan", "NaT", "None"], "")
        )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_excel.to_excel(writer, index=False)
    return buffer.getvalue()


def mapa_html_bytes(mapa):
    return mapa.get_root().render().encode("utf-8")


def generar_exportables_bytes(df_datos, titulo_base):
    """Conserva la función de exportación original, adaptada a Streamlit."""
    archivos = {}

    archivos[f"{titulo_base}.xlsx"] = registros_a_excel_bytes(df_datos)

    plt.figure(figsize=(10, 4))
    df_plot = df_datos.copy()
    df_plot["fecha_hora"] = pd.to_datetime(df_plot["fecha_limpia"])
    sns.histplot(data=df_plot, x="fecha_hora", bins=20, kde=True)
    plt.title("Línea de Tiempo de Conexiones")
    plt.tight_layout()

    img = io.BytesIO()
    plt.savefig(img, format="png")
    plt.close()
    archivos[f"{titulo_base}.png"] = img.getvalue()

    df_mapa = df_datos.dropna(subset=["latitud", "longitud"])
    archivo_mapa = f"{titulo_base}_Mapa.html"

    if not df_mapa.empty:
        m = folium.Map(
            location=[df_mapa["latitud"].mean(), df_mapa["longitud"].mean()],
            zoom_start=14,
        )

        for _, r in df_mapa.iterrows():
            folium.CircleMarker(
                [r["latitud"], r["longitud"]],
                radius=5,
                fill=True,
                popup=str(r["fecha_limpia"]),
            ).add_to(m)

        archivos[archivo_mapa] = mapa_html_bytes(m)
    else:
        archivos[archivo_mapa] = b"<h3>Sin coordenadas</h3>"

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for nombre, contenido in archivos.items():
            zf.writestr(nombre, contenido)

    return zip_buffer.getvalue()


# ============================================================
# FUNCIÓN 1: BÚSQUEDA GEOGRÁFICA INDIVIDUAL
# ============================================================
def generar_mapa(
    df_master,
    numero_buscado,
    dropdown_top,
    fecha_ini,
    hora_ini,
    fecha_fin,
    hora_fin,
):
    """Versión Streamlit de la función generar_mapa original."""

    if len(numero_buscado) != 10:
        st.warning("⚠️ Ingresa un número de exactamente 10 dígitos.")
        return

    required = {"numero_limpio", "latitud", "longitud", "fecha_limpia"}
    faltantes = required - set(df_master.columns)
    if faltantes:
        st.error(f"Faltan columnas necesarias: {', '.join(sorted(faltantes))}")
        return

    df_usuario = df_master[
        (df_master["numero_limpio"] == numero_buscado)
        & df_master["latitud"].notna()
        & df_master["longitud"].notna()
    ].copy()

    df_usuario["latitud"] = df_usuario["latitud"].astype(float)
    df_usuario["longitud"] = df_usuario["longitud"].astype(float)

    dt_ini, dt_fin = parse_datetime_range(
        fecha_ini, hora_ini, fecha_fin, hora_fin
    )

    if dt_ini is not None:
        df_usuario = df_usuario[
            (df_usuario["fecha_limpia"] >= dt_ini)
            & (df_usuario["fecha_limpia"] <= dt_fin)
        ]

    if df_usuario.empty:
        st.error(
            f"❌ No hay historial de coordenadas para {numero_buscado} "
            "en este rango de tiempo."
        )
        return

    df_usuario = df_usuario.sort_values(by="fecha_limpia")

    df_agrupado = (
        df_usuario.groupby(["latitud", "longitud"])
        .agg(
            visitas=("numero_limpio", "count"),
            primera_visita=("fecha_limpia", "min"),
            ultima_visita=("fecha_limpia", "max"),
            celdas=(
                "celda",
                lambda x: ", ".join(
                    x.dropna().unique().astype(str)
                ),
            ),
        )
        .reset_index()
        .sort_values(by="visitas", ascending=False)
        .reset_index(drop=True)
    )

    tot_conexiones = len(df_usuario)
    top_lugar = df_agrupado.iloc[0]["celdas"]
    top_visitas = df_agrupado.iloc[0]["visitas"]

    st.markdown(
        f"""
        <div style="background-color:#e8f4f8;padding:15px;
        border-left:5px solid #5bc0de;border-radius:5px;">
        <h4>💡 Sinopsis del Comportamiento (Patrón de Vida)</h4>
        <p>El objetivo <b>{numero_buscado}</b> registra una actividad total de
        <b>{tot_conexiones}</b> conexiones geo-posicionadas dentro del periodo
        consultado. Su zona de mayor confort, residencia o trabajo habitual
        corresponde a la cobertura de la celda <b>{top_lugar}</b>, lugar donde
        acumuló el mayor número de impactos (<b>{top_visitas}</b> visitas
        registradas).</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    n = (
        len(df_agrupado)
        if dropdown_top == "Todos"
        else int(dropdown_top.replace("Top ", ""))
    )
    df_mostrar = df_agrupado.head(n)

    mapa = folium.Map(
        location=[
            df_mostrar.iloc[0]["latitud"],
            df_mostrar.iloc[0]["longitud"],
        ],
        zoom_start=14,
    )

    for i, row in df_mostrar.iterrows():
        coord = [row["latitud"], row["longitud"]]
        ranking = i + 1
        color = (
            "red"
            if ranking == 1
            else "orange"
            if ranking <= 3
            else "blue"
        )

        popup_html = (
            f"<div style='min-width:200px;'>"
            f"<b>Rank:</b> #{ranking}<br>"
            f"<b>Visitas:</b> {row['visitas']}<br>"
            f"<b>Celda:</b> {row['celdas']}<br>"
            f"<b>Inicio:</b> {row['primera_visita']}<br>"
            f"<b>Fin:</b> {row['ultima_visita']}</div>"
        )

        folium.CircleMarker(
            location=coord,
            radius=min(20, 8 + row["visitas"]),
            color=color,
            fill=True,
        ).add_to(mapa)

        folium.Marker(
            location=coord,
            popup=folium.Popup(popup_html, max_width=300),
            tooltip=f"Rank #{ranking}",
            icon=folium.Icon(color=color),
        ).add_to(mapa)

    if dropdown_top == "Todos" and len(df_usuario) > 1:
        AntPath(
            locations=df_usuario[
                ["latitud", "longitud"]
            ].values.tolist(),
            delay=1000,
            color="purple",
            weight=4,
        ).add_to(mapa)

    st_folium(mapa, width=None, height=600)
    st.dataframe(df_agrupado, use_container_width=True)


# ============================================================
# FUNCIÓN 2: ANÁLISIS DE RED / CO-UBICACIÓN
# ============================================================
def procesar_analisis(
    df_master,
    chk_colombia,
    fecha_ini,
    hora_ini,
    fecha_fin,
    hora_fin,
    txt_excluir,
):
    """Conserva el algoritmo original de análisis de encuentros."""

    required = {"numero_limpio", "celda", "fecha_limpia"}
    faltantes = required - set(df_master.columns)
    if faltantes:
        st.error(f"Faltan columnas necesarias: {', '.join(sorted(faltantes))}")
        return

    df_analisis = df_master.dropna(
        subset=["numero_limpio", "celda", "fecha_limpia"]
    ).copy()

    if chk_colombia:
        df_analisis = df_analisis[
            df_analisis["numero_limpio"].astype(str).str.match(
                r"^3\d{9}$"
            )
        ]

    dt_ini, dt_fin = parse_datetime_range(
        fecha_ini, hora_ini, fecha_fin, hora_fin
    )

    if dt_ini is not None:
        df_analisis = df_analisis[
            (df_analisis["fecha_limpia"] >= dt_ini)
            & (df_analisis["fecha_limpia"] <= dt_fin)
        ]

    nums_excluir = [
        x.strip()
        for x in txt_excluir.split(",")
        if x.strip()
    ]

    if nums_excluir:
        df_analisis = df_analisis[
            ~df_analisis["numero_limpio"].isin(nums_excluir)
        ]

    if df_analisis.empty:
        st.warning("⚠️ No hay datos para analizar con los filtros aplicados.")
        return

    df_top_celdas = (
        df_analisis.groupby(["celda", "numero_limpio"])
        .size()
        .reset_index(name="conexiones")
    )

    top_10_celdas = (
        df_top_celdas.sort_values(
            by="conexiones",
            ascending=False,
        )
        .head(10)
        .reset_index(drop=True)
    )

    df_analisis["ventana"] = df_analisis["fecha_limpia"].dt.floor("15min")
    agrupado_tiempo = df_analisis.groupby(
        ["celda", "ventana"]
    )["numero_limpio"].unique()

    conteo_tiempo = Counter()

    for numeros in agrupado_tiempo:
        if 1 < len(numeros) <= 150:
            conteo_tiempo.update(
                itertools.combinations(sorted(numeros), 2)
            )

    df_encuentros = pd.DataFrame(
        [
            {
                "Número A": p[0],
                "Número B": p[1],
                "Coincidencias": c,
            }
            for p, c in conteo_tiempo.most_common(10)
        ]
    )

    st.subheader("📊 INFORME DE INTELIGENCIA DE SEÑALES")

    if not top_10_celdas.empty:
        t_c = top_10_celdas.iloc[0]
        st.info(
            f"📌 Zonas Calientes: el mayor volumen de tráfico aislado fue "
            f"generado por {t_c['numero_limpio']} en la celda "
            f"{t_c['celda']}, con {t_c['conexiones']} conexiones."
        )

    if not df_encuentros.empty:
        t_e = df_encuentros.iloc[0]
        st.warning(
            f"🤝 Co-Ubicación: {t_e['Número A']} y {t_e['Número B']} "
            f"compartieron la misma celda en ventanas de 15 minutos "
            f"{t_e['Coincidencias']} veces."
        )
    else:
        st.info(
            "🤝 Co-Ubicación: no se evidenciaron patrones de encuentro "
            "simultáneo en la franja temporal de 15 min."
        )

    st.subheader("🏆 Top 10: Frecuencia de Números por Celda")
    st.dataframe(top_10_celdas, use_container_width=True)

    st.subheader("📍 Top 10: Encuentros Probables (Espacio/Tiempo)")
    st.dataframe(df_encuentros, use_container_width=True)

    # Exportación Celda
    if not top_10_celdas.empty:
        opciones_celda = [
            f"{r['celda']} | {r['numero_limpio']}"
            for _, r in top_10_celdas.iterrows()
        ]

        sel_celda = st.selectbox(
            "Frecuentes:",
            opciones_celda,
            key="sel_celda_export",
        )

        if st.button("⬇️ Descargar ZIP - Frecuente", key="btn_exp_celda"):
            celda, num = sel_celda.split(" | ")
            df_c = df_analisis[
                (df_analisis["celda"] == celda)
                & (df_analisis["numero_limpio"] == num)
            ]
            contenido = generar_exportables_bytes(
                df_c,
                f"Historial_{num}_Celda_{celda}",
            )
            st.download_button(
                "Descargar paquete",
                data=contenido,
                file_name=f"Historial_{num}_Celda_{celda}_Exportacion.zip",
                mime="application/zip",
                key="download_celda",
            )

    # Exportación Encuentro
    if not df_encuentros.empty:
        opciones_enc = [
            f"{r['Número A']} | {r['Número B']}"
            for _, r in df_encuentros.iterrows()
        ]

        sel_enc = st.selectbox(
            "Encuentros:",
            opciones_enc,
            key="sel_enc_export",
        )

        if st.button("⬇️ Descargar ZIP - Encuentro", key="btn_exp_enc"):
            nA, nB = sel_enc.split(" | ")

            dA = df_analisis[
                df_analisis["numero_limpio"] == nA
            ][["celda", "ventana"]]

            dB = df_analisis[
                df_analisis["numero_limpio"] == nB
            ][["celda", "ventana"]]

            inters = pd.merge(
                dA,
                dB,
                on=["celda", "ventana"],
            ).drop_duplicates()

            df_cruz = pd.merge(
                df_analisis,
                inters,
                on=["celda", "ventana"],
            )

            df_cruz = df_cruz[
                df_cruz["numero_limpio"].isin([nA, nB])
            ]

            contenido = generar_exportables_bytes(
                df_cruz,
                f"Encuentro_{nA}_vs_{nB}",
            )

            st.download_button(
                "Descargar paquete",
                data=contenido,
                file_name=f"Encuentro_{nA}_vs_{nB}_Exportacion.zip",
                mime="application/zip",
                key="download_encuentro",
            )


# ============================================================
# FUNCIÓN 3: LECTURA DE HECHOS / PDF / WORD / TXT
# ============================================================
def extraer_texto_documentos(documentos, texto_inicial=""):
    texto_completo = texto_inicial.strip() + "\n"

    for archivo in documentos or []:
        nombre = archivo.name
        contenido = archivo.getvalue()

        try:
            if nombre.lower().endswith(".txt"):
                texto_completo += contenido.decode(
                    "utf-8",
                    errors="ignore",
                ) + "\n"

            elif nombre.lower().endswith(".docx"):
                documento = docx.Document(io.BytesIO(contenido))
                for para in documento.paragraphs:
                    texto_completo += para.text + "\n"

            elif nombre.lower().endswith(".pdf"):
                lector = PyPDF2.PdfReader(io.BytesIO(contenido))
                for pagina in lector.pages:
                    txt_pag = pagina.extract_text()
                    if txt_pag:
                        texto_completo += txt_pag + "\n"

        except Exception as e:
            st.error(f"❌ Error leyendo {nombre}: {e}")

    return texto_completo


def generar_informe(df_master, texto_completo):
    """Conserva la función generar_informe original."""

    if not texto_completo.strip():
        st.warning("⚠️ No ingresaste texto ni subiste documentos válidos.")
        return

    numeros_implicados = list(
        set(re.findall(r"\b\d{10}\b", texto_completo))
    )

    st.subheader("📊 INFORME ANALÍTICO DE INTELIGENCIA")

    if not numeros_implicados:
        st.warning(
            "No se detectaron números de celular de 10 dígitos "
            "en el texto ni en los documentos adjuntos."
        )
        return

    df_caso = df_master[
        df_master["numero_limpio"].isin(numeros_implicados)
    ].copy()

    if df_caso.empty:
        st.error(
            "Los números hallados en el texto "
            f"({', '.join(numeros_implicados)}) NO registran tráfico "
            "en la evidencia cargada."
        )
        return

    total_registros = len(df_caso)
    nums_encontrados = df_caso["numero_limpio"].unique()

    if "fecha_limpia" in df_caso.columns:
        df_caso = df_caso.dropna(
            subset=["fecha_limpia"]
        ).sort_values("fecha_limpia")

        if not df_caso.empty:
            fecha_ini = df_caso["fecha_limpia"].min()
            fecha_fin = df_caso["fecha_limpia"].max()
            lapso = f"Desde el <b>{fecha_ini}</b> hasta el <b>{fecha_fin}</b>"
        else:
            lapso = "No hay datos de fecha disponibles."
    else:
        lapso = "No hay datos de fecha disponibles."

    top_celdas = (
        df_caso["celda"]
        .value_counts()
        .head(3)
        .index
        .tolist()
    )
    celdas_txt = ", ".join(str(c) for c in top_celdas)

    resumen_txt = (
        texto_completo[:400].replace("\n", " ")
        + ("..." if len(texto_completo) > 400 else "")
    )

    st.markdown(
        f"""
        <div style="background-color:#f9f9f9;padding:15px;
        border-radius:8px;border:1px solid #ddd;">
        <h3>📋 Sinopsis Forense Transversal</h3>
        <p><b>Contexto Procesado:</b>
        <i>"{resumen_txt}"</i></p>
        <h4>🕒 TIEMPO</h4>
        <p>El análisis de la evidencia digital abarca un espectro temporal
        activo {lapso}. Durante este periodo, los objetivos generaron un
        total de <b>{total_registros}</b> registros de conexión.</p>
        <h4>📍 LUGAR</h4>
        <p>La actividad se concentró principalmente en las celdas:
        <b>{celdas_txt}</b>.</p>
        <h4>⚙️ MODO</h4>
        <p>Se perfilaron <b>{len(nums_encontrados)}</b> línea(s) telefónica(s)
        con cruce positivo:
        <b>{', '.join(map(str, nums_encontrados))}</b>.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("📈 Análisis Visual del Comportamiento")

    if "fecha_limpia" in df_caso.columns and not df_caso.empty:
        df_graf = df_caso.copy()
        df_graf["fecha_solo_dia"] = df_graf[
            "fecha_limpia"
        ].dt.date

        conteo_dias = (
            df_graf.groupby(
                ["fecha_solo_dia", "numero_limpio"]
            )
            .size()
            .reset_index(name="eventos")
        )

        fig, ax = plt.subplots(figsize=(12, 4))
        sns.lineplot(
            data=conteo_dias,
            x="fecha_solo_dia",
            y="eventos",
            hue="numero_limpio",
            marker="o",
            linewidth=2,
            ax=ax,
        )
        ax.set_title("Frecuencia de Tráfico en el Tiempo (Por Número)")
        ax.set_xlabel("Fecha")
        ax.set_ylabel("Cantidad de Conexiones (Impactos)")
        ax.grid(True, linestyle="--", alpha=0.7)
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    if {"latitud", "longitud"}.issubset(df_caso.columns):
        df_mapa = df_caso.dropna(
            subset=["latitud", "longitud"]
        )

        if not df_mapa.empty:
            st.subheader(
                "Ubicaciones Geoespaciales Frecuentadas por los Implicados"
            )

            mapa_informe = folium.Map(
                location=[
                    df_mapa["latitud"].mean(),
                    df_mapa["longitud"].mean(),
                ],
                zoom_start=12,
            )

            colores = [
                "red",
                "blue",
                "green",
                "purple",
                "orange",
                "darkred",
                "cadetblue",
                "black",
            ]

            color_map = {
                num: colores[i % len(colores)]
                for i, num in enumerate(nums_encontrados)
            }

            agrupado_mapa = (
                df_mapa.groupby(
                    ["numero_limpio", "latitud", "longitud"]
                )
                .size()
                .reset_index(name="visitas")
            )

            for _, row in agrupado_mapa.iterrows():
                num = row["numero_limpio"]
                color = color_map[num]

                folium.CircleMarker(
                    location=[
                        row["latitud"],
                        row["longitud"],
                    ],
                    radius=min(
                        25,
                        8 + row["visitas"] * 1.5,
                    ),
                    color=color,
                    fill=True,
                    fill_opacity=0.4,
                    tooltip=(
                        f"<b>Línea:</b> {num}<br>"
                        f"<b>Impactos:</b> {row['visitas']}"
                    ),
                ).add_to(mapa_informe)

            st_folium(mapa_informe, width=None, height=600)


# ============================================================
# FUNCIÓN 4: VICTIMARIOS POR CO-DESPLAZAMIENTO
# ============================================================
def rastrear_victimarios(
    df_master,
    num_victima,
    ventana_str,
    chk_colombia,
    fecha_ini,
    hora_ini,
    fecha_fin,
    hora_fin,
):
    """Conserva el algoritmo de co-desplazamiento original."""

    if len(num_victima) != 10:
        st.warning("⚠️ Ingresa el número de la víctima a 10 dígitos.")
        return

    df_limpio = df_master.dropna(
        subset=["numero_limpio", "celda", "fecha_limpia"]
    ).copy()

    dt_ini, dt_fin = parse_datetime_range(
        fecha_ini, hora_ini, fecha_fin, hora_fin
    )

    if dt_ini is not None:
        df_limpio = df_limpio[
            (df_limpio["fecha_limpia"] >= dt_ini)
            & (df_limpio["fecha_limpia"] <= dt_fin)
        ]

    df_victima = df_limpio[
        df_limpio["numero_limpio"] == num_victima
    ].copy()

    if df_victima.empty:
        st.error(
            f"❌ La víctima ({num_victima}) no presenta tráfico. "
            "Revisa los filtros."
        )
        return

    df_otros = df_limpio[
        df_limpio["numero_limpio"] != num_victima
    ].copy()

    if chk_colombia:
        df_otros = df_otros[
            df_otros["numero_limpio"].astype(str).str.match(
                r"^3\d{9}$"
            )
        ]

    td_tolerancia = pd.Timedelta(ventana_str)
    resultados = []

    for _, v_row in df_victima.iterrows():
        celda_v = v_row["celda"]
        tiempo_v = v_row["fecha_limpia"]

        lim_inf = tiempo_v - td_tolerancia
        lim_sup = tiempo_v + td_tolerancia

        match = df_otros[
            (df_otros["celda"] == celda_v)
            & (df_otros["fecha_limpia"] >= lim_inf)
            & (df_otros["fecha_limpia"] <= lim_sup)
        ]

        for _, s_row in match.iterrows():
            resultados.append(
                {
                    "numero_limpio": s_row["numero_limpio"],
                    "celda": celda_v,
                    "fecha_sospechoso": s_row["fecha_limpia"],
                    "registro_original": s_row["registro_original"],
                }
            )

    intersecciones = pd.DataFrame(resultados)

    if intersecciones.empty:
        st.success(
            "✅ Se analizaron los puntos de la víctima; ningún número "
            "la siguió de cerca."
        )
        return

    sospechosos = (
        intersecciones.groupby("numero_limpio")
        .agg(
            celdas_compartidas=("celda", "nunique"),
            impactos_totales=("celda", "count"),
        )
        .reset_index()
        .sort_values(
            by=["celdas_compartidas", "impactos_totales"],
            ascending=[False, False],
        )
        .head(10)
    )

    top_sospechoso = sospechosos.iloc[0]

    st.warning(
        f"💡 Sinopsis de Seguimiento: de acuerdo con la ruta cronológica "
        f"de la víctima ({num_victima}), el número "
        f"{top_sospechoso['numero_limpio']} aparece como el principal "
        f"coincidente. Compartió {top_sospechoso['celdas_compartidas']} "
        f"celdas distintas y registró "
        f"{top_sospechoso['impactos_totales']} impactos dentro de una "
        f"tolerancia de {ventana_str}."
    )

    st.subheader("🎯 Top 10: Posibles Victimarios")
    st.dataframe(sospechosos, use_container_width=True)

    s_sel = st.selectbox(
        "Seleccionar número para exportar",
        sospechosos["numero_limpio"].tolist(),
        key="sospechoso_export",
    )

    if st.button("⬇️ Descargar evidencia ZIP", key="download_persecucion"):
        cruce_final = intersecciones[
            intersecciones["numero_limpio"] == s_sel
        ].drop_duplicates(
            subset=["celda", "fecha_sospechoso"]
        )

        archivos = {}
        archivos[
            f"Persecucion_{num_victima}_vs_{s_sel}.xlsx"
        ] = registros_a_excel_bytes(cruce_final)

        c_vic = df_victima.dropna(
            subset=["latitud", "longitud"]
        )
        c_sos = df_otros[
            df_otros["numero_limpio"] == s_sel
        ].dropna(subset=["latitud", "longitud"])

        arch_map = f"Rutas_{num_victima}_vs_{s_sel}.html"

        if not c_vic.empty:
            m = folium.Map(
                location=[
                    c_vic["latitud"].mean(),
                    c_vic["longitud"].mean(),
                ],
                zoom_start=13,
            )

            AntPath(
                locations=c_vic[
                    ["latitud", "longitud"]
                ].values.tolist(),
                color="green",
                weight=5,
                tooltip="Víctima",
            ).add_to(m)

            if not c_sos.empty:
                AntPath(
                    locations=c_sos[
                        ["latitud", "longitud"]
                    ].values.tolist(),
                    color="red",
                    weight=4,
                    tooltip="Sospechoso",
                ).add_to(m)

            archivos[arch_map] = mapa_html_bytes(m)
        else:
            archivos[arch_map] = b"<h3>Sin coordenadas</h3>"

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(
            zip_buffer,
            "w",
            zipfile.ZIP_DEFLATED,
        ) as zf:
            for nombre, contenido in archivos.items():
                zf.writestr(nombre, contenido)

        st.download_button(
            "Descargar paquete",
            data=zip_buffer.getvalue(),
            file_name=f"Evidencia_Persecucion_{s_sel}.zip",
            mime="application/zip",
            key="download_persecucion_final",
        )


# ============================================================
# FUNCIÓN 5: PERFILAMIENTO DE RUTINAS / KML
# ============================================================
def analizar_rutinas(
    df_master,
    num_vic,
    drop_top,
    chk_colombia,
    fecha_ini,
    hora_ini,
    fecha_fin,
    hora_fin,
):
    """Conserva el algoritmo original de perfilamiento."""

    if len(num_vic) != 10:
        st.warning("⚠️ Ingresa el número del victimario a 10 dígitos.")
        return

    df_limpio = df_master.dropna(
        subset=["numero_limpio", "celda", "fecha_limpia"]
    ).copy()

    dt_ini, dt_fin = parse_datetime_range(
        fecha_ini, hora_ini, fecha_fin, hora_fin
    )

    if dt_ini is not None:
        df_limpio = df_limpio[
            (df_limpio["fecha_limpia"] >= dt_ini)
            & (df_limpio["fecha_limpia"] <= dt_fin)
        ]

    if chk_colombia:
        df_limpio = df_limpio[
            df_limpio["numero_limpio"].astype(str).str.match(
                r"^3\d{9}$"
            )
        ]

    df_obj = df_limpio[
        df_limpio["numero_limpio"] == num_vic
    ].copy()

    if df_obj.empty:
        st.error(
            f"❌ Sin registros: el victimario ({num_vic}) no presenta "
            "tráfico en los filtros indicados."
        )
        return

    df_obj["hora"] = df_obj["fecha_limpia"].dt.hour

    def clasificar_horario(h):
        if h >= 22 or h <= 2:
            return "1. NOCHE (Pernocta: 22:00 - 02:59)"
        elif 3 <= h <= 8:
            return "2. MAÑANA/MADRUGADA (03:00 - 08:59)"
        else:
            return "3. DÍA (Actividad: 09:00 - 21:59)"

    df_obj["franja"] = df_obj["hora"].apply(clasificar_horario)

    col_desc = (
        "descripcion"
        if "descripcion" in df_obj.columns
        else (
            "direccion"
            if "direccion" in df_obj.columns
            else "celda"
        )
    )

    rutinas = (
        df_obj.groupby(["franja", "celda"])
        .agg(
            nombre_lugar=(col_desc, "first"),
            visitas=("celda", "count"),
            primer_registro=("fecha_limpia", "min"),
            ultimo_registro=("fecha_limpia", "max"),
            latitud=("latitud", "mean"),
            longitud=("longitud", "mean"),
        )
        .reset_index()
    )

    rutinas = rutinas.sort_values(
        by=["franja", "visitas"],
        ascending=[True, False],
    )

    n_top = int(drop_top.replace("Top ", ""))
    top_rutinas = rutinas.groupby("franja").head(n_top)

    def obtener_top1(df, franja_nombre):
        filtro = df[df["franja"] == franja_nombre]
        if not filtro.empty:
            return (
                f"Celda <b>{filtro.iloc[0]['celda']}</b> "
                f"({filtro.iloc[0]['nombre_lugar']}) con "
                f"{filtro.iloc[0]['visitas']} conexiones."
            )
        return "<i>Sin actividad registrada.</i>"

    noche_txt = obtener_top1(
        top_rutinas,
        "1. NOCHE (Pernocta: 22:00 - 02:59)",
    )
    manana_txt = obtener_top1(
        top_rutinas,
        "2. MAÑANA/MADRUGADA (03:00 - 08:59)",
    )
    dia_txt = obtener_top1(
        top_rutinas,
        "3. DÍA (Actividad: 09:00 - 21:59)",
    )

    st.markdown(
        f"""
        <div style="background-color:#e8f4f8;padding:15px;
        border-left:5px solid #0275d8;border-radius:5px;">
        <h4>💡 Sinopsis de Patrón de Vida</h4>
        <p>El análisis geoespacial segmentado del victimario
        <b>{num_vic}</b> arroja los siguientes lugares de alta
        recurrencia según la hora:</p>
        <ul>
        <li>🌙 <b>Lugar de Pernocta:</b> {noche_txt}</li>
        <li>🌅 <b>Inicio de Jornada:</b> {manana_txt}</li>
        <li>☀️ <b>Zona de Operación:</b> {dia_txt}</li>
        </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader(f"📊 Desglose de Rutinas ({drop_top})")
    st.dataframe(
        top_rutinas[
            [
                "franja",
                "celda",
                "nombre_lugar",
                "visitas",
                "primer_registro",
                "ultimo_registro",
            ]
        ],
        use_container_width=True,
    )

    df_mapa = top_rutinas.dropna(
        subset=["latitud", "longitud"]
    )

    if not df_mapa.empty:
        mapa = folium.Map(
            location=[
                df_mapa["latitud"].mean(),
                df_mapa["longitud"].mean(),
            ],
            zoom_start=13,
        )

        colores = {
            "1. NOCHE (Pernocta: 22:00 - 02:59)": "darkblue",
            "2. MAÑANA/MADRUGADA (03:00 - 08:59)": "lightblue",
            "3. DÍA (Actividad: 09:00 - 21:59)": "orange",
        }

        for _, row in df_mapa.iterrows():
            color = colores.get(row["franja"], "gray")

            html_pop = (
                f"<b>{row['franja']}</b><br>"
                f"Celda: {row['celda']}<br>"
                f"Lugar: {row['nombre_lugar']}<br>"
                f"Visitas: {row['visitas']}"
            )

            folium.Marker(
                location=[
                    row["latitud"],
                    row["longitud"],
                ],
                popup=folium.Popup(
                    html_pop,
                    max_width=250,
                ),
                tooltip=(
                    f"{row['franja']} "
                    f"(Visitas: {row['visitas']})"
                ),
                icon=folium.Icon(
                    color=color,
                    icon="info-sign",
                ),
            ).add_to(mapa)

        st.subheader("🗺️ Visualización y Exportación 3D")
        st_folium(mapa, width=None, height=600)

    if simplekml is not None:
        if st.button("🌎 Generar KML para Google Earth", key="btn_kml"):
            datos_kml = df_obj.dropna(
                subset=["latitud", "longitud"]
            ).sort_values("fecha_limpia")

            kml = simplekml.Kml(name=f"Rastreo_{num_vic}")

            style_noche = simplekml.Style()
            style_noche.iconstyle.icon.href = (
                "http://maps.google.com/mapfiles/kml/"
                "paddle/blu-circle.png"
            )

            style_dia = simplekml.Style()
            style_dia.iconstyle.icon.href = (
                "http://maps.google.com/mapfiles/kml/"
                "paddle/ylw-circle.png"
            )

            for _, r in datos_kml.iterrows():
                pnt = kml.newpoint(
                    name=str(r["celda"]),
                    coords=[
                        (r["longitud"], r["latitud"])
                    ],
                )
                pnt.description = (
                    f"<b>Fecha:</b> {r['fecha_limpia']}<br>"
                    f"<b>Franja:</b> {r['franja']}"
                )

                if "NOCHE" in r["franja"]:
                    pnt.style = style_noche
                else:
                    pnt.style = style_dia

            kml_buffer = io.BytesIO()
            kml.save(kml_buffer)
            st.download_button(
                "⬇️ Descargar KML",
                data=kml_buffer.getvalue(),
                file_name=f"Ruta_Victimario_{num_vic}.kml",
                mime="application/vnd.google-earth.kml+xml",
                key="download_kml",
            )
    else:
        st.info(
            "Instala simplekml para habilitar la exportación KML."
        )

    gc.collect()


# ============================================================
# INTERFAZ PRINCIPAL
# ============================================================
with st.sidebar:
    st.header("📂 Datos CDR")

    uploaded_files = st.file_uploader(
        "Sube XLSB, XLSX, XLS o CSV",
        type=["xlsb", "xlsx", "xls", "csv"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        payload = tuple(
            (f.name, f.getvalue())
            for f in uploaded_files
        )

        # Solo procesar los archivos cuando realmente cambien.
        if st.session_state.payload_cdr != payload:
            with st.spinner("Procesando y consolidando archivos..."):
                df_loaded, errores_loaded = load_and_process_files(payload)

            st.session_state.payload_cdr = payload
            st.session_state.df_master = df_loaded
            st.session_state.errores = errores_loaded

            # Una nueva carga invalida resultados anteriores.
            st.session_state.geo_ejecutado = False
            st.session_state.red_ejecutado = False
            st.session_state.informe_ejecutado = False
            st.session_state.co_ejecutado = False
            st.session_state.perfil_ejecutado = False

        df_master = st.session_state.df_master
        errores = st.session_state.errores

        if df_master is not None and not df_master.empty:
            st.success(
                f"✅ Archivos consolidados. "
                f"Registros: {len(df_master):,}"
            )
        else:
            st.error("No se encontraron datos válidos.")

        for error in errores:
            st.warning(f"⚠️ {error}")

df_master = st.session_state.df_master

if df_master is None or df_master.empty:
    st.info(
        "📂 Sube primero los archivos CDR desde la barra lateral "
        "para activar todas las funciones."
    )
    st.stop()

fecha_min, fecha_max = create_date_range(df_master)

tabs = st.tabs(
    [
        "📍 Búsqueda Geográfica",
        "📊 Análisis de Red",
        "📝 Hechos y Documentos",
        "🕵️ Co-Desplazamiento",
        "👤 Perfil de Rutinas",
    ]
)


# ------------------------------------------------------------
# TAB 1
# ------------------------------------------------------------
with tabs[0]:
    st.header("📍 Búsqueda Geográfica Individual")

    numero = st.text_input(
        "Número",
        placeholder="Ej: 3157658841",
        max_chars=10,
        key="geo_numero",
    )

    col1, col2 = st.columns(2)

    with col1:
        top_geo = st.selectbox(
            "Mostrar",
            ["Todos", "Top 10", "Top 5", "Top 3", "Top 1"],
            key="geo_top",
        )

    with col2:
        if fecha_min and fecha_max:
            f_ini = st.date_input(
                "Fecha inicial",
                value=fecha_min,
                key="geo_fini",
            )
            f_fin = st.date_input(
                "Fecha final",
                value=fecha_max,
                key="geo_ffin",
            )
        else:
            f_ini = f_fin = None

    c1, c2 = st.columns(2)
    with c1:
        h_ini = st.text_input(
            "Hora inicial",
            "00:00",
            key="geo_hini",
        )
    with c2:
        h_fin = st.text_input(
            "Hora final",
            "23:59",
            key="geo_hfin",
        )

    if st.button("🗺️ Trazar Mapa", type="primary", key="geo_btn"):
        st.session_state.geo_numero_resultado = numero.strip()
        st.session_state.geo_top_resultado = top_geo
        st.session_state.geo_fini_resultado = f_ini
        st.session_state.geo_hini_resultado = h_ini
        st.session_state.geo_ffin_resultado = f_fin
        st.session_state.geo_hfin_resultado = h_fin
        st.session_state.geo_ejecutado = True

    if st.session_state.geo_ejecutado:
        generar_mapa(
            df_master,
            st.session_state.geo_numero_resultado,
            st.session_state.geo_top_resultado,
            st.session_state.geo_fini_resultado,
            st.session_state.geo_hini_resultado,
            st.session_state.geo_ffin_resultado,
            st.session_state.geo_hfin_resultado,
        )


# ------------------------------------------------------------
# TAB 2
# ------------------------------------------------------------
with tabs[1]:
    st.header("📊 Panel de Filtros para Análisis")

    chk_colombia = st.checkbox(
        "Omitir Plataformas (solo móviles Colombia que inician con 3)",
        value=True,
        key="red_colombia",
    )

    c1, c2 = st.columns(2)
    with c1:
        red_fini = st.date_input(
            "Fecha inicial",
            value=fecha_min,
            key="red_fini",
        )
        red_hini = st.text_input(
            "Hora inicial",
            "00:00",
            key="red_hini",
        )
    with c2:
        red_ffin = st.date_input(
            "Fecha final",
            value=fecha_max,
            key="red_ffin",
        )
        red_hfin = st.text_input(
            "Hora final",
            "23:59",
            key="red_hfin",
        )

    excluir = st.text_input(
        "Excluir números",
        placeholder="Ej: 3151234567, 3001234567",
        key="red_excluir",
    )

    if st.button("▶️ Generar Análisis", type="primary", key="red_btn"):
        st.session_state.red_chk_colombia = chk_colombia
        st.session_state.red_fini_resultado = red_fini
        st.session_state.red_hini_resultado = red_hini
        st.session_state.red_ffin_resultado = red_ffin
        st.session_state.red_hfin_resultado = red_hfin
        st.session_state.red_excluir_resultado = excluir
        st.session_state.red_ejecutado = True

    if st.session_state.red_ejecutado:
        procesar_analisis(
            df_master,
            st.session_state.red_chk_colombia,
            st.session_state.red_fini_resultado,
            st.session_state.red_hini_resultado,
            st.session_state.red_ffin_resultado,
            st.session_state.red_hfin_resultado,
            st.session_state.red_excluir_resultado,
        )


# ------------------------------------------------------------
# TAB 3
# ------------------------------------------------------------
with tabs[2]:
    st.header("📝 Redacción de Hechos y Lectura de Documentos")

    st.write(
        "Escribe el contexto del caso o sube oficios/informes "
        "en PDF, Word o TXT. Se extraen automáticamente números "
        "de 10 dígitos y se cruzan con la evidencia CDR."
    )

    hechos = st.text_area(
        "Hechos / contexto",
        height=150,
        key="hechos_texto",
    )

    docs = st.file_uploader(
        "📁 Subir documentos",
        type=["txt", "pdf", "docx"],
        accept_multiple_files=True,
        key="docs_informe",
    )

    if st.button(
        "📄 Generar Informe Analítico",
        type="primary",
        key="informe_btn",
    ):
        st.session_state.informe_texto = extraer_texto_documentos(
            docs,
            hechos,
        )
        st.session_state.informe_ejecutado = True

    if st.session_state.informe_ejecutado:
        generar_informe(
            df_master,
            st.session_state.informe_texto,
        )


# ------------------------------------------------------------
# TAB 4
# ------------------------------------------------------------
with tabs[3]:
    st.header(
        "🕵️‍♂️ Búsqueda de Victimarios por Co-Desplazamiento"
    )

    num_victima = st.text_input(
        "Núm. Víctima",
        placeholder="Ej: 3157658841",
        max_chars=10,
        key="co_victima",
    )

    ventana = st.selectbox(
        "Tolerancia temporal",
        ["15min", "30min", "1H", "2H"],
        index=2,
        key="co_ventana",
    )

    co_colombia = st.checkbox(
        "Ignorar Plataformas (solo móviles Colombia)",
        value=True,
        key="co_colombia",
    )

    c1, c2 = st.columns(2)
    with c1:
        co_fini = st.date_input(
            "Fecha inicial",
            value=fecha_min,
            key="co_fini",
        )
        co_hini = st.text_input(
            "Hora inicial",
            "00:00",
            key="co_hini",
        )
    with c2:
        co_ffin = st.date_input(
            "Fecha final",
            value=fecha_max,
            key="co_ffin",
        )
        co_hfin = st.text_input(
            "Hora final",
            "23:59",
            key="co_hfin",
        )

    if st.button(
        "🔎 Buscar Victimarios",
        type="primary",
        key="co_btn",
    ):
        st.session_state.co_numero_resultado = num_victima.strip()
        st.session_state.co_ventana_resultado = ventana
        st.session_state.co_colombia_resultado = co_colombia
        st.session_state.co_fini_resultado = co_fini
        st.session_state.co_hini_resultado = co_hini
        st.session_state.co_ffin_resultado = co_ffin
        st.session_state.co_hfin_resultado = co_hfin
        st.session_state.co_ejecutado = True

    if st.session_state.co_ejecutado:
        rastrear_victimarios(
            df_master,
            st.session_state.co_numero_resultado,
            st.session_state.co_ventana_resultado,
            st.session_state.co_colombia_resultado,
            st.session_state.co_fini_resultado,
            st.session_state.co_hini_resultado,
            st.session_state.co_ffin_resultado,
            st.session_state.co_hfin_resultado,
        )


# ------------------------------------------------------------
# TAB 5
# ------------------------------------------------------------
with tabs[4]:
    st.header(
        "👤 Perfilamiento de Rutinas y Zonas de Pernocta"
    )

    num_vic = st.text_input(
        "Victimario",
        placeholder="Ej: 3157658841",
        max_chars=10,
        key="perfil_num",
    )

    drop_top = st.selectbox(
        "Mostrar",
        ["Top 10", "Top 5", "Top 3", "Top 1"],
        index=2,
        key="perfil_top",
    )

    perfil_colombia = st.checkbox(
        "Ignorar Plataformas (solo móviles Colombia)",
        value=True,
        key="perfil_colombia",
    )

    c1, c2 = st.columns(2)
    with c1:
        p_fini = st.date_input(
            "Fecha inicial",
            value=fecha_min,
            key="perfil_fini",
        )
        p_hini = st.text_input(
            "Hora inicial",
            "00:00",
            key="perfil_hini",
        )
    with c2:
        p_ffin = st.date_input(
            "Fecha final",
            value=fecha_max,
            key="perfil_ffin",
        )
        p_hfin = st.text_input(
            "Hora final",
            "23:59",
            key="perfil_hfin",
        )

    if st.button(
        "👤 Generar Perfil de Vida",
        type="primary",
        key="perfil_btn",
    ):
        st.session_state.perfil_numero_resultado = num_vic.strip()
        st.session_state.perfil_top_resultado = drop_top
        st.session_state.perfil_colombia_resultado = perfil_colombia
        st.session_state.perfil_fini_resultado = p_fini
        st.session_state.perfil_hini_resultado = p_hini
        st.session_state.perfil_ffin_resultado = p_ffin
        st.session_state.perfil_hfin_resultado = p_hfin
        st.session_state.perfil_ejecutado = True

    if st.session_state.perfil_ejecutado:
        analizar_rutinas(
            df_master,
            st.session_state.perfil_numero_resultado,
            st.session_state.perfil_top_resultado,
            st.session_state.perfil_colombia_resultado,
            st.session_state.perfil_fini_resultado,
            st.session_state.perfil_hini_resultado,
            st.session_state.perfil_ffin_resultado,
            st.session_state.perfil_hfin_resultado,
        )


st.divider()
st.caption(
    "Aplicación convertida desde el notebook original. "
    "Las funciones de limpieza, geolocalización, análisis de red, "
    "lectura documental, co-desplazamiento, perfilamiento y exportación "
    "se mantienen en una interfaz Streamlit."
)
