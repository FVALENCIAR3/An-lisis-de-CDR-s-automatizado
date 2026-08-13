
import io
import re
import zipfile
from datetime import datetime, date, time

import pandas as pd
import streamlit as st
import folium
from folium.plugins import AntPath
from streamlit_folium import st_folium
import matplotlib.pyplot as plt
import seaborn as sns
import itertools
from collections import Counter
import PyPDF2
import docx
import simplekml


# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Análisis automatizado de CDR",
    page_icon="📡",
    layout="wide",
)

st.title("📡 Análisis de CDR's Automatizado")
st.caption(
    "Herramienta para consolidación, análisis temporal/geográfico, "
    "co-ubicación, perfilamiento de rutinas y lectura de documentos."
)


# ============================================================
# FUNCIONES DE PREPROCESAMIENTO
# ============================================================
def limpiar_fechas(s):
    s = s.astype(str).str.strip()
    mask_14 = s.str.match(r"^\d{14}$")
    fechas_14 = pd.to_datetime(
        s[mask_14], format="%Y%m%d%H%M%S", errors="coerce"
    )
    try:
        fechas_otros = pd.to_datetime(
            s[~mask_14], format="mixed", dayfirst=True, errors="coerce"
        )
    except Exception:
        fechas_otros = pd.to_datetime(
            s[~mask_14], dayfirst=True, errors="coerce"
        )
    return pd.concat([fechas_14, fechas_otros]).sort_index()


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


def cargar_dataframe(uploaded_files):
    df_list = []
    errores = []

    for file in uploaded_files:
        try:
            content = file.getvalue()
            nombre = file.name.lower()

            if nombre.endswith(".csv"):
                try:
                    df = pd.read_csv(
                        io.BytesIO(content),
                        sep=";",
                        encoding="utf-8-sig",
                        low_memory=False,
                    )
                    if len(df.columns) <= 1:
                        df = pd.read_csv(
                            io.BytesIO(content),
                            sep=",",
                            encoding="utf-8-sig",
                            low_memory=False,
                        )
                except Exception:
                    df = pd.read_csv(
                        io.BytesIO(content),
                        sep=None,
                        engine="python",
                    )
            elif nombre.endswith((".xlsx", ".xls")):
                df = pd.read_excel(io.BytesIO(content))
            else:
                errores.append(f"{file.name}: formato no soportado.")
                continue

            df.columns = df.columns.astype(str).str.strip().str.lower()

            col_tel = next(
                (
                    c for c in [
                        "numero",
                        "numero_origen",
                        "numero_que_marca",
                        "numero_que_navega",
                        "originador",
                    ]
                    if c in df.columns
                ),
                None,
            )

            col_fec = next(
                (
                    c for c in [
                        "fecha_hora_inicio",
                        "fecha_trafico",
                        "fecha_hora_inicio_llamada",
                        "fecha_hora_inicio_sesion",
                        "fecha_hora",
                    ]
                    if c in df.columns
                ),
                None,
            )

            col_lat = next(
                (c for c in ["latitud", "latitud_n"] if c in df.columns),
                None,
            )

            col_lon = next(
                (c for c in ["longitud", "longitud_w"] if c in df.columns),
                None,
            )

            col_cel = next(
                (
                    c for c in [
                        "celda_decimal",
                        "cell_id_voz",
                        "celda",
                        "celda_inicio_llamada",
                        "bts_id",
                        "celda_hex",
                    ]
                    if c in df.columns
                ),
                None,
            )

            df_temp = pd.DataFrame(index=df.index)

            if col_tel:
                df_temp["numero_limpio"] = (
                    df[col_tel]
                    .astype(str)
                    .str.replace(r"\D", "", regex=True)
                    .str[-10:]
                )
            else:
                df_temp["numero_limpio"] = pd.NA

            if col_fec:
                df_temp["fecha_limpia"] = limpiar_fechas(df[col_fec])
            else:
                df_temp["fecha_limpia"] = pd.NaT

            if col_lat and col_lon:
                df_temp["latitud"] = pd.to_numeric(
                    df[col_lat].astype(str).str.replace(",", ".", regex=False),
                    errors="coerce",
                )
                df_temp["longitud"] = pd.to_numeric(
                    df[col_lon].astype(str).str.replace(",", ".", regex=False),
                    errors="coerce",
                )
            else:
                df_temp["latitud"] = pd.NA
                df_temp["longitud"] = pd.NA

            if col_cel:
                df_temp["celda"] = df[col_cel].apply(normalizar_celda)
            else:
                df_temp["celda"] = pd.NA

            df_temp["registro_original"] = df.to_dict("records")
            df_list.append(df_temp)

        except Exception as exc:
            errores.append(f"{file.name}: {exc}")

    if not df_list:
        return None, errores

    df_master = pd.concat(df_list, ignore_index=True)

    df_master["fecha_limpia"] = pd.to_datetime(
        df_master["fecha_limpia"], errors="coerce"
    )
    df_master["numero_limpio"] = df_master["numero_limpio"].astype("string")
    df_master["celda"] = df_master["celda"].astype("string")

    return df_master, errores


def rango_fechas(df):
    fechas = df["fecha_limpia"].dropna()
    if fechas.empty:
        return date.today(), date.today()
    return fechas.min().date(), fechas.max().date()


def aplicar_filtro_fecha(df, fecha_ini, hora_ini, fecha_fin, hora_fin):
    if "fecha_limpia" not in df.columns:
        return df.copy()

    dt_ini = pd.Timestamp.combine(fecha_ini, hora_ini)
    dt_fin = pd.Timestamp.combine(fecha_fin, hora_fin)

    return df[
        (df["fecha_limpia"] >= dt_ini)
        & (df["fecha_limpia"] <= dt_fin)
    ].copy()


def df_export_original(df):
    if df.empty:
        return pd.DataFrame()

    registros = df["registro_original"].tolist()
    salida = pd.DataFrame(registros).astype(str)

    for col in salida.columns:
        salida[col] = (
            salida[col]
            .str.replace(r"\.0$", "", regex=True)
            .replace(["nan", "NaT", "None"], "")
        )
    return salida


def excel_bytes(df):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer.getvalue()


def mapa_html_bytes(mapa):
    return mapa.get_root().render().encode("utf-8")


def zip_bytes(files_dict):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files_dict.items():
            zf.writestr(name, data)
    buffer.seek(0)
    return buffer.getvalue()


def filtros_ui(df, key_prefix):
    f_min, f_max = rango_fechas(df)

    c1, c2, c3, c4 = st.columns(4)
    fecha_ini = c1.date_input(
        "Fecha inicial",
        value=f_min,
        key=f"{key_prefix}_fi",
    )
    hora_ini = c2.time_input(
        "Hora inicial",
        value=time(0, 0),
        key=f"{key_prefix}_hi",
    )
    fecha_fin = c3.date_input(
        "Fecha final",
        value=f_max,
        key=f"{key_prefix}_ff",
    )
    hora_fin = c4.time_input(
        "Hora final",
        value=time(23, 59),
        key=f"{key_prefix}_hf",
    )

    return fecha_ini, hora_ini, fecha_fin, hora_fin


# ============================================================
# CARGA DE DATOS
# ============================================================
with st.sidebar:
    st.header("📂 Datos")
    archivos = st.file_uploader(
        "Sube archivos CSV o Excel",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="cdr_upload",
    )

    if st.button("🗑️ Limpiar datos", use_container_width=True):
        st.session_state.pop("df_master", None)
        st.rerun()

if archivos:
    nuevos = [f.name for f in archivos]
    if st.session_state.get("loaded_names") != nuevos:
        with st.spinner("Procesando y consolidando archivos..."):
            df_loaded, errores = cargar_dataframe(archivos)
        if df_loaded is not None:
            st.session_state["df_master"] = df_loaded
            st.session_state["loaded_names"] = nuevos
            st.session_state["load_errors"] = errores

df_master = st.session_state.get("df_master")

if df_master is None:
    st.info(
        "👈 Sube uno o varios archivos CSV/Excel desde la barra lateral "
        "para comenzar."
    )
    st.markdown(
        """
        ### Flujo de trabajo
        1. **Carga y consolidación** de CDR.
        2. **Búsqueda geográfica individual**.
        3. **Análisis de red y co-ubicación**.
        4. **Lectura de hechos/documentos**.
        5. **Búsqueda de posibles co-desplazamientos**.
        6. **Perfilamiento de rutinas y exportación KML**.
        """
    )
    st.stop()

if st.session_state.get("load_errors"):
    for error in st.session_state["load_errors"]:
        st.warning(error)

# ============================================================
# RESUMEN
# ============================================================
st.success(
    f"✅ Datos consolidados: **{len(df_master):,} registros** "
    f"de **{len(st.session_state.get('loaded_names', []))} archivo(s)**."
)

with st.expander("🔎 Vista previa de los datos"):
    st.dataframe(df_master.head(100), use_container_width=True)

tabs = st.tabs(
    [
        "📍 Geográfica",
        "📊 Análisis de red",
        "📝 Hechos y documentos",
        "🕵️ Co-desplazamiento",
        "👤 Patrón de vida",
    ]
)


# ============================================================
# TAB 1 - BÚSQUEDA GEOGRÁFICA
# ============================================================
with tabs[0]:
    st.header("📍 Búsqueda Geográfica Individual")

    numero = st.text_input(
        "Número",
        placeholder="Ej: 3157658841",
        key="geo_num",
    )

    top_opcion = st.selectbox(
        "Mostrar",
        ["Todos", "Top 10", "Top 5", "Top 3", "Top 1"],
        key="geo_top",
    )

    fi, hi, ff, hf = filtros_ui(df_master, "geo")

    if st.button("🗺️ Trazar mapa", type="primary", key="geo_btn"):
        if not re.fullmatch(r"\d{10}", numero.strip()):
            st.error("Ingresa un número de exactamente 10 dígitos.")
        elif not {"latitud", "longitud"}.issubset(df_master.columns):
            st.error("La base no contiene columnas de coordenadas.")
        else:
            d = df_master[
                (df_master["numero_limpio"] == numero.strip())
                & df_master["latitud"].notna()
                & df_master["longitud"].notna()
            ].copy()

            d = aplicar_filtro_fecha(d, fi, hi, ff, hf)

            if d.empty:
                st.warning(
                    f"No hay historial de coordenadas para {numero} "
                    "en el rango seleccionado."
                )
            else:
                d = d.sort_values("fecha_limpia")

                agrupado = (
                    d.groupby(["latitud", "longitud"])
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
                    .sort_values("visitas", ascending=False)
                    .reset_index(drop=True)
                )

                top = agrupado.iloc[0]
                st.info(
                    f"💡 El objetivo **{numero}** registra **{len(d):,}**
                    conexiones geo-posicionadas. La ubicación con mayor
                    recurrencia corresponde a la(s) celda(s)
                    **{top['celdas']}**, con **{top['visitas']}** visitas."
                )

                n = (
                    len(agrupado)
                    if top_opcion == "Todos"
                    else int(top_opcion.replace("Top ", ""))
                )
                mostrar = agrupado.head(n)

                mapa = folium.Map(
                    location=[
                        mostrar.iloc[0]["latitud"],
                        mostrar.iloc[0]["longitud"],
                    ],
                    zoom_start=14,
                )

                for i, row in mostrar.iterrows():
                    rank = i + 1
                    color = (
                        "red"
                        if rank == 1
                        else "orange"
                        if rank <= 3
                        else "blue"
                    )
                    coord = [row["latitud"], row["longitud"]]

                    popup = (
                        f"<b>Rank:</b> #{rank}<br>"
                        f"<b>Visitas:</b> {row['visitas']}<br>"
                        f"<b>Celda:</b> {row['celdas']}<br>"
                        f"<b>Inicio:</b> {row['primera_visita']}<br>"
                        f"<b>Fin:</b> {row['ultima_visita']}"
                    )

                    folium.CircleMarker(
                        location=coord,
                        radius=min(20, 8 + row["visitas"]),
                        color=color,
                        fill=True,
                        fill_opacity=0.4,
                    ).add_to(mapa)

                    folium.Marker(
                        location=coord,
                        popup=folium.Popup(popup, max_width=300),
                        tooltip=f"Rank #{rank}",
                        icon=folium.Icon(color=color),
                    ).add_to(mapa)

                if top_opcion == "Todos" and len(d) > 1:
                    AntPath(
                        locations=d[["latitud", "longitud"]].values.tolist(),
                        delay=1000,
                        color="purple",
                        weight=4,
                    ).add_to(mapa)

                st_folium(
                    mapa,
                    use_container_width=True,
                    height=600,
                    key="map_geo",
                )

                st.dataframe(mostrar, use_container_width=True)


# ============================================================
# TAB 2 - ANÁLISIS DE RED
# ============================================================
with tabs[1]:
    st.header("📊 Informe de Inteligencia de Señales")

    solo_colombia = st.checkbox(
        "Omitir plataformas (solo móviles Colombia que inician con 3)",
        value=True,
        key="red_colombia",
    )

    fi, hi, ff, hf = filtros_ui(df_master, "red")

    excluir = st.text_input(
        "Excluir números",
        placeholder="Ej: 3151234567, 3001234567",
        key="red_excluir",
    )

    if st.button("▶️ Generar análisis", type="primary", key="red_btn"):
        d = df_master.dropna(
            subset=["numero_limpio", "celda", "fecha_limpia"]
        ).copy()

        if solo_colombia:
            d = d[d["numero_limpio"].str.match(r"^3\d{9}$", na=False)]

        d = aplicar_filtro_fecha(d, fi, hi, ff, hf)

        excluir_nums = [
            x.strip() for x in excluir.split(",") if x.strip()
        ]
        if excluir_nums:
            d = d[~d["numero_limpio"].isin(excluir_nums)]

        if d.empty:
            st.warning("No hay datos para analizar con los filtros aplicados.")
        else:
            top_celdas = (
                d.groupby(["celda", "numero_limpio"])
                .size()
                .reset_index(name="conexiones")
                .sort_values("conexiones", ascending=False)
                .head(10)
                .reset_index(drop=True)
            )

            d["ventana"] = d["fecha_limpia"].dt.floor("15min")
            agrupado_tiempo = d.groupby(
                ["celda", "ventana"]
            )["numero_limpio"].unique()

            conteo = Counter()
            for numeros in agrupado_tiempo:
                if 1 < len(numeros) <= 150:
                    conteo.update(
                        itertools.combinations(sorted(numeros), 2)
                    )

            encuentros = pd.DataFrame(
                [
                    {
                        "Número A": p[0],
                        "Número B": p[1],
                        "Coincidencias": c,
                    }
                    for p, c in conteo.most_common(10)
                ]
            )

            st.subheader("📌 Sinopsis analítica")

            if not top_celdas.empty:
                t = top_celdas.iloc[0]
                st.info(
                    f"**Zonas calientes:** el mayor volumen aislado fue "
                    f"generado por el número **{t['numero_limpio']}**, "
                    f"anclado a la celda **{t['celda']}**, con "
                    f"**{t['conexiones']} conexiones**."
                )

            if not encuentros.empty:
                e = encuentros.iloc[0]
                st.warning(
                    f"**Co-ubicación:** las líneas **{e['Número A']}** y "
                    f"**{e['Número B']}** compartieron la misma celda "
                    f"en ventanas de 15 minutos en "
                    f"**{e['Coincidencias']} ocasiones**."
                )
            else:
                st.success(
                    "No se evidenciaron patrones de encuentro simultáneo "
                    "en ventanas de 15 minutos."
                )

            st.subheader("🏆 Top 10: frecuencia de números por celda")
            st.dataframe(top_celdas, use_container_width=True)

            st.subheader("📍 Top 10: encuentros probables")
            if encuentros.empty:
                st.write("Sin encuentros detectados.")
            else:
                st.dataframe(encuentros, use_container_width=True)

            st.subheader("⬇️ Exportar evidencia")

            if not top_celdas.empty:
                opciones = [
                    f"{r['celda']} | {r['numero_limpio']}"
                    for _, r in top_celdas.iterrows()
                ]
                seleccion = st.selectbox(
                    "Registro frecuente",
                    opciones,
                    key="red_sel_celda",
                )

                if st.button(
                    "Preparar ZIP del registro frecuente",
                    key="red_exp_celda",
                ):
                    celda, num = seleccion.split(" | ", 1)
                    dc = d[
                        (d["celda"] == celda)
                        & (d["numero_limpio"] == num)
                    ].copy()

                    xlsx = excel_bytes(df_export_original(dc))

                    fig, ax = plt.subplots(figsize=(10, 4))
                    sns.histplot(
                        data=dc,
                        x="fecha_limpia",
                        bins=20,
                        kde=True,
                        ax=ax,
                    )
                    ax.set_title("Línea de Tiempo de Conexiones")
                    fig.tight_layout()

                    img = io.BytesIO()
                    fig.savefig(img, format="png")
                    plt.close(fig)

                    mapa_exp = folium.Map(
                        location=[
                            dc["latitud"].dropna().mean()
                            if dc["latitud"].notna().any()
                            else 0,
                            dc["longitud"].dropna().mean()
                            if dc["longitud"].notna().any()
                            else 0,
                        ],
                        zoom_start=14,
                    )

                    dm = dc.dropna(subset=["latitud", "longitud"])
                    for _, r in dm.iterrows():
                        folium.CircleMarker(
                            [r["latitud"], r["longitud"]],
                            radius=5,
                            color="red",
                            fill=True,
                            popup=str(r["fecha_limpia"]),
                        ).add_to(mapa_exp)

                    paquete = zip_bytes(
                        {
                            f"Historial_{num}_Celda_{celda}.xlsx": xlsx,
                            f"Historial_{num}_Celda_{celda}.png": img.getvalue(),
                            f"Historial_{num}_Celda_{celda}_Mapa.html":
                                mapa_html_bytes(mapa_exp),
                        }
                    )

                    st.download_button(
                        "⬇️ Descargar ZIP",
                        paquete,
                        file_name=f"Historial_{num}_Celda_{celda}_Exportacion.zip",
                        mime="application/zip",
                        key="red_download_celda",
                    )


# ============================================================
# TAB 3 - HECHOS Y DOCUMENTOS
# ============================================================
with tabs[2]:
    st.header("📝 Redacción de Hechos y Lectura de Documentos")
    st.write(
        "Escribe el contexto del caso o sube oficios/informes en "
        "PDF, Word o TXT. Se extraen automáticamente números de "
        "10 dígitos y se cruzan con la evidencia cargada."
    )

    hechos = st.text_area(
        "Hechos / contexto",
        placeholder="Escribe aquí el contexto del caso...",
        height=180,
        key="hechos_text",
    )

    documentos = st.file_uploader(
        "Subir documentos",
        type=["txt", "pdf", "docx"],
        accept_multiple_files=True,
        key="hechos_docs",
    )

    if st.button(
        "📄 Generar informe analítico",
        type="primary",
        key="hechos_btn",
    ):
        texto = hechos.strip() + "\n"

        for archivo in documentos or []:
            try:
                contenido = archivo.getvalue()
                nombre = archivo.name.lower()

                if nombre.endswith(".txt"):
                    texto += contenido.decode(
                        "utf-8", errors="ignore"
                    ) + "\n"

                elif nombre.endswith(".docx"):
                    documento = docx.Document(io.BytesIO(contenido))
                    for para in documento.paragraphs:
                        texto += para.text + "\n"

                elif nombre.endswith(".pdf"):
                    lector = PyPDF2.PdfReader(io.BytesIO(contenido))
                    for pagina in lector.pages:
                        pagina_txt = pagina.extract_text()
                        if pagina_txt:
                            texto += pagina_txt + "\n"

            except Exception as exc:
                st.error(f"Error leyendo {archivo.name}: {exc}")

        if not texto.strip():
            st.warning(
                "No ingresaste texto ni subiste documentos válidos."
            )
        else:
            numeros = sorted(set(re.findall(r"\b\d{10}\b", texto)))

            st.subheader("📊 Informe analítico de inteligencia")

            if not numeros:
                st.warning(
                    "No se detectaron números de celular de 10 dígitos "
                    "en el texto ni en los documentos."
                )
            else:
                st.write("**Números detectados:**", ", ".join(numeros))

                caso = df_master[
                    df_master["numero_limpio"].isin(numeros)
                ].copy()

                if caso.empty:
                    st.error(
                        "Los números hallados en el texto no registran "
                        "tráfico en la evidencia cargada."
                    )
                else:
                    caso = caso.dropna(
                        subset=["fecha_limpia"]
                    ).sort_values("fecha_limpia")

                    total = len(caso)
                    nums_encontrados = caso[
                        "numero_limpio"
                    ].dropna().unique()

                    fecha_ini = caso["fecha_limpia"].min()
                    fecha_fin = caso["fecha_limpia"].max()

                    top_celdas = (
                        caso["celda"]
                        .value_counts()
                        .head(3)
                        .index.astype(str)
                        .tolist()
                    )

                    resumen = texto[:400].replace("\n", " ")
                    if len(texto) > 400:
                        resumen += "..."

                    st.info(
                        f"**Contexto procesado:** {resumen}"
                    )

                    st.markdown(
                        f"""
                        **🕒 TIEMPO**  
                        El análisis abarca desde **{fecha_ini}** hasta
                        **{fecha_fin}**, con un total de
                        **{total} registros** de conexión.

                        **📍 LUGAR**  
                        Las celdas con mayor recurrencia son:
                        **{", ".join(top_celdas)}**.

                        **⚙️ MODO**  
                        Se identificaron **{len(nums_encontrados)}**
                        línea(s) con cruce positivo:
                        **{", ".join(nums_encontrados)}**.
                        """
                    )

                    if not caso.empty:
                        graf = (
                            caso.assign(
                                fecha_solo_dia=caso[
                                    "fecha_limpia"
                                ].dt.date
                            )
                            .groupby(
                                ["fecha_solo_dia", "numero_limpio"]
                            )
                            .size()
                            .reset_index(name="eventos")
                        )

                        fig, ax = plt.subplots(figsize=(12, 4))
                        sns.lineplot(
                            data=graf,
                            x="fecha_solo_dia",
                            y="eventos",
                            hue="numero_limpio",
                            marker="o",
                            linewidth=2,
                            ax=ax,
                        )
                        ax.set_title(
                            "Frecuencia de tráfico en el tiempo"
                        )
                        ax.set_xlabel("Fecha")
                        ax.set_ylabel("Cantidad de conexiones")
                        ax.grid(True, linestyle="--", alpha=0.7)
                        fig.tight_layout()
                        st.pyplot(fig)
                        plt.close(fig)

                    dm = caso.dropna(
                        subset=["latitud", "longitud"]
                    )

                    if not dm.empty:
                        st.subheader(
                            "📍 Ubicaciones geoespaciales"
                        )

                        mapa = folium.Map(
                            location=[
                                dm["latitud"].mean(),
                                dm["longitud"].mean(),
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
                            n: colores[i % len(colores)]
                            for i, n in enumerate(nums_encontrados)
                        }

                        agrupado = (
                            dm.groupby(
                                [
                                    "numero_limpio",
                                    "latitud",
                                    "longitud",
                                ]
                            )
                            .size()
                            .reset_index(name="visitas")
                        )

                        for _, row in agrupado.iterrows():
                            num = row["numero_limpio"]
                            folium.CircleMarker(
                                location=[
                                    row["latitud"],
                                    row["longitud"],
                                ],
                                radius=min(
                                    25,
                                    8 + row["visitas"] * 1.5,
                                ),
                                color=color_map.get(num, "blue"),
                                fill=True,
                                fill_opacity=0.4,
                                tooltip=(
                                    f"Línea: {num} | "
                                    f"Impactos: {row['visitas']}"
                                ),
                            ).add_to(mapa)

                        st_folium(
                            mapa,
                            use_container_width=True,
                            height=550,
                            key="hechos_map",
                        )


# ============================================================
# TAB 4 - CO-DESPLAZAMIENTO
# ============================================================
with tabs[3]:
    st.header("🕵️ Búsqueda de posibles co-desplazamientos")

    victima = st.text_input(
        "Número de referencia",
        placeholder="Ej: 3157658841",
        key="co_victima",
    )

    tolerancia = st.selectbox(
        "Tolerancia temporal",
        ["15min", "30min", "1H", "2H"],
        index=2,
        key="co_tol",
    )

    ignorar_plataformas = st.checkbox(
        "Ignorar plataformas: solo móviles Colombia",
        value=True,
        key="co_colombia",
    )

    fi, hi, ff, hf = filtros_ui(df_master, "co")

    if st.button(
        "🔎 Buscar coincidencias",
        type="primary",
        key="co_btn",
    ):
        if not re.fullmatch(r"\d{10}", victima.strip()):
            st.error("Ingresa el número a 10 dígitos.")
        else:
            d = df_master.dropna(
                subset=["numero_limpio", "celda", "fecha_limpia"]
            ).copy()

            d = aplicar_filtro_fecha(d, fi, hi, ff, hf)

            victim = victima.strip()
            dv = d[d["numero_limpio"] == victim].copy()

            if ignorar_plataformas:
                otros = d[
                    (d["numero_limpio"] != victim)
                    & d["numero_limpio"].str.match(
                        r"^3\d{9}$", na=False
                    )
                ].copy()
            else:
                otros = d[d["numero_limpio"] != victim].copy()

            if dv.empty:
                st.warning(
                    f"La línea {victim} no presenta tráfico "
                    "en los filtros seleccionados."
                )
            else:
                td = pd.Timedelta(tolerancia)
                resultados = []

                for _, vr in dv.iterrows():
                    celda_v = vr["celda"]
                    tiempo_v = vr["fecha_limpia"]

                    limite_inf = tiempo_v - td
                    limite_sup = tiempo_v + td

                    match = otros[
                        (otros["celda"] == celda_v)
                        & (
                            otros["fecha_limpia"]
                            >= limite_inf
                        )
                        & (
                            otros["fecha_limpia"]
                            <= limite_sup
                        )
                    ]

                    for _, sr in match.iterrows():
                        resultados.append(
                            {
                                "numero_limpio": sr[
                                    "numero_limpio"
                                ],
                                "celda": celda_v,
                                "fecha_sospechoso": sr[
                                    "fecha_limpia"
                                ],
                                "registro_original": sr[
                                    "registro_original"
                                ],
                            }
                        )

                inter = pd.DataFrame(resultados)

                if inter.empty:
                    st.success(
                        "No se encontraron números que coincidieran "
                        "con la trayectoria dentro de la tolerancia."
                    )
                else:
                    sospechosos = (
                        inter.groupby("numero_limpio")
                        .agg(
                            celdas_compartidas=(
                                "celda",
                                "nunique",
                            ),
                            impactos_totales=(
                                "celda",
                                "count",
                            ),
                        )
                        .reset_index()
                        .sort_values(
                            [
                                "celdas_compartidas",
                                "impactos_totales",
                            ],
                            ascending=[False, False],
                        )
                        .head(10)
                    )

                    top = sospechosos.iloc[0]

                    st.warning(
                        f"🎯 Principal coincidencia: **{top['numero_limpio']}**. "
                        f"Compartió **{top['celdas_compartidas']} celdas** "
                        f"en **{top['impactos_totales']} impactos**, "
                        f"dentro de una ventana máxima de **{tolerancia}**."
                    )

                    st.dataframe(
                        sospechosos,
                        use_container_width=True,
                    )

                    opciones = sospechosos[
                        "numero_limpio"
                    ].tolist()

                    seleccionado = st.selectbox(
                        "Seleccionar número para exportar",
                        opciones,
                        key="co_sel",
                    )

                    if st.button(
                        "📦 Preparar evidencia ZIP",
                        key="co_export",
                    ):
                        cruce = inter[
                            inter["numero_limpio"]
                            == seleccionado
                        ].drop_duplicates(
                            subset=[
                                "celda",
                                "fecha_sospechoso",
                            ]
                        )

                        exp = df_export_original(
                            cruce.assign(
                                registro_original=cruce[
                                    "registro_original"
                                ]
                            )
                        )

                        # Excel
                        xlsx = excel_bytes(exp)

                        # Mapa
                        cv = dv.dropna(
                            subset=["latitud", "longitud"]
                        )
                        cs = otros[
                            otros["numero_limpio"]
                            == seleccionado
                        ].dropna(
                            subset=["latitud", "longitud"]
                        )

                        if not cv.empty:
                            mapa = folium.Map(
                                location=[
                                    cv["latitud"].mean(),
                                    cv["longitud"].mean(),
                                ],
                                zoom_start=13,
                            )

                            AntPath(
                                locations=cv[
                                    ["latitud", "longitud"]
                                ].values.tolist(),
                                color="green",
                                weight=5,
                                tooltip="Referencia",
                            ).add_to(mapa)

                            if not cs.empty:
                                AntPath(
                                    locations=cs[
                                        ["latitud", "longitud"]
                                    ].values.tolist(),
                                    color="red",
                                    weight=4,
                                    tooltip="Coincidencia",
                                ).add_to(mapa)
                        else:
                            mapa = folium.Map(
                                location=[4.57, -74.29],
                                zoom_start=5,
                            )

                        paquete = zip_bytes(
                            {
                                f"Persecucion_{victim}_vs_{seleccionado}.xlsx":
                                    xlsx,
                                f"Rutas_{victim}_vs_{seleccionado}.html":
                                    mapa_html_bytes(mapa),
                            }
                        )

                        st.download_button(
                            "⬇️ Descargar paquete ZIP",
                            paquete,
                            file_name=(
                                f"Evidencia_Persecucion_"
                                f"{seleccionado}.zip"
                            ),
                            mime="application/zip",
                            key="co_download",
                        )


# ============================================================
# TAB 5 - PATRÓN DE VIDA
# ============================================================
with tabs[4]:
    st.header("👤 Perfilamiento de Rutinas y Zonas de Pernocta")
    st.write(
        "Analiza el comportamiento por franjas horarias y permite "
        "exportar los registros geográficos a Google Earth."
    )

    objetivo = st.text_input(
        "Victimario / objetivo",
        placeholder="Ej: 3157658841",
        key="rut_objetivo",
    )

    top_rutina = st.selectbox(
        "Mostrar por franja",
        ["Top 10", "Top 5", "Top 3", "Top 1"],
        index=2,
        key="rut_top",
    )

    solo_colombia_rut = st.checkbox(
        "Ignorar plataformas: solo móviles Colombia",
        value=True,
        key="rut_colombia",
    )

    fi, hi, ff, hf = filtros_ui(df_master, "rut")

    if st.button(
        "👤 Generar perfil de vida",
        type="primary",
        key="rut_btn",
    ):
        if not re.fullmatch(r"\d{10}", objetivo.strip()):
            st.error("Ingresa el número a 10 dígitos.")
        else:
            d = df_master.dropna(
                subset=["numero_limpio", "celda", "fecha_limpia"]
            ).copy()

            d = aplicar_filtro_fecha(d, fi, hi, ff, hf)

            if solo_colombia_rut:
                d = d[
                    d["numero_limpio"].str.match(
                        r"^3\d{9}$", na=False
                    )
                ]

            obj = d[d["numero_limpio"] == objetivo.strip()].copy()

            if obj.empty:
                st.warning(
                    f"El objetivo {objetivo} no presenta tráfico "
                    "en los filtros indicados."
                )
            else:
                obj["hora"] = obj["fecha_limpia"].dt.hour

                def clasificar(h):
                    if h >= 22 or h <= 2:
                        return "1. NOCHE (Pernocta: 22:00 - 02:59)"
                    if 3 <= h <= 8:
                        return "2. MAÑANA/MADRUGADA (03:00 - 08:59)"
                    return "3. DÍA (Actividad: 09:00 - 21:59)"

                obj["franja"] = obj["hora"].apply(clasificar)

                col_desc = (
                    "descripcion"
                    if "descripcion" in obj.columns
                    else (
                        "direccion"
                        if "direccion" in obj.columns
                        else "celda"
                    )
                )

                rutinas = (
                    obj.groupby(["franja", "celda"])
                    .agg(
                        nombre_lugar=(col_desc, "first"),
                        visitas=("celda", "count"),
                        primer_registro=(
                            "fecha_limpia",
                            "min",
                        ),
                        ultimo_registro=(
                            "fecha_limpia",
                            "max",
                        ),
                        latitud=("latitud", "mean"),
                        longitud=("longitud", "mean"),
                    )
                    .reset_index()
                    .sort_values(
                        ["franja", "visitas"],
                        ascending=[True, False],
                    )
                )

                n_top = int(
                    top_rutina.replace("Top ", "")
                )
                top_df = (
                    rutinas.groupby("franja", group_keys=False)
                    .head(n_top)
                    .reset_index(drop=True)
                )

                def top_text(franja):
                    q = top_df[top_df["franja"] == franja]
                    if q.empty:
                        return "Sin actividad registrada."
                    r = q.iloc[0]
                    return (
                        f"Celda **{r['celda']}** "
                        f"({r['nombre_lugar']}) con "
                        f"**{r['visitas']} conexiones**."
                    )

                st.info(
                    f"""
                    **💡 Sinopsis del patrón de vida**

                    - 🌙 **Pernocta:** {top_text(
                        "1. NOCHE (Pernocta: 22:00 - 02:59)"
                    )}
                    - 🌅 **Madrugada/mañana:** {top_text(
                        "2. MAÑANA/MADRUGADA (03:00 - 08:59)"
                    )}
                    - ☀️ **Día:** {top_text(
                        "3. DÍA (Actividad: 09:00 - 21:59)"
                    )}
                    """
                )

                st.subheader(f"📊 Desglose de rutinas ({top_rutina})")
                st.dataframe(
                    top_df[
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

                dm = top_df.dropna(
                    subset=["latitud", "longitud"]
                )

                if not dm.empty:
                    mapa = folium.Map(
                        location=[
                            dm["latitud"].mean(),
                            dm["longitud"].mean(),
                        ],
                        zoom_start=13,
                    )

                    colores = {
                        "1. NOCHE (Pernocta: 22:00 - 02:59)": "darkblue",
                        "2. MAÑANA/MADRUGADA (03:00 - 08:59)": "lightblue",
                        "3. DÍA (Actividad: 09:00 - 21:59)": "orange",
                    }

                    for _, r in dm.iterrows():
                        color = colores.get(
                            r["franja"], "blue"
                        )

                        popup = (
                            f"<b>Franja:</b> {r['franja']}<br>"
                            f"<b>Celda:</b> {r['celda']}<br>"
                            f"<b>Lugar:</b> {r['nombre_lugar']}<br>"
                            f"<b>Visitas:</b> {r['visitas']}<br>"
                            f"<b>Inicio:</b> {r['primer_registro']}<br>"
                            f"<b>Fin:</b> {r['ultimo_registro']}"
                        )

                        folium.Marker(
                            [
                                r["latitud"],
                                r["longitud"],
                            ],
                            popup=folium.Popup(
                                popup,
                                max_width=350,
                            ),
                            tooltip=r["franja"],
                            icon=folium.Icon(color=color),
                        ).add_to(mapa)

                    st_folium(
                        mapa,
                        use_container_width=True,
                        height=550,
                        key="rut_map",
                    )

                    # KML
                    kml = simplekml.Kml()

                    for _, r in dm.iterrows():
                        p = kml.newpoint(
                            name=f"{r['franja']} - Celda {r['celda']}"
                        )
                        p.coords = [
                            (float(r["longitud"]), float(r["latitud"]))
                        ]
                        p.description = (
                            f"Objetivo: {objetivo}<br>"
                            f"Franja: {r['franja']}<br>"
                            f"Celda: {r['celda']}<br>"
                            f"Visitas: {r['visitas']}<br>"
                            f"Inicio: {r['primer_registro']}<br>"
                            f"Fin: {r['ultimo_registro']}"
                        )

                    st.download_button(
                        "🌎 Descargar KML para Google Earth",
                        kml.kml(),
                        file_name=f"Perfil_Vida_{objetivo}.kml",
                        mime="application/vnd.google-earth.kml+xml",
                        key="rut_kml",
                    )
                else:
                    st.warning(
                        "Los registros no cuentan con coordenadas "
                        "para representar el perfil en el mapa."
                    )
