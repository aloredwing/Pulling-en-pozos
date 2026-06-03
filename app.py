import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

st.set_page_config(
    page_title="Análisis de Pulling por Pozo",
    layout="wide"
)

st.title("Análisis de pozos con caída a Pulling")
st.write("Carga tu Excel, selecciona la batería y analiza los pozos que más veces caen a Pulling desde la primera fecha detectada en la data.")

archivo = st.file_uploader("Carga tu archivo Excel", type=["xlsx"])


def limpiar_columnas(df):
    df.columns = df.columns.astype(str).str.strip()
    return df


def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def buscar_columna(df, posibles_nombres):
    columnas = list(df.columns)
    for nombre in posibles_nombres:
        for col in columnas:
            if col.strip().lower() == nombre.strip().lower():
                return col
    return None


def convertir_fecha_segura(serie):
    """
    Convierte fechas aunque vengan como:
    31/05/2026
    2026-05-31
    20260531
    número serial de Excel
    """

    serie_original = serie.copy()

    fecha_1 = pd.to_datetime(
        serie_original,
        errors="coerce",
        dayfirst=True
    )

    numerica = pd.to_numeric(serie_original, errors="coerce")

    fecha_excel = pd.to_datetime(
        numerica,
        errors="coerce",
        unit="D",
        origin="1899-12-30"
    )

    fecha_yyyymmdd = pd.to_datetime(
        serie_original.astype(str).str.replace(".0", "", regex=False),
        format="%Y%m%d",
        errors="coerce"
    )

    fecha_final = fecha_1.copy()

    fecha_final = fecha_final.fillna(fecha_excel)
    fecha_final = fecha_final.fillna(fecha_yyyymmdd)

    fecha_final = fecha_final.where(
        (fecha_final.dt.year >= 2000) & (fecha_final.dt.year <= 2100)
    )

    return fecha_final


if archivo is not None:

    df = pd.read_excel(archivo, sheet_name=0)
    df = limpiar_columnas(df)

    col_bateria = buscar_columna(df, ["Bateria", "Batería", "Battery"])
    col_pozo = buscar_columna(df, ["Pozo", "Well"])
    col_fecha = buscar_columna(df, ["FechaDif", "Fecha", "Fecha Dif", "Fecha de intervención", "FechaIntervencion"])
    col_servicio = buscar_columna(df, ["TServicio", "Servicio", "Tipo Servicio", "Tipo de Servicio"])
    col_causa = buscar_columna(df, ["Causa"])
    col_motivo = buscar_columna(df, ["Motivo"])
    col_falla = buscar_columna(df, ["Falla"])
    col_ubicacion = buscar_columna(df, ["Ubicación", "Ubicacion"])
    col_cfalla = buscar_columna(df, ["CFalla", "C Falla", "Causa Final", "CausaFalla", "CFalla "])

    columnas_minimas = {
        "Batería": col_bateria,
        "Pozo": col_pozo,
        "Fecha": col_fecha,
        "Servicio": col_servicio
    }

    faltantes = [nombre for nombre, col in columnas_minimas.items() if col is None]

    if faltantes:
        st.error("Faltan columnas necesarias en el Excel:")
        st.write(faltantes)
        st.write("Columnas encontradas en tu archivo:")
        st.write(list(df.columns))
        st.stop()

    df[col_bateria] = df[col_bateria].apply(limpiar_texto)
    df[col_pozo] = df[col_pozo].apply(limpiar_texto)
    df[col_servicio] = df[col_servicio].apply(limpiar_texto)

    for col in [col_causa, col_motivo, col_falla, col_ubicacion, col_cfalla]:
        if col is not None:
            df[col] = df[col].apply(limpiar_texto)

    df[col_fecha] = convertir_fecha_segura(df[col_fecha])
    df = df.dropna(subset=[col_fecha])

    if df.empty:
        st.error("No se detectaron fechas válidas en el Excel.")
        st.stop()

    df["Año"] = df[col_fecha].dt.year
    df["Mes"] = df[col_fecha].dt.month
    df["Periodo"] = df[col_fecha].dt.strftime("%Y-%m")
    df["Fecha_Texto"] = df[col_fecha].dt.strftime("%d/%m/%Y")

    fecha_minima = df[col_fecha].min().date()
    fecha_maxima = df[col_fecha].max().date()

    st.success(f"Rango detectado en la data: {fecha_minima.strftime('%d/%m/%Y')} al {fecha_maxima.strftime('%d/%m/%Y')}")

    df_pulling = df[
        df[col_servicio].str.contains("pulling", case=False, na=False)
    ].copy()

    if df_pulling.empty:
        st.warning("No se encontraron registros de Pulling en la columna de servicio.")
        st.stop()

    st.sidebar.header("Filtros")

    rango_fechas = st.sidebar.date_input(
        "Rango de fechas para analizar",
        value=(fecha_minima, fecha_maxima),
        min_value=fecha_minima,
        max_value=fecha_maxima
    )

    if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
        fecha_inicio, fecha_fin = rango_fechas
    else:
        fecha_inicio = fecha_minima
        fecha_fin = fecha_maxima

    df_pulling = df_pulling[
        (df_pulling[col_fecha].dt.date >= fecha_inicio) &
        (df_pulling[col_fecha].dt.date <= fecha_fin)
    ].copy()

    if df_pulling.empty:
        st.warning("No hay registros de Pulling dentro del rango de fechas seleccionado.")
        st.stop()

    baterias = sorted(df_pulling[col_bateria].dropna().unique())

    bateria_sel = st.sidebar.selectbox(
        "Selecciona batería",
        baterias
    )

    df_bat = df_pulling[df_pulling[col_bateria] == bateria_sel].copy()

    años = sorted(df_bat["Año"].dropna().unique())

    años_sel = st.sidebar.multiselect(
        "Selecciona años",
        años,
        default=años
    )

    df_bat = df_bat[df_bat["Año"].isin(años_sel)].copy()

    if df_bat.empty:
        st.warning("No hay registros para la batería, años y fechas seleccionadas.")
        st.stop()

    st.subheader(f"Batería seleccionada: {bateria_sel}")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Fecha inicial analizada", fecha_inicio.strftime("%d/%m/%Y"))
    col2.metric("Fecha final analizada", fecha_fin.strftime("%d/%m/%Y"))
    col3.metric("Total eventos Pulling", len(df_bat))
    col4.metric("Pozos con Pulling", df_bat[col_pozo].nunique())

    st.divider()

    resumen_pozo = (
        df_bat
        .groupby(col_pozo, as_index=False)
        .agg(
            Veces_Pulling=(col_pozo, "count"),
            Primera_Caida=(col_fecha, "min"),
            Ultima_Caida=(col_fecha, "max")
        )
        .sort_values("Veces_Pulling", ascending=False)
    )

    resumen_pozo["Primera_Caida"] = resumen_pozo["Primera_Caida"].dt.strftime("%d/%m/%Y")
    resumen_pozo["Ultima_Caida"] = resumen_pozo["Ultima_Caida"].dt.strftime("%d/%m/%Y")

    pozo_top = resumen_pozo.iloc[0][col_pozo]
    veces_top = int(resumen_pozo.iloc[0]["Veces_Pulling"])

    st.subheader("Pozo con más caídas a Pulling")

    col5, col6 = st.columns(2)

    col5.metric("Pozo", pozo_top)
    col6.metric("Cantidad de veces", veces_top)

    st.divider()

    st.subheader("Ranking de pozos que más caen a Pulling")

    st.dataframe(
        resumen_pozo,
        use_container_width=True
    )

    fig = px.bar(
        resumen_pozo.head(20),
        x=col_pozo,
        y="Veces_Pulling",
        text="Veces_Pulling",
        title="Top 20 pozos con más eventos de Pulling"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    st.subheader("Cantidad de Pulling por pozo y año")

    resumen_anual = (
        df_bat
        .groupby(["Año", col_pozo], as_index=False)
        .agg(Veces_Pulling=(col_pozo, "count"))
        .sort_values(["Año", "Veces_Pulling"], ascending=[True, False])
    )

    st.dataframe(
        resumen_anual,
        use_container_width=True
    )

    fig2 = px.bar(
        resumen_anual,
        x=col_pozo,
        y="Veces_Pulling",
        color="Año",
        text="Veces_Pulling",
        barmode="group",
        title="Eventos de Pulling por pozo y año"
    )

    st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    st.subheader(f"Detalle del pozo que más cayó a Pulling: {pozo_top}")

    columnas_detalle = [
        "Fecha_Texto",
        "Año",
        "Periodo",
        col_bateria,
        col_pozo,
        col_servicio
    ]

    for col in [col_causa, col_motivo, col_falla, col_ubicacion, col_cfalla]:
        if col is not None:
            columnas_detalle.append(col)

    detalle_top = (
        df_bat[df_bat[col_pozo] == pozo_top]
        .sort_values(col_fecha)
        [columnas_detalle]
    )

    st.dataframe(
        detalle_top,
        use_container_width=True
    )

    st.divider()

    st.subheader("Buscar cualquier pozo de la batería")

    pozos = sorted(df_bat[col_pozo].dropna().unique())

    pozo_sel = st.selectbox(
        "Selecciona pozo",
        pozos
    )

    detalle_pozo = (
        df_bat[df_bat[col_pozo] == pozo_sel]
        .sort_values(col_fecha)
        [columnas_detalle]
    )

    st.metric(
        f"Veces que el pozo {pozo_sel} cayó a Pulling",
        len(detalle_pozo)
    )

    st.dataframe(
        detalle_pozo,
        use_container_width=True
    )

    st.divider()

    st.subheader("Resumen de causas, motivos y fallas del pozo seleccionado")

    columnas_resumen_motivo = []

    for col in [col_causa, col_motivo, col_falla, col_ubicacion, col_cfalla]:
        if col is not None:
            columnas_resumen_motivo.append(col)

    if columnas_resumen_motivo:
        resumen_motivos = (
            df_bat[df_bat[col_pozo] == pozo_sel]
            .groupby(columnas_resumen_motivo, as_index=False)
            .agg(Veces=(col_pozo, "count"))
            .sort_values("Veces", ascending=False)
        )

        st.dataframe(
            resumen_motivos,
            use_container_width=True
        )
    else:
        st.info("No se encontraron columnas de causa, motivo, falla o ubicación.")

    st.divider()

    st.subheader("Descargar análisis filtrado")

    df_exportar = df_bat[columnas_detalle].copy()

    csv = df_exportar.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="Descargar detalle filtrado en CSV",
        data=csv,
        file_name=f"analisis_pulling_{bateria_sel}.csv",
        mime="text/csv"
    )

else:
    st.info("Primero carga tu archivo Excel.")
