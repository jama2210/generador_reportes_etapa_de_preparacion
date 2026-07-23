import tempfile

import streamlit as st

from src.lector_excel import leer_excel
from src.generador_zip import generar_zip


st.set_page_config(
    page_title="Generador Informes MINEDUC",
    page_icon="📄",
    layout="wide"
)

st.title(
    "📄 Generador De Informes Apoyo A La Mejora"
)

archivo = st.file_uploader(
    "Seleccione archivo Excel",
    type=["xlsx"]
)

# Solo procesar si existe archivo
if archivo:

    df = leer_excel(archivo)

    st.subheader("Diagnóstico")

    st.write("Columnas detectadas:")

    st.write(df.columns.tolist())

    total_registros = len(df)

    # Validación columnas obligatorias
    columnas_obligatorias = [
        "Nombre",
        "Indique su región",
        "Deprov",
        "Tipo Asesoría"
    ]

    faltantes = [
        c for c in columnas_obligatorias
        if c not in df.columns
    ]

    if faltantes:

        st.error(
            f"Faltan columnas requeridas: {', '.join(faltantes)}"
        )

        st.stop()

    total_asesores = df["Nombre"].nunique()

    total_directa = len(
        df[
            df["Tipo Asesoría"]
            .astype(str)
            .str.strip()
            == "Directa EE"
        ]
    )

    total_red = len(
        df[
            df["Tipo Asesoría"]
            .astype(str)
            .str.strip()
            == "Red EE"
        ]
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Registros",
        total_registros
    )

    c2.metric(
        "Profesionales",
        total_asesores
    )

    c3.metric(
        "Directa EE",
        total_directa
    )

    c4.metric(
        "Red EE",
        total_red
    )

    st.divider()

    resumen = (
        df.groupby(
            [
                "Nombre",
                "Indique su región",
                "Deprov",
                "Tipo Asesoría"
            ]
        )
        .size()
        .reset_index(name="Registros")
    )

    st.subheader(
        "Resumen De Generación"
    )

    st.dataframe(
        resumen,
        use_container_width=True
    )

    st.info(
        f"Se generarán {len(resumen)} documentos Word y serán empaquetados en un archivo ZIP."
    )

    if st.button(
        "🚀 Generar Informes"
    ):

        with st.spinner(
            "Generando documentos..."
        ):

            temp_zip = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".zip"
            )

            generar_zip(
                df,
                temp_zip.name
            )

            with open(
                temp_zip.name,
                "rb"
            ) as f:

                st.download_button(
                    label="📥 Descargar ZIP",
                    data=f,
                    file_name="Informes_Mineduc.zip",
                    mime="application/zip"
                )

        st.success(
            "Informes generados correctamente."
        )