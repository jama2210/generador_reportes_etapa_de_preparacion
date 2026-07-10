import tempfile

import pandas as pd
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

if archivo:

    df = leer_excel(archivo)

    total_registros = len(df)

    total_asesores = df["Nombre"].nunique()

    total_directa = len(
        df[df["Tipo Asesoría"] == "Directa EE"]
    )

    total_red = len(
        df[df["Tipo Asesoría"] == "Red EE"]
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
        "Resumen de generación"
    )

    st.dataframe(
        resumen,
        use_container_width=True
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