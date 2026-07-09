import streamlit as st
import tempfile

from src.lector_excel import leer_excel
from src.generador_word import generar_word

st.set_page_config(
    page_title="Generador Informes MINEDUC",
    layout="wide"
)

st.title("Generador de Informes MINEDUC Etapa de Preparación")

archivo = st.file_uploader(
    "Seleccione Excel",
    type=["xlsx", "xls"]
)

if archivo:

    df = leer_excel(archivo)

    st.success(
        f"Registros encontrados: {len(df)}"
    )

    st.dataframe(
        df[
            [
                "Nombre Asesoría",
                "Tipo Asesoría"
            ]
        ]
    )

    if st.button(
        "Generar Word"
    ):

        salida = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".docx"
        )

        generar_word(
            df,
            salida.name
        )

        with open(
            salida.name,
            "rb"
        ) as f:

            st.download_button(
                "Descargar Word",
                f,
                file_name="Informe_Mineduc.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )