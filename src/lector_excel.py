import pandas as pd


def leer_excel(archivo):

    df = pd.read_excel(archivo)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    renombres = {
        "Tipo de Asesoría.1": "Tipo Asesoría",
        "Tipo de Asesoría": "Tipo Asesoría",
        "Tipo Asesoria": "Tipo Asesoría",
        "Tipo Asesoria": "Tipo Asesoría",
    }

    for origen, destino in renombres.items():

        if origen in df.columns:

            df.rename(
                columns={
                    origen: destino
                },
                inplace=True
            )

    return df