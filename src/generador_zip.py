import os
import tempfile
import zipfile

from src.generador_word import generar_word
from src.utils import nombre_seguro


def generar_zip(df, zip_path):

    temp_dir = tempfile.mkdtemp()

    grupos = df.groupby(
        [
            "Indique su región",
            "Deprov",
            "Tipo Asesoría",
            "Nombre"
        ]
    )

    archivos = []

    for (
        region,
        deprov,
        modalidad,
        asesor
    ), grupo in grupos:

        carpeta = os.path.join(
            temp_dir,
            nombre_seguro(region),
            nombre_seguro(deprov),
            nombre_seguro(modalidad)
        )

        os.makedirs(
            carpeta,
            exist_ok=True
        )

        ruta_docx = os.path.join(
            carpeta,
            f"{nombre_seguro(asesor)}.docx"
        )

        generar_word(
            grupo,
            ruta_docx
        )

        archivos.append(ruta_docx)

    with zipfile.ZipFile(
        zip_path,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        for root, _, files in os.walk(temp_dir):

            for file in files:

                ruta = os.path.join(
                    root,
                    file
                )

                arcname = os.path.relpath(
                    ruta,
                    temp_dir
                )

                zipf.write(
                    ruta,
                    arcname
                )