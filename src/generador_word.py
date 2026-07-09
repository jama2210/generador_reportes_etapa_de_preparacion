from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from src.utils import limpiar_valor


NO_RESPONDIDO = "(No respondido en el formulario)"


def agregar_campo(tabla, nombre, valor):

    fila = tabla.add_row().cells

    fila[0].text = nombre

    valor = limpiar_valor(valor)

    p = fila[1].paragraphs[0]

    if valor == NO_RESPONDIDO:

        run = p.add_run(valor)

        run.bold = True
        run.font.color.rgb = RGBColor(255, 102, 0)

    else:

        p.add_run(str(valor))


def generar_word(df, salida):

    doc = Document()

    #################################
    # PORTADA
    #################################

    titulo = doc.add_heading(
        "Informe Individual de Asesoría MINEDUC",
        level=0
    )

    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    region = limpiar_valor(
        df.iloc[0]["Indique su región"]
    )

    deprov = limpiar_valor(
        df.iloc[0]["Deprov"]
    )

    modalidad = limpiar_valor(
        df.iloc[0]["Tipo Asesoría"]
    )

    asesor = limpiar_valor(
        df.iloc[0]["Nombre"]
    )

    doc.add_paragraph(f"Región: {region}")
    doc.add_paragraph(f"DEPROV: {deprov}")
    doc.add_paragraph(f"Modalidad: {modalidad}")
    doc.add_paragraph("")
    doc.add_paragraph(f"Asesor: {asesor}")

    doc.add_page_break()

    #################################
    # INDICE
    #################################

    doc.add_heading("Índice", level=1)

    for _, row in df.iterrows():

        nombre = limpiar_valor(
            row["Nombre Asesoría"]
        )

        doc.add_paragraph(nombre)

    doc.add_page_break()

    #################################
    # REGISTROS
    #################################

    for _, row in df.iterrows():

        nombre_registro = limpiar_valor(
            row["Nombre Asesoría"]
        )

        doc.add_heading(
            f"Registro asociado a: {nombre_registro}",
            level=1
        )

        #################################################

        doc.add_heading(
            "Información General",
            level=2
        )

        tabla = doc.add_table(
            rows=1,
            cols=2
        )

        tabla.style = "Light Grid Accent 1"

        cab = tabla.rows[0].cells
        cab[0].text = "Campo"
        cab[1].text = "Valor"

        agregar_campo(
            tabla,
            "Correo",
            row["Correo electrónico"]
        )

        agregar_campo(
            tabla,
            "Nombre",
            row["Nombre"]
        )

        agregar_campo(
            tabla,
            "Supervisor",
            row["Supervisor"]
        )

        agregar_campo(
            tabla,
            "Director",
            row["Nombre Director"]
        )

        agregar_campo(
            tabla,
            "Tipo Asesoría",
            row["Tipo Asesoría"]
        )

        #################################################

        doc.add_heading(
            "Brechas Críticas Lectura",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row[
                    "Indique las brechas críticas identificadas en lectura"
                ]
            )
        )

        #################################################

        doc.add_heading(
            "Brechas Críticas Matemática",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row[
                    "Indique las brechas críticas identificadas en matemática"
                ]
            )
        )

        #################################################

        doc.add_heading(
            "Hallazgos DIA",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row[
                    "Principales hallazgos del análisis DIA (Socioemocional / Académico)"
                ]
            )
        )

        #################################################

        doc.add_heading(
            "Fortalezas",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row[
                    "Fortalezas del Establecimiento (prácticas exitosas, recursos o liderazgos consolidados)"
                ]
            )
        )

        #################################################

        doc.add_heading(
            "Oportunidades de Mejora",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row[
                    "Oportunidades de mejora del Establecimiento (nudos críticos o debilidades a subsanar)"
                ]
            )
        )

        #################################################

        doc.add_heading(
            "PME",
            level=2
        )

        tabla_pme = doc.add_table(
            rows=1,
            cols=2
        )

        tabla_pme.style = "Table Grid"

        cab = tabla_pme.rows[0].cells

        cab[0].text = "Elemento"
        cab[1].text = "Valor"

        agregar_campo(
            tabla_pme,
            "Acciones Lenguaje",
            row[
                "El PME vigente contempla acciones específicas para el abordaje de la asignatura de lenguaje"
            ]
        )

        agregar_campo(
            tabla_pme,
            "Descripción Lenguaje",
            row[
                "Indique una breve descripción de la acción"
            ]
        )

        agregar_campo(
            tabla_pme,
            "Acciones Matemática",
            row[
                "El PME vigente contempla acciones específicas para el abordaje de la asignatura de matemática"
            ]
        )

        agregar_campo(
            tabla_pme,
            "Descripción Matemática",
            row[
                "Indique una breve descripción de la acción2"
            ]
        )

        agregar_campo(
            tabla_pme,
            "Observaciones",
            row[
                "Observaciones / Sugerencias de ajuste al PME"
            ]
        )

        doc.add_page_break()

    doc.save(salida)