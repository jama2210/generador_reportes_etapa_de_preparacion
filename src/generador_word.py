from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from src.utils import limpiar_valor

NO_RESPONDIDO = "(No respondido en el formulario)"


# ==================================================
# ESTILOS
# ==================================================

def configurar_estilos(doc):

    estilos = [
        "Normal",
        "Heading 1",
        "Heading 2",
        "Heading 3"
    ]

    for nombre_estilo in estilos:

        style = doc.styles[nombre_estilo]

        style.font.name = "Aptos"

        if nombre_estilo == "Normal":
            style.font.size = Pt(11)


# ==================================================
# TOC
# ==================================================

def insertar_toc(paragraph):

    run = paragraph.add_run()

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')

    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = r'TOC \o "1-3" \h \z \u'

    fld_separate = OxmlElement('w:fldChar')
    fld_separate.set(qn('w:fldCharType'), 'separate')

    text = OxmlElement('w:t')
    text.text = "Actualice el índice al abrir el documento (Ctrl+A y F9)"

    fld_separate.append(text)

    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_separate)
    run._r.append(fld_end)


# ==================================================
# ENCABEZADO
# ==================================================

def agregar_encabezado(doc):

    seccion = doc.sections[0]

    header = seccion.header

    parrafo = header.paragraphs[0]

    run = parrafo.add_run()

    try:

        run.add_picture(
            "assets/ApoyoalaMejora.png",
            width=Inches(1.5)
        )

    except:
        pass


# ==================================================
# CAMPOS
# ==================================================

def agregar_campo(tabla, nombre, valor):

    fila = tabla.add_row().cells

    fila[0].text = str(nombre)

    valor = limpiar_valor(valor)

    p = fila[1].paragraphs[0]

    if valor == NO_RESPONDIDO:

        run = p.add_run(valor)

        run.bold = True

        run.font.color.rgb = RGBColor(
            255,
            102,
            0
        )

    else:

        p.add_run(str(valor))


# ==================================================
# GENERADOR
# ==================================================

def generar_word(df, salida):

    doc = Document()

    configurar_estilos(doc)

    agregar_encabezado(doc)

    # =====================================
    # DATOS PORTADA
    # =====================================

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

    # =====================================
    # TITULO
    # =====================================

    titulo = doc.add_paragraph()

    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    run = titulo.add_run(
        "INFORME INDIVIDUAL ETAPA DE PREPARACIÓN"
    )

    run.bold = True
    run.font.name = "Aptos"
    run.font.size = Pt(18)

    doc.add_paragraph("")

    # =====================================
    # TABLA PORTADA
    # =====================================

    tabla_portada = doc.add_table(
        rows=4,
        cols=2
    )

    tabla_portada.style = "Table Grid"

    tabla_portada.cell(0, 0).text = "Región"
    tabla_portada.cell(0, 1).text = region

    tabla_portada.cell(1, 0).text = "DEPROV"
    tabla_portada.cell(1, 1).text = deprov

    tabla_portada.cell(2, 0).text = "Modalidad"
    tabla_portada.cell(2, 1).text = modalidad

    tabla_portada.cell(3, 0).text = "Asesor"
    tabla_portada.cell(3, 1).text = asesor

    doc.add_page_break()

    # =====================================
    # INDICE REAL WORD
    # =====================================

    doc.add_heading(
        "Índice",
        level=1
    )

    insertar_toc(
        doc.add_paragraph()
    )

    doc.add_page_break()

    # =====================================
    # REGISTROS
    # =====================================

    for _, row in df.iterrows():

        nombre_registro = limpiar_valor(
            row["Nombre Asesoría"]
        )

        doc.add_heading(
            f"Registro asociado a: {nombre_registro}",
            level=1
        )

        # ----------------------------------

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
            row.get("Correo electrónico", "")
        )

        agregar_campo(
            tabla,
            "Nombre",
            row.get("Nombre", "")
        )

        agregar_campo(
            tabla,
            "Supervisor",
            row.get("Supervisor", "")
        )

        agregar_campo(
            tabla,
            "Director",
            row.get("Nombre Director", "")
        )

        agregar_campo(
            tabla,
            "Tipo Asesoría",
            row.get("Tipo Asesoría", "")
        )

        # ----------------------------------

        doc.add_heading(
            "Brechas Críticas Lectura",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row.get(
                    "Indique las brechas críticas identificadas en lectura",
                    ""
                )
            )
        )

        # ----------------------------------

        doc.add_heading(
            "Brechas Críticas Matemática",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row.get(
                    "Indique las brechas críticas identificadas en matemática",
                    ""
                )
            )
        )

        # ----------------------------------

        doc.add_heading(
            "Hallazgos DIA",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row.get(
                    "Principales hallazgos del análisis DIA (Socioemocional / Académico)",
                    ""
                )
            )
        )

        # ----------------------------------

        doc.add_heading(
            "Fortalezas",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row.get(
                    "Fortalezas del Establecimiento (prácticas exitosas, recursos o liderazgos consolidados)",
                    ""
                )
            )
        )

        # ----------------------------------

        doc.add_heading(
            "Oportunidades de Mejora",
            level=2
        )

        doc.add_paragraph(
            limpiar_valor(
                row.get(
                    "Oportunidades de mejora del Establecimiento (nudos críticos o debilidades a subsanar)",
                    ""
                )
            )
        )

        # ----------------------------------

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
            row.get(
                "El PME vigente contempla acciones específicas para el abordaje de la asignatura de lenguaje",
                ""
            )
        )

        agregar_campo(
            tabla_pme,
            "Descripción Lenguaje",
            row.get(
                "Indique una breve descripción de la acción",
                ""
            )
        )

        agregar_campo(
            tabla_pme,
            "Acciones Matemática",
            row.get(
                "El PME vigente contempla acciones específicas para el abordaje de la asignatura de matemática",
                ""
            )
        )

        agregar_campo(
            tabla_pme,
            "Descripción Matemática",
            row.get(
                "Indique una breve descripción de la acción2",
                ""
            )
        )

        agregar_campo(
            tabla_pme,
            "Observaciones PME",
            row.get(
                "Observaciones / Sugerencias de ajuste al PME",
                ""
            )
        )

        doc.add_page_break()

    doc.save(salida)