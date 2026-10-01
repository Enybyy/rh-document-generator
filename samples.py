"""Synthetic examples; no employee data."""
from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook

ROWS = [
    ["00000001", "Valeria Campos Medina", "Analista administrativa", "Lima", "01/10/2026", "31/12/2026", 2800],
    ["00000002", "Mateo Salazar Torres", "Asistente de operaciones", "Arequipa", "01/10/2026", "31/12/2026", 2400],
    ["00000003", "Lucía Paredes Ríos", "Coordinadora de proyecto", "Cusco", "01/10/2026", "31/12/2026", 3600],
]
HEADERS = ["DNI", "NOMBRE", "CARGO", "CIUDAD", "FECHA INICIO", "FECHA FIN", "PAGO"]


def excel():
    book = Workbook()
    sheet = book.active
    sheet.title = "Personal de muestra"
    sheet.append(HEADERS)
    for row in ROWS:
        sheet.append(row)
    for col in "ABCDEFG":
        sheet.column_dimensions[col].width = 29 if col != "B" else 36
    sheet.freeze_panes = "A2"
    stream = BytesIO()
    book.save(stream)
    return stream.getvalue()


def template():
    """Two-page editable reference agreement; no legal validation is implied."""
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2.2)
    section.left_margin = section.right_margin = Cm(2.4)
    section.header_distance = section.footer_distance = Cm(1.1)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Times New Roman", Pt(11)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 1.2
    normal.paragraph_format.space_after = Pt(9)
    normal.paragraph_format.widow_control = True
    for name, size in [("Title", 19), ("Heading 1", 12)]:
        style = doc.styles[name]
        style.font.name, style.font.size = "Times New Roman", Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.bold = True
        style.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 0)
        style.paragraph_format.space_after = Pt(7)
        style.paragraph_format.keep_with_next = True
    header = section.header.paragraphs[0]
    header.text = "Contrato de servicios · [NOMBRE]"
    header.runs[0].font.size = Pt(8)
    header.runs[0].font.color.rgb = RGBColor.from_string("666666")
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = footer.add_run("Modelo de referencia · Datos de ejemplo     Página ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string("666666")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    footer.add_run(" de 2").font.size = Pt(8)

    def paragraph(text):
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return p

    doc.add_paragraph("Contrato de prestación de servicios", "Title")
    subtitle = doc.add_paragraph("[CIUDAD] · [FECHA_INICIO]")
    subtitle.paragraph_format.space_after = Pt(16)
    subtitle.runs[0].italic = True
    doc.add_heading("Partes del acuerdo", level=1)
    paragraph("Intervienen la empresa contratante, cuya razón social, RUC, domicilio y "
              "representante se completarán antes de la firma, y [NOMBRE], identificado/a "
              "con DNI [DNI], con domicilio por completar en [CIUDAD], en adelante el/la prestador/a.")
    doc.add_heading("01 Objeto del servicio", level=1)
    paragraph("El/la prestador/a desarrollará el servicio de [CARGO]. El alcance, los "
              "entregables y los criterios de aceptación se detallarán en el anexo de "
              "trabajo acordado por ambas partes.")
    doc.add_heading("02 Vigencia", level=1)
    paragraph("El periodo previsto comprende desde el [FECHA_INICIO] hasta el [FECHA_FIN]. "
              "Cualquier ampliación o modificación se documentará por escrito.")
    doc.add_heading("03 Honorarios y pago", level=1)
    paragraph("Los honorarios de referencia ascienden a S/ [PAGO]. Las partes completarán "
              "la periodicidad, el calendario de pago, las condiciones de conformidad y "
              "el tratamiento tributario aplicable antes de suscribir este documento.")
    doc.add_page_break()
    doc.add_paragraph("Términos del servicio", "Title")
    doc.add_paragraph("[NOMBRE] · DNI [DNI]")
    doc.add_heading("04 Obligaciones de las partes", level=1)
    paragraph("El/la prestador/a realizará los entregables acordados y comunicará "
              "oportunamente cualquier incidencia. La empresa facilitará la información "
              "necesaria, designará un responsable de revisión y comunicará la conformidad "
              "u observaciones según el calendario que ambas partes establezcan.")
    doc.add_heading("05 Confidencialidad", level=1)
    paragraph("La información y los documentos recibidos para ejecutar el servicio "
              "se utilizarán exclusivamente para su finalidad acordada. Las partes "
              "definirán las medidas de protección, los accesos autorizados y la "
              "devolución o eliminación de materiales al finalizar el servicio.")
    doc.add_heading("06 Conformidad y firmas", level=1)
    paragraph("Tras completar los datos pendientes y revisar las condiciones, las "
              "partes podrán suscribir el documento en [CIUDAD]. "
              "Fecha de firma: ____________________.")
    signatures = doc.add_paragraph()
    signatures.paragraph_format.space_before = Pt(38)
    signatures.paragraph_format.line_spacing = 1.5
    signatures.add_run("______________________________\n").bold = True
    signatures.add_run("Por la empresa contratante\nRepresentante y cargo: ______________________________")
    signatures = doc.add_paragraph()
    signatures.paragraph_format.space_before = Pt(22)
    signatures.paragraph_format.line_spacing = 1.5
    signatures.add_run("______________________________\n").bold = True
    signatures.add_run("[NOMBRE] · DNI [DNI]")
    stream = BytesIO()
    doc.save(stream)
    return stream.getvalue()
