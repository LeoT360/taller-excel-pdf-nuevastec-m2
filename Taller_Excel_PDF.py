from openpyxl import load_workbook
from tabulate import tabulate
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

wb = load_workbook('ventas_tecnologia_v2.xlsx', data_only=True)
sheet = wb.active

headers = [cell.value for cell in sheet[1]]

data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(list(row))

alignments = ("center", "left", "center", "right", "right")

print(f"\n==================== REPORTE DE VENTAS EN CONSOLA ({len(data)} PRODUCTOS) ====================")
print(tabulate( data,  headers=headers,   tablefmt="fancy_grid",  colalign=alignments))

doc = SimpleDocTemplate(  "reporte_ejecutivo_ventas.pdf",  pagesize=letter)

styles = getSampleStyleSheet()

title_style = ParagraphStyle(
    "TitleStyle", parent=styles["Title"],
    fontName="Helvetica-Bold", fontSize=18,
    alignment=TA_CENTER, textColor=colors.HexColor("#17365D")
)
body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9)
header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.whitesmoke, alignment=1)

story = [
    Paragraph("Informe Ejecutivo de Ventas Tecno", title_style),
    Paragraph(f"<b>Total de Productos Procesados:</b> {len(data)}", body_style),
    Spacer(1, 12)
]

header_paragraphs = [Paragraph(f"<b>{h}</b>", header_style) for h in headers]
data_paragraphs = [
    [
        Paragraph(str(r[0]), ParagraphStyle('C', parent=body_style, alignment=1)),
        Paragraph(str(r[1]), ParagraphStyle('L', parent=body_style, alignment=0)),
        Paragraph(str(r[2]), ParagraphStyle('C', parent=body_style, alignment=1)),
        Paragraph(f"${r[3]:,.2f}", ParagraphStyle('R', parent=body_style, alignment=2)),
        Paragraph(str(r[4]), ParagraphStyle('R', parent=body_style, alignment=2))
    ]
    for r in data
]

col_widths = [40, 160, 100, 80, 80]
pdf_table = Table([header_paragraphs] + data_paragraphs, colWidths=col_widths, repeatRows=1)
pdf_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1),
     [colors.white, colors.lightgrey])
]))

story.append(pdf_table)
doc.build(story)
print("\n[ÉXITO] Archivo 'reporte_ejecutivo_ventas.pdf' generado correctamente.")