import os
import win32com.client
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from io import BytesIO
from reportlab.lib.pagesizes import A2, A5
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer
from reportlab.lib.units import mm, cm
from flask import Flask, request, jsonify
import pythoncom




def draw_top_table():
    data_top = [
        ["TABLE OF CONTROL", "", "", "", "", ""],
        ["Station Name:", "MALWAN", "Division:", "PRAYAGRAJ", "CODE", "MWH"],
        ["Station ID:", "37111", "Section:", "NC Railway", "", ""],
    ]

    col_widths_top = [25*mm, 30*mm, 25*mm, 40*mm, 40*mm, 80*mm]
    row_heights_top = [5*mm, 6*mm, 6*mm]

    table_top = Table(data_top, colWidths=col_widths_top, rowHeights=row_heights_top)

    style_top = TableStyle([
        ('GRID', (0,1), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (0, 6), 8),
        ('FONTSIZE', (1,0), (-1,-1), 6),
        ('FONTNAME', (0,1), (0,6), 'Helvetica-Bold'),
        ('FONTNAME', (0,1), (0,2), 'Helvetica-Bold'),
        ('FONTNAME', (2,1), (2,2), 'Helvetica-Bold'),
        ('FONTNAME', (4,1), (4,1), 'Helvetica-Bold'),
        ('SPAN', (0,0), (5,0)),
        ('SPAN', (4,1), (4,2)),
        ('SPAN', (5,1), (5,2)),
        ('BACKGROUND', (0,1), (0,2), colors.lightgrey),
        ('BACKGROUND', (2,1), (2,2), colors.lightgrey),
        ('BACKGROUND', (4,1), (4,2), colors.lightgrey),
    ])

    table_top.setStyle(style_top)
    return table_top


def draw_bottom_tables(current_page, total_pages):
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    files_dir = os.path.join(base_dir, "files")

    sign1_path = os.path.join(files_dir, "sign1.png")
    sign2_path = os.path.join(files_dir, "sign2.png")
    logo_path = os.path.join(files_dir, "logo.png")

    sign1 = Image(sign1_path, width=2*cm, height=1*cm)
    sign2 = Image(sign2_path, width=2*cm, height=1*cm)
    logos = Image(logo_path, width=2*cm, height=1*cm)

    # Adjusted widths to fit within 842 pts (landscape A4)
    toc1_col_widths = [50, 50, 50, 60, 60]     # ~475
    toc2_col_widths = [50, 50, 50]             # ~300
    toc3_col_widths = [40, 60, 12, 65, 40, 23, 23]
    toc1_row_heights = [19, 18, 18, 19, 19, 35, 19, 18]
    toc2_row_heights = [15, 15, 29, 25, 50, 15, 15]
    toc3_row_heights = [9*mm, 5.5*mm, 5.5*mm, 8.5*mm, 8.5*mm, 9*mm, 3*mm, 9*mm]


    # Adjust TOC1
    data_toc1 = [
        ["MALWAN(MWH), PRAYAGRAJ DIVISION, NC RAILWAY", "", "", "", ""],
        ["REF : SIP : SSP.NCR.PRYJ.MWH2.0.0", "", "", "", ""],
        ["REF: RFID_TAG_LAYOUT_MWH_2.0.0", "", "", "", ""],
        ["TABLE OF CONTROL: 37111_MWH_TOC_Ver2_0_0", "", "", "", ""],
        ["", "PREPARED BY", "CHECKED BY", "VERIFIED BY", ""],
        ["SIGN", "", "", "", ""],
        ["NAME", "", "", "", ""],
        ["DESIGNATION", "", "", "", ""]
    ]
    table_toc1 = Table(data_toc1, colWidths=toc1_col_widths, rowHeights=toc1_row_heights)
    table_toc1.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 4),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, 3), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
        ("SPAN", (0, 0), (-1, 0)),
        ("SPAN", (0, 1), (-1, 1)),
        ("SPAN", (0, 2), (-1, 2)),
        ("SPAN", (0, 3), (-1, 3)),
        ("SPAN", (4, 5), (4, 6)),
        ("SPAN", (4, 7), (4, 7)),
        ("TOPPADDING", (0, 0), (-1, -1), 10)
    ]))

    # Adjust TOC2
    data_toc2 = [
        ["Remarks", "", ""],
        ["REVISION HISTORY IS AVAILABLE ON PAGE-12", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["DY.CSTE/P-1/PRYJ", "SSTE/Pr/PRYJ", "SSE/SIG/PRYJ"],
        ["", "", ""],
        ["NOTE:", "", ""]
    ]
    table_toc2 = Table(data_toc2, colWidths=toc2_col_widths, rowHeights=toc2_row_heights)
    table_toc2.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
        ("ALIGN", (0, 0), (-1, 5), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
        ("FONTNAME", (0, 5), (-1, 5), "Helvetica-Bold"),
        ("FONTNAME", (0, 6), (2, 6), "Helvetica-Bold"),
        ("FONTSIZE", (0, 6), (2, 6), 5),
        ("FONTSIZE", (0, 0), (-1, -1), 4),
        ("ALIGN", (0, 6), (0, 6), "LEFT"),
        ("SPAN", (0, 0), (-1, 0)),
        ("SPAN", (0, 1), (-1, 1)),
        ("TOPPADDING", (0, 0), (-1, -1), 10)
    ]))

    # Adjust TOC3
    data_toc3 = [
        ["SIP: D-3067/1", "REF: SIG.PLAN.No.", "", "MALWAN(MWH)", "", "", ""],
        ["", "REVISION", "", "SIGNAL & TELECOMMUNICATION", "", "", ""],
        ["", "DRAWING", "", "KAVACH Indian Railway ATP", "", "", ""],
        ["", "DY.CSTE/D&D/PRYJ", "C\nH\nE\nC\nK\nE\nD", "PRAYAGRAJ DIVISION, NC RAILWAY", "", "", ""],
        ["", "SSTE/D&D/PRYJ", "", "TABLE OF CONTROL:KAVACH_TOC_MWH_4.0.0", "", "", ""],
        ["", "ASTE/D&D/PRYJ", "", "", "KAVACH\nD-3067/1", "", ""],
        ["", "", "", "N.C.R RLY.DRG.NO:", "", "SHEET", "SHEETS"],
        ["", "\nSSE/D&D/PRYJ", "", "", "", str(current_page), str(total_pages)],
    ]
    table_toc3 = Table(data_toc3, colWidths=toc3_col_widths, rowHeights=toc3_row_heights)
    table_toc3.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE', (0,0), (-1,-1), 4),
        ("FONTNAME", (0, 0), (7, 7), "Helvetica-Bold"),
        ('SPAN', (0,0), (0,0)),
        ('SPAN', (0,6), (0,7)),
        ('SPAN', (1,0), (2,0)),
        ('SPAN', (1,1), (2,1)),
        ('SPAN', (1,2), (2,2)),
        ('SPAN', (1,6), (1,7)),
        ('SPAN', (3,0), (6,0)),
        ('SPAN', (3,1), (6,1)),
        ('SPAN', (3,2), (6,2)),
        ('SPAN', (3,3), (6,3)),
        ('SPAN', (3,4), (6,4)),
        ('SPAN', (3,6), (3,6)),
        ('SPAN', (2,3), (2,7)),
        ('SPAN', (4,5), (4,7)),
        ('SPAN', (3,7), (3,7)),
        ('SPAN', (5,6), (5,6)),
        ('SPAN', (6,6), (6,6)),
        ('SPAN', (5,7), (5,7)),
        ('SPAN', (6,7), (6,7)),
        ("TOPPADDING", (0, 0), (-1, -1), 10)
    ]))

    # Combine the three TOCs
    bottom_table = Table([[table_toc1, table_toc2, table_toc3]],
                         colWidths=[sum(toc1_col_widths),
                                    sum(toc2_col_widths, 4),
                                    sum(toc3_col_widths)])
    bottom_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    return bottom_table



def generate_toc_pdf_with_header_footer():
    try:
        pythoncom.CoInitialize()
        input_path = request.form.get("input_path")
        output_path = request.form.get("output_path")

        # === Validate inputs ===
        if not input_path or not os.path.exists(input_path):
            return jsonify({"error": "Invalid or missing Excel path"}), 400
        if not output_path:
            return jsonify({"error": "Missing output path"}), 400

        # === Step 1: Save Excel as PDF ===
        input_path = os.path.abspath(input_path)
        output_pdf_path = os.path.join(os.path.abspath(output_path), "37111_MWH_TOC_Ver2_0_0.pdf")

        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(input_path)

        for sheet in workbook.Worksheets:
            sheet.PageSetup.TopMargin = excel_app.CentimetersToPoints(2.7)
            sheet.PageSetup.BottomMargin = excel_app.CentimetersToPoints(7)
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1.2
            sheet.PageSetup.FitToPagesTall = False

        workbook.ExportAsFixedFormat(0, output_pdf_path)
        workbook.Close(False)
        excel_app.Quit()

        # === Step 2: Generate Overlay Page ===
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=A2)
        width_a2, height_a2 = A2

        bottom_table = draw_bottom_tables(current_page=1, total_pages=6)
        bottom_table.wrapOn(c, width_a2, height_a2)
        bottom_table.drawOn(c, 1.77 * cm, 0.6 * cm)

        top_table = draw_top_table()
        top_table.wrapOn(c, width_a2, height_a2)
        top_table.drawOn(c, 1.77 * cm, 19.2 * cm)

        c.save()
        packet.seek(0)
        overlay_page = PdfReader(packet).pages[0]

        # === Step 3: Read generated PDF and create first page ===
        reader = PdfReader(output_pdf_path)
        total_pages = len(reader.pages) + 1
        writer = PdfWriter()

        from PyPDF2 import PageObject
        from reportlab.platypus import Paragraph, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet

        # Get size of actual PDF page
        first_page_size = reader.pages[0].mediabox
        width = float(first_page_size.width)
        height = float(first_page_size.height)

        # === Create the first page with table ===
        buffer = BytesIO()
        first_canvas = canvas.Canvas(buffer, pagesize=(width, height))

        styles = getSampleStyleSheet()
        styleN = styles["Normal"]
        styleB = styles["Heading4"]
        styleB.fontName = "Helvetica-Bold"

        data = [
            ["SIGNAL & TELECOMMUNICATION", ''],
            ["TRAIN COLLISION AVOIDANCE SYSTEM (TCAS) - KAVACH", ''],
            ["TABLE OF CONTROL (TOC)", ''],
            ['Zone', 'NC Railway'],
            ['Division', 'Prayagraj Division, NC Railway'],
            ['Section', '-'],
            ['Station Name', 'MALWAN'],
            ['Station Code', 'MWH'],
            ['Station ID', '37111'],
            ['REF SIP No.', 'SI-3067/1'],
            ['REf RFID Tag Layout No.', 'RFID_TAG_LAYOUT_MWH_2.0.0'],
            ['TOC Version', '2.0.0']
        ]

        table = Table(data, colWidths=[250, 350], rowHeights=[30]*12)
        table.setStyle(TableStyle([
            ('SPAN', (0,0), (-1,0)),
            ('SPAN', (0,1), (-1,1)),
            ('SPAN', (0,2), (-1,2)),
            ('ALIGN', (0,0), (-1, -1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (-1,2), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,2), 11),
            ('FONTNAME', (0,3), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,3), (-1,-1), 10),
            ('GRID', (0,0), (-1,-1), 0.8, colors.black),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))

        table.wrapOn(first_canvas, width, height)
        table.drawOn(first_canvas, 4 * cm, height - 16 * cm)

        first_canvas.save()
        buffer.seek(0)
        first_page_with_table = PdfReader(buffer).pages[0]
        writer.add_page(first_page_with_table)

        # === Step 4: Merge Overlay with remaining PDF pages ===
        for page in reader.pages:
            page.merge_page(overlay_page)
            writer.add_page(page)

        # === Save final output ===
        with open(output_pdf_path, "wb") as f:
            writer.write(f)

        return jsonify({
            "message": "PDF generated successfully with header, footer, and first page table.",
            "output_pdf_path": output_pdf_path
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500



# def generate_toc_pdf_with_header_footer():
#     try:
#         pythoncom.CoInitialize()
#         input_path = request.form.get("input_path")
#         output_path = request.form.get("output_path")

#         # === Validate inputs ===
#         if not input_path or not os.path.exists(input_path):
#             return jsonify({"error": "Invalid or missing Excel path"}), 400
#         if not output_path:
#             return jsonify({"error": "Missing output path"}), 400

#         # === Step 1: Save Excel as PDF ===
#         input_path = os.path.abspath(input_path)
#         output_pdf_path = os.path.join(os.path.abspath(output_path), "37111_MWH_TOC_Ver2_0_0.pdf")

#         excel_app = win32com.client.Dispatch("Excel.Application")
#         excel_app.Visible = False
#         workbook = excel_app.Workbooks.Open(input_path)

#         for sheet in workbook.Worksheets:
#             sheet.PageSetup.TopMargin = excel_app.CentimetersToPoints(2.7)
#             sheet.PageSetup.BottomMargin = excel_app.CentimetersToPoints(7)
#             sheet.PageSetup.Zoom = False
#             sheet.PageSetup.FitToPagesWide = 1.2
#             sheet.PageSetup.FitToPagesTall = False

#         workbook.ExportAsFixedFormat(0, output_pdf_path)
#         workbook.Close(False)
#         excel_app.Quit()

#         # === Step 2: Generate Overlay Page ===
#         packet = BytesIO()
#         c = canvas.Canvas(packet, pagesize=A2)
#         width, height = A2

#         bottom_table = draw_bottom_tables()
#         bottom_table.wrapOn(c, width, height)
#         bottom_table.drawOn(c, 1.77 * cm, 0.6 * cm)

#         top_table = draw_top_table()
#         top_table.wrapOn(c, width, height)
#         top_table.drawOn(c, 1.77 * cm, 19.2 * cm)

#         c.save()
#         packet.seek(0)
#         overlay_page = PdfReader(packet).pages[0]

#         # === Step 3: Merge Overlay with PDF ===
#         reader = PdfReader(output_pdf_path)
#         writer = PdfWriter()

#         first_page_size = reader.pages[0].mediabox
#         width = float(first_page_size.width)
#         height = float(first_page_size.height)

#         # Create a blank page with the same size
#         blank_page = PageObject.create_blank_page(width=width, height=height)
#         writer.add_page(blank_page)

#         # Merge the overlay starting from second page
#         for page in reader.pages:
#             page.merge_page(overlay_page)
#             writer.add_page(page)

#         with open(output_pdf_path, "wb") as f:
#             writer.write(f)

#         return jsonify({
#             "message": "PDF generated successfully with header and footer.",
#             "output_pdf_path": output_pdf_path
#         })

#     except Exception as e:
#         return jsonify({"error": str(e)}), 500



# def generate_toc_pdf_with_header_footer():
#     try:
#         data = request.get_json()
#         input_path = data.get("input_path")

#         if not input_path or not os.path.exists(input_path):
#             return jsonify({"error": "Invalid or missing Excel path"}), 400

#         # === Step 1: Save Excel as PDF ===
#         input_path = os.path.abspath(input_path)
#         output_pdf_path = os.path.splitext(input_path)[0] + '.pdf'

#         excel_app = win32com.client.Dispatch("Excel.Application")
#         excel_app.Visible = False
#         workbook = excel_app.Workbooks.Open(input_path)

#         for sheet in workbook.Worksheets:
#             sheet.PageSetup.TopMargin = excel_app.CentimetersToPoints(2.7)
#             sheet.PageSetup.BottomMargin = excel_app.CentimetersToPoints(7)

#         workbook.ExportAsFixedFormat(0, output_pdf_path)
#         workbook.Close(False)
#         excel_app.Quit()

#         # === Step 2: Generate Overlay Page ===
#         packet = BytesIO()
#         c = canvas.Canvas(packet, pagesize=A3)
#         width, height = A3

#         bottom_table = draw_bottom_tables()
#         bottom_table.wrapOn(c, width, height)
#         bottom_table.drawOn(c, 1.77 * cm, 0.6 * cm)

#         top_table = draw_top_table()
#         top_table.wrapOn(c, width, height)
#         top_table.drawOn(c, 1.77 * cm, 19.2 * cm)

#         c.save()
#         packet.seek(0)
#         overlay_page = PdfReader(packet).pages[0]

#         # === Step 3: Merge Overlay with PDF ===
#         reader = PdfReader(output_pdf_path)
#         writer = PdfWriter()

#         for page in reader.pages:
#             page.merge_page(overlay_page)
#             writer.add_page(page)

#         with open(output_pdf_path, "wb") as f:
#             writer.write(f)

#         return jsonify({
#             "message": "PDF generated successfully with header and footer.",
#             "output_pdf_path": output_pdf_path
#         })

#     except Exception as e:
#         return jsonify({"error": str(e)}), 500








# def generate_overlay_page(page_size=A3):
#     packet = BytesIO()
#     c = canvas.Canvas(packet, pagesize=page_size)
#     width, height = page_size

#     # === Draw bottom table ===
#     bottom_table = draw_bottom_tables()
#     bottom_table.wrapOn(c, width, height)
#     bottom_table.drawOn(c, 1.77 * cm, 0.6 * cm)

#     # === Draw top table ===
#     top_table = draw_top_table()
#     top_table.wrapOn(c, width, height)
#     top_table.drawOn(c, 1.77 * cm, 19.2 * cm)

#     c.save()
#     packet.seek(0)
#     return PdfReader(packet).pages[0]


# def add_overlay_to_pdf(input_pdf_path):
#     overlay_page = generate_overlay_page()

#     reader = PdfReader(input_pdf_path)
#     writer = PdfWriter()

#     for page in reader.pages:
#         page.merge_page(overlay_page)
#         writer.add_page(page)

#     # Overwrite the original PDF with the final version
#     with open(input_pdf_path, "wb") as f:
#         writer.write(f)

#     print(f"Final PDF with header and footer saved at: {input_pdf_path}")


# def save_excel_as_pdf(input_path):
#     input_path = os.path.abspath(input_path)
#     output_pdf_path = os.path.splitext(input_path)[0] + '.pdf'

#     excel_app = win32com.client.Dispatch("Excel.Application")
#     excel_app.Visible = False

#     workbook = excel_app.Workbooks.Open(input_path)

#     for sheet in workbook.Worksheets:
#         sheet.PageSetup.TopMargin = excel_app.CentimetersToPoints(2.7)
#         sheet.PageSetup.BottomMargin = excel_app.CentimetersToPoints(7)

#     workbook.ExportAsFixedFormat(0, output_pdf_path)
#     workbook.Close(False)
#     excel_app.Quit()

#     return output_pdf_path


# # === Run it ===
# input_path = r"D:\jay-robotix\Design Document Automation\Converted_Documents\Malwan\TOC_Malwan_station.xlsx"
# pdf_path = save_excel_as_pdf(input_path)
# add_overlay_to_pdf(pdf_path)











# def draw_precise_table_of_control(output_path):
#     doc = SimpleDocTemplate(output_path, pagesize=(A2[1], A2[0]))
#     elements = []

#     # Top Table: TABLE OF CONTROL
#     data_top = [
#         ["TABLE OF CONTROL", "", "", "", "", ""],
#         ["Station Name:", "MALWAN", "Division:", "PRAYAGRAJ", "CODE", "MWH"],
#         ["Station ID:", "37111", "Section:", "NC Railway", "", ""],
#     ]

#     col_widths_top = [65*mm, 80*mm, 65*mm, 80*mm, 80*mm, 135*mm]
#     row_heights_top = [10*mm, 12*mm, 12*mm]

#     table_top = Table(data_top, colWidths=col_widths_top, rowHeights=row_heights_top)

#     style_top = TableStyle([
#         ('GRID', (0,1), (-1,-1), 0.5, colors.black),
#         ('ALIGN', (0,0), (-1,-1), 'CENTER'),
#         ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
#         ('FONTSIZE', (0,0), (0, 6), 15),
#         ('FONTSIZE', (1,0), (-1,-1), 11),
#         ('FONTNAME', (0,1), (0,6), 'Helvetica-Bold'),
#         ('FONTNAME', (0,1), (0,2), 'Helvetica-Bold'),
#         ('FONTNAME', (2,1), (2,2), 'Helvetica-Bold'),
#         ('FONTNAME', (4,1), (4,1), 'Helvetica-Bold'),
#         ('SPAN', (0,0), (5,0)),
#         ('SPAN', (4,1), (4,2)),
#         ('SPAN', (5,1), (5,2)),
#         ('BACKGROUND', (0,1), (0,2), colors.lightgrey),
#         ('BACKGROUND', (2,1), (2,2), colors.lightgrey),
#         ('BACKGROUND', (4,1), (4,2), colors.lightgrey),
#     ])

#     table_top.setStyle(style_top)
#     elements.append(table_top)
#     elements.append(Spacer(1, 20*mm))  # Space between top and bottom tables

#     # Bottom Tables: TOC1, TOC2, TOC3 side by side
#     # TOC1 Table
#     sign1 = Image("sign1.png", width=3*cm, height=1*cm)
#     sign2 = Image("sign2.png", width=3*cm, height=1*cm)
#     logos = Image("logo.png", width=3*cm, height=1.5*cm)

#     data_toc1 = [
#         ["MALWAN(MWH), PRAYAGRAJ DIVISION, NC RAILWAY", "", "", "", ""],
#         ["REF : SIP : SSP.NCR.PRYJ.MWH2.0.1", "", "", "", ""],
#         ["REF: RFID_TAG_LAYOUT_MWH_2.0.0", "", "", "", ""],
#         ["TABLE OF CONTROL: ICT.NCR.PRYJ.MWH 2.0.1", "", "", "", ""],
#         ["", "PREPARED BY", "CHECKED BY", "VERIFIED BY", ""],
#         ["SIGN", sign1, sign2, "", logos],
#         ["NAME", "S DAS", "S JITENDRA VIJAY", "A.K. SRIVASTAVA", ""],
#         ["DESIGNATION", "Sr.S/W ENGINEER", "MANAGER V&V", "PROOF CONSULTANT", "KERNEX–KEC"]
#     ]

#     toc1_col_widths = [70, 100, 130, 130, 130]
#     toc1_row_heights = [31, 25, 25, 30, 30, 40, 25, 25]

#     table_toc1 = Table(data_toc1, colWidths=toc1_col_widths, rowHeights=toc1_row_heights)

#     style_toc1 = TableStyle([
#         ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
#         ("FONTSIZE", (0, 0), (-1, -1), 9),
#         ("ALIGN", (0, 0), (-1, -1), "CENTER"),
#         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
#         ("ALIGN", (0, 0), (-1, 3), "LEFT"),
#         ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
#         ("SPAN", (0, 0), (-1, 0)),
#         ("SPAN", (0, 1), (-1, 1)),
#         ("SPAN", (0, 2), (-1, 2)),
#         ("SPAN", (0, 3), (-1, 3)),
#         ("SPAN", (4, 5), (4, 6)),
#         ("SPAN", (4, 7), (4, 7)),
#     ])

#     table_toc1.setStyle(style_toc1)

#     # TOC2 Table
#     data_toc2 = [
#         ["Remarks", "", ""],
#         ["REVISION HISTORY IS AVAILABLE ON PAGE-12", "", ""],
#         ["", "", ""],
#         ["", "", ""],
#         [
#             "उप.मु.सि.एवं\nदू.सं.अभि/प्रो-1/\n प्रयागराज",
#             "वरि.सि.एवं दू.सं\nअभि/ प्रो/\nप्रयागरा",
#             "वरि.ख. अभि.\nसि./प्रयागराज"
#         ],
#         ["DY.CSTE/P-1/PRYJ", "SSTE/Pr/PRYJ", "SSE/SIG/PRYJ"],
#         ["NOTE:", "", ""]
#     ]

#     toc2_col_widths = [120, 120, 120]
#     toc2_row_heights = [20, 20, 50, 30, 70, 20, 20]

#     table_toc2 = Table(data_toc2, colWidths=toc2_col_widths, rowHeights=toc2_row_heights)

#     style_toc2 = TableStyle([
#         ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
#         ("ALIGN", (0, 0), (-1, 5), "CENTER"),
#         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
#         ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
#         ("FONTNAME", (0, 5), (-1, 5), "Helvetica-Bold"),
#         ("FONTNAME", (0, 6), (0, 6), "Helvetica-Bold"),
#         ("FONTNAME", (0, 6), (2, 6), "Helvetica-Bold"),
#         ("FONTSIZE", (0, 6), (2, 6), 10),
#         ("FONTSIZE", (0, 0), (-1, -1), 9),
#         ("ALIGN", (0, 6), (0, 6), "LEFT"),
#         ("SPAN", (0, 0), (-1, 0)),
#         ("SPAN", (0, 1), (-1, 1)),
#     ])

#     table_toc2.setStyle(style_toc2)

#     # TOC3 Table
#     data_toc3 = [
#         ["SIP: D-3067/1", "संदर्भ: सिग्नल ओरख सं:\nREF: SIG.PLAN.No.", "", "MALWAN(MWH)", "", "", ""],
#         ["", "संशोधन REVISION", "", "SIGNAL & TELECOMMUNICATION", "", "", ""],
#         ["", "आरेखकार DRAWING", "", "KAVACH Indian Railway ATP", "", "", ""],
#         ["", "कु.मु.सि.दू.सं.अ.\nDY.CSTE/D&D/PRYJ", "C\nH\nE\nC\nK\nE\nD", "PRAYAGRAJ DIVISION, NC RAILWAY", "", "", ""],
#         ["", "व.सि.दू.सं.अ.\nSSTE/D&D/PRYJ", "", "TABLE OF CONTROL:KAVACH_TOC_MWH_4.0.0", "", "", ""],
#         ["", "स.सि.दू.सं.अ.\nASTE/D&D/PRYJ", "", "एनसीआर.डी.आरजी", "KAVACH\nD-3067/1", "शीट", "कुलशीट"],
#         ["", "व.ख.अभि.सं.अ.", "", "N.C.R RLY.DRG.NO:", "", "SHEET", "SHEETS"],
#         ["", "\nSSE/D&D/PRYJ", "", "", "", "1", "12"],
#     ]

#     toc3_col_widths = [30*mm, 35*mm, 10*mm, 45*mm, 30*mm, 15*mm, 15*mm]
#     toc3_row_heights = [12*mm, 8.5*mm, 8.5*mm, 12*mm, 12*mm, 12*mm, 5*mm, 12*mm]

#     table_toc3 = Table(data_toc3, colWidths=toc3_col_widths, rowHeights=toc3_row_heights)

#     style_toc3 = TableStyle([
#         ('GRID', (0,0), (-1,-1), 0.5, colors.black),
#         ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
#         ('ALIGN', (0,0), (-1,-1), 'CENTER'),
#         ('FONTSIZE', (0,0), (-1,-1), 8),
#         ("FONTNAME", (0, 0), (7, 7), "Helvetica-Bold"),
#         ('SPAN', (0,0), (0,0)),
#         ('SPAN', (0,6), (0,7)),
#         ('SPAN', (1,0), (2,0)),
#         ('SPAN', (1,1), (2,1)),
#         ('SPAN', (1,2), (2,2)),
#         ('SPAN', (1,6), (1,7)),
#         ('SPAN', (3,0), (6,0)),
#         ('SPAN', (3,1), (6,1)),
#         ('SPAN', (3,2), (6,2)),
#         ('SPAN', (3,3), (6,3)),
#         ('SPAN', (3,4), (6,4)),
#         ('SPAN', (3,6), (3,6)),
#         ('SPAN', (2,3), (2,7)),
#         ('SPAN', (4,5), (4,7)),
#         ('SPAN', (3,7), (3,7)),
#         ('SPAN', (5,6), (5,6)),
#         ('SPAN', (6,6), (6,6)),
#         ('SPAN', (5,7), (5,7)),
#         ('SPAN', (6,7), (6,7)),
#     ])

#     table_toc3.setStyle(style_toc3)

#     # Combine TOC1, TOC2, TOC3 in a single row
#     bottom_table = Table([[table_toc1, table_toc2, table_toc3]], colWidths=[sum(toc1_col_widths, 10), sum(toc2_col_widths, 10), sum(toc3_col_widths)])
#     bottom_table.setStyle(TableStyle([
#         ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
#         ('ALIGN', (0,0), (-1,-1), 'CENTER'),
#     ]))
#     elements.append(bottom_table)

#     # Build the document
#     doc.build(elements)

