from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Table, TableStyle, Paragraph
import pandas as pd
import re
from collections import Counter, defaultdict


def draw_title_block_on_pdf(c, x0, y0, width, total_height):
    """
    Draws a reference block, title block, and revision history table (to the right).
    `x0`, `y0` is the bottom-left corner of the combined block.
    """
    ref_block_height = 80
    title_block_height = 80
    rev_block_width = width
    rev_block_height = title_block_height

    # ------------------ Draw REF Block ------------------
    ref_block_y = y0 + title_block_height + 10
    c.setLineWidth(1)
    c.rect(x0, ref_block_y, width, ref_block_height)

    ref_text = """*REF.SIP:STATION  &  BLOCK  SECTION
(MWH           SIP:D–3067/1  DATED  31.03.2022)
(MWH–KKS SECTION     SIP:3064/1  DATED  31.03.2022)
(KSQ–MWH SECTION     SIP:SSP.NCR.PRYJ.KSQ–MWH.01  DATED  20.04.2022)
*REF. TOC: RCT–3067/1"""

    c.setFont("Helvetica", 8)
    line_height = 10
    lines = ref_text.split("\n")
    total_text_height = len(lines) * line_height
    start_y = ref_block_y + (ref_block_height + total_text_height) / 2 - line_height

    for i, line in enumerate(lines):
        c.drawString(x0 + 4, start_y - i * line_height, line)

    # ------------------ Draw Title Block ------------------
    left_width = width * 0.5
    right_width = width * 0.5

    upper_height = title_block_height * 0.6
    middle_height = title_block_height * 0.2
    lower_height = title_block_height * 0.2

    # Outer boxes of title
    c.rect(x0, y0, width, title_block_height)                                # Full block
    c.rect(x0, y0 + lower_height + middle_height, left_width, upper_height)  # Top-left
    c.rect(x0, y0 + lower_height, left_width, middle_height)                 # Middle-left
    c.rect(x0, y0, left_width, lower_height)                                 # Bottom-left
    c.rect(x0 + left_width, y0, right_width, title_block_height)             # Right

    # Title text
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(x0 + left_width / 2, y0 + lower_height + middle_height + upper_height / 2 - 5, "MALWAN(MWH)")

    c.setFont("Helvetica", 9)
    c.drawCentredString(x0 + left_width / 2, y0 + lower_height / 2 - 2, "NOT DRAWN     TO SCALE")

    c.setFont("Helvetica", 10)
    c.drawCentredString(x0 + left_width + right_width / 2, y0 + title_block_height / 2 - 5, "RAIL YARD_LAYOUT_MWH_2.0.1")

    # ------------------ Draw Revision History Block ------------------
    rev_x = x0 + 10 + width  # right to title block
    rev_y = y0
    rev_his_width = rev_block_width + 50
    col_widths = [rev_his_width * 0.2, rev_his_width * 0.6, rev_his_width * 0.2]
    row_height = rev_block_height / 3

    # Outer box
    c.rect(rev_x, rev_y, rev_his_width, rev_block_height)

    # Horizontal lines
    c.line(rev_x, rev_y + row_height, rev_x + rev_his_width, rev_y + row_height)
    c.line(rev_x, rev_y + 2 * row_height, rev_x + rev_his_width, rev_y + 2 * row_height)

    # Vertical lines
    c.line(rev_x + col_widths[0], rev_y, rev_x + col_widths[0], rev_y + rev_block_height)
    c.line(rev_x + col_widths[0] + col_widths[1], rev_y, rev_x + col_widths[0] + col_widths[1], rev_y + rev_block_height)

    # Data
    c.setFont("Helvetica", 8)
    data = [
        ["2.0.1", "Initial Version(4.0 Amdt–5.0) updated as per Guidlines change/comments", "21.10.2024"],
        ["2.0.0", "Initial Version(4.0 Amdt–4.0)", "18.12.2023"],
        ["REV. NO", "REVISION HISTORY", "DATE"],
    ]

    for i, row in enumerate(data[::-1]):  # bottom to top
        for j, text in enumerate(row):
            c.drawCentredString(
                rev_x + sum(col_widths[:j]) + col_widths[j] / 2,
                rev_y + i * row_height + row_height / 2 - 3,
                text
            )
    
    
    
    # ------------------ Draw LEGEND Table ------------------
    legend_x = x0 + width + rev_his_width + 30  # Right of revision history table
    legend_y = y0
    legend_width = width * 0.5
    legend_height = title_block_height * 2  # Same height as title block

    c.setFont("Helvetica", 8)

    # Outer border
    c.rect(legend_x, legend_y, legend_width, legend_height)

    # Header
    c.line(legend_x, legend_y + legend_height - 15, legend_x + legend_width, legend_y + legend_height - 15)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(legend_x + legend_width / 2, legend_y + legend_height - 12, "LEGEND")

    # Rows (each approx. 12 px high)
    legend_items = [
        ("N", "NORMAL TAG"),
        ("S", "SIGNAL FOOT TAG"),
        ("T", "TIN DISCRIMINATION TURNOUT TAG"),
        ("G", "GATE TAG"),
        ("X", "EXIT TAG"),
        ("L", "ADJACENT LINE INFO. TAG"),
        ("A", "ADJUSTMENT/JUNCTION TAG"),
    ]

    row_height = 18
    triangle_size = 12
    for i, (code, label) in enumerate(legend_items):
        y_pos = legend_y + legend_height - 15 - (i + 1) * row_height + 3
        x_pos = legend_x + 5

        # Draw triangle pointing right
        c.setLineWidth(1)
        c.line(x_pos, y_pos, x_pos + triangle_size, y_pos + triangle_size / 2)
        c.line(x_pos + triangle_size, y_pos + triangle_size / 2, x_pos, y_pos + triangle_size)
        c.line(x_pos, y_pos + triangle_size, x_pos, y_pos)

        # Letter inside triangle
        c.setFont("Helvetica-Bold", 6)
        c.drawCentredString(x_pos + triangle_size / 2 + 0.5, y_pos + triangle_size / 2 - 1, code)

        # Label text
        c.setFont("Helvetica", 7)
        c.drawString(x_pos + triangle_size + 5, y_pos + 1, label)
        
        
        
        # ------------------ Draw NOTES Block ------------------
        notes_x = legend_x + legend_width + 20  # Right of LEGEND
        notes_y = y0
        notes_width = width  # Double width (as per image)
        notes_height = title_block_height * 1.3  # Same height as title block

        # Outer border
        c.rect(notes_x, notes_y, notes_width, notes_height)

        # Vertical divider (for "LEGEND" label cell)
        label_width = width * 0.2   
        c.line(notes_x + label_width, notes_y, notes_x + label_width, notes_y + notes_height)

        # "LEGEND" text vertically/horizontally centered
        c.setFont("Helvetica", 9)
        c.drawCentredString(
            notes_x + label_width / 2,
            notes_y + notes_height / 2 - 4,
            "LEGEND"
        )

        # Notes text
        note_text = (
            "RFID TAG LAYOUT CHAINAGE AND SIP CHAINAGE ARE\n"
            "MISMATCH HENCE WE HAVE MENTIONED AS PER\n"
            "DRONE/PHYSICAL SURVEY DATA.\n\n"
            "THIS LAYOUT PREPARED BASED ON TCAS 4.0\n"
            "VERSION INPUTS."
        )

        # Font setup
        c.setFont("Helvetica", 9)
        line_height = 10

        # Prepare lines
        lines = note_text.strip().split("\n")
        num_lines = len(lines)

        # Total text block height
        text_block_height = num_lines * line_height

        # Y-position to start drawing so that block is vertically centered
        start_y = notes_y + (notes_height - text_block_height) / 2 + (line_height * (num_lines - 1))

        # X-position to start so that block is horizontally centered inside right cell
        text_x = notes_x + label_width + (notes_width - label_width) / 2 - 100  # Adjust -100 if needed

        # Draw each line
        for i, line in enumerate(lines):
            c.drawString(
                text_x,
                start_y - i * line_height,
                line
            )




def draw_bottom_right_block_on_pdf(c, x0, y0, width, height):
    """
    Draws the bottom-right title block (SIGNAL & TELECOMMUNICATIONS etc.)
    :param c: Canvas object
    :param x0: X-coordinate of bottom-left corner
    :param y0: Y-coordinate of bottom-left corner
    :param width: Total width of the block
    :param height: Total height of the block
    """
    from reportlab.lib.units import mm

    # Define heights
    upper_height = height * 0.55
    middle_height = height * 0.30
    lower_height = height * 0.15

    # Outer border
    c.rect(x0, y0, width, height)

    # Horizontal lines separating sections
    c.line(x0, y0 + lower_height, x0 + width, y0 + lower_height)
    c.line(x0, y0 + lower_height + middle_height, x0 + width, y0 + lower_height + middle_height)

    # Draw "SIGNAL & TELECOMMUNICATIONS" in upper block (centered)
    c.setFont("Helvetica", 9)
    c.drawCentredString(
        x0 + width / 2,
        y0 + lower_height + middle_height + upper_height / 2 - 4,
        "SIGNAL & TELECOMMUNICATIONS"
    )

    # Draw "MALWAN (MWH)" and "RAIL YARD LAYOUT" in middle block
    # Compute vertical centering
    malwan_font_size = 12
    layout_font_size = 10
    line_spacing = 2
    total_text_height = malwan_font_size + layout_font_size + line_spacing
    middle_y_center = y0 + lower_height + (middle_height / 2)

    start_y = middle_y_center + total_text_height / 2 - malwan_font_size

    c.setFont("Helvetica-Bold", malwan_font_size)
    c.drawString(x0 + 5, start_y, "MALWAN (MWH)")

    c.setFont("Helvetica", layout_font_size)
    layout_y = start_y - layout_font_size - line_spacing
    c.drawString(x0 + 5, layout_y, "RAIL YARD LAYOUT")

    # Underline
    c.line(x0 + 5, layout_y - 1.5, x0 + 70, layout_y - 1.5)

    # Lower section columns: [left blank], [drawing no], [SHEET], [SHEETS]
    col1 = x0
    col2 = x0 + width * 0.5
    col3 = x0 + width * 0.7
    col4 = x0 + width * 0.85

    # Vertical lines
    c.line(col2, y0, col2, y0 + lower_height)
    c.line(col3, y0, col3, y0 + lower_height)
    c.line(col4, y0, col4, y0 + lower_height)

    # Drawing number in center of col2-col3
    c.setFont("Helvetica", 9)
    c.drawCentredString((col2 + col3) / 2, y0 + lower_height / 2 - 4, "R–3067/1")

    # SHEET and SHEETS text (headers)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString((col3 + col4) / 2, y0 + lower_height * 0.65, "SHEET")
    c.drawCentredString((col4 + x0 + width) / 2, y0 + lower_height * 0.65, "SHEETS")

    # Values
    c.setFont("Helvetica", 8)
    c.drawCentredString((col3 + col4) / 2, y0 + lower_height * 0.25, "1")
    c.drawCentredString((col4 + x0 + width) / 2, y0 + lower_height * 0.25, "1")
    
    
    
        # === Left Reference Block ===
    left_block_width = width * 2
    left_block_x0 = x0 - left_block_width
    left_block_height = lower_height * 1.5
    c.rect(left_block_x0, y0, left_block_width, left_block_height)

    # Split into 2 columns
    col_mid = left_block_x0 + left_block_width * 0.5
    c.line(col_mid, y0, col_mid, y0 + left_block_height)

    # Left column text (left-aligned)
    c.setFont("Helvetica", 7.5)
    c.drawString(left_block_x0 + 4, y0 + left_block_height * 0.55, "*REF.SIP: SI–3067/1")
    c.drawString(left_block_x0 + 4, y0 + left_block_height * 0.15, "*REF.TOC:  RCT–3067/1")

    # Right column text (centered)
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(
        col_mid + (left_block_x0 + left_block_width - col_mid) / 2,
        y0 + left_block_height * 0.35,
        "REF: SIG.PLAN.No.:"
    )






def draw_direction_arrows(c, x_center, y_center, length=40, spacing=15):
    """
    Draws two directional arrows with text: DN↔UP and Reverse↔Nominal
    :param c: Canvas object
    :param x_center: Center X position
    :param y_center: Center Y position between the two arrow rows
    :param length: Length of each arrow line
    :param spacing: Vertical spacing between the arrows
    """
    arrow_size = 3  # arrowhead size
    half_length = length / 2

    # First row (Top): DN <-> UP
    y1 = y_center + spacing / 2
    c.line(x_center - half_length, y1, x_center + half_length, y1)
    # Arrows
    c.line(x_center - half_length + arrow_size, y1 + arrow_size, x_center - half_length, y1)
    c.line(x_center - half_length + arrow_size, y1 - arrow_size, x_center - half_length, y1)
    c.line(x_center + half_length - arrow_size, y1 + arrow_size, x_center + half_length, y1)
    c.line(x_center + half_length - arrow_size, y1 - arrow_size, x_center + half_length, y1)
    # Labels
    c.setFont("Helvetica", 10)
    c.drawRightString(x_center - half_length - 5, y1 - 4, "DN")
    c.drawString(x_center + half_length + 5, y1 - 4, "UP")

    # Second row (Bottom): Reverse <-> Nominal
    y2 = y_center - spacing / 2
    c.line(x_center - half_length, y2, x_center + half_length, y2)
    # Arrows
    c.line(x_center - half_length + arrow_size, y2 + arrow_size, x_center - half_length, y2)
    c.line(x_center - half_length + arrow_size, y2 - arrow_size, x_center - half_length, y2)
    c.line(x_center + half_length - arrow_size, y2 + arrow_size, x_center + half_length, y2)
    c.line(x_center + half_length - arrow_size, y2 - arrow_size, x_center + half_length, y2)
    # Labels
    c.drawRightString(x_center - half_length - 5, y2 - 4, "Reverse")
    c.drawString(x_center + half_length + 5, y2 - 4, "Nominal")



def get_all_station_data(excel_path):
    df = pd.read_excel(excel_path, sheet_name="RFID TAG & TIN Ranges")
    df.columns = [str(col).strip().upper() for col in df.columns]

    required_cols = ["STATION CODE", "STATION ID", "RFID TAG RANGE", "RFID TIN RANGE"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing column: {col}")

    return df[required_cols].dropna().to_dict(orient="records")

def calculate_spare_range(actual_range, allotted_range):
    def parse_range(r):
        if not r:
            return set()
        parts = r.split(",")
        numbers = set()
        for part in parts:
            part = part.strip()
            if "-" in part:
                try:
                    start, end = map(int, part.split("-"))
                    numbers.update(range(start, end + 1))
                except:
                    continue  # ignore malformed range
            else:
                try:
                    numbers.add(int(part))
                except:
                    continue  # ignore malformed single value
        return numbers


    actual_set = parse_range(actual_range)
    allotted_set = parse_range(allotted_range)
    spare_set = sorted(actual_set - allotted_set)

    if not spare_set:
        return "-----"

    # Convert back to compressed range string
    ranges = []
    start = prev = spare_set[0]
    for num in spare_set[1:]:
        if num == prev + 1:
            prev = num
        else:
            ranges.append(f"{start}-{prev}" if start != prev else str(start))
            start = prev = num
    ranges.append(f"{start}-{prev}" if start != prev else str(start))

    return ", ".join(ranges)


def draw_combined_station_and_border_table(
    c, x0, y0, width, height=None,
    alloted_tags=None, alloted_tins=None,
    station_id=None, border_tags=None,
    station_info_path=None
):
    styles = getSampleStyleSheet()
    para_style = ParagraphStyle(
        name='WrapStyle',
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=1,
    )

    station_data_list = get_all_station_data(station_info_path)

    # First two columns: labels
    headers = ["STATION", "CODE"] + [f'{station["STATION CODE"]}' for station in station_data_list]
    station_ids = ["", "ID"] + [str(station["STATION ID"]).split(".")[0] for station in station_data_list]
    rfid_range_row = ["RFID", "RANGE"] + [station["RFID TAG RANGE"] for station in station_data_list]
    tin_range_row = ["TIN", "RANGE"] + [station["RFID TIN RANGE"] for station in station_data_list]
    rfid_allotted_row = ["", "ALLOTTED"]
    tin_allotted_row = ["", "ALLOTTED"]

    for station in station_data_list:
        sid = str(station["STATION ID"]).split(".")[0]
        if station_id and str(station_id) == sid:
            rfid_allotted_row.append(alloted_tags)
            tin_allotted_row.append(alloted_tins)
        else:
            rfid_allotted_row.append("-----")
            tin_allotted_row.append("-----")
            
    rfid_spare_row = ["", "SPARE"]
    tin_spare_row = ["", "SPARE"]

    for station in station_data_list:
        sid = str(station["STATION ID"]).split(".")[0]
        if station_id and str(station_id) == sid:
            rfid_spare = calculate_spare_range(station["RFID TAG RANGE"], alloted_tags)
            tin_spare = calculate_spare_range(station["RFID TIN RANGE"], alloted_tins)
            rfid_spare_row.append(rfid_spare)
            tin_spare_row.append(tin_spare)
        else:
            rfid_spare_row.append("-----")
            tin_spare_row.append("-----")



    station_raw_data = [
        headers,
        station_ids,
        rfid_range_row,
        rfid_allotted_row,
        rfid_spare_row,
        tin_range_row,
        tin_allotted_row,
        tin_spare_row
    ]

    station_col_widths = [80, 80] + [150] * len(station_data_list)

    wrapped_data = []
    row_heights = []

    for row in station_raw_data:
        wrapped_row = []
        row_height = 0
        for i, cell in enumerate(row):
            para = Paragraph(str(cell), para_style)
            wrapped_row.append(para)
            avail_width = station_col_widths[i]
            _, cell_height = para.wrap(avail_width, 10000)
            row_height = max(row_height, cell_height)
        wrapped_data.append(wrapped_row)
        row_heights.append(row_height + 15)

    table = Table(wrapped_data, colWidths=station_col_widths, rowHeights=row_heights)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (0, 2), (0, 4)),
        ("SPAN", (0, 5), (0, 7)),
    ]))

    station_table_width = sum(station_col_widths)
    table.wrapOn(c, width, height or sum(row_heights))
    table.drawOn(c, x0, y0)
    
    
    # ------------------ RIGHT: Border Line Tags Table ------------------
    border_data = [
        ["BORDER LINE TAGS", "", "", "", ""],
        ["", "DIRECTION", "TAG ID", "DISTANCE FROM\nMID POINT (MWH)", "SIG STRENGTH AS PER\nRSSI SURVEY"],
        ["KURASTI KALAN SIDE", "UP", f"R-{border_tags[0]}", "3.168KM", "ABOVE –60db"],
        ["", "DN", f"R-{border_tags[1]}", "2.779KM", "ABOVE –60db"],
        ["KANSPUR GUGAULI SIDE", "UP", f"R-{border_tags[2]}", "3.984KM", "ABOVE –60db"],
        ["", "DN", f"R-{border_tags[3]}", "4.586KM", "ABOVE –60db"],
    ]
    border_col_widths = [150, 150, 150, 150, 150]
    border_row_heights = [36, 36, 32, 32, 32, 32]

    border_table = Table(border_data, colWidths=border_col_widths, rowHeights=border_row_heights)
    border_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("SPAN", (0, 0), (-1, 0)),
        ("SPAN", (0, 2), (0, 3)),
        ("SPAN", (0, 4), (0, 5)),
    ]))

    border_table_width = sum(border_col_widths)
    border_table_x = x0 + station_table_width + 20
    border_table.wrapOn(c, border_table_width, sum(border_row_heights))
    border_table.drawOn(c, border_table_x, y0)


    
    


def extract_tag_and_tin_ranges(file_path):
    # Regex to match tag numbers like 947/M, 947/D
    tag_pattern = re.compile(r"(\d{3,4})[\/_]([MD])", re.IGNORECASE)
    tin_keywords = ["TIN 1", "TIN 2", "TIN in Nominal Direction", "TIN in Reverse Direction"]
    station_keywords = ["Station ID in Nominal Direction", "Station ID in Reverse Direction"]

    tag_numbers = set()
    tin_numbers = set()
    station_id_list = []
    tag_station_map = defaultdict(dict)  # Format: {947: {"nominal": "37111", "reverse": "37112"}}

    try:
        xls = pd.ExcelFile(file_path)
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return "", "", "", "", ""

    for sheet in xls.sheet_names:
        # Read the sheet, try different header rows to find the correct one
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        header_row = None
        tag_columns = {}

        # Find the row with tag headers (e.g., 947/M, 947/D)
        for i in range(min(5, len(df))):  # Check first 5 rows
            row = df.iloc[i].astype(str)
            tag_count = 0
            for col_idx, val in enumerate(row):
                match = tag_pattern.search(val.strip())
                if match:
                    tag_num = int(match.group(1))
                    tag_columns[col_idx] = tag_num
                    tag_numbers.add(tag_num)
                    tag_count += 1
            if tag_count > 5:  # Assume row is header if multiple tags found
                header_row = i
                break

        if header_row is None:
            df = pd.read_excel(xls, sheet_name=sheet, header=1)
            for col in df.columns:
                match = tag_pattern.search(str(col).strip())
                if match:
                    tag_num = int(match.group(1))
                    tag_numbers.add(tag_num)
                    tag_columns[df.columns.get_loc(col)] = tag_num
            if not tag_columns:
                print(f"No tags found in columns for sheet {sheet}")
                continue

        # Process "NT" sheet for station IDs
        if sheet.upper() == "NT":
            nominal_row = None
            reverse_row = None
            # Find station ID rows
            for i in range(len(df)):
                row = df.iloc[i].astype(str)
                row_str = " ".join(val.strip().upper() for val in row if pd.notna(val))
                if "STATION ID IN NOMINAL DIRECTION" in row_str:
                    if i + 1 < len(df):
                        nominal_row = df.iloc[i]
                        # print(f"Nominal station ID row at index {i}: {nominal_row.values}")
                if "STATION ID IN REVERSE DIRECTION" in row_str:
                    if i + 1 < len(df):
                        reverse_row = df.iloc[i]
                        # print(f"Reverse station ID row at index {i}: {reverse_row.values}")

            # Map tags to station IDs
            if nominal_row is not None and reverse_row is not None:
                for col_idx, tag_num in tag_columns.items():
                    try:
                        nominal_id = str(int(float(nominal_row.iloc[col_idx])))
                        reverse_id = str(int(float(reverse_row.iloc[col_idx])))
                        tag_station_map[tag_num]['nominal'] = nominal_id
                        tag_station_map[tag_num]['reverse'] = reverse_id
                        station_id_list.extend([nominal_id, reverse_id])
                    except (ValueError, TypeError) as e:
                        print(f"Error processing tag {tag_num} at column {col_idx}: {e}")
                        continue
            else:
                print("Nominal or Reverse station ID row not found in NT sheet")

        # Extract TIN numbers
        for i in range(len(df)):
            row = df.iloc[i].astype(str)
            row_str = " ".join(val.strip().upper() for val in row if pd.notna(val))
            if any(keyword.upper() in row_str for keyword in tin_keywords):
                for val in row:
                    try:
                        num = int(float(val))
                        if num >= 100:
                            tin_numbers.add(num)
                    except (ValueError, TypeError):
                        continue

        # Generic Station ID capture
        for i in range(len(df)):
            row = df.iloc[i].astype(str)
            row_str = " ".join(val.strip().upper() for val in row if pd.notna(val))
            if "STATION ID" in row_str:
                for val in row:
                    val_str = str(val).strip()
                    if val_str.isdigit() and len(val_str) >= 4:
                        station_id_list.append(val_str)

    def group_into_ranges(numbers):
        sorted_nums = sorted(numbers)
        if not sorted_nums:
            return ""
        ranges = []
        start = end = sorted_nums[0]
        for num in sorted_nums[1:]:
            if num == end + 1:
                end = num
            else:
                ranges.append((start, end))
                start = end = num
        ranges.append((start, end))
        return ", ".join(f"{s}-{e}" if s != e else f"{s}" for s, e in ranges)

    alloted_tags = group_into_ranges(tag_numbers)
    alloted_tins = group_into_ranges(tin_numbers)
    unique_station_ids = sorted(set(station_id_list))
    station_ids = ", ".join(unique_station_ids)
    most_station_id = Counter(station_id_list).most_common(1)[0][0] if station_id_list else ""

    # Tags with different Nominal and Reverse station IDs
    diff_station_tags = []
    for tag_num, dirs in tag_station_map.items():
        if 'nominal' in dirs and 'reverse' in dirs and dirs['nominal'] != dirs['reverse']:
            diff_station_tags.append(tag_num)
    border_tags = diff_station_tags

    return alloted_tags, alloted_tins, station_ids, most_station_id, border_tags

# Example usage
# tag_str, tin_str, station_str, common_station, diff_tags = extract_tag_and_tin_ranges("D:/jay-robotix/Design Document Automation/Output_Documents/Malwan/tables/TD_Malwan_station.xlsx")

# print("TAGs:", tag_str)
# print("TINs:", tin_str)
# print("Stations:", station_str)
# print("Most Common Station:", common_station)
# print("Tags with Different Nominal/Reverse Station IDs:", diff_tags)




def get_station_data_by_id(excel_path, station_id):
    df = pd.read_excel(excel_path, sheet_name="Sheet1", header=1)
    df.columns = [str(col).strip().upper() for col in df.columns]
    station_id_col = next((col for col in df.columns if "STATION" in col and "ID" in col), None)
    if not station_id_col:
        raise ValueError("No column found containing 'Station ID'")
    matched_rows = df[df[station_id_col] == station_id]
    return matched_rows.to_dict(orient="records")



# excel_file = "Footer_data.xlsx"
# station_id = 37111
# station_data = get_station_data_by_id(excel_file, station_id)

# for row in station_data:
#     print(row)
