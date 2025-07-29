from reportlab.lib.units import mm


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
    c.setFont("Helvetica-Bold", 10)
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

    row_height = 12
    triangle_size = 7
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
