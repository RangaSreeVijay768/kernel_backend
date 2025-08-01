from flask import Flask, request, send_file, jsonify
import pandas as pd
import re
import os
import textwrap
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from calculations_nt import generate_pages_for_tag as calculate_values_nt, TagInfo as TagInfoNT, TagDir as TagDirNT
from calculations_aline import calculate_values as calculate_values_aline, TagInfo as TagInfoAline, TagDir as TagDirAline
from calculations_adj import generate_pages_for_tag as calculate_values_adj, TagInfo as TagInfoAdj, TagDir as TagDirAdj

app = Flask(__name__)

def extract_tag_columns(columns):
    tag_cols = []
    for c in columns:
        col_str = str(c).strip()
        if re.match(r"^\d+/[MD]$", col_str):
            tag_cols.append(c)
        else:
            print(f"Column '{col_str}' does not match regex r'^\\d+/[MD]$'")
    return tag_cols

def get_style_elements():
    thin = colors.black
    fill = colors.HexColor("#D9D9D9")
    font_name = "Times-Roman"
    font_size = 8
    bold_font = "Times-Bold"
    return thin, fill, font_name, font_size, bold_font


def format_pdf_table(df, tag_title, sheet_name):
    from math import ceil
    try:
        from footer import extract_template_with_placeholders
        from openpyxl.utils import range_boundaries
    except ImportError:
        print("Warning: Footer module not found. Footer insertion will be skipped.")
        footer_template = {"template": [], "merged_cells": [], "start_row": 1, "start_col": 1}

    thin, fill, font_name, font_size, bold_font = get_style_elements()
    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(
        name='HeaderStyle',
        fontName='Times-Bold',
        fontSize=14,
        alignment=1,  # TA_CENTER
        spaceAfter=6,
        leading=16
    )
    footer_style = ParagraphStyle(
        name='FooterStyle',
        fontName='Times-Roman',  # Match Excel's default font
        fontSize=11,  # Default size, will be overridden for specific fields
        alignment=1,  # Center-aligned
        leading=12,
        spaceBefore=2,
        spaceAfter=2
    )

    predefined_cols = ["FIELD NAME / DESCRIPTION", "BIT POSITION", "Size (Bits)"]
    tag_columns = [col for col in df.columns if col not in predefined_cols]
    chunk_size = 10
    total_chunks = ceil(len(tag_columns) / chunk_size)
    elements = []

    page_width = 397 * mm
    page_height = 210 * mm
    margin = 10 * mm
    table_width = page_width - 2 * margin

    # Adjusted column widths to fit within page (total ~277mm), matching sample PDF
    col_widths = [80 * mm, 35 * mm, 30 * mm] + [27 * mm] * 10  # Up to 10 tag columns
    total_width = sum(col_widths[:13])  # Max 13 columns (3 predefined + 10 tags)
    if total_width > table_width:
        scale = table_width / total_width
        col_widths = [w * scale for w in col_widths]

    # Footer column widths (15 columns to match Excel A:O)
    footer_col_widths = [table_width / 15] * 15  # Equal widths for simplicity, adjust if needed

    # Load footer template
    if 'footer_template' not in locals():
        footer_template = extract_template_with_placeholders("template.xlsx", "A2:O6")
    
    

    for i in range(total_chunks):
        footer_values = {
            "division": {"value": "PRAYAGRAJ Division\n", "size": 16},
            "railways": {"value": "NC RAILWAYS\n", "size": 16},
            "station-name": {"value": "MALWAN(MWH)\n", "size": 16},
            "station-id": {"value": "Section Id: 37111\n", "size": 16},
            "date": {"value": "11-07-2025", "size": 8},
            "total-pages": {"value": str(total_chunks), "size": 8},
            "current-page": {"value": str(i + 1), "size": 8}
        }
        start = i * chunk_size
        end = start + chunk_size
        tag_chunk = tag_columns[start:end]
        temp_df = df[predefined_cols + tag_chunk]

        # Header
        header_text = f"{tag_title}"
        header = Paragraph(header_text, header_style)
        elements.append(header)
        elements.append(Spacer(1, 4 * mm))

        # Prepare table data
        def wrap_text(text, max_chars):
            if pd.isna(text):
                return ''
            text = str(text)
            wrapped_lines = textwrap.wrap(text, width=max_chars)
            return '\n'.join(wrapped_lines)

        # Header row
        data = [temp_df.columns.tolist()]

        # Data rows
        for _, row in temp_df.iterrows():
            wrapped_row = []
            for idx, (col, cell) in enumerate(zip(temp_df.columns, row)):
                text = str(cell) if pd.notnull(cell) else ''
                if idx == 0:  # First column only
                    text = wrap_text(text, 50)
                wrapped_row.append(text)
            data.append(wrapped_row)

        # Create table
        table = Table(data, colWidths=col_widths[:len(temp_df.columns)], hAlign='LEFT')
        table_style = TableStyle([
            ('FONT', (0, 0), (-1, 0), bold_font, font_size),  # Header font
            ('FONT', (0, 1), (2, -1), bold_font, font_size),  # Predefined columns bold
            ('FONT', (3, 1), (-1, -1), font_name, font_size),  # Other cells normal
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # All centered (adjusted to match Excel)
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 0.25, thin),
            ('BOX', (0, 0), (-1, -1), 0.25, thin),
            ('BACKGROUND', (0, 0), (2, 0), fill),  # Header fill for predefined columns
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ])
        table.setStyle(table_style)

        # Adjust row heights based on content
# Adjust row heights based on the tallest cell in the row
        for row_idx, row in enumerate(data):
            max_lines = 1
            for cell in row:
                if isinstance(cell, str):
                    line_count = cell.count('\n') + 1
                    max_lines = max(max_lines, line_count)
            # Assume 3.5 mm per line (font_size 8 + padding)
            table._argH[row_idx] = max(9, max_lines * 3.5) * mm

        elements.append(table)
        elements.append(Spacer(1, 6 * mm))

        # Prepare footer data
        if footer_template["template"]:
            footer_data = []
            for row in footer_template["template"]:
                footer_row = []
                for cell_data in row:
                    value = cell_data["value"]
                    font_size = 8  # Default
                    # Replace placeholders
                    for key, val in footer_values.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in value:
                            if isinstance(val, dict):
                                replacement = str(val.get("value", "")).replace("\n", "<br/>")
                                font_size = val.get("size", 8)
                            else:
                                replacement = str(val).replace("\n", "<br/>")
                            value = value.replace(placeholder, replacement)
                    # Wrap text if necessary
                    wrapped_value = wrap_text(value, 30)  # Adjust based on column width
                    # Use Paragraph to apply dynamic font size
                    para = Paragraph(wrapped_value, ParagraphStyle(
                        name=f'FooterCell_{id(wrapped_value)}',
                        fontName='Times-Bold',
                        fontSize=font_size,
                        alignment=1,  # Center
                        leading=font_size + 2
                    ))
                    footer_row.append(para)
                footer_data.append(footer_row)

            # Create footer table
            footer_table = Table(footer_data, colWidths=footer_col_widths, hAlign='LEFT')
            footer_style = TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('INNERGRID', (0, 0), (-1, -1), 0.25, thin),
                ('BOX', (0, 0), (-1, -1), 0.25, thin),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 1),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ])

            # Apply merged cells
            for merge in footer_template["merged_cells"]:
                min_col, min_row, max_col, max_row = range_boundaries(merge)
                footer_style.add('SPAN', (min_col - footer_template["start_col"], min_row - footer_template["start_row"]),
                                (max_col - footer_template["start_col"], max_row - footer_template["start_row"]))

            # Set row heights (50 points in Excel ≈ 17.64 mm)
            for row_idx in range(len(footer_data)):
                footer_table._argH[row_idx] = 12.64 * mm

            footer_table.setStyle(footer_style)
            elements.append(Spacer(1, 3 * mm))
            elements.append(footer_table)
            elements.append(PageBreak())

    return elements

def process_input_sheet(sheet_name, df):
    print(f"\nProcessing sheet: {sheet_name}")

    df = df.dropna(how='all').reset_index(drop=True)
    if df.empty:
        print(f"Sheet {sheet_name} is empty after dropping NA rows.")
        return None

    df.columns = [str(col).strip() for col in df.columns]
    
    df['original_index'] = df[df.columns[0]]
    if df['original_index'].duplicated().any():
        print(f"Duplicate index labels found in {sheet_name}: {df['original_index'][df['original_index'].duplicated()].tolist()}")
        df['original_index'] = df['original_index'].astype(str) + '_' + df.groupby('original_index').cumcount().astype(str)
        df['original_index'] = df['original_index'].str.replace('_0$', '', regex=True)

    df.set_index('original_index', inplace=True)
    df.index = df.index.astype(str).str.strip()

    tag_cols = extract_tag_columns(df.columns)
    print(f"Tag columns found: {tag_cols}")
    if not tag_cols:
        print(f"No valid tag columns in sheet {sheet_name}. Regex used: r'^\\d+/[MD]$'")
        return None

    field_desc = list(df.index)
    bit_positions = df.loc[:, "BIT POSITION"] if "BIT POSITION" in df.columns else [""] * len(field_desc)
    sizes = df.loc[:, "Size (Bits)"] if "Size (Bits)" in df.columns else [""] * len(field_desc)

    output_df = pd.DataFrame({
        "FIELD NAME / DESCRIPTION": field_desc + ["CRC", "PAGE -X", "PAGE -Y"],
        "BIT POSITION": list(bit_positions) + ["Y63 - Y34", "", ""],
        "Size (Bits)": list(sizes) + [30, 64, 64]
    })

    nt_field_mapping = {
        'uctypeofTag': 'Type of Tag (9- Normal, 10 - LC, 11- Adj.Line ,12-Junction)',
        'uc_version': 'Version(As per Spec 4.0)',
        'uiUniqueID': 'Unique ID of RFID Tag Set',
        'fAbsLoc': 'Absolute Loc In  meters',
        'ucTagPlacement': 'Tag Placement(0-InL, 1- SIG(N), 2-SIG(R), 3-Tout, 4-Exit(N),5-Exit(R), 6-SIG(N/R),7-exit tag both directions, 8-Dead stopin Nominal,9- Dead stop in Reverse)',
        'AbsoluteLocationReset': 'Tag Duplication (0-Main tag,1-Dup tag)',
        'nominal': {
            'ucTin': 'TIN in Nominal Direction',
            'StationId': 'Station ID in Nominal Direction',
            'comMark': 'Communication in Nominal Direction  (0- Required, 1- Not Required).',
            'secType': 'Section type in Nominal Direction ( 0-Station, 1- Abs Blk, 2-Auto, 3- VBlk)'
        },
        'reverse': {
            'ucTin': 'TIN in Reverse Direction',
            'StationId': 'Station ID in Reverse Direction',
            'comMark': 'Communication in Reverse Direction. (0- Required, 1- Not Required).',
            'secType': 'Section type in Reverse Direction ( 0-Station, 1- Abs Blk, 2-Auto, 3- VBlk)'
        }
    }

    aline_field_mapping = {
        'uctypeofTag': 'Type of Tag (9- Normal, 10 - LC, 11- Adj.Line ,12-Junction)',
        'uc_version': 'Version(As per Spec 4.0)',
        'uiUniqueID': 'Unique ID of RFID Tag Set',
        'fAbsLoc': 'Absolute Loc In  meters',
        'nominal': {
            'ucTin': 'TIN in Nominal Direction'
        },
        'reverse': {
            'ucTin': 'TIN in Reverse Direction'
        },
        'AdjLine1_tin': 'Adjacent Line-1 TIN',
        'AdjLine2_tin': 'Adjacent Line-2 TIN',
        'AdjLine3_tin': 'Adjacent Line-3 TIN',
        'AdjLine4_tin': 'Adjacent Line-4 TIN',
        'AdjLine5_tin': 'Adjacent Line-5 TIN',
        'ucTagDuplication': 'Tag Duplication (0-Main tag,1-Dup tag)'
    }

    adj_field_mapping = {
        'uctypeofTag': 'Type of Tag (9- Normal, 10 - LC, 11- Adj.Line ,12-Junction)',
        'uc_version': 'Version(As per Spec 4.0)',
        'uiUniqueID': 'Unique ID of RFID Tag Set',
        'fAbsLoc1': 'Absolute Loc In  meters-1',
        'fAbsLoc2': 'Absolute Loc In  meters-2',
        'nominal': {
            'ucTin': 'TIN 1'
        },
        'reverse': {
            'ucTin': 'TIN 2'
        },
        'dirResetAbsLoc1': 'Direction reset absolute location-1',
        'dirResetAbsLoc2': 'Direction reset absolute location-2',
        'locCorrectionType': 'Location Correction Type',
        'reserved': 'Reserved',
        'secTypeNominal': 'Section type in nominal direction',
        'secTypeReverse': 'Section type in reverse direction',
        'reserved1': 'Reserved_1',
        'tagType': 'Tag type',
        'comMarkNominal': 'Communication in nominal direction',
        'comMarkReverse': 'Communication in reverse direction'
    }

    tag_data = {}
    for col in tag_cols:
        values = []
        try:
            for row in field_desc:
                try:
                    val = df.at[row, col]
                    if isinstance(val, pd.Series):
                        val = val.iloc[0]
                    values.append(val)
                except (KeyError, IndexError):
                    values.append("")

            if sheet_name == 'NT':
                uctypeofTag = int(df.at[nt_field_mapping['uctypeofTag'], col]) if nt_field_mapping['uctypeofTag'] in df.index else 9
                uc_version = int(df.at[nt_field_mapping['uc_version'], col]) if nt_field_mapping['uc_version'] in df.index else 1
                uiUniqueID = int(df.at[nt_field_mapping['uiUniqueID'], col]) if nt_field_mapping['uiUniqueID'] in df.index else 0
                fAbsLoc = int(df.at[nt_field_mapping['fAbsLoc'], col]) if nt_field_mapping['fAbsLoc'] in df.index and df.at[nt_field_mapping['fAbsLoc'], col] != 'UNKNOWN' else 0
                ucTagPlacement = int(df.at[nt_field_mapping['ucTagPlacement'], col]) if nt_field_mapping['ucTagPlacement'] in df.index else 0
                AbsoluteLocationReset = int(df.at[nt_field_mapping['AbsoluteLocationReset'], col]) if nt_field_mapping['AbsoluteLocationReset'] in df.index else 0

                nominal_ucTin = int(df.at[nt_field_mapping['nominal']['ucTin'], col]) if nt_field_mapping['nominal']['ucTin'] in df.index else 0
                nominal_StationId = int(df.at[nt_field_mapping['nominal']['StationId'], col]) if nt_field_mapping['nominal']['StationId'] in df.index else 0
                nominal_comMark = int(df.at[nt_field_mapping['nominal']['comMark'], col]) if nt_field_mapping['nominal']['comMark'] in df.index else 0
                nominal_secType = int(df.at[nt_field_mapping['nominal']['secType'], col]) if nt_field_mapping['nominal']['secType'] in df.index else 0

                reverse_ucTin = int(df.at[nt_field_mapping['reverse']['ucTin'], col]) if nt_field_mapping['reverse']['ucTin'] in df.index else 0
                reverse_StationId = int(df.at[nt_field_mapping['reverse']['StationId'], col]) if nt_field_mapping['reverse']['StationId'] in df.index else 0
                reverse_comMark = int(df.at[nt_field_mapping['reverse']['comMark'], col]) if nt_field_mapping['reverse']['comMark'] in df.index else 0
                reverse_secType = int(df.at[nt_field_mapping['reverse']['secType'], col]) if nt_field_mapping['reverse']['secType'] in df.index else 0

                tag = TagInfoNT(
                    uctypeofTag=uctypeofTag,
                    uc_version=uc_version,
                    uiUniqueID=uiUniqueID,
                    fAbsLoc=fAbsLoc,
                    ucTagPlacement=ucTagPlacement,
                    AbsoluteLocationReset=AbsoluteLocationReset,
                    stDir=[
                        TagDirNT(nominal_ucTin, nominal_StationId, nominal_comMark, nominal_secType),
                        TagDirNT(reverse_ucTin, reverse_StationId, reverse_comMark, reverse_secType)
                    ]
                )

                page1, page2, crc = calculate_values_nt(tag)
                def format_page(page: bytearray) -> str:
                    trimmed = page[:8]
                    while trimmed and trimmed[-1] == 0x00 and len(trimmed) > 6:
                        trimmed = trimmed[:-1]
                    return ''.join(f"{b:02x}" for b in trimmed)

                crc_str = f"{crc:08x}"
                page_x = format_page(page1)
                page_y = format_page(page2)

            elif sheet_name == 'AlineT':
                uctypeofTag = int(df.at[aline_field_mapping['uctypeofTag'], col]) if aline_field_mapping['uctypeofTag'] in df.index else 11
                uc_version = int(df.at[aline_field_mapping['uc_version'], col]) if aline_field_mapping['uc_version'] in df.index else 1
                uiUniqueID = int(df.at[aline_field_mapping['uiUniqueID'], col]) if aline_field_mapping['uiUniqueID'] in df.index else 0
                fAbsLoc = int(df.at[aline_field_mapping['fAbsLoc'], col]) if aline_field_mapping['fAbsLoc'] in df.index and df.at[aline_field_mapping['fAbsLoc'], col] != 'UNKNOWN' else 0
                nominal_ucTin = int(df.at[aline_field_mapping['nominal']['ucTin'], col]) if aline_field_mapping['nominal']['ucTin'] in df.index else 0
                reverse_ucTin = int(df.at[aline_field_mapping['reverse']['ucTin'], col]) if aline_field_mapping['reverse']['ucTin'] in df.index else 0
                AdjLine1_tin = int(df.at[aline_field_mapping['AdjLine1_tin'], col]) if aline_field_mapping['AdjLine1_tin'] in df.index else 0
                AdjLine2_tin = int(df.at[aline_field_mapping['AdjLine2_tin'], col]) if aline_field_mapping['AdjLine2_tin'] in df.index else 0
                AdjLine3_tin = int(df.at[aline_field_mapping['AdjLine3_tin'], col]) if aline_field_mapping['AdjLine3_tin'] in df.index else 0
                AdjLine4_tin = int(df.at[aline_field_mapping['AdjLine4_tin'], col]) if aline_field_mapping['AdjLine4_tin'] in df.index else 0
                AdjLine5_tin = int(df.at[aline_field_mapping['AdjLine5_tin'], col]) if aline_field_mapping['AdjLine5_tin'] in df.index else 0
                ucTagDuplication = int(df.at[aline_field_mapping['ucTagDuplication'], col]) if aline_field_mapping['ucTagDuplication'] in df.index else 0
                
                tag = TagInfoAline(
                    uctypeofTag=uctypeofTag,
                    uc_version=uc_version,
                    uiUniqueID=uiUniqueID,
                    fAbsLoc=fAbsLoc,
                    stDir=[
                        TagDirAline(nominal_ucTin),
                        TagDirAline(reverse_ucTin)
                    ],
                    AdjLine1_tin=AdjLine1_tin,
                    AdjLine2_tin=AdjLine2_tin,
                    AdjLine3_tin=AdjLine3_tin,
                    AdjLine4_tin=AdjLine4_tin,
                    AdjLine5_tin=AdjLine5_tin,
                    ucTagDuplication=ucTagDuplication
                )

                result = calculate_values_aline(tag)
                crc_str = f"{result.crc:08X}".lower()
                page_x = result.page_x.hex().lower()
                page_y = result.page_y.hex().lower()

            elif sheet_name == 'AdjT':
                uctypeofTag = int(df.at[adj_field_mapping['uctypeofTag'], col]) if adj_field_mapping['uctypeofTag'] in df.index else 12
                uc_version = int(df.at[adj_field_mapping['uc_version'], col]) if adj_field_mapping['uc_version'] in df.index else 1
                uiUniqueID = int(df.at[adj_field_mapping['uiUniqueID'], col]) if adj_field_mapping['uiUniqueID'] in df.index else 0
                fAbsLoc1 = int(df.at[adj_field_mapping['fAbsLoc1'], col]) if adj_field_mapping['fAbsLoc1'] in df.index and df.at[adj_field_mapping['fAbsLoc1'], col] != 'UNKNOWN' else 0
                fAbsLoc2 = int(df.at[adj_field_mapping['fAbsLoc2'], col]) if adj_field_mapping['fAbsLoc2'] in df.index and df.at[adj_field_mapping['fAbsLoc2'], col] != 'UNKNOWN' else 0
                nominal_ucTin = int(df.at[adj_field_mapping['nominal']['ucTin'], col]) if adj_field_mapping['nominal']['ucTin'] in df.index else 0
                reverse_ucTin = int(df.at[adj_field_mapping['reverse']['ucTin'], col]) if adj_field_mapping['reverse']['ucTin'] in df.index else 0
                dirResetAbsLoc1 = int(df.at[adj_field_mapping['dirResetAbsLoc1'], col]) if adj_field_mapping['dirResetAbsLoc1'] in df.index else 0
                dirResetAbsLoc2 = int(df.at[adj_field_mapping['dirResetAbsLoc2'], col]) if adj_field_mapping['dirResetAbsLoc2'] in df.index else 0
                locCorrectionType = int(df.at[adj_field_mapping['locCorrectionType'], col]) if adj_field_mapping['locCorrectionType'] in df.index else 0
                reserved = int(df.at[adj_field_mapping['reserved'], col]) if adj_field_mapping['reserved'] in df.index else 0
                secTypeNominal = int(df.at[adj_field_mapping['secTypeNominal'], col]) if adj_field_mapping['secTypeNominal'] in df.index else 0
                secTypeReverse = int(df.at[adj_field_mapping['secTypeReverse'], col]) if adj_field_mapping['secTypeReverse'] in df.index else 0
                reserved1 = int(df.at[adj_field_mapping['reserved1'], col]) if adj_field_mapping['reserved1'] in df.index else 0
                tagType = int(df.at[adj_field_mapping['tagType'], col]) if adj_field_mapping['tagType'] in df.index else 0
                comMarkNominal = int(df.at[adj_field_mapping['comMarkNominal'], col]) if adj_field_mapping['comMarkNominal'] in df.index else 0
                comMarkReverse = int(df.at[adj_field_mapping['comMarkReverse'], col]) if adj_field_mapping['comMarkReverse'] in df.index else 0

                tag = TagInfoAdj(
                    uctypeofTag=uctypeofTag,
                    uc_version=uc_version,
                    uiUniqueID=uiUniqueID,
                    stDir=[
                        TagDirAdj(nominal_ucTin, secTypeNominal, dirResetAbsLoc1, comMarkNominal, fAbsLoc1),
                        TagDirAdj(reverse_ucTin, secTypeReverse, dirResetAbsLoc2, comMarkReverse, fAbsLoc2)
                    ],
                    locCorrectionType=locCorrectionType,
                    reserved=reserved,
                    reserved1=reserved1,
                    tagType=tagType
                )
                
                page1, page2, crc = calculate_values_adj(tag)
                def format_page(page: bytearray) -> str:
                    trimmed = page[:8]
                    while trimmed and trimmed[-1] == 0x00 and len(trimmed) > 6:
                        trimmed = trimmed[:-1]
                    return ''.join(f"{b:02x}" for b in trimmed)
                
                crc_str = f"{crc:08x}"
                page_x = format_page(page1)
                page_y = format_page(page2)

            else:
                crc_str = 0
                page_x = '0 0 0 0 0 0 0 0'
                page_y = '0 0 0 0 0 0 0 0'

            values += [crc_str, page_x, page_y]
            tag_data[str(col)] = values

        except Exception as e:
            print(f"Error processing column {col} in sheet {sheet_name}: {e}")
            values = [""] * len(field_desc) + [0, '0 0 0 0 0 0 0 0', '0 0 0 0 0 0 0 0']
            tag_data[str(col)] = values

    if tag_data:
        tag_df = pd.DataFrame(tag_data, index=output_df.index)
        output_df = pd.concat([output_df, tag_df], axis=1)

    return output_df

def process_pdf():
    try:
        if 'input_path' not in request.form or 'output_path' not in request.form:
            return jsonify({"error": "Missing input or output path"}), 400

        input_path = request.form['input_path']
        output_path = request.form['output_path']
        
        if not input_path.endswith('.xlsx'):
            return jsonify({"error": "Invalid input file format. Please provide an .xlsx file path"}), 400

        if not os.path.isfile(input_path):
            return jsonify({"error": "Input file does not exist"}), 400

        if not os.path.isdir(os.path.dirname(output_path)):
            return jsonify({"error": "Invalid output directory"}), 400
        
        drive = os.path.splitdrive(output_path)[0]
        if drive and not os.path.exists(drive):
            return jsonify({"error": f"Drive {drive} does not exist or is not accessible"}), 400

        # Create full output directory if not present
        if not os.path.exists(output_path):
            os.makedirs(output_path)

        # Generate output file name by appending '_formatted' to input file name
        input_filename = os.path.splitext(os.path.basename(input_path))[0]
        base_output_filename = f"{input_filename}_formatted"
        output_extension = '.pdf'
        output_file = os.path.join(output_path, f"{base_output_filename}{output_extension}")
        
        # Check for existing files and append (1), (2), etc., if necessary
        if os.path.isfile(output_file):
            os.remove(output_file)

        # Setup PDF document
        doc = SimpleDocTemplate(
            output_file,
            pagesize=(A3[1], A3[0]),  # Landscape modex
            rightMargin=10 * mm,
            leftMargin=10 * mm,
            topMargin=10 * mm,
            bottomMargin=0 * mm
        )
        elements = []

        all_sheets = pd.read_excel(input_path, sheet_name=None, header=1)
        all_raw = pd.read_excel(input_path, sheet_name=None, header=None)
        processed_sheets = 0


        for sheet_name, df in all_sheets.items():
            raw_df = all_raw[sheet_name]
            tag_title = str(raw_df.iloc[0, 0]) if not raw_df.empty else "TAG"

            formatted_df = process_input_sheet(sheet_name, df)
            if formatted_df is None:
                print(f"Skipping sheet {sheet_name} — no valid tag data.")
                continue

            sheet_elements = format_pdf_table(formatted_df, tag_title, sheet_name)
            elements.extend(sheet_elements)
            processed_sheets += 1

        if processed_sheets == 0:
            return jsonify({"error": "No sheets processed. Output file not saved."}), 400

        doc.build(elements)
        print(f"✅ Final PDF saved: {output_file}")

        return send_file(
            output_file,
            as_attachment=True,
            download_name=os.path.basename(output_file),
            mimetype="application/pdf"
        )

    except Exception as e:
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)