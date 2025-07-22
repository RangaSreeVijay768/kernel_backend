from flask import Flask, request, send_file, jsonify
import pandas as pd
import re
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from calculations_nt import calculate_values as calculate_values_nt, TagInfo as TagInfoNT, TagDir as TagDirNT
from calculations_aline import calculate_values as calculate_values_aline, TagInfo as TagInfoAline, TagDir as TagDirAline
from calculations_adj import calculate_values as calculate_values_adj, TagInfo as TagInfoAdj, TagDir as TagDirAdj
from math import ceil

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
    return {
        'border': colors.black,
        'fill': colors.HexColor("#D9D9D9"),
        'font': 'Times-Roman',
        'bold_font': 'Times-Bold',
        'font_size': 11,
        'header_font_size': 14,
        'align': 'CENTER'
    }

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

                result = calculate_values_nt(tag)
                crc_str = f"{result.crc:08X}".lower()
                page_x = result.page_x.hex().lower()
                page_y = result.page_y.hex().lower()

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
                    fAbsLoc1=fAbsLoc1,
                    fAbsLoc2=fAbsLoc2,
                    stDir=[
                        TagDirAdj(nominal_ucTin),
                        TagDirAdj(reverse_ucTin)
                    ],
                    dirResetAbsLoc1=dirResetAbsLoc1,
                    dirResetAbsLoc2=dirResetAbsLoc2,
                    locCorrectionType=locCorrectionType,
                    reserved=reserved,
                    secTypeNominal=secTypeNominal,
                    secTypeReverse=secTypeReverse,
                    reserved1=reserved1,
                    tagType=tagType,
                    comMarkNominal=comMarkNominal,
                    comMarkReverse=comMarkReverse
                )

                result = calculate_values_adj(tag)
                crc_str = f"{result.crc:08X}".lower()
                page_x = result.page_x.hex().lower()
                page_y = result.page_y.hex().lower()

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

def format_pdf_elements(elements, df, tag_title, chunk_index, total_chunks):
    styles = get_style_elements()
    chunk_size = 10
    predefined_cols = ["FIELD NAME / DESCRIPTION", "BIT POSITION", "Size (Bits)"]
    tag_columns = [col for col in df.columns if col not in predefined_cols]
    start = chunk_index * chunk_size
    end = start + chunk_size
    tag_chunk = tag_columns[start:end]
    temp_df = df[predefined_cols + tag_chunk]

    # Convert mm to points (1 mm = 2.83465 points)
    mm_to_pt = 2.83465
    col_widths = [20 * mm_to_pt, 20 * mm_to_pt, 20 * mm_to_pt, 35 * mm_to_pt, 35 * mm_to_pt] + [20 * mm_to_pt] * len(tag_chunk)
    row_heights = []

    # Calculate row heights
    for row_idx in range(temp_df.shape[0] + 1):  # +1 for header
        max_lines = 1
        for col_idx, value in enumerate(temp_df.iloc[row_idx - 1] if row_idx > 0 else temp_df.columns):
            if value:
                text = str(value)
                line_breaks = text.count('\n')
                wrapped_lines = len(text) // 40 + 1
                total_lines = max(line_breaks + 1, wrapped_lines)
                max_lines = max(max_lines, total_lines)
        height = (max_lines * 15 + 20 if max_lines > 2 else 40) * mm_to_pt / 2.83465
        row_heights.append(height)

    # Header
    header_style = ParagraphStyle(
        name='Header',
        fontName=styles['bold_font'],
        fontSize=styles['header_font_size'],
        alignment=1,  # Center
        spaceAfter=10 * mm_to_pt
    )
    elements.append(Paragraph(tag_title, header_style))

    # Table data
    data = [temp_df.columns.tolist()] + temp_df.values.tolist()
    table = Table(data, colWidths=col_widths, rowHeights=row_heights)

    # Table style
    table_style = TableStyle([
        ('FONT', (0, 0), (-1, -1), styles['font'], styles['font_size']),
        ('FONT', (0, 0), (-1, 0), styles['bold_font'], styles['font_size']),
        ('FONT', (0, 1), (2, -1), styles['bold_font'], styles['font_size']),
        ('ALIGN', (0, 0), (-1, -1), styles['align']),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, styles['border']),
        ('BACKGROUND', (0, 0), (2, 0), styles['fill']),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])
    # Merge cells for FIELD NAME / DESCRIPTION
    table_style.add('SPAN', (0, 1), (2, 1))
    for row_idx in range(1, len(data)):
        table_style.add('SPAN', (0, row_idx), (2, row_idx))
    table.setStyle(table_style)
    elements.append(table)

    # Footer
    try:
        from footer import extract_template_with_placeholders
        footer_template = extract_template_with_placeholders("template.xlsx", "A2:O6")
        footer_values = {
            "division": {"value": "Secunderabad", "size": 22},
            "railways": {"value": "South Central Railway", "size": 22},
            "station-name": {"value": "Malwan", "size": 22},
            "station-id": {"value": "MW123", "size": 22},
            "date": {"value": "11-07-2025", "size": 22},
            "total-pages": {"value": str(total_chunks), "size": 22},
            "current-page": {"value": str(chunk_index + 1), "size": 22}
        }
        footer_data = []
        for row in footer_template["template"]:
            footer_row = []
            for cell in row:
                if isinstance(cell, dict) and "placeholder" in cell:
                    value = footer_values.get(cell["placeholder"], {"value": ""})["value"]
                    footer_row.append(str(value))
                else:
                    footer_row.append(str(cell))
            footer_data.append(footer_row)
        if footer_data:
            footer_table = Table(footer_data, colWidths=[30 * mm_to_pt] * len(footer_data[0]))
            footer_style = TableStyle([
                ('FONT', (0, 0), (-1, -1), styles['font'], styles['font_size']),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, styles['border']),
            ])
            footer_table.setStyle(footer_style)
            elements.append(Spacer(1, 20 * mm_to_pt))
            elements.append(footer_table)
    except ImportError:
        print(f"Footer skipped for chunk {chunk_index + 1} in sheet with title: {tag_title}")

    elements.append(Spacer(1, 30 * mm_to_pt))

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

        # Generate output file name
        input_filename = os.path.splitext(os.path.basename(input_path))[0]
        base_output_filename = f"{input_filename}_formatted"
        output_extension = '.pdf'
        output_file = os.path.join(output_path, f"{base_output_filename}{output_extension}")
        
        # Check for existing files
        counter = 1
        while os.path.isfile(output_file):
            output_filename = f"{base_output_filename}({counter}){output_extension}"
            output_file = os.path.join(output_path, output_filename)
            counter += 1

        # Initialize PDF
        doc = SimpleDocTemplate(
            output_file,
            pagesize=A4,
            leftMargin=15 * mm,
            rightMargin=15 * mm,
            topMargin=15 * mm,
            bottomMargin=15 * mm
        )
        elements = []

        # Process Excel sheets
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

            chunk_size = 10
            tag_columns = [col for col in formatted_df.columns if col not in ["FIELD NAME / DESCRIPTION", "BIT POSITION", "Size (Bits)"]]
            total_chunks = ceil(len(tag_columns) / chunk_size)

            for chunk_index in range(total_chunks):
                format_pdf_elements(elements, formatted_df, tag_title, chunk_index, total_chunks)

            processed_sheets += 1

        if processed_sheets == 0:
            return jsonify({"error": "No sheets processed. Output file not saved."}), 400

        # Build PDF
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



