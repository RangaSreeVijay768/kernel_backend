from flask import request, jsonify, send_file
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from bottom_left_tables import draw_title_block_on_pdf, draw_bottom_right_block_on_pdf, draw_direction_arrows
from PIL import Image
import os

def generate_pdf_with_rfid_image():
    try:
        image_path = request.form.get('image_path')
        output_path = request.form.get('output_path')
        top_left_html = request.form.get('top_left_html', '')
        top_center_html = request.form.get('top_center_html', '')
        top_right_html = request.form.get('top_right_html', '')

        if not image_path or not output_path:
            return jsonify({"error": "Missing image path or output path"}), 400

        if not os.path.isfile(image_path):
            return jsonify({"error": "Image file does not exist"}), 400

        if not os.path.isdir(output_path):
            return jsonify({"error": "Invalid output directory"}), 400

        output_file = os.path.join(output_path, "RFID_TIN_Layout.pdf")
        if os.path.exists(output_file):
            os.remove(output_file)

        with Image.open(image_path) as img:
            img_width, img_height = img.size

        outer_margin = 1
        border_thickness = 1
        inner_margin = 16
        content_padding = 20
        top_text_height = 80
        title_block_width = 400
        title_block_height = 300
        gap_below_image = 10
        
        content_height = img_height + top_text_height + gap_below_image + title_block_height
        page_height = content_height + 2 * (inner_margin + content_padding + outer_margin)

        content_width = img_width
        page_width = content_width + 2 * (inner_margin + content_padding + outer_margin)

        c = canvas.Canvas(output_file, pagesize=(page_width, page_height))

        # Draw borders
        outer_rect = {'x0': outer_margin, 'y0': outer_margin, 'x1': page_width - outer_margin, 'y1': page_height - outer_margin}
        inner_rect = {
            'x0': outer_rect['x0'] + inner_margin, 'y0': outer_rect['y0'] + inner_margin,
            'x1': outer_rect['x1'] - inner_margin, 'y1': outer_rect['y1'] - inner_margin
        }
        content_rect = {
            'x0': inner_rect['x0'] + content_padding, 'y0': inner_rect['y0'] + content_padding,
            'x1': inner_rect['x1'] - content_padding, 'y1': inner_rect['y1'] - content_padding
        }

        c.setLineWidth(border_thickness)
        c.rect(outer_rect['x0'], outer_rect['y0'], outer_rect['x1'] - outer_rect['x0'], outer_rect['y1'] - outer_rect['y0'])
        c.rect(inner_rect['x0'], inner_rect['y0'], inner_rect['x1'] - inner_rect['x0'], inner_rect['y1'] - inner_rect['y0'])

        # Draw border circles
        circle_radius = 6
        top_bottom_circles = 12
        left_right_circles = 4

        def draw_side_circles(start_x, start_y, dx, dy, count, spacing, offset):
            for i in range(1, count + 1):
                x = start_x + (spacing * i * dx)
                y = start_y + (spacing * i * dy)
                c.circle(x + offset[0], y + offset[1], circle_radius, stroke=1, fill=0)

        width = outer_rect['x1'] - outer_rect['x0']
        height = outer_rect['y1'] - outer_rect['y0']
        top_spacing = width / (top_bottom_circles + 1)
        side_spacing = height / (left_right_circles + 1)

        draw_side_circles(outer_rect['x0'], outer_rect['y1'], 1, 0, top_bottom_circles, top_spacing, (0, -circle_radius / 0.72))
        draw_side_circles(outer_rect['x0'], outer_rect['y0'], 1, 0, top_bottom_circles, top_spacing, (0, circle_radius / 0.72))
        draw_side_circles(outer_rect['x0'], outer_rect['y0'], 0, 1, left_right_circles, side_spacing, (circle_radius / 0.72, 0))
        draw_side_circles(outer_rect['x1'], outer_rect['y0'], 0, 1, left_right_circles, side_spacing, (-circle_radius / 0.72, 0))

        # Top Text Columns (Left, Center, Right) using HTML
        styles = getSampleStyleSheet()
        def make_style(alignment):
            return ParagraphStyle(
                'CustomHTMLStyle',
                parent=styles['Normal'],
                fontName='Helvetica',
                alignment=alignment,
                leading=16,
                allowOrphans=1,
                allowWidows=1
            )

        text_y = content_rect['y1'] - top_text_height
        col_width = (content_rect['x1'] - content_rect['x0']) / 3

        # Define Frames
        frame_left = Frame(content_rect['x0'], text_y, col_width, top_text_height, showBoundary=0)
        frame_center = Frame(content_rect['x0'] + col_width, text_y, col_width, top_text_height, showBoundary=0)
        frame_right = Frame(content_rect['x0'] + 2 * col_width, text_y, col_width, top_text_height, showBoundary=0)
        
        
        # Draw arrows above left frame
        draw_direction_arrows(c, x_center=frame_left._x1 + (frame_left._x2 - frame_left._x1) / 2,
                            y_center=frame_left._y1 + 10)

        # Draw arrows above right frame
        draw_direction_arrows(c, x_center=frame_right._x1 + (frame_right._x2 - frame_right._x1) / 2,
                            y_center=frame_right._y1 + 10)


        # Add HTML content as Paragraphs
        frame_left.addFromList([Paragraph(top_left_html, make_style(TA_LEFT))], c)
        frame_center.addFromList([Paragraph(top_center_html, make_style(TA_CENTER))], c)
        frame_right.addFromList([Paragraph(top_right_html, make_style(TA_RIGHT))], c)

        # Image below text
        title_block_y = content_rect['y0']
        image_y = title_block_y + title_block_height + gap_below_image

        # Draw image first (above title block)
        c.drawImage(ImageReader(image_path), content_rect['x0'], image_y, width=img_width, height=img_height)

        # Draw title block at bottom
        draw_title_block_on_pdf(
            c,
            x0=content_rect['x0'],
            y0=title_block_y,
            width=title_block_width,
            total_height=title_block_height * 0.7
        )
        
        draw_bottom_right_block_on_pdf(
            c=c,
            x0=content_rect['x1'] - title_block_width,  # right-aligned X
            y0=title_block_y + 5,
            width=title_block_width,
            height=title_block_height * 0.7
        )

        c.showPage()
        c.save()

        return send_file(
            output_file,
            as_attachment=True,
            download_name="schematic_plan.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return jsonify({"error": f"PDF generation failed: {str(e)}"}), 500
