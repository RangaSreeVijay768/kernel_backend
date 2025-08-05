from flask import Flask, request, jsonify
import os
import shutil
from flask_cors import CORS
from controllers.station_controllers import get_stations, get_station_details, copy_to_inputs_folder, copy_files_to_output, open_file, upload_input_file
from components.file_checker import check_documents
from file_generators.tag_data_excel_formatted_generator import process_excel
from file_generators.tag_data_pdf_generator import process_pdf
from file_generators.rfid_pdf_generator import generate_pdf_with_rfid_image
from file_generators.toc_pdf_generator import generate_toc_pdf_with_header_footer

app = Flask(__name__)
CORS(app)
# app.config.from_object(Config)


@app.route('/api/stations', methods=['GET'])
def get_all_stations():
    return get_stations()

@app.route('/api/station_details', methods=['GET'])
def get_details_station():
    return get_station_details()

@app.route("/api/open-file", methods=["POST"])
def openFile():
    return open_file()

@app.route('/api/upload-input-file', methods=['POST'])
def uploadInputs():
    return upload_input_file()

@app.route('/api/run-documents', methods=['POST'])
def copy_to_inputs():
    return copy_to_inputs_folder()

@app.route('/api/copy-files', methods=['POST'])
def copy_files_to_outputs():
    return copy_files_to_output()
    

@app.route('/api/check-documents', methods=['POST'])
def start_doc_watch():
    return check_documents()

@app.route('/api/convert-file', methods=['POST'])
def process_excel_route():
    return process_excel()

@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf_route():
    return process_pdf()

@app.route('/api/generate-rfid-pdf', methods=['POST'])
def generate_rdid_pdf_route():
    return generate_pdf_with_rfid_image()

@app.route('/api/generate-toc-pdf', methods=['POST'])
def generate_toc_pdf():
    return generate_toc_pdf_with_header_footer()


if __name__ == '__main__':
    app.run(debug=True)


# if __name__ == '__main__':
#     app.run(
#         host='0.0.0.0',
#         port=5000,
#         debug=True,
#         ssl_context=('/etc/letsencrypt/live/bhepl.com/cert.pem', '/etc/letsencrypt/live/bhepl.com/key.pem')
#     )
