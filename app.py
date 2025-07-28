from flask import Flask, request, jsonify
import os
import shutil
from flask_cors import CORS
from station_controllers import get_stations, copy_to_inputs_folder, copy_files_to_output
from file_checker import check_documents
from file_converter import process_excel
from pdf_converter import process_pdf

app = Flask(__name__)
CORS(app)
# app.config.from_object(Config)


@app.route('/api/stations', methods=['GET'])
def get_all_stations():
    return get_stations()


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


if __name__ == '__main__':
    app.run(debug=True)
