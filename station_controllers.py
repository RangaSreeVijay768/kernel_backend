from flask import Flask, request, jsonify
import os
import shutil
from flask_cors import CORS
import os
from flask import request, jsonify
from werkzeug.utils import secure_filename
import shutil


def get_stations():
    base_path = request.args.get('path')

    if not base_path or not os.path.isdir(base_path):
        return jsonify({'error': 'Invalid or missing path'}), 400

    try:
        # Get only folders
        folders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]
        return jsonify({'stations': folders})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    

def get_station_details():
    base_path = request.args.get('path')

    if not base_path or not os.path.isdir(base_path):
        return jsonify({'error': 'Invalid or missing path'}), 400

    try:
        result = []
        for root, dirs, files in os.walk(base_path):
            for name in dirs:
                result.append(os.path.join(root, name))
            for name in files:
                result.append(os.path.join(root, name))

        return jsonify({'contents': result})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    


def open_file():
    data = request.get_json()
    file_path = data.get("file_path")

    if not os.path.isfile(file_path):
        return jsonify({"success": False, "message": "File not found"}), 404

    try:
        # For Windows (use startfile or subprocess)
        os.startfile(file_path)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500

# ✅ BACKEND: upload_input_file
def upload_input_file():
    file = request.files.get('file')
    station = request.form.get('station')
    base_path = request.form.get('base_path')

    if not file or not station or not base_path:
        return jsonify({"success": False, "message": "Missing data"}), 400

    try:    
        filename = secure_filename(file.filename)
        target_dir = os.path.join(base_path, "Output_Documents", station, "inputs")
        os.makedirs(target_dir, exist_ok=True)
        file.save(os.path.join(target_dir, filename))
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


def copy_to_inputs_folder():
    data = request.json
    source_base = data.get('source_path')       # path where all stations (folders) exist
    station = data.get('station')               # selected station folder name
    output_path = data.get('output_path')       # where to paste files

    if not source_base or not station or not output_path:
        return jsonify({'error': 'Missing required fields'}), 400

    station_path = os.path.join(source_base, station)

    if not os.path.exists(station_path):
        return jsonify({'error': 'Selected station does not exist'}), 404

    try:
        if not os.path.exists(output_path):
            os.makedirs(output_path)

        for item in os.listdir(station_path):
            src = os.path.join(station_path, item)
            dst = os.path.join(output_path, item)

            if os.path.isdir(src):
                if os.path.exists(dst):
                    shutil.rmtree(dst)  # remove existing folder if present
                shutil.copytree(src, dst)
            elif os.path.isfile(src):
                shutil.copy2(src, dst)

        return jsonify({'message': 'Documents and folders copied successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    

def copy_files_to_output():
    data = request.json
    file_paths = data.get('file_paths', [])
    output_path = data.get('output_path')

    if not file_paths or not output_path:
        return jsonify({'error': 'file_paths and output_path are required'}), 400

    try:
        if not os.path.exists(output_path):
            os.makedirs(output_path)

        for file_path in file_paths:
            if os.path.isfile(file_path):
                filename = os.path.basename(file_path)
                dest_path = os.path.join(output_path, filename)
                shutil.copy2(file_path, dest_path)

        return jsonify({'message': 'Files copied successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    

# def copy_to_inputs_folder():
#     data = request.json
#     source_base = data.get('source_path')       # path where all stations (folders) exist
#     station = data.get('station')               # selected station folder name
#     output_path = data.get('output_path')       # where to paste files

#     if not source_base or not station or not output_path:
#         return jsonify({'error': 'Missing required fields'}), 400

#     station_path = os.path.join(source_base, station)

#     if not os.path.exists(station_path):
#         return jsonify({'error': 'Selected station does not exist'}), 404

#     try:
#         if not os.path.exists(output_path):
#             os.makedirs(output_path)

#         for filename in os.listdir(station_path):
#             src = os.path.join(station_path, filename)
#             dst = os.path.join(output_path, filename)

#             if os.path.isfile(src):
#                 shutil.copy2(src, dst)

#         return jsonify({'message': 'Documents copied successfully'})
#     except Exception as e:
#         return jsonify({'error': str(e)}), 500
    
    