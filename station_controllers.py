from flask import Flask, request, jsonify
import os
import shutil
from flask_cors import CORS


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
    
    