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

        for filename in os.listdir(station_path):
            src = os.path.join(station_path, filename)
            dst = os.path.join(output_path, filename)

            if os.path.isfile(src):
                shutil.copy2(src, dst)

        return jsonify({'message': 'Documents copied successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    