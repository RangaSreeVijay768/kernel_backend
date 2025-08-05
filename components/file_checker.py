from flask import request, jsonify
import os

def check_documents():
    data = request.get_json()
    base_path = data.get('path')

    found_files = []

    if not base_path or not os.path.exists(base_path):
        return jsonify({'found': found_files}), 200

    # Define expected folders and the files they must contain
    expected_structure = {
        'tables': [
            'RCT_Malwan_station.xlsx',
            'TD_Malwan_station.xlsx',
            'TOC_Malwan_station.xlsx'
        ],
        'schematic_plan': [
            'RFID_TIN_layout.bmp'
        ]
    }

    # Check each required folder and its files
    for folder, files in expected_structure.items():
        full_folder_path = os.path.join(base_path, folder)
        if not os.path.isdir(full_folder_path):
            continue  # Skip if the folder doesn't exist

        for file in files:
            file_path = os.path.join(full_folder_path, file)
            if os.path.isfile(file_path):
                found_files.append(file)

    return jsonify({'found': found_files}), 200
