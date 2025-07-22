# import os
# from flask import jsonify
# from watchdog.observers import Observer
# from watchdog.events import FileSystemEventHandler
# import threading
# import time

# REQUIRED_DOCS = ['docA.txt', 'docB.txt', 'docC.txt', 'docD.txt']
# watchers = {}  # path -> {found, completed}

# class DocWatchHandler(FileSystemEventHandler):
#     def __init__(self, path):
#         super().__init__()
#         self.path = path
#         self.found_docs = set(watchers[self.path]['found'])  # Use shared state

#     def on_created(self, event):
#         if not event.is_directory:
#             filename = os.path.basename(event.src_path)
#             if filename in REQUIRED_DOCS:
#                 self.found_docs.add(filename)
#                 watchers[self.path]['found'] = list(self.found_docs)

#             if self.found_docs == set(REQUIRED_DOCS):
#                 watchers[self.path]['completed'] = True


# def start_watcher(path):
#     if path in watchers:
#         return  # already watching

#     # Step 1: Check existing files
#     found_now = set(os.listdir(path)) & set(REQUIRED_DOCS)
#     completed = found_now == set(REQUIRED_DOCS)

#     watchers[path] = {
#         'found': list(found_now),
#         'completed': completed
#     }

#     # Step 2: If not completed, start watching
#     if not completed:
#         event_handler = DocWatchHandler(path)
#         observer = Observer()
#         observer.schedule(event_handler, path=path, recursive=False)
#         observer.start()

#         def monitor():
#             while not watchers[path]['completed']:
#                 time.sleep(1)
#             observer.stop()
#             observer.join()

#         threading.Thread(target=monitor, daemon=True).start()


# def watch_documents():
#     from flask import request
#     data = request.get_json()
#     folder_path = data.get("path")

#     if not os.path.exists(folder_path):
#         return jsonify({'error': 'Path does not exist'}), 400

#     start_watcher(folder_path)
#     return jsonify({'status': 'Watching started'}), 200

# def check_watch_status():
#     from flask import request
#     data = request.get_json()
#     path = data.get("path")

#     status = watchers.get(path, {'found': [], 'completed': False})
#     return jsonify(status), 200




from flask import request, jsonify
import os

REQUIRED_DOCS = ['docA.txt', 'docB.txt', 'docC.txt', 'docD.txt']

def check_documents():
    data = request.get_json()
    folder_path = data.get('path')

    found_files = []
    if not os.path.exists(folder_path):
        return jsonify({'found': found_files}), 200

    for doc in REQUIRED_DOCS:
        if os.path.isfile(os.path.join(folder_path, doc)):
            found_files.append(doc)

    return jsonify({'found': found_files}), 200
