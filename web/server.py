import sys
import os
from flask import Flask, render_template_string, send_from_directory
from flask_socketio import SocketIO
from flask import redirect, url_for
import threading
import time

app = Flask(__name__, template_folder='_output', static_folder='_output')
socketio = SocketIO(app, cors_allowed_origins="*")

BASE_DIR = os.path.abspath("_output")
WEB_DIR = os.path.join(BASE_DIR, "web")

@app.route('/')
def index():
    with open(os.path.join(BASE_DIR, "dashboard.html"), "r", encoding="utf-8") as file:
        content = file.read()
    return render_template_string(content)

@app.route('/web/<path:filename>')
def serve_web_files(filename):
    return send_from_directory(WEB_DIR, filename)

@app.route('/dashboard.html')
def redirect_to_index():
    return redirect(url_for('index'))

def _log(text):
    print(text)

def serverStart(port=2000, debug=True, logger=_log):
    def run():
        try:
            logger(f"[WS] Web server starting on {port} port...")
            # socketio.run заблокирует этот поток до вызова socketio.stop()
            socketio.run(app, host='0.0.0.0', port=port, debug=debug, use_reloader=False)
            logger("[WS] Web server stopped.")
        except Exception as e:
            logger(f"[WS] ERROR: {e}")

    thread = threading.Thread(target=run, daemon=True)
    thread.start()

    def shutdown():
        socketio.stop()
        thread.join()

    return thread, shutdown

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=2000, debug=True)
