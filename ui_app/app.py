import os
import sys
import socket
import threading
import uuid
import logging
import webbrowser
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, abort

# Import the slide generator
from bbgrl.generator.generator import bbgrlslidegeneratorv1


def _detect_base_path() -> Path:
    # Support PyInstaller onefile extraction dir via _MEIPASS
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).parent

def _detect_runtime_dir() -> Path:
    # Where to write logs and other runtime files
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE_PATH = _detect_base_path()
RUNTIME_DIR = _detect_runtime_dir()
TEMPLATE_DIR = BASE_PATH / "templates"

# Ensure logging to a stable location next to the EXE/script
LOG_PATH = RUNTIME_DIR / "ui_app.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("ui_app")

app = Flask(__name__, template_folder=str(TEMPLATE_DIR))

# In-memory job store (simple for single-user desktop use)
JOBS = {}
# JOBS[job_id] = {
#     "percent": int,
#     "message": str,
#     "done": bool,
#     "error": str|None,
#     "output_path": str|None,
# }


def _update(job_id: str, percent: int, message: str):
    job = JOBS.get(job_id)
    if job is None:
        return
    job["percent"] = max(0, min(100, int(percent)))
    job["message"] = message


def _run_generation(job_id: str, date_str: str):
    try:
        JOBS[job_id] = {
            "percent": 0,
            "message": "Starting...",
            "done": False,
            "error": None,
            "output_path": None,
        }

        _update(job_id, 5, "Initializing generator")
        gen = bbgrlslidegeneratorv1()

        # Parse input date from YYYY-MM-DD
        _update(job_id, 10, f"Parsing date {date_str}")
        target_date = datetime.strptime(date_str, "%Y-%m-%d")

        # Fetch data (single call internally does Morning Prayer + Readings)
        _update(job_id, 15, "Fetching Morning Prayer and Readings")
        data = gen.fetch_live_liturgical_data(target_date)
        _update(job_id, 55, "Fetched data: parsing complete")

        # Create presentation
        _update(job_id, 60, "Creating PowerPoint presentation")
        # Use a per-job output folder to avoid conflicts
        out_dir = os.path.join("output_v2")
        output_path = gen.create_presentation_from_template(data, output_dir=out_dir)

        # All done!
        _update(job_id, 100, "Done. Ready to download")
        JOBS[job_id]["done"] = True
        JOBS[job_id]["output_path"] = output_path
    except Exception as e:
        logger.exception("Generation failed")
        JOBS[job_id]["done"] = True
        JOBS[job_id]["error"] = str(e)
        _update(job_id, 100, "Failed. See error")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start():
    date_str = request.form.get("date") or request.json.get("date") if request.is_json else None
    if not date_str:
        return jsonify({"error": "Missing date"}), 400
    try:
        # Validate date format early
        datetime.strptime(date_str, "%Y-%m-%d")
    except Exception:
        return jsonify({"error": "Invalid date format. Use YYYY-MM-DD."}), 400

    job_id = uuid.uuid4().hex
    t = threading.Thread(target=_run_generation, args=(job_id, date_str), daemon=True)
    t.start()
    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job id"}), 404
    resp = {
        "percent": job["percent"],
        "message": job["message"],
        "done": job["done"],
        "error": job["error"],
    }
    if job["done"] and not job["error"]:
        resp["download_url"] = f"/download/{job_id}"
    return jsonify(resp)


@app.route("/download/<job_id>")
def download(job_id: str):
    job = JOBS.get(job_id)
    if not job or not job.get("output_path") or job.get("error"):
        abort(404)
    path = job["output_path"]
    if not os.path.exists(path):
        abort(404)
    # Send as attachment so the browser downloads it
    filename = os.path.basename(path)
    return send_file(path, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    # Find a free localhost port starting from 5000
    def find_free_port(start_port: int = 5000, max_tries: int = 50) -> int:
        port = start_port
        for _ in range(max_tries):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                try:
                    s.bind(("127.0.0.1", port))
                    return port
                except OSError:
                    port += 1
        return start_port

    port = find_free_port()
    url = f"http://127.0.0.1:{port}"
    logger.info("Starting UI on %s", url)

    # Open the browser shortly after server starts
    def _open_browser():
        try:
            webbrowser.open(url, new=1)
        except Exception:
            logger.exception("Failed to open browser")

    threading.Timer(0.8, _open_browser).start()
    app.run(host="127.0.0.1", port=port, debug=False)
