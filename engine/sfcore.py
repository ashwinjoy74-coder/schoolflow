# Stateless SchoolFlow API for the website.
#
# The browser keeps the school's files (requirements CSV and timetable workbook)
# and sends them with every call. Each call unpacks them into a private temporary
# folder, runs the desktop app's own Api method on them, returns the result and
# any files that changed, then deletes the folder. Nothing is kept on the server,
# so the same code runs on a laptop (server.py) or a serverless host (api/index.py).

import base64
import json
import os
import re
import shutil
import tempfile
import threading
import time

os.environ.setdefault("SCHOOLFLOW_DATA", os.path.join(tempfile.gettempdir(), "schoolflow"))

import school_timetable_generator as generator  # noqa: E402
import schoolflow_paths as paths  # noqa: E402
from schoolflow_api import MIN_PERIODS_BY_DAY_COUNT, Api  # noqa: E402

MAX_SOLVER_SECONDS = int(os.environ.get("SCHOOLFLOW_MAX_SOLVER_SECONDS", "240"))
REMOTE_METHODS = {
    "summary", "requirements", "save_requirements", "check_requirements", "generate", "timetables",
    "lesson_options", "set_lesson", "split_options", "split_candidates", "split",
}
DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


class _CallApi(Api):
    """The desktop Api, pointed at one call's temporary files and never saving settings."""

    def __init__(self, input_path, output_path, days, periods, time_limit):
        self._settings = dict(paths.DEFAULT_SETTINGS, input_path=input_path, output_path=output_path,
                              days=days, periods_per_day=periods, solver_time_limit=time_limit)
        self._window = None
        self._dirty = False
        self._job = {"running": False, "log": [], "result": None, "started": None, "finished": None}
        self._job_lock = threading.Lock()
        self._timetable_cache = None

    def _save_settings(self):
        pass

    def open_path(self, path):
        return {"ok": True}


def _clean_name(name, default, extension):
    name = os.path.basename(str(name or "")).strip()
    name = re.sub(r"[^\w .()-]", "_", name)[:80] or default
    return name if name.lower().endswith(extension) else name + extension


def info():
    try:
        import ortools
        ortools_version = ortools.__version__
    except Exception:
        ortools_version = "not installed"
    import openpyxl
    return {
        "ok": True,
        "version": paths.APP_VERSION,
        "engines": "Google OR-Tools " + ortools_version + " · openpyxl " + openpyxl.__version__,
        "classes": [str(grade) for grade in generator.CLASSES],
        "divisions": list(generator.DIVISIONS),
        "minPeriods": {str(key): value for key, value in MIN_PERIODS_BY_DAY_COUNT.items()},
        "maxSolverSeconds": MAX_SOLVER_SECONDS,
    }


def sample():
    with open(paths.SAMPLE_REQUIREMENTS, "r", encoding="utf-8-sig") as file:
        return {"ok": True, "text": file.read()}


def handle(payload):
    """Run one API call. payload = {method, args, files: {...}, settings: {...}}."""
    method = str(payload.get("method") or "")
    if method == "info":
        return info()
    if method == "sample":
        return sample()
    if method not in REMOTE_METHODS:
        return {"ok": False, "error": "Unknown action: " + method}

    args = payload.get("args") or []
    files = payload.get("files") or {}
    settings = payload.get("settings") or {}
    days = [day for day in DAY_NAMES if day in (settings.get("days") or DAY_NAMES[:5])]
    periods = int(settings.get("periods_per_day") or 7)
    time_limit = max(10, min(int(settings.get("solver_time_limit") or 180), MAX_SOLVER_SECONDS))
    req_name = _clean_name(settings.get("requirements_name"), "requirements.csv", ".csv")
    wb_name = _clean_name(settings.get("workbook_name"), "school_timetable.xlsx", ".xlsx")

    folder = tempfile.mkdtemp(prefix="schoolflow-")
    try:
        input_path = os.path.join(folder, req_name)
        output_path = os.path.join(folder, wb_name)
        req_text = files.get("requirements")
        if isinstance(req_text, str):
            with open(input_path, "w", encoding="utf-8-sig", newline="") as file:
                file.write(req_text)
        wb_in = None
        if files.get("workbook"):
            wb_in = base64.b64decode(files["workbook"])
            with open(output_path, "wb") as file:
                file.write(wb_in)
            modified = files.get("workbookModified")
            if isinstance(modified, (int, float)) and modified > 0:
                os.utime(output_path, (modified, modified))

        api = _CallApi(input_path, output_path, days, periods, time_limit)
        if method == "generate":
            started = api.generate()
            if not started.get("ok"):
                return started
            while api._job["running"]:
                time.sleep(0.1)
            result = api.job(0)
        else:
            result = getattr(api, method)(*args)

        changed = {}
        if os.path.isfile(input_path):
            with open(input_path, "r", encoding="utf-8-sig") as file:
                text = file.read()
            if text != req_text:
                changed["requirements"] = text
        if os.path.isfile(output_path):
            with open(output_path, "rb") as file:
                data = file.read()
            if data != wb_in:
                changed["workbook"] = base64.b64encode(data).decode("ascii")
                changed["workbookModified"] = time.time()

        # Show the browser's own file names instead of the temporary folder.
        encoded = json.dumps(result, default=str)
        for real, shown in ((input_path, "Requirements/" + req_name), (output_path, "Exports/" + wb_name),
                            (folder, "Requirements")):
            encoded = encoded.replace(json.dumps(real)[1:-1], shown)
        result = json.loads(encoded)
        if isinstance(result, dict):
            result["_files"] = changed
        return result
    finally:
        shutil.rmtree(folder, ignore_errors=True)
