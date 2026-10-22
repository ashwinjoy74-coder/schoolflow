# The Python side of the app window. Every public method here can be called from
# the interface's JavaScript as `await window.pywebview.api.<name>(...)` and
# returns plain JSON-friendly data. Errors are returned as {"ok": False, "error": ...}.

import csv
import os
import sys
import tempfile
import threading
import time

import school_timetable_generator as generator
import schoolflow_lessons as lessons
import schoolflow_paths as paths
from schoolflow_workbook import read_timetables

REQUIRED_COLUMNS = ["class", "division", "subject", "teacher", "periods_per_week"]
MIN_PERIODS_BY_DAY_COUNT = {2: 12, 3: 7, 4: 5, 5: 4, 6: 4, 7: 3}


# Wrap a method so any error reaches the interface as a readable message.
def _safe(method):
    def wrapper(self, *args):
        try:
            result = method(self, *args)
            if isinstance(result, dict) and "ok" not in result:
                result["ok"] = True
            return result
        except (ValueError, RuntimeError, OSError, KeyError) as error:
            return {"ok": False, "error": str(error)}
    wrapper.__name__ = method.__name__
    return wrapper


def _dialog_constant(name, legacy):
    import webview
    dialog = getattr(webview, "FileDialog", None)
    return getattr(dialog, name) if dialog is not None else getattr(webview, legacy)


class Api:
    def __init__(self):
        self._settings = paths.load_settings()
        if not os.path.isfile(self._settings["input_path"]):
            self._settings["input_path"] = paths.DEFAULT_INPUT
        if not self._settings["output_path"]:
            self._settings["output_path"] = paths.DEFAULT_OUTPUT
        self._window = None
        self._dirty = False
        self._job = {"running": False, "log": [], "result": None, "started": None, "finished": None}
        self._job_lock = threading.Lock()
        self._timetable_cache = None

    def _attach(self, window):
        self._window = window

    def _save_settings(self):
        paths.save_settings(self._settings)

    # ---- Start-up and shared state

    @_safe
    def bootstrap(self):
        try:
            import ortools
            ortools_version = ortools.__version__
        except Exception:
            ortools_version = "not installed"
        import openpyxl
        return {
            "version": paths.APP_VERSION,
            "mode": paths.MODE,
            "dataFolder": paths.DATA_FOLDER,
            "requirementsFolder": paths.REQUIREMENTS_FOLDER,
            "exportsFolder": paths.EXPORTS_FOLDER,
            "settingsFile": paths.SETTINGS_FILE,
            "engines": "Google OR-Tools " + ortools_version + " · openpyxl " + openpyxl.__version__,
            "settings": self._settings,
            "classes": [str(grade) for grade in generator.CLASSES],
            "divisions": list(generator.DIVISIONS),
            "minPeriods": {str(key): value for key, value in MIN_PERIODS_BY_DAY_COUNT.items()},
        }

    @_safe
    def save_settings(self, patch):
        for key, value in (patch or {}).items():
            if key in paths.DEFAULT_SETTINGS and isinstance(value, type(paths.DEFAULT_SETTINGS[key])):
                self._settings[key] = value
        self._save_settings()
        return {"settings": self._settings}

    # Facts for the Home page: requirement counts and the latest workbook.
    @_safe
    def summary(self):
        result = {"requirements": None, "workbook": None}
        path = self._settings["input_path"]
        if os.path.isfile(path):
            rows, divisions, periods = 0, set(), 0
            subjects, teachers, groups = set(), set(), set()
            with open(path, "r", newline="", encoding="utf-8-sig") as file:
                for row in csv.DictReader(file):
                    subject = (row.get("subject") or "").strip()
                    teacher = (row.get("teacher") or "").strip()
                    if not subject and not teacher:
                        continue
                    group = (row.get("group") or "").strip()
                    if group:  # group activity rows are counted as groups, not subject rows
                        groups.add(((row.get("class") or "").strip(), group.casefold()))
                        teachers.add(teacher.casefold())
                        continue
                    rows += 1
                    divisions.add((row.get("class") or "").strip() + "-" + (row.get("division") or "").strip().upper())
                    try:
                        count = int((row.get("periods_per_week") or "0").strip())
                    except ValueError:
                        count = 0
                    periods += count
                    subjects.add(subject.casefold())
                    teachers.add(teacher.casefold())
            result["requirements"] = {"rows": rows, "divisions": len(divisions), "teachers": len(teachers),
                                      "subjects": len(subjects), "periods": periods, "groups": len(groups)}
        output = self._settings["output_path"]
        if os.path.isfile(output):
            result["workbook"] = {"path": output, "modified": os.path.getmtime(output)}
        return result

    # ---- Requirements

    @_safe
    def requirements(self):
        path = self._settings["input_path"]
        if not os.path.isfile(path):
            return {"path": path, "fieldnames": REQUIRED_COLUMNS, "rows": [], "missing": True}
        with open(path, "r", newline="", encoding="utf-8-sig") as file:
            reader = csv.DictReader(file)
            fieldnames = reader.fieldnames or []
            missing = [name for name in REQUIRED_COLUMNS if name not in fieldnames]
            if missing:
                raise ValueError("The CSV is missing these columns: " + ", ".join(missing))
            rows = [{key: (value or "").strip() for key, value in row.items() if key}
                    for row in reader if any((row.get(key) or "").strip() for key in REQUIRED_COLUMNS)]
        for row in rows:
            row["division"] = row.get("division", "").upper()
        return {"path": path, "fieldnames": fieldnames, "rows": rows}

    def _write_csv(self, path, fieldnames, rows):
        temporary = path + ".tmp"
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        try:
            with open(temporary, "w", newline="", encoding="utf-8-sig") as file:
                writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(rows)
            os.replace(temporary, path)
        except PermissionError:
            raise ValueError("The requirements file is open in another program (probably Excel). Close it and save again.")
        finally:
            if os.path.exists(temporary):
                os.remove(temporary)

    @_safe
    def save_requirements(self, fieldnames, rows):
        fieldnames = list(fieldnames or REQUIRED_COLUMNS)
        for column in REQUIRED_COLUMNS:
            if column not in fieldnames:
                fieldnames.append(column)
        self._write_csv(self._settings["input_path"], fieldnames, rows)
        return {"saved": len(rows)}

    # Validate rows with the scheduler's own rules (without solving).
    @_safe
    def check_requirements(self, fieldnames, rows):
        if self._job["running"]:
            raise ValueError("A timetable is being generated. Try again when it finishes.")
        handle, temporary = tempfile.mkstemp(suffix=".csv")
        os.close(handle)
        try:
            self._write_csv(temporary, list(fieldnames or REQUIRED_COLUMNS), rows)
            return generator.check_requirements(temporary, self._settings["days"], self._settings["periods_per_day"])
        finally:
            if os.path.exists(temporary):
                os.remove(temporary)

    @_safe
    def choose_requirements(self):
        current = self._settings["input_path"]
        folder = os.path.dirname(current) if os.path.isdir(os.path.dirname(current)) else paths.REQUIREMENTS_FOLDER
        chosen = self._window.create_file_dialog(_dialog_constant("OPEN", "OPEN_DIALOG"), directory=folder,
                                                 file_types=("CSV files (*.csv)", "All files (*.*)"))
        if not chosen:
            return {"changed": False}
        self._settings["input_path"] = chosen[0] if isinstance(chosen, (list, tuple)) else chosen
        self._save_settings()
        return {"changed": True, "path": self._settings["input_path"]}

    @_safe
    def choose_output(self):
        current = self._settings["output_path"]
        folder = os.path.dirname(current) if os.path.isdir(os.path.dirname(current)) else paths.EXPORTS_FOLDER
        chosen = self._window.create_file_dialog(_dialog_constant("SAVE", "SAVE_DIALOG"), directory=folder,
                                                 save_filename=os.path.basename(current),
                                                 file_types=("Excel workbook (*.xlsx)",))
        if not chosen:
            return {"changed": False}
        path = chosen[0] if isinstance(chosen, (list, tuple)) else chosen
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self._settings["output_path"] = path
        self._save_settings()
        return {"changed": True, "path": path}

    @_safe
    def restore_sample(self):
        paths.restore_sample_requirements()
        self._settings["input_path"] = paths.DEFAULT_INPUT
        self._save_settings()
        return {"path": paths.DEFAULT_INPUT}

    # ---- Generation (runs in the background; the interface polls job())

    @_safe
    def generate(self):
        with self._job_lock:
            if self._job["running"]:
                raise ValueError("A timetable is already being generated.")
            settings = self._settings
            output = settings["output_path"]
            if not os.path.isfile(settings["input_path"]):
                raise ValueError("The requirements file was not found:\n" + settings["input_path"])
            if not output.lower().endswith(".xlsx"):
                raise ValueError("The workbook name must end in .xlsx.")
            if os.path.isfile(output):
                try:
                    with open(output, "a+b"):
                        pass
                except PermissionError:
                    raise ValueError("Close " + os.path.basename(output) + " in Excel, then generate again.")
            days = list(settings["days"])
            periods = int(settings["periods_per_day"])
            if len(days) < 2:
                raise ValueError("Choose at least two school days.")
            if periods < MIN_PERIODS_BY_DAY_COUNT.get(len(days), 12) or periods > 12:
                raise ValueError("Choose a valid number of periods per day for " + str(len(days)) + " days.")
            self._job = {"running": True, "log": [], "result": None, "started": time.time(), "finished": None}
            self._log("Started: " + str(len(days)) + " days × " + str(periods) + " periods.", "accent")
        threading.Thread(target=self._run, args=(settings["input_path"], output, days, periods,
                                                 settings["solver_time_limit"]), daemon=True).start()
        return {"started": True}

    def _log(self, message, kind="info"):
        self._job["log"].append({"time": time.strftime("%H:%M:%S"), "text": message, "kind": kind})

    def _run(self, input_path, output, days, periods, time_limit):
        try:
            result = generator.generate(input_path, output, days, periods, time_limit=time_limit, log=self._log)
            result["ok"] = True
            result["seconds"] = round(time.time() - self._job["started"], 1)
            self._log("Done in " + str(result["seconds"]) + " s — saved " + os.path.basename(output), "success")
        except Exception as error:
            result = {"ok": False, "error": str(error),
                      "settings": str(len(days)) + " days × " + str(periods) + " periods = "
                      + str(len(days) * periods) + " slots per division"}
            for line in str(error).splitlines():
                if line.strip():
                    self._log(line, "error")
        self._job["result"] = result
        self._job["finished"] = time.time()
        self._job["running"] = False
        if result["ok"] and self._settings.get("open_after_generate"):
            self.open_path(output)

    @_safe
    def job(self, since=0):
        job = self._job
        elapsed = ((job["finished"] or time.time()) - job["started"]) if job["started"] else 0
        return {"running": job["running"], "elapsed": round(elapsed, 1), "log": job["log"][int(since or 0):],
                "logCount": len(job["log"]), "result": job["result"]}

    # ---- Timetables and lesson editing

    @_safe
    def timetables(self):
        path = self._settings["output_path"]
        if not os.path.isfile(path):
            return {"missing": True, "path": path}
        mtime = os.path.getmtime(path)
        if self._timetable_cache is None or self._timetable_cache[0] != (path, mtime):
            data = read_timetables(path)
            for grid in data["classes"] + data["teachers"]:
                grid.pop("sort_key", None)
            self._timetable_cache = ((path, mtime), data)
        return dict(self._timetable_cache[1])

    @_safe
    def lesson_options(self, class_name, day_index, period_index):
        return lessons.lesson_options(self._settings["output_path"], class_name, day_index, period_index)

    @_safe
    def set_lesson(self, class_name, day_index, period_index, subject, teacher):
        return lessons.set_lesson(self._settings["output_path"], class_name, day_index, period_index, subject, teacher)

    @_safe
    def split_options(self, grade):
        return {"subjects": lessons.split_options(self._settings["output_path"], grade)}

    @_safe
    def split_candidates(self, grade, subject, existing_teacher, division):
        return {"teachers": lessons.split_candidates(self._settings["output_path"], grade, subject,
                                                     existing_teacher, division)}

    @_safe
    def split(self, grade, subject, existing_teacher, new_teacher, periods, division):
        return lessons.split(self._settings["output_path"], grade, subject, existing_teacher, new_teacher,
                             periods, division)

    # ---- Files and window

    @_safe
    def open_path(self, path):
        if not path or not os.path.exists(path):
            raise ValueError("This file or folder does not exist yet:\n" + str(path))
        if sys.platform.startswith("win"):
            os.startfile(path)
        else:
            import subprocess
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        return {}

    def set_dirty(self, dirty):
        self._dirty = bool(dirty)
        return True

    # Called by the window's closing event: confirm before losing work.
    def _confirm_close(self):
        reasons = []
        if self._dirty:
            reasons.append("You have unsaved requirement changes.")
        if self._job["running"]:
            reasons.append("A timetable is still being generated.")
        if not reasons:
            return True
        return self._window.create_confirmation_dialog("Quit SchoolFlow?", " ".join(reasons) + " Quit anyway?")
