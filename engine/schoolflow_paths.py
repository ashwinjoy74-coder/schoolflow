# SchoolFlow file locations and saved preferences — web edition.
#
# The web server keeps one school's data in a single data folder:
#   <data>/Requirements/requirements.csv   the working requirements
#   <data>/Exports/school_timetable.xlsx   the generated workbook
#   <data>/schoolflow_settings.json        preferences
#
# The data folder is Web/data by default; set SCHOOLFLOW_DATA to use another one.
# On first run the requirements are copied from the desktop app's Requirements
# folder (one level up) when it exists, otherwise from the bundled sample.

import json
import os
import shutil

APP_NAME = "SchoolFlow"
APP_VERSION = "2.0.0-web"
MODE = "web"

ENGINE_FOLDER = os.path.dirname(os.path.abspath(__file__))
WEB_FOLDER = os.path.dirname(ENGINE_FOLDER)
DESKTOP_FOLDER = os.path.dirname(WEB_FOLDER)

SITE_FOLDER = os.path.join(WEB_FOLDER, "site")
SAMPLE_REQUIREMENTS = os.path.join(WEB_FOLDER, "sample", "requirements.csv")
DESKTOP_REQUIREMENTS = os.path.join(DESKTOP_FOLDER, "Requirements", "requirements.csv")

DATA_FOLDER = os.path.abspath(os.environ.get("SCHOOLFLOW_DATA") or os.path.join(WEB_FOLDER, "data"))
SETTINGS_FOLDER = DATA_FOLDER
REQUIREMENTS_FOLDER = os.path.join(DATA_FOLDER, "Requirements")
EXPORTS_FOLDER = os.path.join(DATA_FOLDER, "Exports")
DEFAULT_INPUT = os.path.join(REQUIREMENTS_FOLDER, "requirements.csv")
DEFAULT_OUTPUT = os.path.join(EXPORTS_FOLDER, "school_timetable.xlsx")
SETTINGS_FILE = os.path.join(SETTINGS_FOLDER, "schoolflow_settings.json")

DEFAULT_SETTINGS = {
    "input_path": DEFAULT_INPUT,
    "output_path": DEFAULT_OUTPUT,
    "days": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "periods_per_day": 7,
    "solver_time_limit": 180,
    "open_after_generate": False,
    "theme": "system",
}


# Create the data folders and seed the requirements on first run.
def ensure_data_folders():
    os.makedirs(REQUIREMENTS_FOLDER, exist_ok=True)
    os.makedirs(EXPORTS_FOLDER, exist_ok=True)
    if not os.path.isfile(DEFAULT_INPUT):
        source = DESKTOP_REQUIREMENTS if os.path.isfile(DESKTOP_REQUIREMENTS) else SAMPLE_REQUIREMENTS
        if os.path.isfile(source):
            shutil.copyfile(source, DEFAULT_INPUT)


# Replace the working requirements CSV with the bundled sample data.
def restore_sample_requirements(target=DEFAULT_INPUT):
    if not os.path.isfile(SAMPLE_REQUIREMENTS):
        raise FileNotFoundError("The bundled sample requirements file is missing.")
    os.makedirs(os.path.dirname(target), exist_ok=True)
    shutil.copyfile(SAMPLE_REQUIREMENTS, target)


# Load saved preferences, falling back to defaults for missing or broken values.
def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8-sig") as file:
            saved = json.load(file)
        if isinstance(saved, dict):
            for key, default in DEFAULT_SETTINGS.items():
                if key in saved and isinstance(saved[key], type(default)):
                    settings[key] = saved[key]
    except (OSError, ValueError):
        pass
    return settings


# Save preferences atomically so a crash cannot leave a half-written file.
def save_settings(settings):
    os.makedirs(SETTINGS_FOLDER, exist_ok=True)
    temporary = SETTINGS_FILE + ".tmp"
    with open(temporary, "w", encoding="utf-8") as file:
        json.dump(settings, file, indent=2)
    os.replace(temporary, SETTINGS_FILE)
