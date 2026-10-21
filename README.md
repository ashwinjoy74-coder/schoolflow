# SchoolFlow

Timetable planner for schools: enter who teaches what, and SchoolFlow builds
clash-free class and teacher timetables (Google OR-Tools CP-SAT) and an Excel
workbook. Landing page at `/`, app at `/app/`.

Each visitor's school data stays in their own browser; the server works on it
for one request and keeps nothing.

## Run locally (Windows)

Double-click `start_web.bat` (needs Python 3.10+). It installs `ortools` and
`openpyxl` on first run and opens http://127.0.0.1:8000/.

Any OS: `pip install -r requirements.txt` then `python server.py`.

## Deploy (Vercel)

Import this repo in Vercel; `vercel.json` configures everything (no build step,
static files from `site/`, the API from `api/index.py`). Pushes to the main
branch redeploy automatically. On the free Hobby plan the solver is capped at
4 minutes per run.

## Layout

| Path | What it is |
|---|---|
| `site/` | Landing page (`index.html`) and the app UI (`site/app/`) |
| `api/index.py` | Vercel serverless entry point for `POST /api/call` |
| `server.py` | Local web server (same API, standard library only) |
| `engine/` | Scheduler and lesson-editing code from the desktop app |
| `sample/` | Sample school used by "Restore sample requirements" |
