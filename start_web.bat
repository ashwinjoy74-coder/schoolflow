@echo off
title SchoolFlow Web
cd /d "%~dp0"
where python >nul 2>nul || (echo Python 3.10+ is needed. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH". & pause & exit /b 1)
python -c "import ortools, openpyxl" 2>nul || (echo Installing SchoolFlow packages... & python -m pip install -r requirements.txt || (pause & exit /b 1))
python server.py %*
pause
