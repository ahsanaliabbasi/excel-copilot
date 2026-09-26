@echo off
rem Starts the Excel Copilot server and opens it in your normal browser (file dialog + drag-and-drop work best there).
cd /d "%~dp0backend"
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:8000"
"%~dp0backend\venv\Scripts\python.exe" -m uvicorn main:app --port 8000
