@echo off
cd /d "%~dp0"
echo Installing required Python packages...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo Package installation failed. Check Python and internet connection.
  pause
  exit /b 1
)
echo Starting ExamDesk...
python app.py
pause
