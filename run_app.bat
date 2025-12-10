@echo off
echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Error installing dependencies. Please check your Python installation.
    pause
    exit /b
)

echo Starting AutoArchiver...
python batch_archiver.py
pause
