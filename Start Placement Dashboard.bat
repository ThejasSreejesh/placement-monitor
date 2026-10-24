@echo off
REM ===== One-click launcher for the Placement Offers dashboard =====
cd /d "%~dp0"
echo Starting Placement Offers dashboard...
echo A browser window will open automatically. Keep this window open while using it.
echo (Close this window to stop the dashboard.)
python app.py
pause
