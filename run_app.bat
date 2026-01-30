@echo off
chcp 65001
echo ===================================================
echo  Gantt Monthly Progress Report - Web App Launcher
echo ===================================================
echo.
echo [1/2] Checking and Installing libraries...
echo.
pip install -r requirements.txt
echo.
echo [2/2] Launching Application...
echo.
streamlit run app.py --server.address 0.0.0.0
pause
