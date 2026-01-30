@echo off
echo Installing PyInstaller...
pip install pyinstaller
echo.
echo Building SCH Tool.exe...
echo This process may take a few minutes.
echo.
python -m PyInstaller --noconfirm --onefile --windowed --name "SCH Tool" ^
 --hidden-import=streamlit ^
 --hidden-import=pandas ^
 --hidden-import=plotly ^
 --hidden-import=openpyxl ^
 --collect-all=streamlit ^
 --collect-all=altair ^
 --collect-all=pyarrow ^
 --add-data "app.py;." ^
 --add-data "requirements.txt;." ^
 run_exe.py

echo.
echo Build Complete!
