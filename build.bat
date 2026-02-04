@echo off
echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo Building executable...
pyinstaller --onefile --windowed --name "DisplaySnap" --icon=NONE main.py

echo Done! Executable is in dist\DisplaySnap.exe
pause
