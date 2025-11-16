@echo off
cd /d %~dp0


python -m pip install pyinstaller

REM Package with all icon files
pyinstaller --onefile --windowed --add-data "icon;icon" --icon=icon/favicon48X48.ico audio_processor.py

if exist build rmdir /s /q build
move /y dist\audio_processor.exe .
if exist dist rmdir /s /q dist

echo Build completed! The program now includes main window icon functionality.
pause