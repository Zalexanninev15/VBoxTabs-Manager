@echo off
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --onedir --windowed --name "VBoxTabs Manager" --clean --log-level "ERROR" --optimize "2" --noupx --hide-console "hide-late" --hidden-import "PySide6" VBoxTabs-Manager.py
rmdir /S /Q build