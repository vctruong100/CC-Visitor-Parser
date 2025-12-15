@echo off
setlocal
py -m pip install --upgrade pip
py -m pip install pandas openpyxl
py planned_gui.py
endlocal
