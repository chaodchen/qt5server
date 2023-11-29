@echo off
pyinstaller --onefile main.py --hidden-import PyQt5 --hidden-import xlwings
pause