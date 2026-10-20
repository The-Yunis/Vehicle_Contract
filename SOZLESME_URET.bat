@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
echo Sozlesmeler uretiliyor... (Excel dosyasini kaydettiginizden emin olun)
python araclar\uret.py
echo.
if exist CIKTI start "" CIKTI
pause
