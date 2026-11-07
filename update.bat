@echo off
chcp 65001 >nul
cd /d "%~dp0"
python tools\update_site.py
if errorlevel 1 (echo خطا؛ انتشار انجام نشد & pause & exit /b 1)
git add docs
git commit -m "update %date% %time%"
git push
echo انجام شد
pause
