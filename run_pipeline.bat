@echo off
cd /d "%~dp0"
python -m py_compile pipeline_final.py
if errorlevel 1 goto :end
python -u pipeline_final.py
:end
pause
