@echo off
cd /d "%~dp0"
python -m py_compile run_pipeline.py
if errorlevel 1 goto :end
python -u run_pipeline.py
:end
pause
