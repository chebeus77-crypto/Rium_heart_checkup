@echo off
cd /d "%~dp0"
echo ===================================================
echo   Starting Heart Report Generator...
echo ===================================================
uv run streamlit run app.py
pause
