@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

where python >nul 2>&1
if errorlevel 1 (
  echo Khong tim thay Python trong PATH. Vui long cai dat Python 3.10+ va danh dau "Add to PATH".
  pause
  exit /b 1
)

python -m pip install --disable-pip-version-check --no-input -r requirements.txt
if errorlevel 1 (
  python -m pip install --disable-pip-version-check --no-input pymupdf pandas reportlab Pillow streamlit
)

wscript //nologo run_app_gui.vbs
exit /b 0
