@echo off
echo Stopping Python/Streamlit servers...
taskkill /F /IM python.exe >nul 2>&1
taskkill /F /IM streamlit.exe >nul 2>&1
echo Done.
pause
