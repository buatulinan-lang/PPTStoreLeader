@echo off
REM Klik dua kali file ini di Windows untuk menjalankan aplikasi.
cd /d "%~dp0"
if not exist ".venv" (
  echo Menyiapkan aplikasi untuk pertama kali ^(butuh internet, sekali saja^)...
  python -m venv .venv
  .venv\Scripts\python -m pip install --upgrade pip >nul
  .venv\Scripts\pip install -r requirements.txt
)
.venv\Scripts\streamlit run app.py
pause
