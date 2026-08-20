#!/bin/bash
# Klik dua kali file ini di macOS untuk menjalankan aplikasi.
cd "$(dirname "$0")"
if [ ! -d ".venv" ]; then
  echo "Menyiapkan aplikasi untuk pertama kali (butuh internet, sekali saja)..."
  python3 -m venv .venv
  ./.venv/bin/pip install --upgrade pip >/dev/null
  ./.venv/bin/pip install -r requirements.txt
fi
./.venv/bin/streamlit run app.py
