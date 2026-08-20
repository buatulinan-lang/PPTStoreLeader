# -*- coding: utf-8 -*-
"""Cek cepat: apakah semua file & pustaka yang dibutuhkan aplikasi sudah lengkap.

Jalankan:  python cek_lingkungan.py
"""
import os
import sys

AKAR = os.path.dirname(os.path.abspath(__file__))
WAJIB_FILE = ["app.py", "requirements.txt", "mflash/__init__.py", "mflash/loader.py",
              "mflash/metrics.py", "mflash/context.py", "mflash/charts.py", "mflash/theme.py",
              "mflash/deck.py", "mflash/template.py"]
DISARANKAN = ["assets/bg.jpg", "assets/logo_mflash.png", "assets/logo_madinah.png"]
PUSTAKA = ["streamlit", "pandas", "numpy", "openpyxl", "pptx", "plotly", "PIL"]

print(f"Python  : {sys.version.split()[0]}")
print(f"Folder  : {AKAR}\n")

gagal = False
print("File wajib:")
for f in WAJIB_FILE:
    ada = os.path.exists(os.path.join(AKAR, f))
    gagal = gagal or not ada
    print(f"  {'OK  ' if ada else 'HILANG'}  {f}")

print("\nAset template (slide tanpa latar/logo bila hilang):")
for f in DISARANKAN:
    print(f"  {'OK  ' if os.path.exists(os.path.join(AKAR, f)) else 'HILANG'}  {f}")

print("\nPustaka:")
for m in PUSTAKA:
    try:
        mod = __import__(m)
        print(f"  OK    {m} {getattr(mod, '__version__', '')}")
    except Exception as e:  # noqa: BLE001
        gagal = True
        print(f"  GAGAL {m} — {type(e).__name__}: {e}")

print("\nUji impor paket inti:")
sys.path.insert(0, AKAR)
try:
    from mflash import loader, metrics, context, deck, template  # noqa: F401
    print("  OK    semua modul mflash terbaca")
except Exception as e:  # noqa: BLE001
    gagal = True
    import traceback
    print(f"  GAGAL — {type(e).__name__}: {e}")
    traceback.print_exc()

print("\n" + ("❌ Ada yang perlu diperbaiki (lihat baris HILANG/GAGAL di atas)."
              if gagal else "✅ Lengkap. Aplikasi siap dijalankan."))
