#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tek komutla çalıştırılabilir paket üretir (Windows exe / macOS app / Linux).

    python build.py

Sonuç:
    Windows : dist/EtiketYazdirma/EtiketYazdirma.exe   (klasörün tamamını taşıyın)
    macOS   : dist/EtiketYazdirma.app
    Linux   : dist/EtiketYazdirma/EtiketYazdirma
    + hepsi için dağıtıma hazır zip: dist/EtiketYazdirma-<platform>.zip

Paket, derlendiği platforma özeldir: Windows exe'si Windows'ta, Mac uygulaması
Mac'te derlenir. Paketi alan kişinin Python kurmasına gerek yoktur.

Önce gerekli paketleri (requirements.txt + pyinstaller) kurar, sonra
customtkinter'ın tema/ikon dosyalarını bulup pakete ekler. Bu adım atlanırsa
exe açılırken 'customtkinter/assets bulunamadı' hatası verir.
"""

import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
PY = sys.executable
AD = "EtiketYazdirma"


def calistir(cmd):
    print("\n> " + " ".join(cmd), flush=True)
    r = subprocess.run(cmd)
    if r.returncode != 0:
        sys.exit("\nHATA: komut başarısız (çıkış kodu %d). Yukarıdaki mesajlara bakın." % r.returncode)


def main():
    print("[1/3] Paketler kuruluyor (requirements.txt + pyinstaller)…")
    calistir([PY, "-m", "pip", "install", "--quiet", "-r", "requirements.txt", "pyinstaller"])

    print("[2/3] customtkinter klasörü bulunuyor…")
    import customtkinter  # pip'ten sonra import edilmeli
    ctk_dir = os.path.dirname(customtkinter.__file__)
    print("      " + ctk_dir)
    ayirac = ";" if os.name == "nt" else ":"

    print("[3/3] PyInstaller çalışıyor…")
    for d in ("build", "dist"):
        shutil.rmtree(d, ignore_errors=True)
    cmd = [PY, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed", "--onedir",
           "--name", AD,
           "--add-data", ctk_dir + ayirac + "customtkinter",
           "--collect-submodules", "openpyxl",
           "etiket_gui.py"]
    if os.name == "nt":
        cmd += ["--hidden-import", "win32print", "--hidden-import", "win32timezone"]
    calistir(cmd)

    if sys.platform == "darwin":
        cikti = os.path.join("dist", AD + ".app")
    elif os.name == "nt":
        cikti = os.path.join("dist", AD, AD + ".exe")
    else:
        cikti = os.path.join("dist", AD, AD)
    print("\nTamam: %s" % os.path.abspath(cikti))

    # Dağıtım için zip
    platform_adi = {"nt": "windows", "posix": "macos" if sys.platform == "darwin" else "linux"}[os.name]
    zip_taban = os.path.join("dist", "%s-%s" % (AD, platform_adi))
    kok = AD + ".app" if sys.platform == "darwin" else AD
    zip_yolu = shutil.make_archive(zip_taban, "zip", root_dir="dist", base_dir=kok)
    print("Zip   : %s  (%.1f MB)" % (os.path.abspath(zip_yolu), os.path.getsize(zip_yolu) / 1e6))
    print("\nZip'i gönderin; alan kişi açıp %s dosyasını çalıştırır, Python gerekmez."
          % (AD + (".exe" if os.name == "nt" else ".app")))
    if os.name == "nt":
        print("İlk açılışta SmartScreen uyarısı çıkarsa: 'Ek bilgi' -> 'Yine de çalıştır'.")
    print("Excel dosyasını ve etiket_gui_ayarlar.json'u uygulamanın yanına koyabilirsiniz.")


if __name__ == "__main__":
    main()
