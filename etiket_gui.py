#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
COP-31 Demirbaş Etiket Yazdırma — grafik arayüz (CustomTkinter)

Excel okuma, ZPL üretimi ve yazıcıya gönderme işlerini etiket_bas.py yapar;
bu dosya yalnızca arayüzdür. Komut satırı aracı (etiket_bas.py) aynen
çalışmaya devam eder.

Çalıştırmak için:  python etiket_gui.py
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
except ImportError:
    sys.exit("HATA: 'customtkinter' paketi kurulu değil.\n"
             "      Kurmak için: pip install -r requirements.txt")

import etiket_bas as E

UYGULAMA_ADI = "COP-31 Etiket Yazdırma"
ONIZLEME_PANEL_GENISLIK = 560          # px; önizleme bu genişliğe sığdırılır
ONIZLEME_RENDER_OLCEK = 3              # render ölçeği (1 nokta = 3 px), sonra küçültülür

VARSAYILAN_SATIRLAR = [list(s) for s in E.ETIKET_SATIRLARI]
VARSAYILAN_SATIRLAR_BARKODLU = [list(s) for s in E.ETIKET_SATIRLARI_BARKODLU]


def uygulama_dizini():
    """Ayar dosyası exe'nin (ya da scriptin) yanında tutulur."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


AYAR_DOSYASI = os.path.join(uygulama_dizini(), "etiket_gui_ayarlar.json")


def sayi(metin, varsayilan=None, tam=False):
    """'3,5' veya '3.5' -> 3.5. Boş/bozuksa varsayılan."""
    s = str(metin).strip().replace(",", ".")
    if not s:
        return varsayilan
    try:
        return int(float(s)) if tam else float(s)
    except ValueError:
        return varsayilan


def hata_metni(exc):
    """sys.exit('HATA: ...') ile fırlatılan SystemExit'ten mesajı al."""
    m = str(exc.code) if isinstance(exc, SystemExit) else str(exc)
    return m[len("HATA: "):] if m.startswith("HATA: ") else m


# ==========================================================================
#  Etiket satırı widget'ı: [şablon] [Y mm] [punto mm] [sil]
# ==========================================================================

class SatirWidget:
    def __init__(self, master, satir_no, degisince, silince, sablon="", y=0.0, punto=0.0):
        self.master = master
        self.e_sablon = ctk.CTkEntry(master, width=232, placeholder_text="şablon")
        self.e_y = ctk.CTkEntry(master, width=58, justify="center")
        self.e_punto = ctk.CTkEntry(master, width=58, justify="center")
        self.b_sil = ctk.CTkButton(master, text="✕", width=30, fg_color="transparent",
                                   border_width=1, text_color=("gray20", "gray80"),
                                   command=lambda: silince(self))
        self.e_sablon.insert(0, sablon)
        self.e_y.insert(0, ("%g" % y) if y else "")
        self.e_punto.insert(0, ("%g" % punto) if punto else "")
        for w in (self.e_sablon, self.e_y, self.e_punto):
            w.bind("<KeyRelease>", lambda _e: degisince())
        self.yerlestir(satir_no)

    def yerlestir(self, satir_no):
        self.e_sablon.grid(row=satir_no, column=0, padx=(0, 4), pady=2, sticky="w")
        self.e_y.grid(row=satir_no, column=1, padx=2, pady=2)
        self.e_punto.grid(row=satir_no, column=2, padx=2, pady=2)
        self.b_sil.grid(row=satir_no, column=3, padx=(4, 0), pady=2)

    def kaldir(self):
        for w in (self.e_sablon, self.e_y, self.e_punto, self.b_sil):
            w.destroy()

    def degerler(self):
        return [self.e_sablon.get(), sayi(self.e_y.get(), 0.0), sayi(self.e_punto.get(), 0.0)]


# ==========================================================================
#  Ana pencere
# ==========================================================================

class Uygulama(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title(UYGULAMA_ADI)
        self.geometry("1320x860")
        self.minsize(1120, 720)

        self.wb = None
        self.excel_yolu = ""
        self.sheet_adlari = []
        self.sheet_sayilari = {}
        self.etiketler = []
        self.uyarilar = []
        self.secim = []
        self.onizleme_idx = 0
        self._onizleme_img = None       # label'ın şu an gösterdiği CTkImage
        self._onceki_img = None         # bir önceki; Tk eski görüntüyü silmesin diye bir tur tutulur
        self._bos_img = None            # "etiket yok" durumunda 1x1 yer tutucu
        self._yenile_isi = None
        self._sheet_onbellek = {}

        self.satir_widgetlari = []
        self.satirlar_normal = [list(s) for s in VARSAYILAN_SATIRLAR]
        self.satirlar_barkodlu = [list(s) for s in VARSAYILAN_SATIRLAR_BARKODLU]
        self._tablo_modu = "normal"

        self._arayuzu_kur()
        self._ayarlari_yukle()
        self.protocol("WM_DELETE_WINDOW", self._kapat)

    # ------------------------------------------------------------------
    #  Arayüz kurulumu
    # ------------------------------------------------------------------

    def _arayuzu_kur(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        self.sol = ctk.CTkScrollableFrame(self, width=470, label_text="Ayarlar")
        self.sol.grid(row=0, column=0, sticky="nsw", padx=(10, 5), pady=(10, 5))

        self.sag = ctk.CTkFrame(self)
        self.sag.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=(10, 5))

        self.log_kutusu = ctk.CTkTextbox(self, height=120, font=ctk.CTkFont(family="Courier", size=12))
        self.log_kutusu.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
        self.log_kutusu.configure(state="disabled")

        self._sol_paneli_kur()
        self._sag_paneli_kur()

    def _baslik(self, metin):
        ctk.CTkLabel(self.sol, text=metin, font=ctk.CTkFont(size=14, weight="bold"),
                     anchor="w").pack(fill="x", padx=6, pady=(14, 2))

    def _satir_cercevesi(self):
        f = ctk.CTkFrame(self.sol, fg_color="transparent")
        f.pack(fill="x", padx=6, pady=1)
        return f

    def _girdi(self, master, etiket, genislik=70, varsayilan="", yan=True):
        ctk.CTkLabel(master, text=etiket).pack(side="left", padx=(0, 4))
        e = ctk.CTkEntry(master, width=genislik, justify="center")
        e.insert(0, str(varsayilan))
        e.pack(side="left", padx=(0, 12))
        e.bind("<KeyRelease>", lambda _e: self._yenile_planla())
        return e

    def _sol_paneli_kur(self):
        # ---- 1. Excel
        self._baslik("1. Excel dosyası")
        f = self._satir_cercevesi()
        self.e_excel = ctk.CTkEntry(f, width=350, placeholder_text="Excel dosyası seçin…")
        self.e_excel.pack(side="left", padx=(0, 6))
        ctk.CTkButton(f, text="Seç…", width=70, command=self._excel_sec).pack(side="left")

        # ---- 2. Sheet
        self._baslik("2. Sheet")
        f = self._satir_cercevesi()
        self.cb_sheet = ctk.CTkComboBox(f, width=430, values=["(önce Excel seçin)"],
                                        state="readonly", command=lambda _v: self._sheet_degisti())
        self.cb_sheet.pack(side="left")
        self.l_sheet_bilgi = ctk.CTkLabel(self.sol, text="", anchor="w", justify="left",
                                          wraplength=440, text_color=("gray30", "gray70"))
        self.l_sheet_bilgi.pack(fill="x", padx=8, pady=(2, 0))

        # ---- 3. Aralık
        self._baslik("3. Satır aralığı")
        f = self._satir_cercevesi()
        self.seg_aralik = ctk.CTkSegmentedButton(f, values=["Excel satırı", "Etiket sırası"],
                                                 command=lambda _v: self._yenile_planla())
        self.seg_aralik.set("Excel satırı")
        self.seg_aralik.pack(side="left")
        f = self._satir_cercevesi()
        self.e_bas = self._girdi(f, "Başlangıç", 80)
        self.e_son = self._girdi(f, "Bitiş", 80)
        self.l_aralik_bilgi = ctk.CTkLabel(self.sol, text="Boş bırakılırsa sheet'in tamamı.",
                                           anchor="w", text_color=("gray30", "gray70"))
        self.l_aralik_bilgi.pack(fill="x", padx=8)

        # ---- 4. Ürün kodu
        self._baslik("4. Ürün kodu")
        f = self._satir_cercevesi()
        self.e_kod = self._girdi(f, "Kısa kod", 160)
        ctk.CTkLabel(f, text="boş = otomatik", text_color=("gray30", "gray70")).pack(side="left")
        self.l_kod_bilgi = ctk.CTkLabel(self.sol, text="", anchor="w", justify="left",
                                        wraplength=440, text_color=("gray30", "gray70"))
        self.l_kod_bilgi.pack(fill="x", padx=8)

        # ---- 5. Etiket boyutu
        self._baslik("5. Etiket boyutu (mm)")
        f = self._satir_cercevesi()
        self.e_gen = self._girdi(f, "Genişlik", 60, "%g" % E.ETIKET_GENISLIK_MM)
        self.e_yuk = self._girdi(f, "Yükseklik", 60, "%g" % E.ETIKET_YUKSEKLIK_MM)
        self.e_kenar = self._girdi(f, "Kenar", 50, "%g" % E.KENAR_BOSLUK_X_MM)
        f = self._satir_cercevesi()
        ctk.CTkLabel(f, text="Yazıcı çözünürlüğü").pack(side="left", padx=(0, 6))
        self.cb_dpi = ctk.CTkComboBox(f, width=140, values=["203 dpi (8 dot/mm)", "300 dpi (12 dot/mm)"],
                                      state="readonly", command=lambda _v: self._yenile_planla())
        self.cb_dpi.set("203 dpi (8 dot/mm)")
        self.cb_dpi.pack(side="left")

        # ---- 6. Etiket satırları
        self._baslik("6. Etiket satırları")
        ctk.CTkLabel(self.sol, anchor="w", justify="left", wraplength=440,
                     text_color=("gray30", "gray70"),
                     text="Alanlar: {sartname} {seri} {urun} {sira} {sira_ham} {kod} {isim} {urun_id}"
                     ).pack(fill="x", padx=8)
        self.tablo = ctk.CTkFrame(self.sol, fg_color="transparent")
        self.tablo.pack(fill="x", padx=6, pady=(4, 0))
        ctk.CTkLabel(self.tablo, text="Şablon", anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(self.tablo, text="Üst (mm)").grid(row=0, column=1)
        ctk.CTkLabel(self.tablo, text="Punto (mm)").grid(row=0, column=2)
        f = self._satir_cercevesi()
        ctk.CTkButton(f, text="+ Satır ekle", width=110, command=self._satir_ekle).pack(side="left", padx=(0, 6))
        ctk.CTkButton(f, text="Varsayılana dön", width=130, fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"), command=self._satirlari_sifirla).pack(side="left")
        f = self._satir_cercevesi()
        self.ck_esitle = ctk.CTkCheckBox(f, text="Tüm satırlar aynı punto", command=self._yenile_planla)
        self.ck_esitle.pack(side="left")
        f = self._satir_cercevesi()
        self.e_basamak = self._girdi(f, "Sıra no basamak", 44, str(E.SIRA_BASAMAK))
        self.e_minfont = self._girdi(f, "En küçük punto (mm)", 50, "%g" % E.MIN_FONT_MM)
        self._tabloyu_doldur(self.satirlar_normal)

        # ---- 7. Barkod / QR
        self._baslik("7. Seri no barkodu")
        f = self._satir_cercevesi()
        self.seg_kod = ctk.CTkSegmentedButton(f, values=["Yok", "Code 128", "QR"],
                                              command=lambda _v: self._kod_modu_degisti())
        self.seg_kod.set("Yok")
        self.seg_kod.pack(side="left")
        self.f_barkod = self._satir_cercevesi()
        self.e_bk_y = self._girdi(self.f_barkod, "Barkod üst", 55, "%g" % E.BARKOD_Y_MM)
        self.e_bk_h = self._girdi(self.f_barkod, "Yükseklik", 55, "%g" % E.BARKOD_YUKSEKLIK_MM)
        self.f_qr = self._satir_cercevesi()
        self.e_qr_x = self._girdi(self.f_qr, "QR sol", 50, "%g" % E.QR_X_MM)
        self.e_qr_y = self._girdi(self.f_qr, "üst", 50, "%g" % E.QR_Y_MM)
        self.e_qr_b = self._girdi(self.f_qr, "büyütme", 40, str(E.QR_BUYUKLUK))
        self.f_barkod.pack_forget()
        self.f_qr.pack_forget()

        # ---- 8. Diğer
        self._baslik("8. Diğer")
        f = self._satir_cercevesi()
        self.ck_trascii = ctk.CTkCheckBox(f, text="Türkçe karakterleri ASCII'ye çevir",
                                          command=self._yenile_planla)
        self.ck_trascii.pack(side="left", padx=(0, 14))
        self.e_kopya = self._girdi(f, "Kopya", 44, "1")

        # ---- 9. Yazıcı
        self._baslik("9. Yazıcı")
        f = self._satir_cercevesi()
        self.seg_yazici = ctk.CTkSegmentedButton(f, values=["Kurulu yazıcı", "Ağ (IP)"],
                                                 command=lambda _v: self._yazici_modu_degisti())
        self.seg_yazici.set("Kurulu yazıcı")
        self.seg_yazici.pack(side="left")
        self.f_yazici = self._satir_cercevesi()
        self.cb_yazici = ctk.CTkComboBox(self.f_yazici, width=340, values=["(yazıcı bulunamadı)"],
                                         state="readonly")
        self.cb_yazici.pack(side="left", padx=(0, 6))
        ctk.CTkButton(self.f_yazici, text="Yenile", width=70, command=self._yazicilari_yenile).pack(side="left")
        self.f_ip = self._satir_cercevesi()
        ctk.CTkLabel(self.f_ip, text="IP adresi").pack(side="left", padx=(0, 4))
        self.e_ip = ctk.CTkEntry(self.f_ip, width=180, placeholder_text="192.168.1.50")
        self.e_ip.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(self.f_ip, text="Port").pack(side="left", padx=(0, 4))
        self.e_port = ctk.CTkEntry(self.f_ip, width=70, justify="center")
        self.e_port.insert(0, "9100")
        self.e_port.pack(side="left")
        self.f_ip.pack_forget()

        # ---- Eylemler
        ctk.CTkFrame(self.sol, fg_color="transparent", height=10).pack()
        f = self._satir_cercevesi()
        ctk.CTkButton(f, text="ZPL dosyasına kaydet", width=160, fg_color="transparent",
                      border_width=1, text_color=("gray20", "gray80"),
                      command=self._zpl_kaydet).pack(side="left", padx=(0, 6))
        ctk.CTkButton(f, text="Test: ilk etiketi bas", width=150,
                      command=lambda: self._yazdir(test=True)).pack(side="left", padx=(0, 6))
        f = self._satir_cercevesi()
        ctk.CTkButton(f, text="YAZDIR", height=44, font=ctk.CTkFont(size=16, weight="bold"),
                      fg_color="#1f8f4e", hover_color="#176d3b",
                      command=lambda: self._yazdir(test=False)).pack(fill="x", pady=(6, 12))

    def _sag_paneli_kur(self):
        self.sag.grid_columnconfigure(0, weight=1)
        self.sag.grid_rowconfigure(2, weight=1)

        ust = ctk.CTkFrame(self.sag, fg_color="transparent")
        ust.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 0))
        ctk.CTkLabel(ust, text="Önizleme", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        self.sw_oto = ctk.CTkSwitch(ust, text="Otomatik yenile", command=self._yenile_planla)
        self.sw_oto.select()
        self.sw_oto.pack(side="right")
        ctk.CTkButton(ust, text="Yenile", width=70, command=self._onizlemeyi_yenile
                      ).pack(side="right", padx=(0, 12))
        ctk.CTkButton(ust, text="PNG kaydet", width=100, fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"), command=self._png_kaydet
                      ).pack(side="right", padx=(0, 12))

        nav = ctk.CTkFrame(self.sag, fg_color="transparent")
        nav.grid(row=1, column=0, sticky="ew", padx=10, pady=(6, 0))
        ctk.CTkButton(nav, text="◀", width=36, command=lambda: self._gez(-1)).pack(side="left")
        self.l_nav = ctk.CTkLabel(nav, text="—", width=320)
        self.l_nav.pack(side="left", padx=8)
        ctk.CTkButton(nav, text="▶", width=36, command=lambda: self._gez(+1)).pack(side="left")
        ctk.CTkButton(nav, text="Son", width=50, fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"), command=lambda: self._gez(10**9)
                      ).pack(side="left", padx=(12, 0))
        ctk.CTkButton(nav, text="İlk", width=50, fg_color="transparent", border_width=1,
                      text_color=("gray20", "gray80"), command=lambda: self._gez(-10**9)
                      ).pack(side="left", padx=(6, 0))

        self.l_onizleme = ctk.CTkLabel(self.sag, text="Excel dosyası seçince önizleme burada görünür.",
                                       fg_color=("gray85", "gray20"), corner_radius=8)
        self.l_onizleme.grid(row=2, column=0, sticky="n", padx=10, pady=10)

        self.t_ozet = ctk.CTkTextbox(self.sag, height=150, font=ctk.CTkFont(family="Courier", size=12))
        self.t_ozet.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.t_ozet.configure(state="disabled")

    # ------------------------------------------------------------------
    #  Yardımcılar
    # ------------------------------------------------------------------

    def log(self, metin):
        self.log_kutusu.configure(state="normal")
        self.log_kutusu.insert("end", metin.rstrip() + "\n")
        self.log_kutusu.see("end")
        self.log_kutusu.configure(state="disabled")

    def _ozet_yaz(self, metin):
        self.t_ozet.configure(state="normal")
        self.t_ozet.delete("1.0", "end")
        self.t_ozet.insert("1.0", metin)
        self.t_ozet.configure(state="disabled")

    def _hata(self, baslik, exc):
        m = hata_metni(exc)
        self.log("HATA: " + m)
        messagebox.showerror(baslik, m, parent=self)

    def _tabloyu_doldur(self, satirlar):
        for w in self.satir_widgetlari:
            w.kaldir()
        self.satir_widgetlari = []
        for i, (sab, y, p) in enumerate(satirlar, 1):
            self.satir_widgetlari.append(
                SatirWidget(self.tablo, i, self._yenile_planla, self._satir_sil, sab, y, p))

    def _tablodan_oku(self):
        return [w.degerler() for w in self.satir_widgetlari]

    def _satir_ekle(self):
        n = len(self.satir_widgetlari) + 1
        son_y = self.satir_widgetlari[-1].degerler()[1] if self.satir_widgetlari else 0
        self.satir_widgetlari.append(
            SatirWidget(self.tablo, n, self._yenile_planla, self._satir_sil,
                        "", son_y + 6.0, 4.0))
        self._yenile_planla()

    def _satir_sil(self, w):
        w.kaldir()
        self.satir_widgetlari.remove(w)
        for i, sw in enumerate(self.satir_widgetlari, 1):
            sw.yerlestir(i)
        self._yenile_planla()

    def _satirlari_sifirla(self):
        if self._tablo_modu == "normal":
            self.satirlar_normal = [list(s) for s in VARSAYILAN_SATIRLAR]
            self._tabloyu_doldur(self.satirlar_normal)
        else:
            self.satirlar_barkodlu = [list(s) for s in VARSAYILAN_SATIRLAR_BARKODLU]
            self._tabloyu_doldur(self.satirlar_barkodlu)
        self._yenile_planla()

    def _kod_modu(self):
        return {"Yok": None, "Code 128": "barkod", "QR": "qr"}[self.seg_kod.get()]

    def _kod_modu_degisti(self):
        # Tablo içeriğini mevcut listeye yaz, diğer listeyi yükle
        yeni = "barkodlu" if self._kod_modu() else "normal"
        if yeni != self._tablo_modu:
            if self._tablo_modu == "normal":
                self.satirlar_normal = self._tablodan_oku()
            else:
                self.satirlar_barkodlu = self._tablodan_oku()
            self._tablo_modu = yeni
            self._tabloyu_doldur(self.satirlar_normal if yeni == "normal" else self.satirlar_barkodlu)
        m = self._kod_modu()
        self.f_barkod.pack_forget()
        self.f_qr.pack_forget()
        if m == "barkod":
            self.f_barkod.pack(fill="x", padx=6, pady=1, after=self.seg_kod.master)
        elif m == "qr":
            self.f_qr.pack(fill="x", padx=6, pady=1, after=self.seg_kod.master)
        self._yenile_planla()

    def _yazici_modu_degisti(self):
        if self.seg_yazici.get() == "Ağ (IP)":
            self.f_yazici.pack_forget()
            self.f_ip.pack(fill="x", padx=6, pady=1, after=self.seg_yazici.master)
        else:
            self.f_ip.pack_forget()
            self.f_yazici.pack(fill="x", padx=6, pady=1, after=self.seg_yazici.master)

    def _yazicilari_yenile(self):
        liste = E.yazicilari_listele()
        if liste:
            self.cb_yazici.configure(values=liste)
            if self.cb_yazici.get() not in liste:
                self.cb_yazici.set(liste[0])
            self.log("Kurulu yazıcılar: " + ", ".join(liste))
        else:
            self.cb_yazici.configure(values=["(yazıcı bulunamadı)"])
            self.cb_yazici.set("(yazıcı bulunamadı)")
            self.log("Kurulu yazıcı bulunamadı. Ağ (IP) seçeneğini kullanabilirsiniz.")

    # ------------------------------------------------------------------
    #  Excel / sheet
    # ------------------------------------------------------------------

    def _excel_sec(self):
        yol = filedialog.askopenfilename(
            parent=self, title="Excel dosyası seçin",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tümü", "*.*")])
        if yol:
            self.e_excel.delete(0, "end")
            self.e_excel.insert(0, yol)
            self._excel_yukle(yol)

    def _excel_yukle(self, yol):
        self.l_sheet_bilgi.configure(text="Excel yükleniyor…")
        self.update_idletasks()
        try:
            self.wb = E.excel_yukle(yol)
        except SystemExit as ex:
            self.wb = None
            self._hata("Excel açılamadı", ex)
            return
        self.excel_yolu = yol
        self._sheet_onbellek = {}
        self.sheet_adlari = list(self.wb.sheetnames)

        # Sheet'leri tara, etiket sayılarını bul
        self.sheet_sayilari = {}
        for i, ad in enumerate(self.sheet_adlari, 1):
            self.l_sheet_bilgi.configure(text="Sheet taranıyor %d/%d: %s" % (i, len(self.sheet_adlari), ad))
            self.update_idletasks()
            try:
                et, _ = E.sheet_oku(self.wb[ad])
                self.sheet_sayilari[ad] = len(et)
            except SystemExit:
                self.sheet_sayilari[ad] = 0

        secenekler = ["%d · %s  (%d etiket)" % (i, ad, self.sheet_sayilari[ad])
                      for i, ad in enumerate(self.sheet_adlari, 1)]
        self.cb_sheet.configure(values=secenekler)
        dolu = [s for s, ad in zip(secenekler, self.sheet_adlari) if self.sheet_sayilari[ad] > 0]
        self.cb_sheet.set(dolu[0] if dolu else secenekler[0])
        toplam = sum(self.sheet_sayilari.values())
        self.log("Excel yüklendi: %s  (%d sheet, toplam %d etiket)"
                 % (os.path.basename(yol), len(self.sheet_adlari), toplam))
        self._sheet_degisti()

    def _secili_sheet(self):
        s = self.cb_sheet.get()
        if not self.wb or "·" not in s:
            return None
        idx = int(s.split("·")[0].strip()) - 1
        return self.sheet_adlari[idx]

    def _sheet_degisti(self):
        self.onizleme_idx = 0
        self.e_bas.delete(0, "end")
        self.e_son.delete(0, "end")
        self._yenile_planla(hemen=True)

    def _sheet_oku(self, ad, kod):
        anahtar = (ad, kod, E.SIRA_BASAMAK)
        if anahtar not in self._sheet_onbellek:
            self._sheet_onbellek[anahtar] = E.sheet_oku(self.wb[ad], kod_zorla=kod or None)
        return self._sheet_onbellek[anahtar]

    # ------------------------------------------------------------------
    #  Ayarları etiket_bas modülüne uygula
    # ------------------------------------------------------------------

    def _ayarlari_uygula(self):
        """Arayüzdeki değerleri etiket_bas sabitlerine yaz. Hatalı değerde ValueError."""
        gen = sayi(self.e_gen.get(), E.ETIKET_GENISLIK_MM)
        yuk = sayi(self.e_yuk.get(), E.ETIKET_YUKSEKLIK_MM)
        kenar = sayi(self.e_kenar.get(), E.KENAR_BOSLUK_X_MM)
        if not (gen and yuk and gen > 5 and yuk > 5):
            raise ValueError("Etiket genişliği/yüksekliği geçersiz.")
        if kenar is None or kenar < 0 or kenar * 2 >= gen:
            raise ValueError("Kenar boşluğu geçersiz.")
        E.ETIKET_GENISLIK_MM = gen
        E.ETIKET_YUKSEKLIK_MM = yuk
        E.KENAR_BOSLUK_X_MM = kenar
        E.DPMM = 12 if self.cb_dpi.get().startswith("300") else 8
        E.olculeri_hesapla()

        satirlar = []
        for sab, y, p in self._tablodan_oku():
            if not sab.strip():
                continue
            if p <= 0:
                raise ValueError("'%s' satırının puntosu geçersiz." % sab)
            satirlar.append((sab, y, p))
        if not satirlar:
            raise ValueError("En az bir etiket satırı gerekli.")
        if self._tablo_modu == "normal":
            E.ETIKET_SATIRLARI = satirlar
        else:
            E.ETIKET_SATIRLARI_BARKODLU = satirlar

        E.SATIRLARI_ESITLE = bool(self.ck_esitle.get())
        E.SIRA_BASAMAK = max(0, sayi(self.e_basamak.get(), 3, tam=True) or 0)
        E.MIN_FONT_MM = max(0.5, sayi(self.e_minfont.get(), 2.0) or 2.0)

        E.BARKOD_Y_MM = sayi(self.e_bk_y.get(), E.BARKOD_Y_MM)
        E.BARKOD_YUKSEKLIK_MM = sayi(self.e_bk_h.get(), E.BARKOD_YUKSEKLIK_MM)
        E.QR_X_MM = sayi(self.e_qr_x.get(), E.QR_X_MM)
        E.QR_Y_MM = sayi(self.e_qr_y.get(), E.QR_Y_MM)
        E.QR_BUYUKLUK = max(1, min(10, sayi(self.e_qr_b.get(), E.QR_BUYUKLUK, tam=True) or 4))

        m = self._kod_modu()
        return {
            "barkod": m == "barkod",
            "qr": m == "qr",
            "tr_ascii": bool(self.ck_trascii.get()),
            "kopya": max(1, sayi(self.e_kopya.get(), 1, tam=True) or 1),
            "kod": self.e_kod.get().strip(),
            "by_label": self.seg_aralik.get() == "Etiket sırası",
        }

    def _secimi_hesapla(self):
        """Sheet'i oku, aralığı uygula. (secim, ayarlar) döndürür; sorun varsa None."""
        ad = self._secili_sheet()
        if not ad:
            return None, None
        try:
            ayar = self._ayarlari_uygula()
        except ValueError as ex:
            self.l_aralik_bilgi.configure(text=str(ex))
            return None, None
        try:
            self.etiketler, self.uyarilar = self._sheet_oku(ad, ayar["kod"])
        except SystemExit as ex:
            self.l_sheet_bilgi.configure(text=hata_metni(ex))
            return None, None

        if not self.etiketler:
            self.l_sheet_bilgi.configure(text="Bu sheet'te basılacak etiket yok. "
                                              + (self.uyarilar[0] if self.uyarilar else ""))
            return None, ayar

        e0, e1 = self.etiketler[0], self.etiketler[-1]
        bilgi = ["%d etiket · Excel satır %d-%d · sıra %s-%s"
                 % (len(self.etiketler), e0.satir, e1.satir, e0.sira, e1.sira),
                 "Şartname no: %s" % (e0.sartname or "(yok)")]
        for u in self.uyarilar:
            if "ürün kısa kodu" not in u:
                bilgi.append("! " + u.split("\n")[0])
        self.l_sheet_bilgi.configure(text="\n".join(bilgi))

        id_li = sum(1 for e in self.etiketler if e.urun_id)
        if ayar["kod"]:
            kaynak = "elle verildi"
        elif id_li == len(self.etiketler):
            kaynak = "Excel'deki 'Ürün id' sütunundan"
        elif id_li:
            kaynak = "%d satırda Excel 'Ürün id', kalan %d satırda %s-sıra" % (
                id_li, len(self.etiketler) - id_li, e0.kod)
        elif E.normalize(ad) in E.URUN_KODLARI:
            kaynak = "URUN_KODLARI tablosundan"
        else:
            kaynak = "sheet adından türetildi — kısa kod girmeniz önerilir"
        self.l_kod_bilgi.configure(text="Ürün satırı: %s  (%s)" % (e0.urun, kaynak))

        bas = sayi(self.e_bas.get(), None, tam=True)
        son = sayi(self.e_son.get(), None, tam=True)
        alan = "etiket sırası" if ayar["by_label"] else "Excel satırı"
        key = (lambda e: e.index) if ayar["by_label"] else (lambda e: e.satir)
        self.l_aralik_bilgi.configure(text="Mevcut %s aralığı: %d-%d. Boş = tamamı."
                                      % (alan, key(e0), key(e1)))
        if bas is not None and son is not None and bas > son:
            self.l_aralik_bilgi.configure(text="Başlangıç bitişten büyük olamaz.")
            return None, ayar
        try:
            secim = E.aralik_uygula(self.etiketler, bas, son, ayar["by_label"], ad)
        except SystemExit as ex:
            self.l_aralik_bilgi.configure(text=hata_metni(ex).split("\n")[0]
                                          + "  (mevcut: %d-%d)" % (key(e0), key(e1)))
            return None, ayar
        return secim, ayar

    # ------------------------------------------------------------------
    #  Önizleme
    # ------------------------------------------------------------------

    def _yenile_planla(self, hemen=False):
        if self._yenile_isi:
            self.after_cancel(self._yenile_isi)
            self._yenile_isi = None
        if hemen:
            self._onizlemeyi_yenile()
        elif self.sw_oto.get():
            self._yenile_isi = self.after(300, self._onizlemeyi_yenile)

    def _onizleme_goster(self, img, metin):
        """Önizleme label'ına görüntü + metin koy.

        Tk, label hâlâ bir PhotoImage'a işaret ederken o nesne çöp toplanırsa
        'image "pyimageN" doesn't exist' hatası verir. Bu yüzden:
          1. önce yeni görüntü label'a atanır (eski referans hâlâ canlıyken),
          2. metin ayrı bir configure ile verilir (CTkLabel text'i image'dan önce uygular),
          3. eski görüntü bir tur daha referansta tutulur,
          4. image=None hiç kullanılmaz; boş durumda 1x1 yer tutucu görüntü atanır.
        """
        if img is None:
            if self._bos_img is None:
                from PIL import Image
                bos = Image.new("RGB", (1, 1), (235, 235, 235))
                self._bos_img = ctk.CTkImage(light_image=bos, dark_image=bos, size=(1, 1))
            yeni = self._bos_img
        else:
            oran = ONIZLEME_PANEL_GENISLIK / img.width
            yeni = ctk.CTkImage(light_image=img, dark_image=img,
                                size=(ONIZLEME_PANEL_GENISLIK, int(img.height * oran)))
        self.l_onizleme.configure(image=yeni)
        self.l_onizleme.configure(text=metin)
        self._onceki_img, self._onizleme_img = self._onizleme_img, yeni

    def _gez(self, adim):
        if not self.secim:
            return
        self.onizleme_idx = max(0, min(len(self.secim) - 1, self.onizleme_idx + adim))
        self._onizlemeyi_yenile()

    def _onizlemeyi_yenile(self):
        self._yenile_isi = None
        self.secim, ayar = self._secimi_hesapla()
        if not self.secim:
            self.l_nav.configure(text="—")
            self._ozet_yaz("")
            self._onizleme_goster(None, "Seçimde etiket yok.")
            return

        self.onizleme_idx = max(0, min(len(self.secim) - 1, self.onizleme_idx))
        e = self.secim[self.onizleme_idx]
        try:
            img = E.onizleme_goruntu([e], barkod=ayar["barkod"], qr=ayar["qr"],
                                     tr_ascii=ayar["tr_ascii"], olcek=ONIZLEME_RENDER_OLCEK,
                                     sutun=1, altyazi_goster=False)
        except SystemExit as ex:
            self._hata("Önizleme", ex)
            return
        self._onizleme_goster(img, "")
        self.l_nav.configure(text="Etiket %d / %d   ·   Excel satır %d   ·   sıra %s"
                             % (self.onizleme_idx + 1, len(self.secim), e.satir, e.sira))

        toplam = len(self.secim) * ayar["kopya"]
        ilk, son = self.secim[0], self.secim[-1]
        satirlar = " | ".join(m for m, _, _ in E.etiket_satirlari(e, ayar["tr_ascii"],
                                                                   barkodlu=ayar["barkod"] or ayar["qr"]))
        _, uyari = E.is_uret([e], barkod=ayar["barkod"], qr=ayar["qr"], tr_ascii=ayar["tr_ascii"])
        self._ozet_yaz(
            "Sheet        : %s\n"
            "Seçim        : %d etiket  x %d kopya  =  %d\n"
            "İlk          : satır %d  %s\n"
            "Son          : satır %d  %s\n"
            "Bu etiket    : %s\n"
            "Etiket       : %g x %g mm  (%d x %d nokta)%s"
            % (self._secili_sheet(), len(self.secim), ayar["kopya"], toplam,
               ilk.satir, ilk.urun, son.satir, son.urun, satirlar,
               E.ETIKET_GENISLIK_MM, E.ETIKET_YUKSEKLIK_MM, E.GENISLIK_DOT, E.YUKSEKLIK_DOT,
               ("\n! " + "\n! ".join(dict.fromkeys(uyari))) if uyari else ""))

    def _png_kaydet(self):
        if not self.secim:
            messagebox.showinfo(UYGULAMA_ADI, "Önce bir seçim yapın.", parent=self)
            return
        yol = filedialog.asksaveasfilename(parent=self, defaultextension=".png",
                                           initialfile="onizleme.png",
                                           filetypes=[("PNG", "*.png")])
        if not yol:
            return
        _, ayar = self._secimi_hesapla()
        try:
            img = E.onizleme_goruntu(self.secim, barkod=ayar["barkod"], qr=ayar["qr"],
                                     tr_ascii=ayar["tr_ascii"])
            img.save(yol)
        except (SystemExit, OSError) as ex:
            self._hata("PNG kaydedilemedi", ex)
            return
        self.log("Önizleme kaydedildi: %s (%d etiket)" % (yol, min(len(self.secim), E.ONIZLEME_MAKS)))

    # ------------------------------------------------------------------
    #  ZPL üretimi / yazdırma
    # ------------------------------------------------------------------

    def _zpl_hazirla(self, test=False):
        secim, ayar = self._secimi_hesapla()
        if not secim:
            messagebox.showwarning(UYGULAMA_ADI, "Basılacak etiket yok. Sheet ve aralığı kontrol edin.",
                                   parent=self)
            return None, None, None
        if test:
            secim = secim[:1]
        try:
            zpl, uyari = E.is_uret(secim, kopya=ayar["kopya"], barkod=ayar["barkod"],
                                   qr=ayar["qr"], tr_ascii=ayar["tr_ascii"])
        except SystemExit as ex:
            self._hata("ZPL üretilemedi", ex)
            return None, None, None
        for u in dict.fromkeys(uyari):
            self.log("! " + u)
        return zpl, secim, ayar

    def _zpl_kaydet(self):
        zpl, secim, ayar = self._zpl_hazirla()
        if not zpl:
            return
        yol = filedialog.asksaveasfilename(parent=self, defaultextension=".zpl",
                                           initialfile="etiketler.zpl",
                                           filetypes=[("ZPL", "*.zpl"), ("Metin", "*.txt")])
        if not yol:
            return
        try:
            with open(yol, "w", encoding="utf-8", newline="\n") as f:
                f.write(zpl)
        except OSError as ex:
            self._hata("ZPL kaydedilemedi", ex)
            return
        self.log("ZPL kaydedildi: %s  (%d etiket x %d kopya, %d bayt)"
                 % (yol, len(secim), ayar["kopya"], len(zpl.encode("utf-8"))))

    def _yazdir(self, test=False):
        zpl, secim, ayar = self._zpl_hazirla(test=test)
        if not zpl:
            return
        ip = yazici = None
        if self.seg_yazici.get() == "Ağ (IP)":
            ip = self.e_ip.get().strip()
            port = sayi(self.e_port.get(), 9100, tam=True) or 9100
            if not ip:
                messagebox.showwarning(UYGULAMA_ADI, "Yazıcı IP adresi girin.", parent=self)
                return
            hedef = "%s:%d" % (ip, port)
        else:
            yazici = self.cb_yazici.get()
            port = 9100
            if not yazici or yazici.startswith("("):
                messagebox.showwarning(UYGULAMA_ADI,
                                       "Kurulu yazıcı seçin (Yenile) ya da Ağ (IP) kullanın.",
                                       parent=self)
                return
            hedef = yazici

        toplam = len(secim) * (1 if test else ayar["kopya"])
        if test:
            mesaj = "TEST: yalnızca ilk etiket basılacak.\n\n%s\n\nYazıcı: %s" % (secim[0].urun, hedef)
        else:
            mesaj = ("%d etiket basılacak (%d satır x %d kopya).\n\n"
                     "İlk : satır %d  %s\nSon : satır %d  %s\n\nYazıcı: %s\n\nDevam edilsin mi?"
                     % (toplam, len(secim), ayar["kopya"],
                        secim[0].satir, secim[0].urun, secim[-1].satir, secim[-1].urun, hedef))
        if not messagebox.askyesno("Baskı onayı", mesaj, parent=self):
            self.log("Baskı iptal edildi.")
            return

        if test:
            zpl, _ = E.is_uret(secim, kopya=1, barkod=ayar["barkod"], qr=ayar["qr"],
                               tr_ascii=ayar["tr_ascii"])
        try:
            E.gonder(zpl, yazici=yazici, ip=ip, port=port)
        except SystemExit as ex:
            self._hata("Yazdırılamadı", ex)
            return
        self.log("Gönderildi: %d etiket -> %s" % (toplam, hedef))
        if not test:
            key = (lambda e: e.index) if ayar["by_label"] else (lambda e: e.satir)
            self.log("Yarıda kalırsa devam için: başlangıç = kaldığınız satır, bitiş = %d"
                     % key(secim[-1]))
        self._ayarlari_kaydet()

    # ------------------------------------------------------------------
    #  Ayar dosyası
    # ------------------------------------------------------------------

    def _ayarlari_topla(self):
        if self._tablo_modu == "normal":
            self.satirlar_normal = self._tablodan_oku()
        else:
            self.satirlar_barkodlu = self._tablodan_oku()
        return {
            "excel": self.e_excel.get(),
            "sheet": self._secili_sheet() or "",
            "aralik_modu": self.seg_aralik.get(),
            "bas": self.e_bas.get(), "son": self.e_son.get(),
            "kod": self.e_kod.get(),
            "genislik": self.e_gen.get(), "yukseklik": self.e_yuk.get(), "kenar": self.e_kenar.get(),
            "dpi": self.cb_dpi.get(),
            "satirlar": self.satirlar_normal,
            "satirlar_barkodlu": self.satirlar_barkodlu,
            "esitle": bool(self.ck_esitle.get()),
            "basamak": self.e_basamak.get(), "min_punto": self.e_minfont.get(),
            "kod_modu": self.seg_kod.get(),
            "bk_y": self.e_bk_y.get(), "bk_h": self.e_bk_h.get(),
            "qr_x": self.e_qr_x.get(), "qr_y": self.e_qr_y.get(), "qr_b": self.e_qr_b.get(),
            "tr_ascii": bool(self.ck_trascii.get()),
            "kopya": self.e_kopya.get(),
            "yazici_modu": self.seg_yazici.get(),
            "yazici": self.cb_yazici.get(),
            "ip": self.e_ip.get(), "port": self.e_port.get(),
            "oto_onizleme": bool(self.sw_oto.get()),
            "pencere": self.geometry(),
        }

    def _ayarlari_kaydet(self):
        try:
            with open(AYAR_DOSYASI, "w", encoding="utf-8") as f:
                json.dump(self._ayarlari_topla(), f, ensure_ascii=False, indent=2)
        except OSError as ex:
            self.log("Ayarlar kaydedilemedi: %s" % ex)

    def _ayarlari_yukle(self):
        self._yazicilari_yenile()
        if not os.path.exists(AYAR_DOSYASI):
            self.log("Hoş geldiniz. Excel dosyası seçerek başlayın.")
            return
        try:
            with open(AYAR_DOSYASI, encoding="utf-8") as f:
                a = json.load(f)
        except (OSError, ValueError) as ex:
            self.log("Ayar dosyası okunamadı, varsayılanlar kullanılıyor: %s" % ex)
            return

        def koy(entry, deger):
            entry.delete(0, "end")
            entry.insert(0, str(deger))

        for anahtar, entry in (("kod", self.e_kod), ("genislik", self.e_gen), ("yukseklik", self.e_yuk),
                               ("kenar", self.e_kenar), ("basamak", self.e_basamak),
                               ("min_punto", self.e_minfont), ("bk_y", self.e_bk_y),
                               ("bk_h", self.e_bk_h), ("qr_x", self.e_qr_x), ("qr_y", self.e_qr_y),
                               ("qr_b", self.e_qr_b), ("kopya", self.e_kopya), ("ip", self.e_ip),
                               ("port", self.e_port), ("bas", self.e_bas), ("son", self.e_son)):
            if anahtar in a:
                koy(entry, a[anahtar])
        if a.get("dpi"):
            self.cb_dpi.set(a["dpi"])
        if a.get("satirlar"):
            self.satirlar_normal = [list(s) for s in a["satirlar"]]
        if a.get("satirlar_barkodlu"):
            self.satirlar_barkodlu = [list(s) for s in a["satirlar_barkodlu"]]
        (self.ck_esitle.select if a.get("esitle") else self.ck_esitle.deselect)()
        (self.ck_trascii.select if a.get("tr_ascii") else self.ck_trascii.deselect)()
        (self.sw_oto.select if a.get("oto_onizleme", True) else self.sw_oto.deselect)()
        if a.get("aralik_modu") in ("Excel satırı", "Etiket sırası"):
            self.seg_aralik.set(a["aralik_modu"])
        self.seg_kod.set(a.get("kod_modu", "Yok") if a.get("kod_modu") in ("Yok", "Code 128", "QR") else "Yok")
        self._tablo_modu = "normal"
        self._tabloyu_doldur(self.satirlar_normal)
        self._kod_modu_degisti()
        if a.get("yazici_modu") in ("Kurulu yazıcı", "Ağ (IP)"):
            self.seg_yazici.set(a["yazici_modu"])
            self._yazici_modu_degisti()
        if a.get("yazici") and a["yazici"] in (self.cb_yazici.cget("values") or []):
            self.cb_yazici.set(a["yazici"])
        if a.get("pencere"):
            try:
                self.geometry(a["pencere"])
            except tk.TclError:
                pass

        yol = a.get("excel", "")
        if yol and os.path.exists(yol):
            koy(self.e_excel, yol)
            self._excel_yukle(yol)
            hedef = a.get("sheet")
            if hedef in self.sheet_adlari:
                i = self.sheet_adlari.index(hedef)
                self.cb_sheet.set(self.cb_sheet.cget("values")[i])
                koy(self.e_bas, a.get("bas", ""))
                koy(self.e_son, a.get("son", ""))
                self._yenile_planla(hemen=True)
        else:
            self.log("Son ayarlar yüklendi. Excel dosyası seçin.")

    def _kapat(self):
        self._ayarlari_kaydet()
        self.destroy()


def main():
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = Uygulama()
    app.mainloop()


if __name__ == "__main__":
    main()
