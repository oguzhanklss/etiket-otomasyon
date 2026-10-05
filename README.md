# COP-31 Demirbaş Etiket Yazdırma

Excel'deki seri numaralarından Zebra ZD220 (203 dpi) için **50 x 30 mm** demirbaş
etiketi basar. ZPL üretir ve yazıcıya **RAW** olarak gönderir (sürücü üzerinden
grafik baskı yapılmaz).

Etiket düzeni:

```
ŞARTNAME NO:3
S/N:35CG6370JV6
NOTEBOOK-012
```

* **ŞARTNAME NO** — sheet adının başındaki numara (`3.HP ELİTEBOOK...` → `3`)
* **S/N** — `Seri No` sütunu
* **Ürün satırı** — Excel'de `Ürün id` sütunu varsa o değer aynen; yoksa
  `URUN_KODLARI` tablosundaki kısa kod + sıra no (`NOTEBOOK-012`)

İki kullanım şekli var:

* **Grafik arayüz** — `python etiket_gui.py` (bölüm 11). Sheet, aralık, satır
  puntoları, barkod, yazıcı; hepsi arayüzden, canlı önizlemeyle.
* **Komut satırı** — `python etiket_bas.py …` (bölüm 2-5). Toplu iş ve
  betiklerde.

İkisi de aynı çekirdeği (`etiket_bas.py`) kullanır; ürettikleri ZPL aynıdır.

---

## 1. Kurulum

Python 3.8 veya üstü gerekir.

```bash
pip install -r requirements.txt
```

Windows'ta RAW baskı için `pywin32` otomatik kurulur. macOS/Linux'ta CUPS'un
`lp` komutu kullanılır, ek paket gerekmez.

---

## 2. Önce sheet'leri listele

Hangi sheet'in kaç etiketi ve hangi Excel satır aralığı olduğunu gösterir:

```bash
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --list
```

```
 No  Sheet adı                             Etiket  Excel satır (sıra no)
 ----------------------------------------------------------------------
 1   3.HP ELİTEBOOK 8G2İ 14 AI                820  2-825  (sıra 1-50)
 2   1. WORKSTATION                           118  2-123  (sıra 1-32)
 ...
```

Uyarıları da görmek için `--detay` ekleyin.

---

## 3. Yazdırma

`--sheet` listedeki **numara** ya da **sheet adı** olabilir.
`--from` / `--to` **Excel satır numarasıdır** (Excel'in solunda gördüğünüz numara),
ikisi de dahildir.

```bash
# Önce kâğıda dökmeden kontrol et
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 2 --to 51 --dry-run

# Tek etiket bas, hizalamayı kontrol et
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --test

# Gerçek baskı (onay sorar)
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 2 --to 51 --printer "ZDesigner ZD220-203dpi ZPL"

# Ağ üzerinden, sürücüsüz
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 2 --to 51 --ip 192.168.1.50
```

`--sheet` vermezseniz script sheet listesini gösterip sheet'i ve satır aralığını
tek tek sorar.

Seçilen etiketlerin **tamamı tek iş olarak** gönderilir; her etiket için ayrı
onay gerekmez.

### Yarıda kalan baskıya devam

Baskı biterken script, kaldığınız yeri gösteren komutu yazdırır. Örneğin
2-200 arası bastıysanız ve 137. satırda kâğıt bittiyse:

```bash
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 137 --to 200
```

Excel satırı yerine "kaçıncı etiket" ile saymayı tercih ederseniz `--by-label`
ekleyin; o zaman `--from 137` "bu sheet'in 137. geçerli etiketi" demektir.

---

## 4. Yazıcısız dijital önizleme

Yazıcı elinizde yokken etiketin nasıl çıkacağını PNG olarak görebilirsiniz.
Render **tamamen çevrimdışıdır**, hiçbir veri dışarıya gönderilmez.

```bash
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 787 --to 789 --onizleme
```

Varsayılan çıktı `onizleme.png`; başka ad vermek için `--onizleme kontrol.png`.
`--onizleme` kullanıldığında yazıcıya hiçbir şey gönderilmez.

Önizleme, ZPL ile **aynı punto hesabını** kullanır; yani satırların küçülüp
küçülmediğini, metnin sığıp sığmadığını buradan görebilirsiniz. Tek fark font:
yazıcıda CG Triumvirate Bold Condensed, önizlemede sistemdeki en yakın
condensed bold font (macOS'ta Arial Narrow Bold) kullanılır, bu yüzden harf
genişlikleri birkaç piksel oynayabilir.

Ayarlar (`etiket_bas.py` başında):

```python
ONIZLEME_OLCEK  = 4    # 1 yazıcı noktası kaç piksel çizilsin
ONIZLEME_MAKS   = 12   # tek PNG'de en fazla kaç etiket
ONIZLEME_SUTUN  = 3    # ızgara sütun sayısı
```

`--barkod` veya `--qr` ile birlikte kullanırsanız barkod da çizilir
(Code 128 gerçek desenle, QR yer tutucu kutu olarak).

---

## 5. Tüm seçenekler

| Seçenek | Açıklama |
|---|---|
| `--list` | Sheet'leri, etiket sayılarını, satır aralıklarını listeler |
| `--detay` | `--list` çıktısına uyarıları ekler |
| `--sheet N\|ad` | Basılacak sheet (liste numarası veya adı) |
| `--from N` `--to N` | Excel satır aralığı (dahil) |
| `--by-label` | `--from/--to`'yu etiket sırası (1..N) olarak yorumlar |
| `--kod KISAKOD` | Ürün satırındaki kısa kodu elle belirler (`HPLAPTOP-001`) |
| `--test` | Sadece ilk etiketi basar |
| `--copies N` | Her etiketten N kopya |
| `--barkod` | Seri noyu ayrıca Code 128 barkod olarak basar |
| `--qr` | Seri noyu ayrıca QR kod olarak basar |
| `--tr-ascii` | Türkçe karakterleri ASCII'ye çevirir |
| `--printer AD` | Yazıcı adı. Verilmezse kurulu yazıcılardan seçtirir |
| `--printers` | Kurulu yazıcıları listeler |
| `--ip ADRES` | Yazıcı IP'si, RAW port 9100 (`--port` ile değiştirilir) |
| `--dry-run` | Yazıcıya göndermez, ZPL'i dosyaya yazar |
| `--out DOSYA` | `--dry-run` çıktı dosyası (varsayılan `etiketler.zpl`) |
| `--onizleme [PNG]` | Yazıcısız PNG önizleme (varsayılan `onizleme.png`) |
| `--yes` / `-y` | Onay sormaz |

---

## 6. Etiket yerleşimini değiştirme

Her ayar `etiket_bas.py` dosyasının başındaki
**"ETİKET AYARLARI"** bölümündedir. Tüm ölçüler milimetredir.

### Satırlar

```python
ETIKET_SATIRLARI = [
    ("ŞARTNAME NO:{sartname}",  3.5, 6.4),   # (şablon, üst konum mm, punto mm)
    ("S/N:{seri}",             11.75, 6.4),
    ("{urun}",                 20.0, 6.4),
]
```

Şablonda kullanılabilecek alanlar:

| Alan | Örnek |
|---|---|
| `{sartname}` | `3` — sheet adının başındaki numara |
| `{seri}` | `35CG6370JV6` |
| `{urun}` | `NOTEBOOK-012` — Ürün id varsa o, yoksa `{kod}-{sira}` |
| `{sira}` | `012` (sıfırla doldurulmuş) |
| `{sira_ham}` | `12` |
| `{kod}` | `NOTEBOOK` |
| `{isim}` | `HP ELİTEBOOK 8G2İ 14 AI` — sheet adı, numarasız |
| `{urun_id}` | Excel'deki `Ürün id` değeri (boş olabilir) |

Satır eklemek/çıkarmak için listeye satır ekleyin ya da silin. Örneğin ürün
adını da basmak için:

```python
ETIKET_SATIRLARI = [
    ("ŞARTNAME NO:{sartname}",  2.0, 5.0),
    ("{isim}",                  8.0, 4.0),
    ("S/N:{seri}",             14.0, 5.5),
    ("{urun}",                 21.0, 5.5),
]
```

### Punto otomatik küçülmesi

Uzun metinler etiketten taşmaz: script punto yüksekliğini `MIN_FONT_MM`'ye
kadar düşürür, ayrıca ZPL `^FB` alanı fazlasını keser.

* `MIN_FONT_MM` — alt sınır (varsayılan 2.0 mm)
* `SATIRLARI_ESITLE = True` yaparsanız üç satır da aynı puntoda basılır
* `ORT_KARAKTER_ORANI` — yazı beklenenden **küçük** çıkıyorsa bu değeri düşürün
  (0.58 → 0.52), **taşıyorsa** yükseltin

### Ürün kısa kodları

Etiketin 3. satırı (`NOTEBOOK-012`) şu öncelikle belirlenir:

1. **`--kod` verildiyse** → `KOD-sıra`
2. Excel'de o satırın `Ürün id` hücresi doluysa → o değer aynen
3. `URUN_KODLARI` tablosunda sheet varsa → `KOD-sıra`
4. Hiçbiri yoksa → `SHEET ADI-sıra` (uzun olur, script uyarır)

#### Komut satırından: `--kod`

Tek seferlik ya da tabloyu hiç düzenlemeden:

```bash
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 2 --to 5 --kod HPLAPTOP
```

```
HPLAPTOP-001
HPLAPTOP-002
HPLAPTOP-003
HPLAPTOP-004
```

`--kod`, Excel'deki `Ürün id` sütunu dolu olsa bile onu **ezer**. Bu, bir
sheet'in bir kısmında `Ürün id` olup kalanında olmadığı durumlarda işe yarar —
örneğin `21. HP USB-C` sheet'inde ilk 30 satırda `Ürün id` var, kalan 740'ta
yok; `--kod RJ45ADAPTER` ile hepsi aynı biçimde basılır.

Script hangi kaynağın kullanıldığını her çalıştırmada yazar:

```
Ürün satırı      : HPLAPTOP-001  (--kod ile verildi)
```

Sıra numarasının kaç basamağa tamamlanacağı `SIRA_BASAMAK` ile ayarlanır
(varsayılan 3 → `001`; `0` yaparsanız dolgu yapılmaz).

#### Kalıcı olarak: `URUN_KODLARI`

Aynı sheet'i sık basıyorsanız her seferinde `--kod` yazmamak için tabloya ekleyin:

```python
URUN_KODLARI = {
    "3.hp elitebook 8g2i 14 ai": "NOTEBOOK",
    "11.monitor":                "MONITOR",
    ...
}
```

Anahtar, sheet adının **normalize edilmiş** hâlidir: küçük harf, Türkçe
karakterler sadeleştirilmiş (`ş→s, ğ→g, ı/İ→i, ö→o, ü→u, ç→c`).
Tanımlı olmayan sheet'lerde script uyarı verir ve eklemeniz gereken satırı
ekrana yazar — kopyalayıp tabloya yapıştırmanız yeterli. `--kod` verildiğinde
bu uyarı çıkmaz, tablo hiç okunmaz.

### Etiket boyutu

```python
ETIKET_GENISLIK_MM  = 50.0
ETIKET_YUKSEKLIK_MM = 30.0
KENAR_BOSLUK_X_MM   = 2.0
DPMM = 8      # 203 dpi. 300 dpi yazıcıda 12 yapın.
```

### Barkod

Varsayılan olarak kapalıdır (örnek etikette yok). `--barkod` ya da `--qr` ile
açılır; kalıcı açmak için `BARKOD_GOSTER = True` yapın.

Barkod açıkken metin satırları `ETIKET_SATIRLARI` yerine **ayrı bir sıkışık
yerleşimden** okunur, böylece barkodun üstüne binmez:

```python
ETIKET_SATIRLARI_BARKODLU = [
    ("ŞARTNAME NO:{sartname}",  1.2, 4.6),
    ("S/N:{seri}",              6.5, 4.6),
    ("{urun}",                 11.8, 4.6),
]
BARKOD_Y_MM         = 18.0
BARKOD_YUKSEKLIK_MM = 9.5
```

Barkodlu düzeni değiştirmek için `ETIKET_SATIRLARI`'nı değil bu listeyi
düzenleyin. `--onizleme --barkod` ile sonucu görebilirsiniz.

Modül genişliği seri no uzunluğuna göre otomatik seçilir (3 → 2 → 1 dot).
Sığmazsa barkod basılmaz ve uyarı verilir.

---

## 7. Türkçe karakterler

ZPL `^CI28` (UTF-8) kullanılır ve dosya UTF-8 gönderilir.

ZD220'nin yerleşik fontu `ş ğ ı İ` karakterlerini basamazsa `--test` çıktısında
boşluk veya kutu görürsünüz. Bu durumda:

* `--tr-ascii` ile çalıştırın (`ŞARTNAME` → `SARTNAME`), ya da
* Yazıcıya Türkçe destekli bir TTF yükleyip `^CW` ile tanımlayın.

---

## 8. Veriyle ilgili bilinen durumlar

Script bunları otomatik yönetir, `--detay` ile görebilirsiniz:

* **Sütun adları sheet'ten sheet'e değişiyor** (`Sıra No` / `PC No` / başlıksız,
  `Seri No` / `S/N`, `Ürün id` / `ID` / `ÜRÜN ID`) ve sütun sıraları farklı.
  Eşleşme sütun **adından** yapılır, konumdan değil.
* **`S.NO` sıra no, `S/NO` seri no** demektir — ikisi ayrı ayrı tanınır.
* **Ara boş satırlar** atlanır.
* **Sayı olarak kaydedilmiş seri numaraları** (`82406307.0`) tam sayıya
  çevrilir; `1234.0` veya bilimsel gösterim basılmaz, baştaki sıfırlar korunur.
* **Sheet içinde ikinci tablo blokları** var:
  * `3.HP ELİTEBOOK` satır 775'te ikinci başlık (`LAPTOP SERİ NO / ADAPTER SERİ NO / ...`)
  * `1. WORKSTATION` satır 91'de ikinci başlık (`TESLİM EDİLEN MALZEMELER`)
  * `21. HP USB-C` satır 33'te sıra no 30'dan 1'e dönüyor

  Script yeni başlık satırını algılayıp sütun eşleşmesini yeniler ve uyarı
  verir. **Blokları ayırmak için `--from/--to`'ya Excel satır numarası verin** —
  sıra numaraları bloklar arasında tekrar ettiği için Excel satırı tek güvenilir
  referanstır.
* **Başlık metni içeren veri satırları** (örn. `Seri No` yazan hücre) atlanır.

---

## 9. Hata mesajları

| Mesaj | Ne yapmalı |
|---|---|
| `Excel dosyası bulunamadı` | Yolu kontrol edin, boşluk içeriyorsa tırnak içine alın |
| `'Seri No' sütunu bulunamadı` | Sheet'te başlık satırı yok veya farklı yazılmış |
| `verdiğiniz aralıkta etiket yok` | Mesajdaki "Mevcut" aralığını kullanın |
| `Sistemde kurulu yazıcı bulunamadı` | `--printers` ile kontrol edin veya `--ip` kullanın |
| `Windows'ta RAW baskı için 'pywin32' gerekir` | `pip install pywin32` |
| `Yazıcıya bağlanılamadı` | IP/port doğru mu, yazıcı açık mı |
| `ürün kısa kodu tanımlı değil` | `--kod KISAKOD` verin ya da `URUN_KODLARI` tablosuna ekleyin (bkz. bölüm 6) |

---

## 10. Windows'a taşıma

Bu klasörü (Excel dosyası dahil) Windows makineye kopyalayın:

```
etiket_bas.py
requirements.txt
README.md
COP-31 DEMİRBAŞLAR.xlsx
```

Windows'ta:

```bat
python -m pip install -r requirements.txt
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --printers
python etiket_bas.py "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --test --printer "ZDesigner ZD220-203dpi ZPL"
```

`--printers` çıktısındaki adı **birebir** (tırnak içinde) `--printer` ile verin.

Python kurulu olmayan bir makinede çalıştırmak isterseniz tek dosyalık exe:

```bat
python -m pip install pyinstaller
pyinstaller --onefile --name etiket_bas etiket_bas.py
```

`dist\etiket_bas.exe` oluşur, kullanımı aynıdır:

```bat
etiket_bas.exe "COP-31 DEMİRBAŞLAR.xlsx" --sheet 1 --from 2 --to 51 --printer "ZDesigner ZD220-203dpi ZPL"
```

---

## 11. Grafik arayüz (etiket_gui.py)

```bash
python etiket_gui.py
```

Sol panelde ayarlar, sağda canlı önizleme. Her değişiklik (satır aralığı,
punto, etiket boyutu, kod…) 0,3 sn içinde önizlemeye yansır; "Otomatik yenile"
kapatılırsa "Yenile" düğmesiyle elle yenilenir.

| Bölüm | Ne yapar |
|---|---|
| **1. Excel** | Dosyayı seçer, tüm sheet'leri tarayıp etiket sayılarını çıkarır |
| **2. Sheet** | Açılır listeden seçilir; satır aralığı, sıra no ve uyarılar altında görünür |
| **3. Satır aralığı** | Excel satırı ya da etiket sırası; boş = tamamı |
| **4. Ürün kodu** | `HPLAPTOP` → `HPLAPTOP-001`. Boş bırakılırsa otomatik (bkz. bölüm 6) |
| **5. Etiket boyutu** | Genişlik / yükseklik / kenar (mm), 203 veya 300 dpi |
| **6. Etiket satırları** | Her satır için şablon, üst konum ve **punto ayrı ayrı**. Satır ekle/sil, varsayılana dön |
| **7. Barkod** | Yok / Code 128 / QR. Açılınca barkodlu düzen tablosu gelir, konum ve boyut ayarlanır |
| **8. Diğer** | Türkçe→ASCII, kopya sayısı |
| **9. Yazıcı** | Kurulu yazıcı listesi ya da IP:port |

Sağ panel: ◀ ▶ ile seçimdeki etiketler arasında gezilir, **PNG kaydet** ilk 12
etiketi tek görüntüde dışa verir. Alttaki özet kutusu seçimi ve uyarıları gösterir.

Düğmeler:

* **ZPL dosyasına kaydet** — yazıcıya göndermeden `.zpl` üretir (komut
  satırındaki `--dry-run` ile aynı)
* **Test: ilk etiketi bas** — hizalama kontrolü için tek etiket
* **YAZDIR** — onay penceresi açar (adet, ilk/son etiket, yazıcı), onaylanırsa
  tüm seçimi tek iş olarak gönderir

Ayarlar her baskıda ve pencere kapanırken `etiket_gui_ayarlar.json` dosyasına
(scriptin/exe'nin yanına) kaydedilir; bir sonraki açılışta Excel, sheet, satır
düzeni ve yazıcı seçimi geri gelir. Dosyayı silerseniz varsayılanlara dönersiniz.

### Çalıştırılabilir paket derleme ve dağıtma

**Paketi alan kişinin Python kurmasına gerek yoktur**; Python exe'nin içine
gömülür. Python yalnızca derlemeyi yapan makinede gerekir.

**Paket, derlendiği platforma özeldir.** Windows exe'si bir Windows makinede
derlenir; Mac'te derlenen `.app` Windows'ta açılmaz.

#### Windows'ta derleme (adım adım)

1. Windows makineye Python 3.8+ kurun — [python.org](https://www.python.org/downloads/windows/),
   kurulumda **"Add python.exe to PATH"** kutusunu işaretleyin.
2. `kaynak-windows-icin.zip`'i açın (ya da şu dosyaları kopyalayın:
   `etiket_bas.py`, `etiket_gui.py`, `build.py`, `requirements.txt`).
3. O klasörde komut istemi açın (adres çubuğuna `cmd` yazıp Enter) ve:

   ```bat
   python build.py
   ```

4. Çıktı:

   ```
   dist\EtiketYazdirma\EtiketYazdirma.exe      <- program
   dist\EtiketYazdirma-windows.zip             <- dağıtıma hazır zip
   ```

`build.py` paketleri kurar, customtkinter tema dosyalarını pakete ekler,
derler ve zipler.

#### Zip'i gönderme

`dist\EtiketYazdirma-windows.zip` dosyasını gönderin. Alan kişi:

1. Zip'i bir klasöre çıkarır (**klasörün tamamı** gerekli; exe'yi tek başına
   kopyalamak çalışmaz).
2. `EtiketYazdirma.exe`'ye çift tıklar.
3. İlk açılışta Windows SmartScreen "Windows bilgisayarınızı korudu" derse:
   **Ek bilgi → Yine de çalıştır**. (İmzasız exe'lerde standart uyarı.)
4. Program içinden Excel dosyasını seçer.

Zip 30-40 MB civarı olur; Gmail/Outlook ek sınırı 25 MB olduğundan Drive,
OneDrive veya WeTransfer bağlantısı gönderin.

#### macOS / Linux

Aynı komut (`python3 build.py`) `dist/EtiketYazdirma.app` ve
`dist/EtiketYazdirma-macos.zip` üretir. İlk açılışta "tanımlanamayan
geliştirici" uyarısı çıkarsa sağ tık → Aç, ya da:

```bash
xattr -dr com.apple.quarantine dist/EtiketYazdirma.app
```

Önizleme fontu: Windows'ta `C:\Windows\Fonts\arialnb.ttf` (Arial Narrow Bold)
varsa o, yoksa Arial Bold kullanılır. Yazıcıdaki fonta en yakın olanlar bunlar.
