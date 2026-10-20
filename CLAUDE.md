# CLAUDE.md — Makine Kontrat

Bu dosya, projede çalışan Claude'un (hangi bilgisayarda olursa olsun) bağlamı bilmesi içindir.
Repo **public**: buraya kişisel veri (TC, IBAN, e-posta, kiralayan bilgisi) yazma.

## Proje ne?

Karayolları karla mücadele işinde kiralanan iş makinaları için **kiralama sözleşmesi üreten Windows masaüstü programı**
(`KMS_Sozlesme.exe`). Eski adı "KMS Sözleşme"; v1.3.0'da "Makine Kontrat" oldu. **Exe adı (`KMS_Sozlesme.exe`, eski
sürümler güncellemeyi bu adla arar), `%APPDATA%\KMS Sozlesme` ve `Belgeler\KMS Sözleşmeler` klasörleri bilerek eski adla kalır** —
değiştirmek güncelleme zincirini koparır / kullanıcı kayıtlarını kaybettirir. Kullanıcı (Yunus) ve şirketteki yönetici arkadaşları kullanıyor.

- Şirket **KİRACI** taraf: ÖZTAÇ PETROL A.Ş. – Hİ-KA NAK. LTD. ŞTİ. İş Ortaklığı. Sözleşme maddeleri **Kiracı lehine** düzenlenir.
  Şirkete gelen "düzeltme notları" Kiralayan gözünden yazılmıştı; körü körüne uygulanmadı.
- Kullanıcı kararları: cezalar (günlük kiranın 5 katı + 10 aylık cezai şart) **olduğu gibi kalır**; yetkili mahkeme **Mersin**;
  sezon tarihleri her sözleşmede elle girilir; çıktı yalnızca **PDF + Word** (Excel kaldırıldı).
- Şirket hukukçusunun orijinal taslaktaki yorumları işlendi: teslim adresi e-posta ile bildirilebilir, sevk süresi
  (varsayılan 48 saat), banka/IBAN zorunlu, teknik özellikler alanı, ek belge listesi. Hukukçu 10.2 cezai şart için
  "kanunen geçerli değil, gözdağı" demiş; bilerek bırakıldı.
- "Coded by Yunus Açıkgöz" imzası (açılış ekranı, yan menü, Hakkında, exe dosya özellikleri) **korunmalı**.

## Mimari

| Dosya | Görevi |
|---|---|
| `araclar/kms.py` | Giriş noktası. 127.0.0.1'de küçük HTTP sunucu + Edge/Chrome `--app` penceresi (`--start-maximized`). Kayıtlar `%APPDATA%\KMS Sozlesme\kayitlar.json`, çıktılar `Belgeler\KMS Sözleşmeler`. Pencere nabız (`/api/nabiz`) göndermeyi kesince süreç kapanır. |
| `araclar/arayuz.html` | Tek dosyalık arayüz (vanilla JS). Canlı doğrulama: TCKN kontrol hanesi, IBAN `TR`+24, bitiş > başlangıç. |
| `araclar/uret.py` | `SABLON.docx`'teki `{{ALAN}}` yer tutucularını doldurur, araç tablosu + her araç için teslim/iade tutanağı ekler, Word COM ile PDF alır. `{{?ALAN}}` boşsa paragraf silinir. |
| `araclar/guncelleme.py` | GitHub Releases'tan otomatik güncelleme: indir → SHA-256 (`digest`) doğrula → çalışan exe'yi `.eski` adıyla kenara al → yenisini koy → yeniden başlat. |
| `araclar/surum.py`, `araclar/yenilikler.json` | Sürüm numarası ve sürüm notları (notlar hem Release açıklaması hem programdaki tek seferlik "Yenilikler" penceresi). |
| `araclar/derle.py` | PyInstaller tek dosya exe derlemesi (yerelde ve CI'da aynı). `--cikti KLASOR` ile başka klasöre derler. |
| `araclar/yayinla.py` | Kontrolleri yapıp `vX.Y.Z` etiketini gönderir. |
| `.github/workflows/surum-yayinla.yml` | Etiket gelince Windows runner'da exe'yi derleyip Release'e yükler. |
| `SABLON.docx` | Sözleşme şablonu (hukukçu yorumları ve renkli vurgular temizlenmiş). |
| `araclar/bicim.py` | Şablonun kurumsal sayfa düzeni (sablon_olustur'un son adımı): A4, sol 2,5 / diğer 2 cm; Calibri 11, 1,15 satır, iki yana yaslı; "MADDE n: BAŞLIK"; x.y. maddeler 1 cm, x.y.z. maddeler 2,2 cm asılı girinti, numaralar kalın; taraf bilgilerinde iki nokta 4,2 cm'de hizalı; altbilgide sözleşme no + "Sayfa X / Y" + paraf alanı. **Biçim değişikliği burada yapılır, SABLON.docx elle düzenlenmez.** |
| `araclar/sablon_olustur.py` | Şablonu şirketin orijinal taslağından üretir. Orijinal taslak hukukçu yorumları içerdiği için **repoda yok** (`.gitignore`). |
| `araclar/form*.py/html`, `excel_olustur.py`, `SOZLESME_*.bat` | Eski denemeler (tarayıcı formu, Excel ile üretim). Artık kullanılmıyor. |

Kod, tanımlayıcılar ve yorumlar **Türkçe**; aynı üslupla devam et.

## Yeni sürüm yayınlama

1. Değişikliği yap.
2. `araclar/surum.py` → `SURUM`'u artır.
3. `araclar/yenilikler.json` başına yeni sürümün notlarını ekle — **JSON'u Python ile yükleyip yaz**, metin değiştirme ile ekleme kırılgan (bir kez sessizce başarısız oldu; `yayinla.py` yakaladı).
4. Commit + push, sonra `python araclar/yayinla.py` (onay için `e`).
5. Actions ~2 dk'da Release'i çıkarır. Açık programlar 30 dk'da bir denetler; kullanıcı Ayarlar → "Güncellemeleri denetle" ile hemen görebilir.

## Bilinen tuzaklar (yaşandı)

- **PyInstaller tek dosya + kendini yeniden başlatma:** Yeni süreç eski sürecin `_PYI*`/`_MEI*` ortam değişkenlerini devralırsa kendini onun alt süreci sanıp silinen geçici klasörü kullanır ve askıda kalır. Çözüm `kms.py`'de: bu değişkenler temizlenir + `PYINSTALLER_RESET_ENVIRONMENT=1`. v1.2.0 bu hatayı taşıyordu; v1.2.1+ düzgün, v1.2.2 ile gerçek GitHub üzerinden doğrulandı.
- **CI'da Türkçe karakter:** Windows runner konsolu cp1252; `PYTHONUTF8=1` olmadan `print("Hazır")` çöker (workflow'da ayarlı).
- **Test derlemeleri:** Kullanıcının açık programı genellikle `DAGITIM\KMS_Sozlesme.exe`'dir. Test exe'lerini **her zaman** `python araclar/derle.py --cikti <geçici klasör>` ile derle, `DAGITIM`'e değil.
- **Sahte güncelleme sunucusuyla test:** `KMS_GUNCELLEME_API` ortam değişkeni sahte `latest.json`'a yönlendirir. Python `http.server` süreçleri farklı adla (`python3.13.exe`) çalışabilir; kapatırken işlem adına değil komut satırına göre bul. Windows aynı porta iki sunucunun bağlanmasına izin verir — eski sunucu kalırsa yanlış dosya sunulur.
- **Exe çalışırken yeniden adlandırılırsa** `sys.executable` eski adı göstermeye devam eder; v1.4.0'da güncelleme bu yüzden yarıda kaldı. v1.4.1'den beri `guncelleme.py` açılıştaki dosya kimliğinden (st_dev/st_ino) exe'yi aynı klasörde bulur; yarım kalan `.yeni` ve `.eski` dosyaları açılışta temizlenir.
- **Ayarlar anında kaydedilir** (v1.4.0). Önceden yalnızca "Ayarları kaydet" düğmesiyle kaydediliyordu; kullanıcı logo anahtarını kapatıp düğmeye basmadığı için logolar çıkmaya devam etti. Oluştur penceresinde de logo anahtarı var (`cikti.logo`, Ayarlar'daki varsayılandan önceliklidir).
- Şablonu yeniden üretmek için orijinal taslak (`İŞ MAKİNASI KİRALAMA SÖZLEŞMESİ*.docx`, repoda yok) proje köküne konmalı; sonra `python araclar/sablon_olustur.py`.
- Test sırasında kullanıcının gerçek kayıtlarına dokunma: `APPDATA`'yı geçici bir klasöre yönlendir; sunucuyu `--sunucu --port N` ile pencere açmadan çalıştır.

## Mac'te geliştirme

Program Windows'a özel (Word COM ile PDF, `os.startfile`, exe). Mac'te:
- `pywin32` kurulmaz; `pip install python-docx openpyxl` yeterli. Exe Mac'te derlenmez — gerek yok, CI derler.
- `python3 araclar/kms.py` arayüzü ve Word üretimini çalıştırır; PDF ve "Aç/Klasörü aç" çalışmaz.
- Yayın akışı (commit → push → `yayinla.py`) Mac'te aynen çalışır.

## Kullanıcıyla çalışma

- Türkçe, samimi ("kanka") konuşur; teknik olmayan açıklama ister. Büyük işlerde önce plan sunup onay almayı sever.
- Dağıtım: arkadaşlarına Releases'taki son exe'yi gönderir; sonrası otomatik güncellenir.
- Sözleşme metnindeki hukuki değişikliklerde imzadan önce şirket hukukçusuna danışılmasını öner.
