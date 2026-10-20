# KMS Sözleşme

Karla mücadele işleri için iş makinası kiralama sözleşmelerini hazırlayan masaüstü programı.
Kiralayan ve araç bilgileri girilir; Kiracı lehine düzenlenmiş şablondan PDF / Word sözleşme,
her araç için teslim/iade tutanağıyla birlikte üretilir.

**İndir:** [Releases](../../releases/latest) sayfasından `KMS_Sozlesme.exe`. Kurulum gerekmez.
Program yeni sürümleri kendisi bulur ve tek tıkla günceller.

## Yeni sürüm yayınlamak

1. Değişiklikleri yap.
2. `araclar/surum.py` içinde `SURUM`'u artır (ör. `1.1.0` → `1.2.0`).
3. `araclar/yenilikler.json` dosyasının en üstüne yeni sürümün notlarını ekle.
   Bu notlar hem GitHub Release'te hem de güncellemeden sonra programda bir kez gösterilir.
4. Commit + push.
5. `python araclar/yayinla.py`

Etiket (ör. `v1.2.0`) GitHub'a gidince **Actions** exe'yi derleyip Release'e yükler (~5 dk).
Programı açık olan herkes en geç birkaç saat içinde, diğerleri ilk açılışta "Güncelle" uyarısını görür.

## Geliştirme

```
pip install -r requirements.txt
python araclar/kms.py              # programı kaynak koddan çalıştır
python araclar/derle.py            # DAGITIM/KMS_Sozlesme.exe üret
```

| Dosya | Görevi |
|---|---|
| `araclar/kms.py` | Program: pencere, yerel sunucu, kayıtlar |
| `araclar/arayuz.html` | Arayüz |
| `araclar/uret.py` | Şablonu doldurup sözleşme + tutanak üretir |
| `araclar/guncelleme.py` | GitHub Releases'tan otomatik güncelleme |
| `SABLON.docx` | Sözleşme şablonu (`{{ALAN}}` yer tutucuları) |
| `araclar/sablon_olustur.py` | Şablonu şirketin orijinal taslağından üretir (taslak repoda değil) |

Kullanıcı kayıtları `%APPDATA%\KMS Sozlesme`, çıktılar `Belgeler\KMS Sözleşmeler` altında tutulur.

---
Coded by Yunus Açıkgöz
