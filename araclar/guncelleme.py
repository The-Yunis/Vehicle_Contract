"""GitHub Releases üzerinden otomatik güncelleme.

Akış: denetle() son sürümü sorar -> kur() yeni exe'yi exe'nin yanına indirir, SHA-256 ile
doğrular, çalışan exe'yi '.eski' adıyla kenara alır (Windows çalışan exe'nin yeniden
adlandırılmasına izin verir), yenisini yerine koyar ve yeniden başlatma için hazır olur.
"""
import hashlib
import json
import os
import sys
import threading
import urllib.request
from pathlib import Path

from surum import EXE_ADI, REPO, SURUM

# KMS_GUNCELLEME_API yalnızca test içindir (sahte bir sürüm sunucusuna yönlendirmek için)
API = os.environ.get("KMS_GUNCELLEME_API") or f"https://api.github.com/repos/{REPO}/releases/latest"
BASLIK = {"User-Agent": f"Makine-Kontrat/{SURUM}", "Accept": "application/vnd.github+json"}


def surum_sayisi(metin):
    parcalar = []
    for p in str(metin).lstrip("vV").split("."):
        rakam = "".join(c for c in p if c.isdigit())
        parcalar.append(int(rakam or 0))
    return tuple(parcalar + [0] * (3 - len(parcalar)))


def _dosya_kimligi(yol):
    s = os.stat(yol)
    return s.st_dev, s.st_ino


# Açılışta exe'nin dosya kimliği (NTFS dosya numarası) alınır: kullanıcı program açıkken
# exe'nin adını değiştirse bile sys.executable eski adı göstermeye devam eder; dosyayı
# bu kimlikle yeniden buluruz.
_ACILIS_EXE = Path(sys.executable).resolve() if getattr(sys, "frozen", False) else None
try:
    _EXE_KIMLIGI = _dosya_kimligi(_ACILIS_EXE) if _ACILIS_EXE else None
except OSError:
    _EXE_KIMLIGI = None


def exe_yolu():
    """Çalışan exe'nin ŞU ANKİ yolu (geliştirme sürümünde None).
    Exe çalışırken yeniden adlandırıldıysa aynı klasörde dosya kimliğinden bulunur."""
    if _ACILIS_EXE is None:
        return None
    if _ACILIS_EXE.exists():
        return _ACILIS_EXE
    if _EXE_KIMLIGI:
        for aday in _ACILIS_EXE.parent.glob("*.exe"):
            try:
                if _dosya_kimligi(aday) == _EXE_KIMLIGI:
                    return aday.resolve()
            except OSError:
                continue
    raise RuntimeError("Program dosyası çalışırken başka bir klasöre taşınmış veya silinmiş. "
                       "Programı kapatıp yeni yerinden açın, sonra tekrar güncelleyin.")


def eskileri_temizle():
    """Önceki güncellemelerden kalan '.eski' / yarım kalmış '.yeni' dosyalarını siler."""
    try:
        exe = exe_yolu()
    except RuntimeError:
        return
    if not exe:
        return
    for artik in [*exe.parent.glob("*.exe.eski*"), *exe.parent.glob("*.exe.yeni")]:
        try:
            artik.unlink()
        except OSError:
            pass  # önceki süreç henüz kapanmadıysa bir sonraki açılışta silinir


class Guncelleyici:
    def __init__(self):
        self.kilit = threading.Lock()
        self.durum = {"asama": "bekliyor", "var": False, "surum": SURUM}
        self.varlik = None

    def guncelle_durum(self, **k):
        with self.kilit:
            self.durum.update(k)

    def ozet(self):
        with self.kilit:
            return dict(self.durum)

    def denetle(self):
        if self.ozet()["asama"] in ("indiriliyor", "kuruluyor", "hazir"):
            return self.ozet()
        self.guncelle_durum(asama="denetleniyor", hata="")
        try:
            istek = urllib.request.Request(API, headers=BASLIK)
            with urllib.request.urlopen(istek, timeout=10) as r:
                rel = json.load(r)
        except Exception as e:
            self.guncelle_durum(asama="bekliyor", hata=f"Güncelleme sunucusuna ulaşılamadı: {e}")
            return self.ozet()
        yeni = rel.get("tag_name", "")
        varlik = next((a for a in rel.get("assets", []) if a.get("name") == EXE_ADI), None)
        var = bool(varlik) and surum_sayisi(yeni) > surum_sayisi(SURUM)
        self.varlik = varlik
        self.guncelle_durum(asama="bekliyor", var=var, yeni=yeni.lstrip("vV"), notlar=rel.get("body") or "",
                            boyut=(varlik or {}).get("size", 0), tarih=rel.get("published_at", ""),
                            gelistirme=_ACILIS_EXE is None)
        return self.ozet()

    def kur(self):
        """Yeni exe'yi indirip yerine koyar. Başarılıysa yeni exe'nin yolunu döndürür."""
        exe = exe_yolu()
        if exe is None:
            raise RuntimeError("Geliştirme sürümünde güncelleme yapılamaz.")
        if not self.varlik:
            raise RuntimeError("İndirilecek güncelleme bulunamadı.")
        yeni = exe.with_name(exe.name + ".yeni")
        boyut = self.varlik.get("size") or 0
        self.guncelle_durum(asama="indiriliyor", yuzde=0, hata="")
        ozet = hashlib.sha256()
        try:
            istek = urllib.request.Request(self.varlik["browser_download_url"], headers={"User-Agent": BASLIK["User-Agent"]})
            with urllib.request.urlopen(istek, timeout=30) as r, open(yeni, "wb") as f:
                inen = 0
                while parca := r.read(256 * 1024):
                    f.write(parca)
                    ozet.update(parca)
                    inen += len(parca)
                    if boyut:
                        self.guncelle_durum(yuzde=int(inen * 100 / boyut))
        except PermissionError:
            raise RuntimeError(f"Programın bulunduğu klasöre yazılamıyor ({exe.parent}). "
                               "Exe'yi Masaüstü veya Belgeler gibi bir klasöre taşıyıp tekrar deneyin.")

        self.guncelle_durum(asama="kuruluyor", yuzde=100)
        beklenen = (self.varlik.get("digest") or "").removeprefix("sha256:")
        if (boyut and yeni.stat().st_size != boyut) or (beklenen and ozet.hexdigest() != beklenen):
            yeni.unlink(missing_ok=True)
            raise RuntimeError("İndirilen dosya doğrulanamadı (eksik veya bozuk). Tekrar deneyin.")

        eski = exe.with_name(exe.name + f".eski{os.getpid()}")
        os.replace(exe, eski)        # çalışan exe kenara alınır
        try:
            os.replace(yeni, exe)
        except OSError:
            os.replace(eski, exe)    # geri al
            raise
        self.guncelle_durum(asama="hazir")
        return exe
