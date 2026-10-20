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
BASLIK = {"User-Agent": f"KMS-Sozlesme/{SURUM}", "Accept": "application/vnd.github+json"}


def surum_sayisi(metin):
    parcalar = []
    for p in str(metin).lstrip("vV").split("."):
        rakam = "".join(c for c in p if c.isdigit())
        parcalar.append(int(rakam or 0))
    return tuple(parcalar + [0] * (3 - len(parcalar)))


def exe_yolu():
    """Çalışan exe (yalnızca derlenmiş sürümde güncelleme yapılır)."""
    return Path(sys.executable).resolve() if getattr(sys, "frozen", False) else None


def eskileri_temizle():
    exe = exe_yolu()
    if not exe:
        return
    for eski in exe.parent.glob(exe.name + ".eski*"):
        try:
            eski.unlink()
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
                            gelistirme=exe_yolu() is None)
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
