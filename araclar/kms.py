"""Makine Kontrat — Araç Kiralama Sözleşme Sistemi (masaüstü uygulaması).

Arayüz, Edge'in uygulama modunda (adres çubuksuz pencere) açılır; arka planda yalnızca
bu bilgisayardan erişilebilen küçük bir yerel sunucu çalışır.
Coded by Yunus Açıkgöz.
"""
import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from guncelleme import Guncelleyici, eskileri_temizle
from surum import SURUM
from uret import kiraci_bilgisi, pdf_yap, sozlesme_isle, tl

AD = "Makine Kontrat"
IMZA = "Coded by Yunus Açıkgöz"
PORT = 8765

# PyInstaller tek dosyada kaynaklar geçici klasöre açılır
KAYNAK = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
VARLIK = KAYNAK / "varliklar"
ARAYUZ = KAYNAK / "arayuz.html"
YENILIKLER = KAYNAK / "yenilikler.json"
VARSAYILAN_SABLON = KAYNAK / "SABLON.docx" if (KAYNAK / "SABLON.docx").exists() \
    else KAYNAK.parent / "SABLON.docx"
LOGOLAR = [VARLIK / "logo_oztac.png", VARLIK / "logo_hika.png"]

# Klasör adları eski adla kalır: mevcut kullanıcıların kayıtları ve çıktıları kaybolmasın
VERI_KLASOR = Path(os.environ.get("APPDATA", Path.home())) / "KMS Sozlesme"
VERI = VERI_KLASOR / "kayitlar.json"
OZEL_SABLON = VERI_KLASOR / "sablon.docx"
GUNLUK = VERI_KLASOR / "hata.log"
VARSAYILAN_CIKTI = Path.home() / "Documents" / "KMS Sözleşmeler"

VARSAYILAN_AYARLAR = {
    "Kiracı Unvanı": "ÖZTAÇ PETROL A.Ş-Hİ-KA NAK.LTD.ŞTİ -İŞ ORTAKLIĞI",
    "Kiracı Adresi": "D 100 GAP MAH. TURGUT ÖZAL BLV. GOLDEN PLAZA 52 521 Batman Merkez/Batman",
    "Kiracı Telefon": "",
    "Kiracı E-posta": "",
    "Kiracı Vergi Dairesi": "BATMAN VERGİ DAİRESİ MÜD.",
    "Kiracı VKN": "7060409340",
    "Kiracı İmza Yetkilisi (Ad Soyad / Unvan)": "",
    "Sevk Talebinden Sonra Teslim Süresi": "48 (kırk sekiz) saat",
    "KDV Oranı (%)": 20,
    "Kullanıcı Adı": "",
    "Kullanıcı Kodu": "",
    "Çıktı Klasörü": str(VARSAYILAN_CIKTI),
    "Çıktılar": {"docx": True, "pdf": True},
    "Son Görülen Sürüm": "",
    "Sözleşmeye Logo Ekle": True,
    "Varsayılan Ödeme Şekli": "cek",          # cek | havale
    "Varsayılan Çek Vadesi (gün)": 90,
}

KILIT = threading.Lock()
SON_NABIZ = [time.time()]
SAYFA_HAZIR = threading.Event()   # arayüz ilk kez tamamen çizildiğinde (/api/hazir)
GUNCELLEYICI = Guncelleyici()
SUNUCU = [None]


def gunluge_yaz(metin):
    try:
        VERI_KLASOR.mkdir(parents=True, exist_ok=True)
        with GUNLUK.open("a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%d.%m.%Y %H:%M:%S}] {metin}\n")
    except OSError:
        pass


def mesaj_kutusu(metin):
    ctypes.windll.user32.MessageBoxW(0, metin, AD, 0x10)


# ---------------------------------------------------------------- kayıtlar

def veri_oku():
    v = json.loads(VERI.read_text(encoding="utf-8")) if VERI.exists() else {}
    v.setdefault("kiralayanlar", [])
    v.setdefault("sozlesmeler", [])
    ayar = dict(VARSAYILAN_AYARLAR)
    ayar.update(v.get("ayarlar", {}))
    v["ayarlar"] = ayar
    return v


def veri_yaz(v):
    VERI_KLASOR.mkdir(parents=True, exist_ok=True)
    gecici = VERI.with_suffix(".tmp")
    gecici.write_text(json.dumps(v, ensure_ascii=False, indent=1), encoding="utf-8")
    if VERI.exists():
        shutil.copyfile(VERI, VERI.with_suffix(".yedek.json"))
    os.replace(gecici, VERI)


def cikti_klasoru(v):
    k = Path(v["ayarlar"].get("Çıktı Klasörü") or VARSAYILAN_CIKTI)
    k.mkdir(parents=True, exist_ok=True)
    return k


def sablon_yolu():
    return OZEL_SABLON if OZEL_SABLON.exists() else VARSAYILAN_SABLON


def kimlik_anahtari(k):
    return re.sub(r"\D", "", (k or {}).get("kimlik", ""))


def sozlesme_kaydet(v, soz, eski_no=None):
    silinecek = {soz["no"], eski_no or soz["no"]}
    v["sozlesmeler"] = [s for s in v["sozlesmeler"] if s["no"] not in silinecek] + [soz]
    v["sozlesmeler"].sort(key=lambda s: s["no"])


def kiralayan_kaydet(v, k, eski_kimlik=None):
    silinecek = {kimlik_anahtari(k), eski_kimlik or kimlik_anahtari(k)}
    v["kiralayanlar"] = [x for x in v["kiralayanlar"] if kimlik_anahtari(x) not in silinecek] + [k]
    v["kiralayanlar"].sort(key=lambda x: x.get("ad", "").casefold())


# ---------------------------------------------------------------- işlemler

def form_to_uretici(soz):
    k = soz.get("kiralayan", {})
    s = {
        "Sözleşme No": soz.get("no", ""),
        "Kiralayan Adı Soyadı / Unvanı": k.get("ad", ""),
        "TCKN (şahıs) / VKN (şirket)": k.get("kimlik", ""),
        "Vergi Dairesi": k.get("vd", ""),
        "Kiralayan Adresi": k.get("adres", ""),
        "Telefon": k.get("tel", ""),
        "E-posta": k.get("eposta", ""),
        "Şirket Yetkilisi": k.get("yetkili", "") if k.get("tur") == "sirket" else "",
        "Banka Adı / Şube": k.get("banka", ""),
        "IBAN": k.get("iban", ""),
        "Teslim (Başlangıç) Tarihi": soz.get("baslangic", ""),
        "Bitiş Tarihi": soz.get("bitis", ""),
        "Teslim Yeri / Adresi": soz.get("teslim_yeri", ""),
        "İmza Tarihi": soz.get("imza", ""),
        "Ödeme Şekli": soz.get("odeme", ""),
        "Çek Vadesi": soz.get("vade", ""),
        "Kiralama Türü": soz.get("kiralama_turu", ""),
    }
    araclar = [(f"{i}. araç", {
        "Plaka": a.get("plaka", ""), "Cinsi": a.get("cins", ""), "Marka": a.get("marka", ""),
        "Model": a.get("model", ""), "Model Yılı": a.get("yil", ""), "Şasi No": a.get("sasi", ""),
        "Teknik Özellikler": a.get("ozellik", ""), "Ataşmanlar": a.get("atasman", ""),
        "Aylık Kira": a.get("bedel") or 0,
    }) for i, a in enumerate(soz.get("araclar", []), 1)]
    return s, araclar


def uret(v, istek):
    soz, secim = istek["sozlesme"], istek.get("cikti") or {"docx": True, "pdf": True}
    if not (secim.get("docx") or secim.get("pdf")):
        return {"ok": False, "hatalar": ["En az bir çıktı türü seçin (PDF veya Word)."], "uyarilar": []}
    ayar = v["ayarlar"]
    R, uyarilar_genel = kiraci_bilgisi(ayar)
    s, araclar = form_to_uretici(soz)
    klasor = cikti_klasoru(v)
    # Oluştur penceresindeki seçim önceliklidir; yoksa Ayarlar'daki varsayılan
    logo_ekle = secim.get("logo", ayar.get("Sözleşmeye Logo Ekle", True))
    logolar = [l for l in LOGOLAR if l.exists()] if logo_ekle else []
    sonuc = sozlesme_isle(s, araclar, R, float(ayar.get("KDV Oranı (%)") or 20), uyarilar_genel,
                          cikti=klasor, sablon=sablon_yolu(), logolar=logolar)
    cevap = {"ok": sonuc["ok"], "hatalar": sonuc["hatalar"],
             "uyarilar": list(dict.fromkeys(sonuc["uyarilar"])), "dosyalar": []}
    eski_no = istek.get("eski_no")
    if sonuc["ok"]:
        docx_yolu = sonuc["hedef"]
        pdf_olmadi = False
        if secim.get("pdf"):
            hata = pdf_yap([docx_yolu], gunluk=gunluge_yaz)
            if hata:
                cevap["uyarilar"].append(hata)
                pdf_olmadi = True
            elif docx_yolu.with_suffix(".pdf").exists():
                cevap["dosyalar"].append(docx_yolu.with_suffix(".pdf").name)
        if secim.get("docx") or pdf_olmadi:  # PDF yapılamadıysa elde en azından Word dosyası kalsın
            cevap["dosyalar"].insert(0, docx_yolu.name)
        else:
            docx_yolu.unlink(missing_ok=True)
        cevap["genel"] = tl(sonuc["genel"])
        soz.update(durum="uretildi", dosyalar=cevap["dosyalar"], uretim=datetime.now().strftime("%d.%m.%Y %H:%M"))
        kiralayan_kaydet(v, soz["kiralayan"])
    else:
        soz.setdefault("durum", "taslak")
        if soz.get("durum") != "uretildi":
            soz["durum"] = "taslak"
    if soz.get("no"):
        sozlesme_kaydet(v, soz, eski_no)
        veri_yaz(v)
    return cevap


def dosya_yolu(v, ad):
    klasor = cikti_klasoru(v).resolve()
    yol = (klasor / Path(ad).name).resolve()
    return yol if yol.parent == klasor and yol.exists() else None


def klasor_sec(baslangic):
    import tkinter
    from tkinter import filedialog
    kok = tkinter.Tk()
    kok.withdraw()
    kok.attributes("-topmost", True)
    secilen = filedialog.askdirectory(initialdir=baslangic, title="Çıktı klasörünü seçin", parent=kok)
    kok.destroy()
    return secilen


# ---------------------------------------------------------------- güncelleme

def yenilikler():
    try:
        return json.loads(YENILIKLER.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guncelle_ve_yeniden_baslat():
    try:
        exe = GUNCELLEYICI.kur()
    except Exception as e:
        gunluge_yaz(traceback.format_exc())
        GUNCELLEYICI.guncelle_durum(asama="hata", hata=str(e))
        return
    time.sleep(0.8)  # arayüz "yeniden başlatılıyor" durumunu görsün
    port = SUNUCU[0].server_address[1]
    SUNUCU[0].shutdown()
    SUNUCU[0].server_close()
    # Yeni sürüm aynı portta açılır; mevcut pencere onu bulup kendini yeniler.
    # PyInstaller'ın ortam değişkenleri yeni sürece geçerse yeni exe kendini bizim alt sürecimiz
    # sanıp bizim (birazdan silinecek) geçici klasörümüzü kullanır ve çöker; temiz başlatıyoruz.
    ortam = {k: v for k, v in os.environ.items() if not k.upper().startswith(("_PYI", "_MEI"))}
    ortam["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    subprocess.Popen([str(exe), "--devam", "--port", str(port)], env=ortam, close_fds=True,
                     creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
    os._exit(0)


def arka_planda_denetle():
    time.sleep(3)
    while True:
        GUNCELLEYICI.denetle()
        time.sleep(30 * 60)


# ---------------------------------------------------------------- http

class Isleyici(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def gonder(self, kod, govde, tur="application/json; charset=utf-8"):
        if not isinstance(govde, bytes):
            govde = json.dumps(govde, ensure_ascii=False).encode("utf-8")
        self.send_response(kod)
        self.send_header("Content-Type", tur)
        self.send_header("Content-Length", str(len(govde)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(govde)

    def do_GET(self):
        yol = self.path.split("?")[0]
        if yol in ("/", "/index.html"):
            return self.gonder(200, ARAYUZ.read_bytes(), "text/html; charset=utf-8")
        if yol.startswith("/varlik/"):
            dosya = VARLIK / Path(yol).name
            if dosya.exists():
                return self.gonder(200, dosya.read_bytes(), "image/png")
        if yol == "/api/kimlik":
            return self.gonder(200, {"uygulama": "KMS", "surum": SURUM})
        if yol == "/api/veri":
            with KILIT:
                v = veri_oku()
            v.update(surum=SURUM, imza=IMZA, sablon_ozel=OZEL_SABLON.exists(), veri_klasoru=str(VERI_KLASOR),
                     yenilikler=yenilikler())
            return self.gonder(200, v)
        if yol == "/api/guncelleme":
            return self.gonder(200, GUNCELLEYICI.ozet())
        self.gonder(404, {"hata": "yok"})

    def do_POST(self):
        koken = self.headers.get("Origin")
        port = self.server.server_address[1]
        if koken and koken not in (f"http://127.0.0.1:{port}", f"http://localhost:{port}"):
            return self.gonder(403, {"hata": "izin yok"})
        try:
            uzunluk = int(self.headers.get("Content-Length") or 0)
            istek = json.loads(self.rfile.read(uzunluk) or b"{}")
            if self.path == "/api/hazir":
                SAYFA_HAZIR.set()
                return self.gonder(200, {"ok": True})
            if self.path == "/api/nabiz":
                SON_NABIZ[0] = time.time()
                return self.gonder(200, {"ok": True})
            if self.path == "/api/guncelleme/denetle":
                return self.gonder(200, GUNCELLEYICI.denetle())
            if self.path == "/api/guncelleme/kur":
                threading.Thread(target=guncelle_ve_yeniden_baslat, daemon=True).start()
                return self.gonder(200, {"ok": True})
            with KILIT:
                return self.gonder(200, self.islem(istek))
        except Exception as e:
            gunluge_yaz(traceback.format_exc())
            self.gonder(200, {"ok": False, "hatalar": [f"Beklenmeyen hata: {e}"], "uyarilar": []})

    def islem(self, istek):
        v = veri_oku()
        y = self.path
        if y == "/api/uret":
            return uret(v, istek)
        if y == "/api/kaydet":
            soz = istek["sozlesme"]
            soz["durum"] = "taslak"
            sozlesme_kaydet(v, soz, istek.get("eski_no"))
            veri_yaz(v)
            return {"ok": True}
        if y == "/api/sil":
            v["sozlesmeler"] = [s for s in v["sozlesmeler"] if s["no"] != istek.get("no")]
            veri_yaz(v)
            return {"ok": True}
        if y == "/api/kiralayan_kaydet":
            kiralayan_kaydet(v, istek["kiralayan"], istek.get("eski_kimlik"))
            veri_yaz(v)
            return {"ok": True}
        if y == "/api/kiralayan_sil":
            v["kiralayanlar"] = [k for k in v["kiralayanlar"] if kimlik_anahtari(k) != istek.get("kimlik")]
            veri_yaz(v)
            return {"ok": True}
        if y == "/api/ayarlar":
            v["ayarlar"].update(istek)
            veri_yaz(v)
            return {"ok": True}
        if y == "/api/ac":
            if istek.get("klasor"):
                os.startfile(cikti_klasoru(v))
                return {"ok": True}
            yol = dosya_yolu(v, istek.get("dosya", ""))
            if not yol:
                return {"ok": False, "hata": "Dosya bulunamadı; taşınmış veya silinmiş olabilir. Yeniden üretin."}
            os.startfile(yol)
            return {"ok": True}
        if y == "/api/klasor_sec":
            secilen = klasor_sec(str(cikti_klasoru(v)))
            return {"ok": bool(secilen), "klasor": str(Path(secilen)) if secilen else ""}
        if y == "/api/sablon":
            islem = istek.get("islem")
            if islem == "duzenle":
                if not OZEL_SABLON.exists():
                    VERI_KLASOR.mkdir(parents=True, exist_ok=True)
                    shutil.copyfile(VARSAYILAN_SABLON, OZEL_SABLON)
                os.startfile(OZEL_SABLON)
                return {"ok": True}
            if islem == "varsayilan":
                if OZEL_SABLON.exists():
                    shutil.move(OZEL_SABLON, OZEL_SABLON.with_name(f"sablon_eski_{datetime.now():%Y%m%d_%H%M%S}.docx"))
                return {"ok": True}
            if islem == "kopyala":
                hedef = cikti_klasoru(v) / "Sablon_Kopyasi.docx"
                shutil.copyfile(sablon_yolu(), hedef)
                os.startfile(hedef)
                return {"ok": True}
        return {"ok": False, "hata": "Bilinmeyen işlem"}


# ---------------------------------------------------------------- pencere ve yaşam döngüsü

def edge_bul():
    adaylar = [Path(os.environ.get(d, "")) / "Microsoft/Edge/Application/msedge.exe"
               for d in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA")]
    adaylar += [Path(os.environ.get(d, "")) / "Google/Chrome/Application/chrome.exe"
                for d in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA")]
    return next((p for p in adaylar if p.exists()), None)


def pencere_ac(adres):
    tarayici = edge_bul()
    if not tarayici:
        webbrowser.open(adres)
        return None
    return subprocess.Popen([
        str(tarayici), f"--app={adres}", f"--user-data-dir={VERI_KLASOR / 'pencere'}",
        "--start-maximized",  # pencere her açılışta tam ekran büyütülmüş gelir
        "--no-first-run", "--no-default-browser-check",
        "--disable-features=Translate,msEdgeSidebarV2", "--disable-sync",
    ])


class AcilisEkrani:
    """Edge penceresi ilk çizimini yapana kadar (o arada pencere şeffaf görünür ve arkadaki
    masaüstü seçilir) çalışma alanını kaplayan beyaz bir açılış ekranı gösterir."""

    LACI, SOLUK, SARI = "#003a5d", "#667085", "#e8b400"

    def __init__(self):
        import ctypes.wintypes
        import tkinter as tk
        self.tk = tk
        self.kok = kok = tk.Tk()
        kok.overrideredirect(True)
        kok.attributes("-topmost", True)
        kok.configure(bg="white")
        alan = ctypes.wintypes.RECT()   # görev çubuğu hariç çalışma alanı (büyütülmüş pencereyle aynı)
        ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(alan), 0)
        kok.geometry(f"{alan.right - alan.left}x{alan.bottom - alan.top}+{alan.left}+{alan.top}")
        orta = tk.Frame(kok, bg="white")
        orta.place(relx=0.5, rely=0.5, anchor="center")
        logolar = tk.Frame(orta, bg="white")
        logolar.pack()
        self.resimler = [tk.PhotoImage(file=str(VARLIK / ad)) for ad in ("acilis_oztac.png", "acilis_hika.png")]
        tk.Label(logolar, image=self.resimler[0], bg="white", bd=0).pack(side="left")
        tk.Frame(logolar, bg="#cfd6e0", width=1, height=56).pack(side="left", padx=34)
        tk.Label(logolar, image=self.resimler[1], bg="white", bd=0).pack(side="left")
        tk.Label(orta, text="Makine Kontrat", font=("Segoe UI", 20, "bold"), fg=self.LACI, bg="white").pack(pady=(30, 0))
        tk.Label(orta, text="Araç Kiralama Sözleşme Sistemi · Karla Mücadele", font=("Segoe UI", 10),
                 fg=self.SOLUK, bg="white").pack(pady=(2, 0))
        self.cubuk = tk.Canvas(orta, width=220, height=3, bg="#e3e8ef", highlightthickness=0)
        self.cubuk.pack(pady=(28, 0))
        self.serit = self.cubuk.create_rectangle(0, 0, 88, 3, fill=self.SARI, width=0)
        imza = tk.Frame(kok, bg="white")
        imza.place(relx=0.5, rely=1.0, y=-26, anchor="s")
        tk.Label(imza, text="Coded by ", font=("Segoe UI", 9), fg=self.SOLUK, bg="white").pack(side="left")
        tk.Label(imza, text="Yunus Açıkgöz", font=("Segoe UI", 9, "bold"), fg=self.LACI, bg="white").pack(side="left")
        self.konum = -88
        kok.update()

    def bekle(self, olay, en_az=1.2, en_cok=25):
        bas = time.time()

        def tik():
            self.konum = self.konum + 6 if self.konum < 220 else -88
            self.cubuk.coords(self.serit, self.konum, 0, self.konum + 88, 3)
            gecen = time.time() - bas
            if (olay.is_set() and gecen >= en_az) or gecen >= en_cok:
                self.kok.after(150, self.kok.destroy)   # sayfa ekrana basılsın, sonra kalk
                return
            self.kok.after(16, tik)

        tik()
        self.kok.mainloop()


def calisan_uygulama_var():
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/kimlik", timeout=1) as r:
            return json.load(r).get("uygulama") == "KMS"
    except Exception:
        return False


def arguman(ad, varsayilan=None):
    return sys.argv[sys.argv.index(ad) + 1] if ad in sys.argv[:-1] else varsayilan


def main():
    eskileri_temizle()
    devam = "--devam" in sys.argv  # güncellemeden sonra: pencere zaten açık
    port = int(arguman("--port", PORT))
    sunucu = None
    for deneme in range(20 if devam else 1):  # eski süreç portu bırakana kadar bekle
        try:
            sunucu = ThreadingHTTPServer(("127.0.0.1", port), Isleyici)
            break
        except OSError:
            time.sleep(0.5)
    if sunucu is None:
        if not devam and calisan_uygulama_var():  # zaten açık: sadece yeni pencere
            pencere_ac(f"http://127.0.0.1:{PORT}/")
            return
        sunucu = ThreadingHTTPServer(("127.0.0.1", 0), Isleyici)
        devam = False
    SUNUCU[0] = sunucu
    adres = f"http://127.0.0.1:{sunucu.server_address[1]}/"
    threading.Thread(target=sunucu.serve_forever, daemon=True).start()
    threading.Thread(target=arka_planda_denetle, daemon=True).start()
    if "--sunucu" in sys.argv:  # geliştirme/test: pencere açmadan çalış
        if sys.stdout:  # konsolsuz exe'de stdout yoktur
            print(adres, flush=True)
        threading.Event().wait()
    acilis = None
    if not devam:
        try:
            acilis = AcilisEkrani()
        except Exception:
            gunluge_yaz("Açılış ekranı gösterilemedi:\n" + traceback.format_exc())
    # Açılış ekranı varsa sayfa kendi açılış animasyonunu atlar ve hazır olunca haber verir
    surec = None if devam else pencere_ac(adres + ("?yerel_acilis=1" if acilis else ""))
    if acilis:
        acilis.bekle(SAYFA_HAZIR)
    SON_NABIZ[0] = time.time()
    # Pencere kapanınca kapan: tarayıcı süreci bitti ve sayfa artık nabız göndermiyor
    while True:
        time.sleep(2)
        sessiz = time.time() - SON_NABIZ[0]
        if surec is not None and surec.poll() is not None and sessiz > 8:
            break
        if sessiz > 180:  # simge durumundaki pencerelerde zamanlayıcılar yavaşlar
            if surec is None or surec.poll() is not None:
                break
    sunucu.shutdown()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        gunluge_yaz(traceback.format_exc())
        mesaj_kutusu(f"Program başlatılamadı. Ayrıntılar:\n{GUNLUK}")
