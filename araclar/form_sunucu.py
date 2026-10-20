"""Sözleşme formu: tarayıcıda açılan yerel form + sözleşme üretimi.

Sadece bu bilgisayardan (127.0.0.1) erişilir, internet gerekmez.
Kayıtlar veri/kayitlar.json dosyasında tutulur.
"""
import json
import os
import re
import shutil
import sys
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

from uret import CIKTI, KOK, kiraci_bilgisi, pdf_yap, sozlesme_isle, tl

PORT = 8765
HTML = Path(__file__).with_name("form.html")
VERI = KOK / "veri" / "kayitlar.json"

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
}


# ---------------------------------------------------------------- kayıt dosyası

def veri_oku():
    if VERI.exists():
        return json.loads(VERI.read_text(encoding="utf-8"))
    return {"ayarlar": dict(VARSAYILAN_AYARLAR), "kiralayanlar": [], "sozlesmeler": []}


def veri_yaz(v):
    VERI.parent.mkdir(exist_ok=True)
    gecici = VERI.with_suffix(".tmp")
    gecici.write_text(json.dumps(v, ensure_ascii=False, indent=1), encoding="utf-8")
    if VERI.exists():  # bir önceki hâli yedek olarak kalsın
        shutil.copyfile(VERI, VERI.with_suffix(".yedek.json"))
    os.replace(gecici, VERI)


def sozlesme_kaydet(v, soz):
    v["sozlesmeler"] = [s for s in v["sozlesmeler"] if s["no"] != soz["no"]] + [soz]
    v["sozlesmeler"].sort(key=lambda s: s["no"])


def kiralayan_kaydet(v, k):
    anahtar = re.sub(r"\D", "", k.get("kimlik", ""))
    v["kiralayanlar"] = [x for x in v["kiralayanlar"]
                         if re.sub(r"\D", "", x.get("kimlik", "")) != anahtar] + [k]
    v["kiralayanlar"].sort(key=lambda x: x.get("ad", ""))


# ---------------------------------------------------------------- form -> üretici

def excel_bicimine(soz):
    """Form verisini üreticinin beklediği (Excel başlıklı) sözlüklere çevirir."""
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
    }
    araclar = []
    for i, a in enumerate(soz.get("araclar", []), 1):
        araclar.append((f"{i}. araç", {
            "Plaka": a.get("plaka", ""), "Cinsi": a.get("cins", ""), "Marka": a.get("marka", ""),
            "Model": a.get("model", ""), "Model Yılı": a.get("yil", ""), "Şasi No": a.get("sasi", ""),
            "Teknik Özellikler": a.get("ozellik", ""), "Ataşmanlar": a.get("atasman", ""),
            "Aylık Kira": a.get("bedel") or 0,
        }))
    return s, araclar


def uret(v, soz):
    ayar = v["ayarlar"]
    R, uyarilar_genel = kiraci_bilgisi(ayar)
    s, araclar = excel_bicimine(soz)
    sonuc = sozlesme_isle(s, araclar, R, float(ayar.get("KDV Oranı (%)") or 20), uyarilar_genel)
    cevap = {"ok": sonuc["ok"], "hatalar": sonuc["hatalar"],
             "uyarilar": list(dict.fromkeys(sonuc["uyarilar"]))}
    if sonuc["ok"]:
        pdf_hatasi = pdf_yap([sonuc["hedef"]])
        cevap.update(dosya=sonuc["hedef"].name, genel=tl(sonuc["genel"]))
        if pdf_hatasi:
            cevap["uyarilar"].append(pdf_hatasi)
        soz["durum"] = "uretildi"
        soz["dosya"] = sonuc["hedef"].name
        soz["uretim"] = datetime.now().strftime("%d.%m.%Y %H:%M")
        kiralayan_kaydet(v, soz["kiralayan"])
    sozlesme_kaydet(v, soz)
    veri_yaz(v)
    return cevap


def cikti_yolu(ad):
    """Sadece CIKTI klasöründeki dosyalara izin ver."""
    yol = (CIKTI / Path(ad).name).resolve()
    return yol if yol.parent == CIKTI.resolve() and yol.exists() else None


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
        if self.path in ("/", "/index.html"):
            self.gonder(200, HTML.read_bytes(), "text/html; charset=utf-8")
        elif self.path == "/api/veri":
            self.gonder(200, veri_oku())
        else:
            self.gonder(404, {"hata": "yok"})

    def do_POST(self):
        # Başka sitelerin bu yerel sunucuya istek atmasını engelle
        koken = self.headers.get("Origin")
        if koken and koken not in (f"http://127.0.0.1:{PORT}", f"http://localhost:{PORT}"):
            return self.gonder(403, {"hata": "izin yok"})
        try:
            uzunluk = int(self.headers.get("Content-Length") or 0)
            istek = json.loads(self.rfile.read(uzunluk) or b"{}")
            v = veri_oku()
            if self.path == "/api/uret":
                return self.gonder(200, uret(v, istek))
            if self.path == "/api/kaydet":
                istek.setdefault("durum", "taslak")
                sozlesme_kaydet(v, istek)
                veri_yaz(v)
                return self.gonder(200, {"ok": True})
            if self.path == "/api/sil":
                v["sozlesmeler"] = [s for s in v["sozlesmeler"] if s["no"] != istek.get("no")]
                veri_yaz(v)
                return self.gonder(200, {"ok": True})
            if self.path == "/api/ayarlar":
                v["ayarlar"].update(istek)
                veri_yaz(v)
                return self.gonder(200, {"ok": True})
            if self.path == "/api/ac":
                if istek.get("klasor"):
                    CIKTI.mkdir(exist_ok=True)
                    os.startfile(CIKTI)
                    return self.gonder(200, {"ok": True})
                yol = cikti_yolu(istek.get("dosya", ""))
                if not yol:
                    return self.gonder(200, {"ok": False, "hata": "Dosya bulunamadı; sözleşmeyi yeniden üretin."})
                os.startfile(yol)
                return self.gonder(200, {"ok": True})
            self.gonder(404, {"hata": "yok"})
        except Exception as e:  # formda gösterilsin, sunucu kapanmasın
            self.gonder(200, {"ok": False, "hatalar": [f"Beklenmeyen hata: {e}"], "uyarilar": []})


def main():
    adres = f"http://127.0.0.1:{PORT}/"
    try:
        sunucu = HTTPServer(("127.0.0.1", PORT), Isleyici)
    except OSError:  # zaten açık: sadece tarayıcıyı aç
        webbrowser.open(adres)
        return
    print(f"Form açıldı: {adres}")
    print("Bu pencere açık kaldığı sürece form çalışır. İşiniz bitince pencereyi kapatın.")
    if "--sessiz" not in sys.argv:  # test için tarayıcı açmadan
        webbrowser.open(adres)
    try:
        sunucu.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
