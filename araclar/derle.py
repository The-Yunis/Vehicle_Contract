"""KMS_Sozlesme.exe'yi derler. Hem yerelde hem GitHub Actions'ta aynı şekilde çalışır.

Kullanım:  python araclar/derle.py [--etiket v1.2.0] [--cikti KLASOR]
--etiket verilirse etiketin surum.py ile aynı olduğu doğrulanır ve sürüm notları
build/notlar.md dosyasına yazılır (GitHub Release açıklaması için).
"""
import json
import sys
from pathlib import Path

import PyInstaller.__main__

ARACLAR = Path(__file__).resolve().parent
KOK = ARACLAR.parent
sys.path.insert(0, str(ARACLAR))
from surum import EXE_ADI, SURUM  # noqa: E402

BUILD = KOK / "build"
CIKTI = KOK / "DAGITIM"
SURUM_DOSYASI = """VSVersionInfo(
  ffi=FixedFileInfo(filevers=({v}, 0), prodvers=({v}, 0), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041F04B0', [
      StringStruct('CompanyName', 'ÖZTAÇ PETROL A.Ş. - Hİ-KA NAK. LTD. ŞTİ. İş Ortaklığı'),
      StringStruct('FileDescription', 'Makine Kontrat - Araç Kiralama Sözleşme Sistemi'),
      StringStruct('FileVersion', '{s}'),
      StringStruct('InternalName', 'Makine_Kontrat'),
      StringStruct('LegalCopyright', 'Coded by Yunus Açıkgöz'),
      StringStruct('OriginalFilename', '{exe}'),
      StringStruct('ProductName', 'Makine Kontrat'),
      StringStruct('ProductVersion', '{s}'),
      StringStruct('Comments', 'Coded by Yunus Açıkgöz')])]),
    VarFileInfo([VarStruct('Translation', [0x041F, 1200])])
  ]
)
"""


def main():
    yenilikler = json.loads((ARACLAR / "yenilikler.json").read_text(encoding="utf-8"))
    if SURUM not in yenilikler:
        sys.exit(f"HATA: yenilikler.json içinde {SURUM} için not yok. Önce yenilikleri yazın.")

    BUILD.mkdir(exist_ok=True)
    global CIKTI
    if "--cikti" in sys.argv:  # ör. test derlemeleri, çalışan exe'nin üzerine yazmamak için
        CIKTI = Path(sys.argv[sys.argv.index("--cikti") + 1]).resolve()
    if "--etiket" in sys.argv:
        etiket = sys.argv[sys.argv.index("--etiket") + 1]
        if etiket.lstrip("vV") != SURUM:
            sys.exit(f"HATA: etiket ({etiket}) ile surum.py ({SURUM}) aynı değil.")
        notlar = "\n".join(f"- {m}" for m in yenilikler[SURUM])
        (BUILD / "notlar.md").write_text(f"## Yenilikler\n\n{notlar}\n\n_Coded by Yunus Açıkgöz_\n", encoding="utf-8")

    v = ", ".join(SURUM.split(".")[:3])
    surum_bilgisi = BUILD / "surum_bilgisi.txt"
    surum_bilgisi.write_text(SURUM_DOSYASI.format(v=v, s=SURUM, exe=EXE_ADI), encoding="utf-8")

    PyInstaller.__main__.run([
        "--noconfirm", "--clean", "--onefile", "--noconsole",
        "--name", Path(EXE_ADI).stem,
        "--icon", str(ARACLAR / "varliklar" / "simge.ico"),
        "--version-file", str(surum_bilgisi),
        "--add-data", f"{ARACLAR / 'arayuz.html'};.",
        "--add-data", f"{ARACLAR / 'yenilikler.json'};.",
        "--add-data", f"{ARACLAR / 'varliklar'};varliklar",
        "--add-data", f"{KOK / 'SABLON.docx'};.",
        "--hidden-import", "pythoncom", "--hidden-import", "win32com.client",
        # openpyxl yalnızca eski Excel yolunda kullanılır; exe'de yok (daha küçük exe, daha hızlı açılış)
        *[a for m in ("numpy", "pandas", "matplotlib", "PIL", "IPython", "openpyxl", "et_xmlfile")
          for a in ("--exclude-module", m)],
        # exe kendini açarken anında görünen küçük açılış görseli
        "--splash", str(ARACLAR / "varliklar" / "yukleniyor.png"),
        "--paths", str(ARACLAR),
        "--workpath", str(BUILD / "pyinstaller"), "--specpath", str(BUILD),
        "--distpath", str(CIKTI),
        str(ARACLAR / "kms.py"),
    ])
    print(f"Hazır: {CIKTI / EXE_ADI} (v{SURUM})")


if __name__ == "__main__":
    main()
