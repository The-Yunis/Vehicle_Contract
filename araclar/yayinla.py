"""Yeni sürümü yayınlar: kontrolleri yapar, sürüm etiketini oluşturup GitHub'a gönderir.
Etiket gidince GitHub Actions exe'yi derleyip Release'e koyar; programlar güncellemeyi görür.

Önce:  araclar/surum.py içindeki SURUM'u artırın, araclar/yenilikler.json'a notları yazın,
       değişiklikleri commit + push edin.
Sonra: python araclar/yayinla.py
"""
import json
import subprocess
import sys
from pathlib import Path

ARACLAR = Path(__file__).resolve().parent
sys.path.insert(0, str(ARACLAR))
from surum import SURUM  # noqa: E402


def git(*args):
    return subprocess.run(["git", *args], cwd=ARACLAR.parent, capture_output=True, text=True, encoding="utf-8")


def main():
    etiket = f"v{SURUM}"
    notlar = json.loads((ARACLAR / "yenilikler.json").read_text(encoding="utf-8")).get(SURUM)
    if not notlar:
        sys.exit(f"HATA: yenilikler.json içinde {SURUM} için not yok.")
    if git("status", "--porcelain").stdout.strip():
        sys.exit("HATA: commit edilmemiş değişiklikler var. Önce commit + push yapın.")
    if git("rev-parse", "-q", "--verify", f"refs/tags/{etiket}").returncode == 0:
        sys.exit(f"HATA: {etiket} zaten var. araclar/surum.py içinde sürümü artırın.")
    git("fetch", "origin")
    if git("rev-list", "--count", "@{u}..HEAD").stdout.strip() not in ("", "0"):
        sys.exit("HATA: gönderilmemiş commit'ler var. Önce 'git push' yapın.")

    print(f"Yayınlanacak sürüm: {etiket}")
    for n in notlar:
        print(f"  - {n}")
    if input("Onaylıyor musunuz? (e/h) ").strip().lower() != "e":
        sys.exit("Vazgeçildi.")
    for komut in (("tag", "-a", etiket, "-m", f"Makine Kontrat {etiket}"), ("push", "origin", etiket)):
        r = git(*komut)
        if r.returncode:
            sys.exit(f"HATA: git {' '.join(komut)}\n{r.stderr}")
    print(f"Gönderildi. Derleme birkaç dakika sürer; GitHub'da Actions sekmesinden izleyebilirsiniz.")


if __name__ == "__main__":
    main()
