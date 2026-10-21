"""Şablon oluşturucu ve sözleşme üretici için ortak yardımcılar."""
import copy
import re

from docx.text.paragraph import Paragraph

ALAN = re.compile(r"\{\{\??[A-Z0-9_]+\}\}")


def metin_degistir(p, eski, yeni):
    """Paragraftaki `eski` metni `yeni` ile değiştirir; metin birden fazla run'a
    bölünmüş olsa bile çalışır (Word'de elle düzenlenen şablonlar için)."""
    while True:
        runs = p.runs
        tam = "".join(r.text for r in runs)
        i = tam.find(eski)
        if i < 0 or not runs:
            return
        j = i + len(eski)
        pos = 0
        ilk = True
        for r in runs:
            bas, son = pos, pos + len(r.text)
            pos = son
            if son <= i or bas >= j:
                continue
            t = r.text
            sol = t[: max(i - bas, 0)]
            sag = t[max(j - bas, 0):] if son > j else ""
            if ilk:
                r.text = sol + yeni + sag
                ilk = False
            else:
                r.text = sag
        if eski in yeni:
            return


def araliga_yaz(p, bas, son, yeni):
    """Paragraf metninin [bas, son) aralığını `yeni` ile değiştirir (run sınırlarını aşsa da)."""
    pos = 0
    ilk = True
    for r in p.runs:
        rbas, rson = pos, pos + len(r.text)
        pos = rson
        if rson <= bas or rbas >= son:
            continue
        t = r.text
        sol = t[:max(bas - rbas, 0)]
        sag = t[max(son - rbas, 0):] if rson > son else ""
        if ilk:
            r.text = sol + yeni + sag
            ilk = False
        else:
            r.text = sag
    if ilk:  # aralık boşsa (saf ekleme): konumdaki run'ın içine yaz
        pos = 0
        for r in p.runs:
            if pos <= bas <= pos + len(r.text):
                r.text = r.text[:bas - pos] + yeni + r.text[bas - pos:]
                return
            pos += len(r.text)


def paragraf_yaz(p, parcalar):
    """Paragrafın tüm run'larını silip (metin, kalın) parçalarıyla yeniden yazar.
    Paragraf biçimi (girinti, hizalama, stil) korunur."""
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for metin, kalin in parcalar:
        r = p.add_run(metin)
        r.bold = kalin or None
    return p


def paragraf_ekle_sonra(p, parcalar):
    """`p`'nin biçimini kopyalayarak hemen arkasına yeni paragraf ekler."""
    yeni = copy.deepcopy(p._p)
    p._p.addnext(yeni)
    return paragraf_yaz(Paragraph(yeni, p._parent), parcalar)


def paragraf_sil(p):
    p._p.getparent().remove(p._p)
