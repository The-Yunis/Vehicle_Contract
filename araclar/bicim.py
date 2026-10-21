"""Sözleşme şablonuna kurumsal sayfa düzeni uygular (sablon_olustur.py'nin son adımı).

Kurallar: A4, sol 2,5 / diğer 2 cm kenar; tek yazı tipi (Calibri 11), satır aralığı 1,15, iki yana yaslı;
"MADDE n: BAŞLIK" başlıkları; asılı girintili madde numaraları (x.y. bir kademe, x.y.z. iki kademe);
taraf bilgilerinde iki nokta tek hizada; altbilgide sözleşme no, sayfa no ve paraf alanı.
"""
import copy
import re

from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

from ortak import metin_degistir, paragraf_sil, paragraf_yaz

YAZI_TIPI = "Calibri"
BOYUT = Pt(11)
METIN_GENISLIGI = Cm(16.5)        # 21 - 2,5 - 2
KADEME1 = Cm(1.0)                 # x.y. maddelerin metin hizası
KADEME2 = Cm(2.2)                 # x.y.z. maddelerin metin hizası
LISTE = Cm(1.6)                   # tire/madde imli listelerin metin hizası

ETIKETLER = ("KİRALAYAN", "KİRACI", "Adresi", "Telefon", "E-posta", "Vergi Dairesi", "Vergi No", "Yetkili")
ETIKET = re.compile(r"^(%s)[\s ]*:[\s ]*" % "|".join(map(re.escape, ETIKETLER)))
MADDE = re.compile(r"^MADDE[\s\-–]*(\d+)\s*[:.\-–]\s*(.+)$")
NUMARA = re.compile(r"^(\d+(?:\.\d+)+)\.?[\s ]*")
BASLIK_DUZELT = {
    "ARACIN TESLİMİ- İADESİ ile  İKAME ARAÇ ŞARTI": "ARACIN TESLİMİ, İADESİ VE İKAME ARAÇ ŞARTI",
    "YÜRÜRLÜLÜK": "YÜRÜRLÜK",
}


def girinti(p, sol, asili=None, hiza=WD_ALIGN_PARAGRAPH.JUSTIFY):
    f = p.paragraph_format
    f.left_indent = sol
    f.first_line_indent = -asili if asili else Cm(0)
    p.alignment = hiza


def kalin_onek(p, onek):
    """Paragraf `onek` ile başlıyorsa önekin ayrı ve kalın bir run olmasını sağlar."""
    if not p.runs or not p.runs[0].text.startswith(onek):
        return
    ilk = p.runs[0]
    if ilk.text != onek:
        kopya = copy.deepcopy(ilk._r)
        ilk._r.addprevious(kopya)
        ilk.text = ilk.text[len(onek):]
        ilk = p.runs[0]
        ilk.text = onek
    ilk.bold = True


def bas_bosluk_temizle(p):
    for r in p.runs:
        if not r.text:
            continue
        r.text = r.text.lstrip(" \u00a0\t")
        if r.text:
            break


def numarali_mi(p):
    pPr = p._p.pPr
    return pPr is not None and pPr.find(qn("w:numPr")) is not None


def satir_sonundan_bol(p):
    """Paragraftaki ilk satır sonunu (<w:br/>) gerçek paragraf ayrımına çevirir."""
    yeni = copy.deepcopy(p._p)
    p._p.addnext(yeni)
    for paragraf, ilk_parca in ((p._p, True), (yeni, False)):
        goruldu = False
        for r in list(paragraf.iter(qn("w:r"))):
            for e in list(r):
                if e.tag == qn("w:rPr"):
                    continue
                if e.tag == qn("w:br") and not goruldu:
                    goruldu = True
                    r.remove(e)
                elif goruldu == ilk_parca:   # ilk parçada br sonrası, ikincide br öncesi silinir
                    r.remove(e)
    return Paragraph(yeni, p._parent)


def run_bicimini_temizle(kok):
    """Yazı tipi, boyut, renk ve vurgu gibi elle verilmiş karakter biçimlerini kaldırır (kalın/italik kalır)."""
    for rPr in kok.iter(qn("w:rPr")):
        for etiket in ("w:rFonts", "w:sz", "w:szCs", "w:color", "w:highlight", "w:shd"):
            for e in rPr.findall(qn(etiket)):
                rPr.remove(e)


def stilleri_ayarla(doc):
    for ad in ("Normal", "Default Paragraph", "List Paragraph"):
        try:
            st = doc.styles[ad]
        except KeyError:
            continue
        st.font.name = YAZI_TIPI
        st.font.size = BOYUT
        st.font.color.rgb = RGBColor(0, 0, 0)
        rFonts = st.element.get_or_add_rPr().get_or_add_rFonts()
        for nitelik in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rFonts.set(qn(nitelik), YAZI_TIPI)
        pf = st.paragraph_format
        pf.space_before, pf.space_after = Pt(0), Pt(6)
        pf.line_spacing = 1.15
        pf.widow_control = True


def alan(paragraf, kod):
    """Altbilgiye PAGE / NUMPAGES gibi bir Word alanı ekler."""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), kod)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    paragraf._p.append(fld)


def altbilgi(doc):
    """Altbilgi: yalnızca ortalı "Sayfa X / Y" (sözleşme no ve paraf alanı v1.7.0'da kaldırıldı)."""
    bolum = doc.sections[0]
    ab = bolum.footer
    ab.is_linked_to_previous = False
    for p in list(ab.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p = ab.paragraphs[0]
    for e in list(p._p):
        if e.tag != qn("w:pPr"):
            p._p.remove(e)
    pPr = p._p.get_or_add_pPr()
    for e in list(pPr):
        if e.tag in (qn("w:pBdr"), qn("w:tabs")):
            pPr.remove(e)
    try:
        stil = doc.styles["Footer"]
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        stil = doc.styles.add_style("Footer", WD_STYLE_TYPE.PARAGRAPH)
    stil.font.size = Pt(8)
    stil.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    p.style = stil
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    p.add_run("Sayfa ")
    alan(p, "PAGE")
    p.add_run(" / ")
    alan(p, "NUMPAGES")
    bolum.footer_distance = Cm(1.0)
    bolum.header_distance = Cm(0.9)


def kurumsal_bicim(doc):
    bolum = doc.sections[0]
    bolum.page_width, bolum.page_height = Cm(21), Cm(29.7)
    bolum.left_margin, bolum.right_margin = Cm(2.5), Cm(2.0)
    bolum.top_margin, bolum.bottom_margin = Cm(2.0), Cm(2.0)
    stilleri_ayarla(doc)
    run_bicimini_temizle(doc.element.body)

    # "10.2. CEZAİ ŞART <satır sonu> metin" gibi paragrafları ikiye böl
    for p in list(doc.paragraphs):
        if NUMARA.match(p.text.strip()) and "\n" in p.text:
            satir_sonundan_bol(p)

    madde = 0
    for i, p in enumerate(list(doc.paragraphs)):
        metin = p.text.strip()
        f = p.paragraph_format
        f.line_spacing = 1.15
        f.space_before, f.space_after = Pt(0), Pt(6)

        if not metin:
            if not p._p.findall(".//" + qn("w:drawing")):
                paragraf_sil(p)
            continue
        if metin.startswith("{{") and metin.endswith("}}"):   # tablo / tutanak yer tutucuları
            continue

        if i == 0:                                              # sözleşme başlığı
            paragraf_yaz(p, [(metin, True)])
            p.runs[0].font.size = Pt(14)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            f.space_after = Pt(18)
            continue

        m = MADDE.match(metin)
        if m:
            madde = int(m.group(1))
            baslik = re.sub(r"\s+", " ", m.group(2)).strip()
            for eski, yeni in BASLIK_DUZELT.items():
                baslik = baslik.replace(re.sub(r"\s+", " ", eski), yeni)
            paragraf_yaz(p, [(f"MADDE {madde}: {baslik}", True)])
            girinti(p, Cm(0), hiza=WD_ALIGN_PARAGRAPH.LEFT)
            f.space_before, f.space_after = Pt(14), Pt(6)
            f.keep_with_next = True
            continue

        e = ETIKET.match(p.text)
        if e:                                                   # taraf bilgileri
            metin_degistir(p, e.group(0), f"{e.group(1)}\t: ")
            f.tab_stops.add_tab_stop(Cm(4.2))
            girinti(p, Cm(4.6), Cm(3.6), WD_ALIGN_PARAGRAPH.LEFT)
            f.space_after = Pt(2)
            if e.group(1) in ("KİRALAYAN", "KİRACI"):
                f.space_before = Pt(8)
            continue

        n = NUMARA.match(p.text.lstrip())
        if n and not numarali_mi(p):                            # 4.1. / 6.1.5. gibi elle numaralı maddeler
            numara = n.group(1)
            on_bosluk = p.text[:len(p.text) - len(p.text.lstrip())]
            metin_degistir(p, on_bosluk + n.group(0), f"{numara}.\t")
            kalin_onek(p, f"{numara}.\t")
            derin = numara.count(".") >= 2
            sol, asili = (KADEME2, KADEME2 - KADEME1) if derin else (KADEME1, KADEME1)
            girinti(p, sol, asili)
            f.tab_stops.add_tab_stop(sol)
            govde = p.text.split("\t", 1)[-1].strip()
            if govde and govde == govde.upper() and len(govde) < 60:   # "6.1. KİRALAYANIN ..." alt başlığı
                for r in p.runs:
                    r.bold = True
                f.keep_with_next = True
                f.space_before = Pt(4)
            continue

        if numarali_mi(p):                                      # Word'ün otomatik numara / madde imleri
            bas_bosluk_temizle(p)
            if madde == 1:                                      # 1.1.
                girinti(p, KADEME1, KADEME1)
                f.tab_stops.add_tab_stop(KADEME1)
            else:                                               # tanımlar ve ekler listesi
                girinti(p, LISTE, LISTE - KADEME1)
                f.tab_stops.add_tab_stop(LISTE)
                f.space_after = Pt(3)
            continue

        if metin.startswith(("- ", "- ")):                # 7.3'teki tireli bentler
            metin_degistir(p, p.text[:p.text.index("-") + 2], "–\t")
            girinti(p, LISTE, LISTE - KADEME1)
            f.tab_stops.add_tab_stop(LISTE)
            f.space_after = Pt(3)
            continue

        bas_bosluk_temizle(p)
        girinti(p, KADEME1)                                     # madde metninin devam paragrafları

    # imza tablosu: tam genişlik, ortalı başlık satırı
    for t in doc.tables:
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, satir in enumerate(t.rows):
            for c in satir.cells:
                for p in c.paragraphs:
                    p.paragraph_format.first_line_indent = Cm(0)
                    p.paragraph_format.left_indent = Cm(0)
                    p.paragraph_format.space_after = Pt(4)
                    if j == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for r in p.runs:
                            r.bold = True
        son = t.rows[-1]
        son.height = Cm(3.0)                                    # imza ve kaşe için yer
        son.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    altbilgi(doc)
