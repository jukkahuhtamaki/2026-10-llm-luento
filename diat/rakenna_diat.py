#!/usr/bin/env python3
# Rakentaa demo-osion diat: python rakenna_diat.py  ->  demot.pptx

import copy

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.dml import MSO_THEME_COLOR as T
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

TEEMA = {
    "dk1": "1F2A2E",   # muste
    "lt1": "FFFFFF",
    "dk2": "4A5A60",   # harmaa teksti
    "lt2": "EEF1EF",   # vaalea pohja
    "accent1": "C8102E",  # puolukka
    "accent2": "4F7A5A",  # sammal
    "accent3": "D9971E",  # meripihka
    "accent4": "1F2A2E",
    "accent5": "C9D8CD",  # vaalea sammal
    "accent6": "F6DCE0",  # vaalea puolukka
    "hlink": "C8102E",
    "folHlink": "4A5A60",
}
OTSIKKOFONTTI = "Cambria"
LEIPAFONTTI = "Calibri"
KOODIFONTTI = "Courier New"

L, K = 13.333, 7.5
REUNA = 0.6

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(L), Inches(K)


# --- teema ja pohjat -------------------------------------------------------

def aseta_teema():
    osa = prs.slide_master.part.part_related_by(RT.THEME)
    puu = etree.fromstring(osa.blob)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    puu.find(f".//{a}clrScheme").set("name", "Puolukka")
    for nimi, arvo in TEEMA.items():
        el = puu.find(f".//{a}clrScheme/{a}{nimi}")
        for lapsi in list(el):
            el.remove(lapsi)
        etree.SubElement(el, f"{a}srgbClr", val=arvo)
    fontit = puu.find(f".//{a}fontScheme")
    fontit.set("name", "Luento")
    fontit.find(f"{a}majorFont/{a}latin").set("typeface", OTSIKKOFONTTI)
    fontit.find(f"{a}minorFont/{a}latin").set("typeface", LEIPAFONTTI)
    osa._blob = etree.tostring(puu, xml_declaration=True, encoding="UTF-8", standalone=True)


def aseta_pohjat():
    master = prs.slide_master
    tyyli = master.element.find(qn("p:txStyles")).find(qn("p:titleStyle"))
    lvl = tyyli.find(qn("a:lvl1pPr"))
    lvl.set("algn", "l")
    rpr = lvl.find(qn("a:defRPr"))
    rpr.set("sz", "3600")
    rpr.set("b", "1")
    for vanha in rpr.findall(qn("a:solidFill")):
        rpr.remove(vanha)
    fill = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:schemeClr"), val="tx1")
    rpr.insert(0, fill)

    paikka = (Inches(REUNA), Inches(0.4), Inches(L - 2 * REUNA), Inches(1.0))
    for ph in master.placeholders:
        if ph.placeholder_format.type == 1:  # otsikko
            ph.left, ph.top, ph.width, ph.height = paikka
            ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            if ph.placeholder_format.type == 1:
                ph.left, ph.top, ph.width, ph.height = paikka


aseta_teema()
aseta_pohjat()
OTSIKKO = prs.slide_layouts[5]   # Title Only
TYHJA = prs.slide_layouts[6]     # Blank


# --- apufunktiot ----------------------------------------------------------

def vari(kohde, vari):
    if isinstance(vari, str):
        kohde.rgb = RGBColor.from_string(vari)
    else:
        kohde.theme_color = vari


def teksti(dia, x, y, w, h, kappaleet, koko=16, vari_=T.TEXT_1, lihava=False,
           fontti=None, tasaus=PP_ALIGN.LEFT, ankkuri=MSO_ANCHOR.TOP,
           nimi=None, rivivali=None, kappalevali=0):
    """kappaleet: merkkijono, tai lista kappaleita; kappale on merkkijono
    tai lista (teksti, {asetukset}) -pareja."""
    laatikko = dia.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if nimi:
        laatikko.name = nimi
    tf = laatikko.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = ankkuri
    if isinstance(kappaleet, str):
        kappaleet = [kappaleet]
    for i, kpl in enumerate(kappaleet):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = tasaus
        if rivivali:
            p.line_spacing = rivivali
        if kappalevali and i < len(kappaleet) - 1:
            p.space_after = Pt(kappalevali)
        if isinstance(kpl, str):
            kpl = [(kpl, {})]
        for sisalto, asetukset in kpl:
            r = p.add_run()
            r.text = sisalto
            f = r.font
            f.size = Pt(asetukset.get("koko", koko))
            f.bold = asetukset.get("lihava", lihava)
            f.italic = asetukset.get("kursiivi", False)
            if asetukset.get("fontti", fontti):
                f.name = asetukset.get("fontti", fontti)
            vari(f.color, asetukset.get("vari", vari_))
    return laatikko


def muoto(dia, tyyppi, x, y, w, h, tayte=None, viiva=None, nimi=None, pyoristys=None):
    s = dia.shapes.add_shape(tyyppi, Inches(x), Inches(y), Inches(w), Inches(h))
    if nimi:
        s.name = nimi
    if tayte is None:
        s.fill.background()
    else:
        s.fill.solid()
        vari(s.fill.fore_color, tayte)
    if viiva is None:
        s.line.fill.background()
    else:
        vari(s.line.color, viiva)
        s.line.width = Pt(1.5)
    s.shadow.inherit = False
    if pyoristys is not None and tyyppi == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = pyoristys
    return s


def kortti(dia, x, y, w, h, tayte=T.BACKGROUND_2, nimi=None):
    return muoto(dia, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, tayte, nimi=nimi, pyoristys=0.08)


def pilleri(dia, x, y, w, h, sisalto, tayte=T.ACCENT_1, tekstivari=T.BACKGROUND_1,
            koko=16, lihava=True, fontti=None, viiva=None, nimi=None):
    s = muoto(dia, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, tayte, viiva, nimi, pyoristys=0.5)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.1)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = sisalto
    r.font.size = Pt(koko)
    r.font.bold = lihava
    if fontti:
        r.font.name = fontti
    vari(r.font.color, tekstivari)
    return s


def nuoli(dia, x1, y1, x2, y2, vari_=T.TEXT_2, paksuus=2.25, nimi=None):
    c = dia.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                                 Inches(x2), Inches(y2))
    if nimi:
        c.name = nimi
    vari(c.line.color, vari_)
    c.line.width = Pt(paksuus)
    ln = c.line._get_or_add_ln()
    etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return c


def otsikkodia(otsikko, osio=None):
    dia = prs.slides.add_slide(OTSIKKO)
    dia.shapes.title.text = otsikko
    return dia


def tumma(dia):
    dia.background.fill.solid()
    dia.background.fill.fore_color.theme_color = T.TEXT_1


def muistiinpanot(dia, teksti_):
    dia.notes_slide.notes_text_frame.text = teksti_


def pylvaat(dia, x, y, w, h, otsikko, sanat, arvot, korostus=0, nimi=None):
    data = CategoryChartData()
    data.categories = sanat
    data.add_series("Todennäköisyys", arvot)
    gf = dia.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y),
                              Inches(w), Inches(h), data)
    if nimi:
        gf.name = nimi
    kaavio = gf.chart
    kaavio.has_legend = False
    kaavio.has_title = True
    kaavio.chart_title.text_frame.text = otsikko
    otsikkofontti = kaavio.chart_title.text_frame.paragraphs[0].runs[0].font
    otsikkofontti.size = Pt(16)
    otsikkofontti.bold = True
    otsikkofontti.name = "+mn-lt"
    vari(otsikkofontti.color, T.TEXT_1)
    kaavio.font.size = Pt(14)
    kaavio.font.name = "+mn-lt"
    vari(kaavio.font.color, T.TEXT_2)

    kat = kaavio.category_axis
    kat.reverse_order = True
    kat.format.line.fill.background()
    kat.has_major_gridlines = False
    kat.tick_labels.font.size = Pt(16)
    kat.tick_labels.font.bold = True
    vari(kat.tick_labels.font.color, T.TEXT_1)
    arv = kaavio.value_axis
    arv.visible = False
    arv.has_major_gridlines = False
    arv.maximum_scale = 1.0
    arv.minimum_scale = 0

    plot = kaavio.plots[0]
    plot.gap_width = 45
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0 %'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(14)
    vari(dl.font.color, T.TEXT_2)
    sarja = plot.series[0]
    for i in range(len(arvot)):
        piste = sarja.points[i]
        piste.format.fill.solid()
        vari(piste.format.fill.fore_color, T.ACCENT_1 if i == korostus else T.ACCENT_5)
    return gf


def lampotilassa(todnak, t):
    painot = [p ** (1 / t) for p in todnak]
    summa = sum(painot) + (1 - sum(todnak)) ** (1 / t)  # loput sanakkeet
    return [round(p / summa, 2) for p in painot]


# --- diat -----------------------------------------------------------------

# 1. Kansi
dia = prs.slides.add_slide(TYHJA)
tumma(dia)
teksti(dia, REUNA, 2.1, 9, 1.4, "Kielimallista agentiksi", koko=54, lihava=True,
       vari_=T.BACKGROUND_1, fontti="+mj-lt", nimi="Otsikko")
teksti(dia, REUNA, 3.5, 9, 0.6, "Demot: sokkotesti ja pieni agentti", koko=24,
       vari_=T.ACCENT_5, nimi="Alaotsikko")
x = REUNA
for sana, tayte, tv in [("Moro", T.ACCENT_1, T.BACKGROUND_1),
                        (" …", T.TEXT_2, T.BACKGROUND_1),
                        ("?", T.ACCENT_3, T.TEXT_1)]:
    lev = 0.4 + 0.22 * len(sana.strip())
    pilleri(dia, x, 4.65, lev, 0.6, sana.strip(), tayte, tv, koko=22, nimi=f"Sanake {sana.strip()}")
    x += lev + 0.15
muistiinpanot(dia, "Demo-osio: ensin sokkotesti kolmella kielimallilla, sitten rakennetaan pieni agentti.")

# 2. Sokkotesti
dia = otsikkodia("Sokkotesti: kuka vastasi?")
kortti(dia, REUNA, 1.9, 3.0, 2.2, T.TEXT_1, nimi="Kysymys")
teksti(dia, REUNA + 0.3, 2.15, 2.4, 0.4, "KYSYMYS", koko=14, lihava=True, vari_=T.ACCENT_5)
teksti(dia, REUNA + 0.3, 2.6, 2.4, 1.3, "Sama kysymys kaikille vastaajille", koko=20,
       lihava=True, vari_=T.BACKGROUND_1)
for i, kirjain in enumerate("ABC"):
    y = 1.9 + i * 0.8
    nuoli(dia, REUNA + 3.1, 3.0, 5.0, y + 0.3, nimi=f"Nuoli {kirjain}")
    pilleri(dia, 5.1, y, 2.6, 0.6, f"Vastaaja {kirjain}", T.BACKGROUND_2, T.TEXT_1,
            koko=18, nimi=f"Vastaaja {kirjain}")
    nuoli(dia, 7.8, y + 0.3, 8.6, 3.0, nimi=f"Nuoli arvaus {kirjain}")
pilleri(dia, 8.7, 2.65, 3.9, 0.7, "Yleisö arvaa", T.ACCENT_1, koko=20, nimi="Yleisö arvaa")
teksti(dia, REUNA, 4.75, 6, 0.4, "VASTAAJAT, SATUNNAISESSA JÄRJESTYKSESSÄ", koko=14,
       lihava=True, vari_=T.TEXT_2)
vastaajat = [("TinyLlama", "1,1 mrd parametria", "toimii läppärillä"),
             ("gpt-oss", "21 mrd parametria", "toimii läppärillä"),
             ("Claude", "keskusteleva palvelu", "claude.ai selaimessa")]
for i, (nimi, koko_, missa) in enumerate(vastaajat):
    x = REUNA + i * 4.1
    kortti(dia, x, 5.25, 3.8, 1.55, nimi=f"Kortti {nimi}")
    teksti(dia, x + 0.3, 5.45, 3.2, 0.5, nimi, koko=22, lihava=True, fontti="+mj-lt")
    teksti(dia, x + 0.3, 5.95, 3.2, 0.7, [koko_, missa], koko=16, vari_=T.TEXT_2)
muistiinpanot(dia,
    "Paikalliset mallit ajetaan komentoriviltä: python3 demot/turing.py\n"
    "Ohjelma näyttää vastaukset järjestyksessä A, B (ja C), ja Enter paljastaa kuka oli kuka.\n"
    "Claudelta kysytään sama kysymys rinnalla claude.ai-palvelussa.")

# 3. Kolme kysymystä
dia = otsikkodia("Kolme kysymystä")
kysymykset = [
    ("1", "Täydennä lause: Moro\u00a0…", "Kieli ja kulttuuri", "Tunnistaako malli suomalaisen tervehdyksen ja ohjelmoijien vitsin?"),
    ("2", "Sää Härmälässä viimeisen 24 tunnin aikana JSON-muodossa", "Ajantasainen tieto",
     "Malli ei näe tämän päivän säätä. Myöntääkö se sen vai keksiikö?"),
    ("3", "Määrittele GenAI:n ja kielimallin suhde.", "Käsitteet",
     "Osaako malli selittää, mitä se itse on?"),
]
for i, (nro, kys, mita, miksi) in enumerate(kysymykset):
    x = REUNA + i * 4.1
    kortti(dia, x, 1.9, 3.8, 4.9, nimi=f"Kysymys {nro}")
    muoto(dia, MSO_SHAPE.OVAL, x + 0.3, 2.2, 0.7, 0.7, T.ACCENT_1, nimi=f"Numero {nro}")
    teksti(dia, x + 0.3, 2.2, 0.7, 0.7, nro, koko=24, lihava=True, vari_=T.BACKGROUND_1,
           tasaus=PP_ALIGN.CENTER, ankkuri=MSO_ANCHOR.MIDDLE, fontti="+mj-lt")
    teksti(dia, x + 0.3, 3.2, 3.2, 1.6, f"”{kys}”", koko=20, lihava=True, fontti="+mj-lt")
    pilleri(dia, x + 0.3, 4.95, 2.6, 0.45, mita, T.ACCENT_2, koko=14)
    teksti(dia, x + 0.3, 5.6, 3.2, 1.1, miksi, koko=16, vari_=T.TEXT_2)
muistiinpanot(dia,
    "1: ”Moro maailma!” on suomenkielinen ”Hello world”, eli ohjelmoijien ensimmäinen ohjelma.\n"
    "2: Härmälässä on Ilmatieteen laitoksen havaintoasema. Ilman työkaluja malli ei voi tietää säätä.\n"
    "3: Vastaus tähän tulee kalvolla ”Generatiivinen tekoäly ja kielimalli”.")

# 4. Seuraava sanake
dia = otsikkodia("Kielimalli arvaa seuraavan sanakkeen")
teksti(dia, REUNA, 1.75, 5.2, 0.4, "SYÖTE", koko=14, lihava=True, vari_=T.TEXT_2)
x = REUNA
for sana in ["Moro", ","]:
    lev = 0.4 + 0.2 * len(sana)
    pilleri(dia, x, 2.25, lev, 0.6, sana, T.TEXT_1, koko=20, nimi=f"Syöte {sana}")
    x += lev + 0.15
pilleri(dia, x, 2.25, 0.8, 0.6, "?", T.BACKGROUND_1, T.ACCENT_1, koko=20,
        viiva=T.ACCENT_1, nimi="Seuraava sanake")
teksti(dia, REUNA, 3.3, 5.2, 2.6, [
    "Kielimalli laskee jokaiselle mahdolliselle sanakkeelle todennäköisyyden.",
    "Sitten se arpoo yhden, lisää sen tekstiin ja toistaa.",
    [("Sanake", {"lihava": True}), (" on sana tai sanan osa, esim. ”Moro”, ”maail” tai ”,”.", {})],
], koko=18, kappalevali=12)
SANAT = ["miten", "mitä", "maailma", "kaikki", "!"]
TODNAK = [0.40, 0.22, 0.12, 0.08, 0.05]
pylvaat(dia, 6.6, 1.65, 6.1, 4.6, "Seuraavan sanakkeen todennäköisyys", SANAT, TODNAK,
        nimi="Kaavio todennäköisyydet")
teksti(dia, 6.6, 6.35, 6.1, 0.4, "Kuvitteellinen esimerkki: luvut eivät ole mitattuja",
       koko=12, vari_=T.TEXT_2, tasaus=PP_ALIGN.RIGHT, nimi="Huomautus")
muistiinpanot(dia,
    "Luvut ovat havainnollistavia, eivät minkään mallin todellisia.\n"
    "Tämä on koko kielimallin ydin: seuraavan sanakkeen ennustaminen, uudelleen ja uudelleen.")

# 5. Lämpötila
dia = otsikkodia("Lämpötila säätää arpajaisia")
for i, (t, kuvaus) in enumerate([(0.3, "Matala lämpötila 0,3: varma ja toisteinen"),
                                 (1.8, "Korkea lämpötila 1,8: yllätyksellinen ja sekava")]):
    x = REUNA + i * 6.2
    pylvaat(dia, x, 1.65, 5.9, 4.0, kuvaus, SANAT, lampotilassa(TODNAK, t),
            nimi=f"Kaavio lämpötila {t}")
    pilleri(dia, x + 0.3, 5.9, 4.6, 0.6, f"python3 agentti.py {t}", T.TEXT_1, koko=16,
            lihava=False, fontti=KOODIFONTTI, nimi=f"Komento {t}")
teksti(dia, REUNA, 6.75, 12, 0.4, "Sama kuvitteellinen esimerkki kuin edellisellä kalvolla, eri lämpötiloissa",
       koko=12, vari_=T.TEXT_2, tasaus=PP_ALIGN.RIGHT)
muistiinpanot(dia,
    "Lämpötila ei muuta mallia, vaan sitä, miten todennäköisyyksistä arvotaan.\n"
    "Matalalla lämpötilalla todennäköisin sanake voittaa lähes aina, korkealla epätodennäköisetkin pääsevät mukaan.\n"
    "Demossa sama kysymys kolme kertaa lämpötilalla 0.0 ja kolme kertaa lämpötilalla 1.8.")

# 6. Kolme kerrosta
dia = otsikkodia("Sama malli, kolme kerrosta")
kerrokset = [
    (REUNA, 1.75, 7.2, 5.15, T.BACKGROUND_2, T.TEXT_1, "Keskusteleva palvelu",
     "esim. claude.ai, ChatGPT, Copilot"),
    (REUNA + 0.4, 2.95, 6.4, 3.6, T.ACCENT_5, T.TEXT_1, "Keskustelumalli",
     "opetettu vastaamaan avustajana"),
    (REUNA + 0.8, 4.15, 5.6, 2.05, T.TEXT_1, T.BACKGROUND_1, "Kielimalli",
     "jatkaa tekstiä sanake kerrallaan"),
]
for x, y, w, h, tayte, tv, nimi, kuvaus in kerrokset:
    kortti(dia, x, y, w, h, tayte, nimi=f"Kerros {nimi}")
    teksti(dia, x + 0.3, y + 0.2, w - 0.6, 0.45, nimi, koko=20, lihava=True, vari_=tv,
           fontti="+mj-lt")
    teksti(dia, x + 0.3, y + 0.65, w - 0.6, 0.4, kuvaus, koko=16,
           vari_=T.ACCENT_5 if tv == T.BACKGROUND_1 else T.TEXT_2)
lisat = [
    ("+ järjestelmäkehote, muisti,", "työkalut, käyttöliittymä"),
    ("+ keskustelupohja,", "avustajan rooli"),
    ("painot: miljardeja", "lukuja"),
]
for (rivi1, rivi2), (x, y, w, h, *_ ) in zip(lisat, kerrokset):
    nuoli(dia, 8.9, y + 0.55, x + w + 0.1, y + 0.55, T.TEXT_2, nimi=f"Nuoli {rivi1}")
    teksti(dia, 9.0, y + 0.3, 3.7, 0.8, [rivi1, rivi2], koko=16, ankkuri=MSO_ANCHOR.MIDDLE)
muistiinpanot(dia,
    "Sama kielimalli on kaikkien kerrosten ytimessä. Ero on siinä, mitä sen ympärille on rakennettu.\n"
    "turing.py --raw näyttää pelkän kielimallin, turing.py keskustelumallin, ja claude.ai keskustelevan palvelun.")

# 7. Mitä mallit vastasivat
dia = otsikkodia("”Täydennä lause: Moro …”")
sarakkeet = [
    ("Kielimalli", "turing.py --raw", "gpt-oss",
     "”We have a puzzle. The question: ”Täydennä lause: Moro ...”. That is Finnish … So we need to fill the rest”"),
    ("Keskustelumalli", "turing.py", "gpt-oss",
     "”Tässä muutama yleinen tapa jatkaa lausetta: 1. Moro! Miten menee? 2. Moro! Mitä kuuluu? …”"),
    ("Keskusteleva palvelu", "claude.ai", "Claude",
     "Kokeillaan luennolla"),
]
for i, (kerros, komento, malli, vastaus) in enumerate(sarakkeet):
    x = REUNA + i * 4.1
    tayte = [T.TEXT_1, T.ACCENT_5, T.BACKGROUND_2][i]
    tv = T.BACKGROUND_1 if i == 0 else T.TEXT_1
    kortti(dia, x, 1.75, 3.8, 1.1, tayte, nimi=f"Otsake {kerros}")
    teksti(dia, x + 0.3, 1.9, 3.2, 0.45, kerros, koko=20, lihava=True, vari_=tv, fontti="+mj-lt")
    teksti(dia, x + 0.3, 2.35, 3.2, 0.4, komento, koko=14, fontti=KOODIFONTTI,
           vari_=T.ACCENT_5 if i == 0 else T.TEXT_2)
    teksti(dia, x, 3.15, 3.8, 0.4, malli.upper(), koko=14, lihava=True, vari_=T.TEXT_2)
    teksti(dia, x, 3.6, 3.8, 3.0, [[(vastaus, {"kursiivi": i == 2})]], koko=18,
           vari_=T.TEXT_2 if i == 2 else T.TEXT_1, nimi=f"Vastaus {kerros}")
muistiinpanot(dia,
    "Vastaukset on saatu omalla koneella 31.10.2026. Lämpötila vaikuttaa, joten livenä tulos voi olla toinen.\n"
    "Pelkkä kielimalli (raw) ei vastaa kysymykseen vaan jatkaa tekstiä: gpt-oss alkoi ”ajatella ääneen” englanniksi.\n"
    "TinyLlama raw-tilassa ei jatkanut mitään, keskustelutilassa se vastasi ”Selain lause: Moro, Moro ...”.\n"
    "Pelkällä ”Moro ...” -syötteellä gpt-oss vastasi romaniaksi: ”Bună! Cu ce te pot ajuta astăzi?”")

# 8. GenAI ja kielimalli
dia = otsikkodia("Generatiivinen tekoäly ja kielimalli")
sisakkaiset = [
    (REUNA, 1.75, 8.0, 5.1, T.BACKGROUND_2, "Tekoäly"),
    (REUNA + 0.35, 2.4, 7.3, 4.1, T.ACCENT_5, "Koneoppiminen"),
    (REUNA + 0.7, 3.05, 6.6, 3.1, T.TEXT_1, "Generatiivinen tekoäly (GenAI)"),
]
for x, y, w, h, tayte, nimi in sisakkaiset:
    kortti(dia, x, y, w, h, tayte, nimi=f"Joukko {nimi}")
    teksti(dia, x + 0.3, y + 0.15, w - 0.6, 0.45, nimi, koko=18, lihava=True,
           vari_=T.BACKGROUND_1 if tayte == T.TEXT_1 else T.TEXT_1, fontti="+mj-lt")
mallit = [("Kielimallit", T.ACCENT_1, T.BACKGROUND_1), ("Kuvamallit", T.TEXT_2, T.BACKGROUND_1),
          ("Ääni ja video", T.TEXT_2, T.BACKGROUND_1)]
for i, (nimi, tayte, tv) in enumerate(mallit):
    pilleri(dia, REUNA + 1.0 + i * 2.05, 3.9, 1.9, 0.6, nimi, tayte, tv, koko=16, nimi=nimi)
teksti(dia, REUNA + 1.0, 4.85, 6.0, 1.1,
       "Kaikki tuottavat uutta sisältöä oppimansa aineiston pohjalta.",
       koko=16, vari_=T.ACCENT_5)
teksti(dia, 9.1, 1.75, 3.6, 0.4, "SUHDE", koko=14, lihava=True, vari_=T.TEXT_2)
teksti(dia, 9.1, 2.2, 3.6, 4.6, [
    [("Kielimalli on yksi GenAI:n muoto", {"lihava": True}), (": se tuottaa tekstiä.", {})],
    [("GenAI on laajempi", {"lihava": True}), (": myös kuvat, ääni, video ja koodi.", {})],
    [("Keskusteleva palvelu", {"lihava": True}),
     (" on kielimalli, jonka ympärille on rakennettu kehote, muisti, työkalut ja käyttöliittymä.", {})],
], koko=18, kappalevali=14)
muistiinpanot(dia, "Tämä on vastaus sokkotestin kolmanteen kysymykseen. Verrataan mallien vastauksiin.")

# 9. Pieni agentti: koodi
dia = otsikkodia("Pieni agentti: 30 riviä Pythonia")
KOODI = [
    'MALLI = "gpt-oss:20b"',
    'LAMPOTILA = float(sys.argv[1]) if len(sys.argv) > 1 else 0.7',
    'JARJESTELMAKEHOTE = "Olet tamperelainen avustaja.',
    '                     Vastaat lyhyesti ja aina suomeksi."',
    '',
    'viestit = [{"role": "system", "content": JARJESTELMAKEHOTE}]',
    '',
    'while True:',
    '    viestit.append({"role": "user", "content": input("Sinä: ")})',
    '    data = json.dumps({',
    '        "model": MALLI,',
    '        "messages": viestit,',
    '        "options": {"temperature": LAMPOTILA},',
    '    }).encode()',
    '    with urllib.request.urlopen(OLLAMA, data) as r:',
    '        vastaus = json.load(r)["message"]',
    '    viestit.append(vastaus)',
    '    print("Agentti:", vastaus["content"])',
]
KOODI_Y, RIVI = 1.95, 0.255
kortti(dia, REUNA, 1.7, 8.55, 5.25, T.TEXT_1, nimi="Koodi tausta")
teksti(dia, REUNA + 0.3, KOODI_Y, 8.0, 4.8, [[(r or " ", {})] for r in KOODI], koko=13,
       vari_=T.BACKGROUND_1, fontti=KOODIFONTTI, rivivali=Pt(18.4), nimi="Koodi")
selitteet = [
    (1, "Lämpötila", "annetaan komentoriviltä", T.ACCENT_3, T.TEXT_1),
    (2, "Järjestelmäkehote", "ohjaa koko keskustelua", T.ACCENT_1, T.BACKGROUND_1),
    (5, "Muisti", "on pelkkä lista viestejä", T.ACCENT_2, T.BACKGROUND_1),
    (16, "Vastaus listaan", "joten malli muistaa sen", T.ACCENT_2, T.BACKGROUND_1),
]
for i, (rivi, nimi, kuvaus, tayte, tv) in enumerate(selitteet):
    y = 1.75 + i * 1.3
    pilleri(dia, 9.75, y, 2.95, 0.55, nimi, tayte, tv, koko=16, nimi=f"Selite {nimi}")
    teksti(dia, 9.75, y + 0.62, 2.95, 0.45, kuvaus, koko=14, vari_=T.TEXT_2)
    nuoli(dia, 9.7, y + 0.28, 9.0, KOODI_Y + rivi * RIVI + 0.13, tayte, 2, nimi=f"Nuoli {nimi}")
muistiinpanot(dia,
    "Koko tiedosto: demot/agentti.py. Kalvolta on jätetty pois tuonnit, virheenkäsittely ja \"stream\": False.\n"
    "Käynnistys: python3 demot/agentti.py 0.7\n"
    "Kokeile: ”Täydennä lause: Moro …” ja heti perään ”Mitä sanoit äsken?”, mikä näyttää, että muisti toimii.")

# 10. Agentin silmukka
dia = otsikkodia("Silmukka, joka muistaa")
solmut = [
    (1.2, 1.9, "Sinä kirjoitat", T.BACKGROUND_2, T.TEXT_1),
    (8.1, 1.9, "Viesti lisätään listaan", T.BACKGROUND_2, T.TEXT_1),
    (8.1, 5.15, "Kielimalli lukee koko listan", T.TEXT_1, T.BACKGROUND_1),
    (1.2, 5.15, "Vastaus lisätään listaan", T.BACKGROUND_2, T.TEXT_1),
]
LEV, KOR = 4.0, 1.1
for x, y, sisalto, tayte, tv in solmut:
    s = kortti(dia, x, y, LEV, KOR, tayte, nimi=f"Vaihe {sisalto}")
    teksti(dia, x, y, LEV, KOR, sisalto, koko=20, lihava=True, vari_=tv,
           tasaus=PP_ALIGN.CENTER, ankkuri=MSO_ANCHOR.MIDDLE)
nuoli(dia, 1.2 + LEV + 0.1, 2.45, 8.0, 2.45, nimi="Nuoli 1")
nuoli(dia, 10.1, 1.9 + KOR + 0.1, 10.1, 5.05, nimi="Nuoli 2")
nuoli(dia, 8.0, 5.7, 1.2 + LEV + 0.1, 5.7, nimi="Nuoli 3")
nuoli(dia, 3.2, 5.05, 3.2, 1.9 + KOR + 0.1, nimi="Nuoli 4")
kortti(dia, 4.2, 3.35, 4.95, 1.4, T.BACKGROUND_1, nimi="Keskiö")
teksti(dia, 4.4, 3.35, 4.55, 1.4, [
    [("viestit", {"fontti": KOODIFONTTI, "lihava": True, "vari": T.ACCENT_1}),
     (" = koko muisti", {"lihava": True})],
    "Ensimmäisenä aina järjestelmäkehote",
], koko=18, tasaus=PP_ALIGN.CENTER, ankkuri=MSO_ANCHOR.MIDDLE, kappalevali=6)
muistiinpanot(dia,
    "Kielimalli itsessään ei muista mitään. Joka kierroksella sille lähetetään koko keskustelu alusta asti.\n"
    "Siksi pitkät keskustelut hidastuvat ja maksavat enemmän: lista kasvaa.")

# 11. Kokeillaan
dia = otsikkodia("Kokeillaan: kehote ja lämpötila")
kokeet = [
    ("Vaihda järjestelmäkehote", T.ACCENT_1, [
        "”Olet tamperelainen avustaja. Vastaat lyhyesti ja aina suomeksi.”",
        "→ ”Vastaat aina runomuodossa.”",
        "→ ”Olet kärsimätön asiakaspalvelija.”",
    ], "Sama malli, sama kysymys, eri persoona."),
    ("Vaihda lämpötilaa", T.ACCENT_3, [
        "python3 agentti.py 0.0",
        "python3 agentti.py 1.8",
        "Kysy sama asia kolme kertaa kummallakin.",
    ], "Matalalla vastaukset toistuvat, korkealla ne hajoavat."),
]
for i, (otsake, korostus, rivit, havainto) in enumerate(kokeet):
    x = REUNA + i * 6.2
    kortti(dia, x, 1.75, 5.9, 5.0, nimi=f"Koe {otsake}")
    muoto(dia, MSO_SHAPE.OVAL, x + 0.3, 2.0, 0.6, 0.6, korostus, nimi=f"Merkki {otsake}")
    teksti(dia, x + 1.1, 2.0, 4.5, 0.6, otsake, koko=22, lihava=True, fontti="+mj-lt",
           ankkuri=MSO_ANCHOR.MIDDLE)
    koodia = i == 1
    teksti(dia, x + 0.3, 2.95, 5.3, 2.3,
           [[(r, {"fontti": KOODIFONTTI if koodia and r.startswith("python") else None})] for r in rivit],
           koko=18 if not koodia else 17, kappalevali=10)
    teksti(dia, x + 0.3, 5.6, 5.3, 0.9, havainto, koko=18, lihava=True, vari_=T.TEXT_2)
muistiinpanot(dia,
    "Järjestelmäkehote muutetaan suoraan tiedostoon demot/agentti.py (JARJESTELMAKEHOTE).\n"
    "Huom: ”Täydennä lause: Moro …” -kysymykseen gpt-oss vastasi suomeksi myös ilman järjestelmäkehotetta. "
    "Pelkkään ”Moro ...” -syötteeseen se vastasi romaniaksi. Kysymyksen muotoilu on siis osa kehotetta.")

# 12. Seuraava askel
dia = otsikkodia("Seuraava askel: agentti saa työkalun")
pilleri(dia, REUNA, 2.0, 5.4, 0.8, "”Sää Härmälässä viimeisen 24 tunnin aikana”",
        T.BACKGROUND_2, T.TEXT_1, koko=16, nimi="Kysymys sää")
kortti(dia, REUNA, 3.35, 3.4, 1.9, T.TEXT_1, nimi="Agentti")
teksti(dia, REUNA, 3.35, 3.4, 1.9, [[("Agentti", {"lihava": True, "koko": 22})],
                                     "kielimalli + silmukka"],
       koko=16, vari_=T.BACKGROUND_1, tasaus=PP_ALIGN.CENTER, ankkuri=MSO_ANCHOR.MIDDLE)
nuoli(dia, REUNA + 1.7, 2.9, REUNA + 1.7, 3.25, nimi="Nuoli kysymys")
kortti(dia, 6.4, 3.35, 3.4, 1.9, T.ACCENT_2, nimi="Työkalu")
teksti(dia, 6.4, 3.35, 3.4, 1.9, [[("Työkalu", {"lihava": True, "koko": 22})],
                                  "hae_saa(paikka)"],
       koko=16, vari_=T.BACKGROUND_1, tasaus=PP_ALIGN.CENTER, ankkuri=MSO_ANCHOR.MIDDLE)
nuoli(dia, REUNA + 3.5, 3.9, 6.3, 3.9, nimi="Nuoli kutsu")
nuoli(dia, 6.3, 4.7, REUNA + 3.5, 4.7, nimi="Nuoli tulos")
teksti(dia, REUNA + 3.55, 3.45, 2.1, 0.4, "pyytää", koko=14, vari_=T.TEXT_2, tasaus=PP_ALIGN.CENTER)
teksti(dia, REUNA + 3.55, 4.8, 2.1, 0.4, "havainnot", koko=14, vari_=T.TEXT_2, tasaus=PP_ALIGN.CENTER)
pilleri(dia, 6.4, 5.85, 3.4, 0.7, "Ilmatieteen laitos, avoin data", T.BACKGROUND_2, T.TEXT_1,
        koko=14, nimi="Lähde")
nuoli(dia, 8.1, 5.35, 8.1, 5.8, nimi="Nuoli lähde")
teksti(dia, 10.3, 2.0, 2.4, 4.6, [
    [("Kielimalli", {"lihava": True}), (" ei tiedä tämän päivän säätä.", {})],
    [("Agentti", {"lihava": True}), (" osaa pyytää työkalua hakemaan sen ja vastaa oikealla datalla.", {})],
], koko=18, kappalevali=16)
muistiinpanot(dia,
    "Demo: python3 demot/agentti_saa.py 0.2 ja kysymys ”Sää Härmälässä viimeisen 24 tunnin aikana JSON-muodossa”.\n"
    "Tämä ratkaisee sokkotestin toisen kysymyksen. Härmälässä on Ilmatieteen laitoksen havaintoasema (Tampere Härmälä).\n"
    "Malli ei itse hae mitään. Se pyytää ohjelmaa ajamaan työkalun, ja ohjelma palauttaa tuloksen viestilistaan.")

# 13. Paljastus (piilotettu)
dia = prs.slides.add_slide(TYHJA)
tumma(dia)
teksti(dia, REUNA, 1.6, 12, 0.5, "PALJASTUS", koko=18, lihava=True, vari_=T.ACCENT_3)
teksti(dia, REUNA, 2.2, 11, 1.8, "Tämän demon koodin kirjoitti tekoäly", koko=48, lihava=True,
       vari_=T.BACKGROUND_1, fontti="+mj-lt", nimi="Otsikko")
for i, (otsake, rivit, tayte) in enumerate([
        ("Ihminen", "idea, kysymykset, sanasto, arviointi", T.ACCENT_5),
        ("Tekoäly", "turing.py, agentti.py ja nämä kalvot", T.ACCENT_1)]):
    x = REUNA + i * 6.0
    pilleri(dia, x, 4.4, 2.2, 0.6, otsake, tayte, T.TEXT_1 if i == 0 else T.BACKGROUND_1,
            koko=18, nimi=f"Tekijä {otsake}")
    teksti(dia, x, 5.15, 5.6, 0.9, rivit, koko=20, vari_=T.BACKGROUND_1)
dia._element.set("show", "0")
muistiinpanot(dia, "Piilotettu kalvo: näytä vasta, kun paljastus tehdään.")


prs.save("demot.pptx")
print("demot.pptx")
