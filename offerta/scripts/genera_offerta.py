#!/usr/bin/env python3
"""
Genera l'offerta KIT FORMASUBITO su carta intestata Overall (assets/carta_intestata.docx).

Struttura del documento:
  pag. 1  lettera: logo FORMASUBITO, data/riferimenti e destinatario, oggetto,
          riepilogo "La vostra proposta in sintesi" (tabella prezzi),
          grafica ridotta facoltativa e prossimi passi
  pag. 2  Condizioni di vendita        (assets/testi/condizioni.txt)
  pag. 3  Informativa privacy + consenso (assets/testi/informativa.txt)

I testi di pag. 2 e 3 si modificano nei .txt, senza toccare lo script:
  "## Titolo"      striscia verde di sezione
  "**grassetto**"  e "*corsivo*" in linea
  "|a|b|c|"        righe di tabella (la prima e' l'intestazione)
  "[FIRMA]" "[FIRMA_CLIENTE]" "[FIRMA_CONSENSO]"  blocchi firma

USO:
  genera_offerta.py --out O --ragione R --indirizzo I --piva P --mail M
                    --spec offerta.json [--data gg/mm/aaaa]

offerta.json:
{
  "kit": ["lavoratori", "preposti", "aggiornamenti"],   # almeno uno
  "mansioni": 4,                     # obbligatorio con "lavoratori"
  "bundle": true,                    # sconto "Tutto e subito" con i 3 KIT (default true)
  "sconti": [{"descrizione": "...", "su": "mansioni_aggiuntive", "percentuale": 100},
             {"descrizione": "...", "importo": 20.00}],
  "grafica": "tu-formi",             # id in GRAFICHE oppure null (nessuna grafica)
  "nota": "testo facoltativo sotto il riepilogo"
}
"su" ammette: lavoratori, mansioni_aggiuntive, preposti, aggiornamenti, totale.
"""

import argparse
import copy
import json
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

ASSETS = Path(__file__).resolve().parent.parent / "assets"

# ============================================================
# CATALOGO (listino, IVA esclusa, canoni annui)
# ============================================================
MANSIONI_INCLUSE = 3
EXTRA_PER_MANSIONE = 19.99
ALIQUOTA_IVA = 22
BUNDLE_SCONTO = 100.00
CATALOGO = {
    "lavoratori": ("KIT Lavoratori – formazione iniziale generale + specifica (fino a 3 mansioni)", 249.00),
    "preposti": ("KIT Preposti – formazione iniziale preposti (12 ore)", 149.00),
    "aggiornamenti": ("KIT Aggiornamenti – lavoratori (6 h ogni 5 anni) e preposti (6 h ogni 2 anni)", 149.00),
}
ORDINE_KIT = ["lavoratori", "preposti", "aggiornamenti"]

# Grafiche per la prima pagina. "prezzi": prezzi stampati nella grafica,
# da confrontare con l'offerta (STEP della SKILL).
GRAFICHE = {
    "infografica": {"file": "infografica.png", "titolo": "Vantaggi + Cosa comprende il KIT", "prezzi": None},
    "tu-formi": {"file": "tu-formi.png", "titolo": "Tu formi. Al resto pensiamo noi (sorgente in grafiche/sorgenti/)", "prezzi": None},
    "cosa-ricevi": {"file": "cosa-ricevi.png", "titolo": "Tutto in una cartella, pronto da usare", "prezzi": None},
    "scegli-kit": {"file": "scegli-kit.png", "titolo": "Scegli il tuo KIT",
                   "prezzi": "Lavoratori 249 €/anno fino a 3 mansioni poi 19 €/mansione; Preposti 149 €; Aggiornamenti 149 €"},
    "tre-kit": {"file": "tre-kit.png", "titolo": "I 3 KIT insieme",
                "prezzi": "547 € → 447 €, risparmi 100 €; 'Valida 10 giorni dalla presentazione di oggi!'"},
}


def formatta_eur(importo: float) -> str:
    """Formato italiano: 249,00 € / 1.249,95 € / -100,00 €"""
    s = f"{importo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} €"


def calcola_righe(spec: dict) -> dict:
    """Dalle scelte del compilatore ricava le righe della tabella e il totale."""
    kit = spec.get("kit") or []
    ignoti = [k for k in kit if k not in CATALOGO]
    if not kit or ignoti:
        raise ValueError(f"'kit' deve contenere almeno uno tra {ORDINE_KIT}; ignoti: {ignoti}")
    kit = [k for k in ORDINE_KIT if k in kit]

    righe, importi = [], {}

    def riga(chiave, voce, qta, unitario, importo):
        righe.append({"chiave": chiave, "voce": voce, "qta": qta, "unitario": unitario, "importo": round(importo, 2)})
        importi[chiave] = round(importo, 2)

    for k in kit:
        voce, prezzo = CATALOGO[k]
        riga(k, voce, "1", formatta_eur(prezzo), prezzo)
        if k == "lavoratori":
            n = int(spec.get("mansioni") or 0)
            if n < 1:
                raise ValueError("Con il KIT Lavoratori serve 'mansioni' >= 1")
            extra = max(0, n - MANSIONI_INCLUSE)
            if extra:
                riga("mansioni_aggiuntive", "KIT Lavoratori – mansioni aggiuntive", str(extra),
                     f"{formatta_eur(EXTRA_PER_MANSIONE)}/cad.", extra * EXTRA_PER_MANSIONE)

    subtotale = sum(importi.values())
    if len(kit) == 3 and spec.get("bundle", True):
        riga("bundle", "Sconto offerta «Tutto e subito» (3 KIT insieme)", "1", "—", -BUNDLE_SCONTO)

    for i, s in enumerate(spec.get("sconti") or []):
        descr = s["descrizione"]
        if "percentuale" in s:
            su = s.get("su", "totale")
            base = subtotale if su == "totale" else importi.get(su)
            if base is None:
                raise ValueError(f"Sconto '{descr}': voce '{su}' non presente nell'offerta")
            pct = float(s["percentuale"])
            riga(f"sconto{i}", descr, "1", f"-{pct:g}%", -base * pct / 100)
        else:
            riga(f"sconto{i}", descr, "1", "—", -abs(float(s["importo"])))

    totale = round(sum(r["importo"] for r in righe), 2)
    if totale < 0:
        raise ValueError("Il totale risulta negativo: controlla gli sconti")
    return {"kit": kit, "righe": righe, "totale": totale}


def _set_cell_shading(cell, color_hex):
    """Imposta lo sfondo di una cella (color_hex senza '#', es. '404040')."""
    tc_pr = cell._tc.get_or_add_tcPr()
    # Rimuovi shading esistente se presente
    existing = tc_pr.find(qn("w:shd"))
    if existing is not None:
        tc_pr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex)
    tc_pr.append(shd)


def _set_cell_borders(cell, size="4", color="BFBFBF"):
    """Imposta bordi sottili grigi su tutti i lati della cella."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is not None:
        tc_pr.remove(tc_borders)
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), size)
        b.set(qn("w:color"), color)
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def _set_cell_border_top_thick(cell, size="12", color="404040"):
    """Bordo superiore spesso (riga totale)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    # Rimuovi top esistente
    existing_top = tc_borders.find(qn("w:top"))
    if existing_top is not None:
        tc_borders.remove(existing_top)
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), size)
    top.set(qn("w:color"), color)
    tc_borders.append(top)




def _svuota_paragrafo_tieni_pPr(paragrafo):
    """Rimuove tutti i run del paragrafo (lascia il pPr / formattazione)."""
    for run in list(paragrafo.runs):
        run._element.getparent().remove(run._element)


def _crea_paragrafo_in_cella(cell, testo, bold=False, size_pt=10, allineamento=None, color_hex=None):
    """
    Sostituisce il contenuto di una cella con un singolo paragrafo formattato.
    Font Gill Sans MT (come il template).
    """
    # Svuota tutti i paragrafi della cella
    for p in cell.paragraphs:
        _svuota_paragrafo_tieni_pPr(p)
    # Rimuovi paragrafi extra (lascia il primo)
    for extra in list(cell.paragraphs[1:]):
        extra._element.getparent().remove(extra._element)

    p = cell.paragraphs[0]
    if allineamento is not None:
        p.alignment = allineamento

    run = p.add_run(testo)
    run.font.name = "Gill Sans MT"
    # Forza anche eastAsia/hAnsi/cs per evitare fallback strani
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), "Gill Sans MT")

    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_hex:
        from docx.shared import RGBColor
        run.font.color.rgb = RGBColor.from_string(color_hex)
    return run


# ============================================================
# TABELLA PREZZI
# ============================================================
COL_HEADER = "0E3A27"   # verde FORMASUBITO
COL_RIGA_ALT = "F6F5EE"  # crema
COL_BORDO = "BFBFBF"


def _cella(cell, testo, align, bold=False, size=9, colore=None, sfondo=None, top_spesso=False):
    if sfondo:
        _set_cell_shading(cell, sfondo)
    _set_cell_borders(cell, size="4", color=COL_BORDO)
    if top_spesso:
        _set_cell_border_top_thick(cell, size="12", color=COL_HEADER)
    _crea_paragrafo_in_cella(cell, testo, bold=bold, size_pt=size, allineamento=align, color_hex=colore)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _tabella_prezzi(doc, info):
    L, C, R = WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT
    righe = info["righe"]

    if len(righe) == 1:  # una sola voce: riga unica con il totale
        tbl = doc.add_table(rows=1, cols=2)
        tbl.autofit = False
        for col, w in zip(tbl.columns, (Cm(11), Cm(5))):
            for cell in col.cells:
                cell.width = w
        r = tbl.rows[0].cells
        _cella(r[0], f"{righe[0]['voce']} – totale annuo (IVA esclusa)", L, bold=True, size=10, top_spesso=True)
        _cella(r[1], formatta_eur(info["totale"]), R, bold=True, size=12, top_spesso=True)
        tbl.style = None
        return tbl

    tbl = doc.add_table(rows=len(righe) + 2, cols=4)
    tbl.autofit = False
    for col, w in zip(tbl.columns, (Cm(8.5), Cm(1.5), Cm(3.0), Cm(3.0))):
        col.width = w
        for cell in col.cells:
            cell.width = w

    for ci, (t, a) in enumerate((("Voce", L), ("Q.tà", C), ("Importo unitario", R), ("Importo", R))):
        _cella(tbl.rows[0].cells[ci], t, a, bold=True, colore="FFFFFF", sfondo=COL_HEADER)

    for ri, r in enumerate(righe, start=1):
        sfondo = COL_RIGA_ALT if ri % 2 == 0 else None
        valori = ((r["voce"], L), (r["qta"], C), (r["unitario"], R), (formatta_eur(r["importo"]), R))
        for ci, (t, a) in enumerate(valori):
            _cella(tbl.rows[ri].cells[ci], t, a, sfondo=sfondo)

    ult = tbl.rows[-1].cells
    etichetta = ult[0].merge(ult[1]).merge(ult[2])
    _cella(etichetta, "Totale annuo (IVA esclusa)", L, bold=True, size=10, top_spesso=True)
    _cella(ult[3], formatta_eur(info["totale"]), R, bold=True, size=12, top_spesso=True)
    tbl.style = None
    return tbl


# ============================================================
# MATTONI DI IMPAGINAZIONE
# ============================================================
VERDE = "0E3A27"
FONT = "Gill Sans MT"
GIALLO, CREMA, ROSSO, GRIGIO = "F5B700", "F6F5EE", "B3261E", "6B6B6B"
PROSSIMI_PASSI = [
    ("Firmate", "l'offerta e rimandatela a formasubito@overallgroup.info"),
    ("Pagamento", "dell'importo di {ivato} ({totale} imponibile + {iva} IVA {aliquota}%) all'IBAN IT71 G083 4059 9600 0000 2176 704 "
                  "(Cassa Padana BCC Società Cooperativa) intestato a OVERALL GROUP SRL, con causale "
                  "\"Attivazione kit formasubito come da offerta del {data}\""),
    ("Consulenza", "in videocall: definiamo insieme il progetto formativo"),
]
DESCRIZIONI = {
    "lavoratori": ("KIT Lavoratori", "Formazione iniziale generale + specifica · fino a 3 mansioni da DVR"),
    "preposti": ("KIT Preposti", "Formazione iniziale preposti"),
    "aggiornamenti": ("KIT Aggiornamenti", "Lavoratori 6 h ogni 5 anni · preposti 6 h ogni 2 anni"),
}


def _run(p, testo, bold=False, italic=False, size=None, colore=None):
    r = p.add_run(testo)
    r.bold, r.italic = bold, italic
    r.font.name = FONT
    if size:
        r.font.size = Pt(size)
    if colore:
        r.font.color.rgb = RGBColor.from_string(colore)
    return r


def _inline(p, testo, size=None):
    """Testo con **grassetto** e *corsivo*."""
    for parte in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", testo):
        if parte.startswith("**"):
            _run(p, parte[2:-2], bold=True, size=size)
        elif parte.startswith("*") and len(parte) > 1:
            _run(p, parte[1:-1], italic=True, size=size)
        elif parte:
            _run(p, parte, size=size)


def _par(container, testo="", size=9, align=WD_ALIGN_PARAGRAPH.JUSTIFY, prima=3, dopo=0):
    p = container.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(prima), Pt(dopo), 1.0
    if testo:
        _inline(p, testo, size=size)
    return p


def _striscia(doc, titolo):
    """Titolo di sezione su striscia verde FORMASUBITO."""
    p = _par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, prima=10, dopo=6)
    p.paragraph_format.keep_with_next = True
    ppr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), VERDE)
    ppr.append(shd)
    r = _run(p, titolo.upper(), bold=True, size=9, colore="FFFFFF")
    sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), "20")
    r._r.get_or_add_rPr().append(sp)
    return p


def _interruzione_pagina(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _tabella_senza_bordi(doc, righe, colonne, larghezze):
    t = doc.add_table(rows=righe, cols=colonne)
    t.autofit = False
    t.style = None
    for col, w in zip(t.columns, larghezze):
        col.width = w
        for c in col.cells:
            c.width = w
    return t


def _cella_testo(cell, righe, size=9, align=WD_ALIGN_PARAGRAPH.LEFT):
    """righe: lista di stringhe con markup in linea; la prima sostituisce il paragrafo vuoto."""
    cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
    for testo in righe:
        _par(cell, testo, size=size, align=align, prima=0)


def _immagine(container, path, larghezza_cm):
    ps = getattr(container, "paragraphs", [])
    vuoto = ps and not ps[-1].runs and hasattr(container, "_tc")
    p = ps[-1] if vuoto else container.add_paragraph()
    p.add_run().add_picture(str(path), width=Cm(larghezza_cm))
    return p


def _blocco_firme(doc, sinistra, destra=None):
    """Una riga: firma del Cliente a sinistra, timbro e firma Overall a destra."""
    t = _tabella_senza_bordi(doc, 1, 2, (Cm(8.5), Cm(8.5)))
    _cella_testo(t.cell(0, 0), [sinistra, "", "________________________________"], size=8)
    if destra:
        _cella_testo(t.cell(0, 1), [destra], size=8)
        _immagine(t.cell(0, 1), ASSETS / "firma_overall.jpeg", 2.4)
    t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    return t


# ============================================================
# PAGINA 1 — LETTERA E RIEPILOGO (impaginazione B)
# ============================================================
def _margini(cell, alto=80, basso=80, sx=140, dx=140):
    """Margini interni della cella in twip."""
    tcpr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for lato, v in (("top", alto), ("bottom", basso), ("left", sx), ("right", dx)):
        el = OxmlElement(f"w:{lato}"); el.set(qn("w:w"), str(v)); el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcpr.append(mar)


def _bordi(cell, sopra=None, sotto=None, sx=None, dx=None):
    """Ogni lato: None = nessun bordo, oppure (colore, spessore in ottavi di punto)."""
    tcpr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for lato, v in (("top", sopra), ("left", sx), ("bottom", sotto), ("right", dx)):
        el = OxmlElement(f"w:{lato}")
        if v:
            el.set(qn("w:val"), "single"); el.set(qn("w:color"), v[0]); el.set(qn("w:sz"), str(v[1]))
        else:
            el.set(qn("w:val"), "nil")
        b.append(el)
    tcpr.append(b)


def _p(cell_o_doc, prima=0, dopo=0, align=WD_ALIGN_PARAGRAPH.LEFT, nuovo=True):
    if not nuovo:
        p = cell_o_doc.paragraphs[0]
    else:
        p = cell_o_doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(prima), Pt(dopo), 1.0
    return p


def _righe_riepilogo(info):
    out = []
    for r in info["righe"]:
        k = r["chiave"]
        if k in DESCRIZIONI:
            nome, descr = DESCRIZIONI[k]
            out.append((nome, descr, r["importo"], False))
        elif k == "mansioni_aggiuntive":
            out.append(("Mansioni aggiuntive KIT Lavoratori", f"{r['qta']} oltre la 3ª × {formatta_eur(EXTRA_PER_MANSIONE)}",
                        r["importo"], False))
        else:
            out.append((r["voce"], None, r["importo"], True))
    return out


def _box_riepilogo(doc, info):
    righe = _righe_riepilogo(info)
    t = _tabella_senza_bordi(doc, len(righe) + 2, 2, (Cm(12.0), Cm(5.0)))
    grigio = ("D9D9D9", 4)
    # intestazione verde
    h0, h1 = t.cell(0, 0), t.cell(0, 1)
    for c in (h0, h1):
        _set_cell_shading(c, VERDE); _margini(c, 120, 120); _bordi(c, sopra=(VERDE, 8), sx=(VERDE, 8) if c is h0 else None,
                                                                    dx=(VERDE, 8) if c is h1 else None)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _run(_p(h0, nuovo=False), "LA VOSTRA PROPOSTA IN SINTESI", bold=True, size=12, colore="FFFFFF")
    _run(_p(h1, nuovo=False, align=WD_ALIGN_PARAGRAPH.RIGHT), "canone annuo · IVA esclusa", size=8, colore="DDE5E0")
    # voci
    for i, (nome, descr, importo, sconto) in enumerate(righe, start=1):
        c0, c1 = t.cell(i, 0), t.cell(i, 1)
        for c in (c0, c1):
            _margini(c, 90, 90); c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _bordi(c0, sotto=grigio, sx=grigio); _bordi(c1, sotto=grigio, dx=grigio)
        colore = ROSSO if sconto else None
        _run(_p(c0, nuovo=False), nome, size=10, colore=colore)
        if descr:
            _run(_p(c0), descr, size=8, colore=GRIGIO)
        _run(_p(c1, nuovo=False, align=WD_ALIGN_PARAGRAPH.RIGHT), formatta_eur(importo), size=10, colore=colore)
    # totale
    n = len(righe) + 1
    t0, t1 = t.cell(n, 0), t.cell(n, 1)
    for c in (t0, t1):
        _set_cell_shading(c, CREMA); _margini(c, 140, 140); c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _bordi(t0, sotto=grigio, sx=grigio); _bordi(t1, sotto=grigio, dx=grigio)
    _run(_p(t0, nuovo=False), "Totale annuo", bold=True, size=11)
    _run(_p(t1, nuovo=False, align=WD_ALIGN_PARAGRAPH.RIGHT), formatta_eur(info["totale"]), bold=True, size=18, colore=VERDE)
    for r in t.rows:
        r._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    return t


def _badge(p, numero):
    r = _run(p, f" {numero} ", bold=True, size=9)
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), GIALLO)
    r._r.get_or_add_rPr().append(shd)


def _passi(container_cell, info, data_str):
    """Tre schede crema con numero giallo, una sotto l'altra."""
    n = container_cell.add_table(rows=len(PROSSIMI_PASSI), cols=1)
    n.autofit = False
    iva = round(info["totale"] * ALIQUOTA_IVA / 100, 2)
    for i, (titolo, testo) in enumerate(PROSSIMI_PASSI):
        c = n.cell(i, 0)
        c.width = Cm(6.4)
        _set_cell_shading(c, CREMA); _margini(c, 110, 110, 160, 160)
        _bordi(c, sopra=("FFFFFF", 24) if i else None, sotto=("FFFFFF", 24) if i < len(PROSSIMI_PASSI) - 1 else None)
        p = _p(c, nuovo=False)
        _badge(p, i + 1)
        _run(p, "  " + titolo.upper(), bold=True, size=10, colore=VERDE)
        _run(_p(c, prima=2), testo.format(
            totale=formatta_eur(info["totale"]), iva=formatta_eur(iva), ivato=formatta_eur(info["totale"] + iva),
            aliquota=ALIQUOTA_IVA, data=data_str), size=9)
    return n


def pagina_lettera(doc, cliente, info, spec, data_str):
    # data, riferimenti e destinatario, staccati dall'intestazione
    _par(doc, prima=0, dopo=12)
    t = _tabella_senza_bordi(doc, 1, 2, (Cm(8.8), Cm(8.2)))
    sx, dx = t.cell(0, 0), t.cell(0, 1)
    for i, (etichetta, valore, coda) in enumerate((
            ("Villafranca di Verona, ", data_str, ""),
            ("Rif. ", "Offerta KIT FORMASUBITO", ""),
            ("Il suo referente: ", "Tazio Pradella", " · 045 2050597"))):
        p = _p(sx, nuovo=(i > 0), dopo=2)
        _run(p, etichetta, size=9, colore=GRIGIO); _run(p, valore, size=9); _run(p, coda, size=9, colore=GRIGIO)
    _run(_p(dx, nuovo=False, dopo=2), "Spett.le", size=8, colore=GRIGIO)
    _run(_p(dx, dopo=2), cliente["ragione"].upper(), bold=True, size=12)
    for riga in [r.strip() for r in cliente["indirizzo"].split("\n") if r.strip()]:
        _run(_p(dx), riga, size=10)
    _run(_p(dx), f"P.IVA {cliente['piva']}", size=10)
    _run(_p(dx, prima=3), cliente["mail"], size=8.5, colore=GRIGIO)

    # oggetto e apertura
    _run(_p(doc, prima=16), "Oggetto: offerta KIT FORMASUBITO – formazione obbligatoria D.Lgs. 81/08 erogata in azienda",
         bold=True, size=10.5)
    p = _p(doc, prima=4, dopo=10, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    _run(p, "In seguito ai colloqui intercorsi, Vi proponiamo i KIT FORMASUBITO: il Datore di Lavoro eroga direttamente "
            "la formazione obbligatoria ai propri lavoratori e preposti, con progetto formativo e documenti già "
            "predisposti da noi sulle mansioni del Vostro DVR.", size=10)

    _box_riepilogo(doc, info)
    if spec.get("nota"):
        _run(_p(doc, prima=3), spec["nota"], size=8, colore=GRIGIO)

    # grafica e prossimi passi
    _p(doc, prima=0, dopo=0)
    gid = spec.get("grafica", "tu-formi")
    if gid and gid not in GRAFICHE:
        raise ValueError(f"Grafica ignota {gid!r}; disponibili: {list(GRAFICHE)}")
    g = _tabella_senza_bordi(doc, 1, 2, (Cm(10.4), Cm(6.6)))
    sxg, dxg = g.cell(0, 0), g.cell(0, 1)
    if gid:
        _set_cell_shading(sxg, CREMA); _margini(sxg, 120, 120, 120, 120)
        sxg.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        largh = 6.5 if gid == "infografica" else 9.8
        p = _p(sxg, nuovo=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        p.add_run().add_picture(str(ASSETS / "grafiche" / GRAFICHE[gid]["file"]), width=Cm(largh))
    _margini(dxg, 0, 0, 200, 0)
    dxg.paragraphs[0].paragraph_format.space_after = Pt(0)
    _passi(dxg, info, data_str)
    _run(_p(doc, prima=4), "Contenuti del servizio e condizioni complete alla pagina 2.", size=8, colore=GRIGIO)


# ============================================================
# PAGINE 2 E 3 — TESTI DA FILE
# ============================================================
def pagina_da_testo(doc, path, size=8):
    tabella = []

    def chiudi_tabella():
        if not tabella:
            return
        n = len(tabella[0])
        t = doc.add_table(rows=len(tabella), cols=n)
        t.autofit = False
        larg = [Cm(5.5), Cm(6.5), Cm(5.0)] if n == 3 else [Cm(17 / n)] * n
        for ri, riga in enumerate(tabella):
            for ci, testo in enumerate(riga):
                c = t.cell(ri, ci)
                c.width = larg[ci]
                _set_cell_borders(c, size="4", color="BFBFBF")
                if ri == 0:
                    _set_cell_shading(c, VERDE)
                    c.paragraphs[0]._p.getparent().remove(c.paragraphs[0]._p)
                    _run(_par(c, prima=0, align=WD_ALIGN_PARAGRAPH.LEFT), testo, bold=True, size=8, colore="FFFFFF")
                else:
                    _cella_testo(c, [testo], size=8)
        t.style = None
        tabella.clear()

    for riga in Path(path).read_text(encoding="utf-8").splitlines():
        riga = riga.rstrip()
        if riga.startswith("|"):
            tabella.append([c.strip() for c in riga.strip("|").split("|")])
            continue
        chiudi_tabella()
        if not riga:
            continue
        if riga.startswith("## "):
            _striscia(doc, riga[3:])
        elif riga == "[FIRMA]":
            _blocco_firme(doc, "Luogo e data, il Cliente (timbro e firma)", "OVERALL GROUP S.r.l.")
        elif riga == "[FIRMA_CLIENTE]":
            _blocco_firme(doc, "Il Cliente (timbro e firma)")
        elif riga == "[FIRMA_CONSENSO]":
            _blocco_firme(doc, "Data e firma")
        else:
            p = _par(doc, riga, size=size, prima=2)
            if riga.startswith("**Approvazione"):
                p.paragraph_format.keep_together = True
                p.paragraph_format.keep_with_next = True
    chiudi_tabella()


# ============================================================
# ORCHESTRAZIONE
# ============================================================
def genera_offerta(out_path, cliente, spec, data_str=None):
    data_str = data_str or date.today().strftime("%d/%m/%Y")
    info = calcola_righe(spec)
    doc = Document(str(ASSETS / "carta_intestata.docx"))
    sect = doc.element.body[-1]           # sectPr con intestazione e piè di pagina
    pagina_lettera(doc, cliente, info, spec, data_str)
    _interruzione_pagina(doc)
    pagina_da_testo(doc, ASSETS / "testi" / "condizioni.txt")
    pagina_da_testo(doc, ASSETS / "testi" / "informativa.txt")
    doc.element.body.append(sect)         # il sectPr deve restare in fondo
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return {**info, "data": data_str, "out_path": out_path}


def main():
    ap = argparse.ArgumentParser()
    for a in ("--out", "--ragione", "--indirizzo", "--piva", "--mail", "--spec"):
        ap.add_argument(a, required=True)
    ap.add_argument("--data", default=None)
    args = ap.parse_args()
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    cliente = {"ragione": args.ragione, "indirizzo": args.indirizzo.replace("\\n", "\n"),
               "piva": args.piva, "mail": args.mail}
    info = genera_offerta(args.out, cliente, spec, args.data)
    print("✅ Offerta generata")
    for r in info["righe"]:
        print(f"  {r['voce'][:60]:<60} {r['qta']:>3}  {formatta_eur(r['importo']):>12}")
    print(f"  {'Totale annuo (IVA esclusa)':<60}      {formatta_eur(info['totale']):>12}")
    print(f"  Grafica: {spec.get('grafica', 'tu-formi') or 'nessuna'}")
    print(f"  Data: {info['data']}   Output: {info['out_path']}")


def _autoverifica():
    # Belli Freschi: 3 KIT, 4 mansioni, mansione aggiuntiva scontata al 100% -> 447,00
    i = calcola_righe({"kit": ["aggiornamenti", "lavoratori", "preposti"], "mansioni": 4,
                       "sconti": [{"descrizione": "s", "su": "mansioni_aggiuntive", "percentuale": 100}]})
    assert i["totale"] == 447.00 and [r["chiave"] for r in i["righe"]][:3] == ["lavoratori", "mansioni_aggiuntive", "preposti"], i
    assert calcola_righe({"kit": ["lavoratori"], "mansioni": 5})["totale"] == 288.98
    assert calcola_righe({"kit": ["lavoratori"], "mansioni": 2})["totale"] == 249.00
    assert calcola_righe({"kit": ["preposti", "aggiornamenti"]})["totale"] == 298.00
    assert calcola_righe({"kit": ["lavoratori", "preposti", "aggiornamenti"], "mansioni": 3, "bundle": False})["totale"] == 547.00
    assert calcola_righe({"kit": ["preposti"], "sconti": [{"descrizione": "x", "importo": 49}]})["totale"] == 100.00
    for bad in ({"kit": []}, {"kit": ["lavoratori"]}, {"kit": ["preposti"], "sconti": [{"descrizione": "x", "su": "lavoratori", "percentuale": 10}]}):
        try:
            calcola_righe(bad)
        except ValueError:
            continue
        raise AssertionError(bad)
    print("autoverifica OK")


if __name__ == "__main__":
    import sys
    if sys.argv[1:] == ["--test"]:
        _autoverifica()
    else:
        main()
