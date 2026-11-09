#!/usr/bin/env python3
"""
Genera l'offerta KIT FORMASUBITO dal fac-simile Word ufficiale
(assets/template_offerta.docx) popolando:
- intestazione cliente (Spett.le + indirizzo + Mail) nella Tabella 0
- data dell'offerta nella Tabella 0
- grafica riassuntiva in prima pagina (scelta tra assets/grafiche/)
- tabella "Corrispettivo economico" con i KIT scelti, gli sconti e il totale

Il resto del documento (articoli A1-A11, GDPR, firme) NON viene modificato.

USO:
  genera_offerta.py --template T --out O --ragione R --indirizzo I --piva P
                    --mail M --spec offerta.json [--data gg/mm/aaaa]

offerta.json:
{
  "kit": ["lavoratori", "preposti", "aggiornamenti"],   # almeno uno
  "mansioni": 4,                     # obbligatorio se c'e' "lavoratori"
  "bundle": true,                    # sconto "Tutto e subito" se ci sono tutti e 3 i KIT (default true)
  "sconti": [                        # facoltativi, decisi dal compilatore
    {"descrizione": "Sconto mansione aggiuntiva", "su": "mansioni_aggiuntive", "percentuale": 100},
    {"descrizione": "Sconto commerciale", "importo": 20.00}
  ],
  "grafiche": ["infografica"],       # id in GRAFICHE, una o piu', in ordine (default: infografica)
  "nota": "testo facoltativo sotto la tabella prezzi"
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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

ASSETS = Path(__file__).resolve().parent.parent / "assets"

# ============================================================
# CATALOGO (listino, IVA esclusa, canoni annui)
# ============================================================
MANSIONI_INCLUSE = 3
EXTRA_PER_MANSIONE = 19.99
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
    "tu-formi": {"file": "tu-formi.png", "titolo": "Tu formi. Al resto pensiamo noi", "prezzi": None},
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


# ============================================================
# HELPER docx — ereditare lo stile dei run di ancora
# ============================================================
def _clona_run_pr(run_sorgente, run_destinazione):
    """
    Copia il blocco <w:rPr> (run properties) da run_sorgente a run_destinazione.
    Garantisce che il nuovo testo erediti font, dimensione, colore esatti.
    """
    sorgente_pr = run_sorgente._element.find(qn("w:rPr"))
    if sorgente_pr is None:
        return
    # Rimuovi eventuale rPr già presente nel destinatario
    dest_pr_existing = run_destinazione._element.find(qn("w:rPr"))
    if dest_pr_existing is not None:
        run_destinazione._element.remove(dest_pr_existing)
    # Clona e inserisci all'inizio (rPr deve precedere il <w:t>)
    nuovo_pr = copy.deepcopy(sorgente_pr)
    run_destinazione._element.insert(0, nuovo_pr)


def _aggiungi_run_ereditando_stile(paragrafo, testo, run_modello, bold=None):
    """
    Aggiunge un run al paragrafo, clonando lo stile da run_modello.
    Se bold è specificato (True/False), sovrascrive il bold ereditato.
    """
    nuovo = paragrafo.add_run(testo)
    _clona_run_pr(run_modello, nuovo)
    if bold is not None:
        nuovo.bold = bold
    return nuovo


def _svuota_paragrafo_tieni_pPr(paragrafo):
    """Rimuove tutti i run del paragrafo (lascia il pPr / formattazione)."""
    for run in list(paragrafo.runs):
        run._element.getparent().remove(run._element)


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


# ============================================================
# PATCH INTESTAZIONE (Tabella 0)
# ============================================================
def patch_intestazione(doc, ragione, indirizzo, piva, mail):
    """
    T0/R0/C1 (== T0/R1/C1, cella verticalmente unita):
      para[0] = "Spett.le"   -> aggiungo ragione sociale
      para[1..6] = vuoti     -> riempio con indirizzo + P.IVA
      para[7] = "Mail: "     -> aggiungo email/PEC
      para[8] = vuoto

    Strategia chiave: NON forzare font Calibri o size custom.
    Clonare lo stile (w:rPr) dei run modello "Spett.le" e "Mail: "
    per ereditare font e corpo esattamente come nel template.
    """
    t0 = doc.tables[0]
    cell = t0.rows[0].cells[1]
    paragraphs = cell.paragraphs

    if not paragraphs[0].text.strip().startswith("Spett.le"):
        raise RuntimeError(
            f"Template inatteso: T0/R0/C1 para[0] = {paragraphs[0].text!r} (atteso 'Spett.le')"
        )
    if not paragraphs[7].text.strip().startswith("Mail:"):
        raise RuntimeError(
            f"Template inatteso: T0/R0/C1 para[7] = {paragraphs[7].text!r} (atteso 'Mail:')"
        )

    # Modelli di stile dai run originali
    run_modello_spettle = paragraphs[0].runs[0]   # "Spett.le"
    run_modello_mail = paragraphs[7].runs[0]      # "Mail: "

    # 1) "Spett.le" rimane da solo nel para[0] (NIENTE break/ragione attaccata).
    #    Il para[1] resta VUOTO come spaziatura voluta dal cliente.
    #    Ragione sociale + indirizzo + P.IVA partono dal para[2] in giù.
    righe_indirizzo = [r.strip() for r in indirizzo.split("\n") if r.strip()]
    righe_extra = [ragione.upper()] + righe_indirizzo + [f"P.IVA/C.F.: {piva}"]

    # Slot disponibili: dal para[2] al para[6] = 5 paragrafi
    # (para[1] resta vuoto come spaziatura; para[7] è "Mail: ")
    slot_disponibili = paragraphs[2:7]
    for i, testo in enumerate(righe_extra[:len(slot_disponibili)]):
        p = slot_disponibili[i]
        _svuota_paragrafo_tieni_pPr(p)
        # La prima riga (ragione sociale) è in bold, le altre no
        is_ragione = (i == 0)
        _aggiungi_run_ereditando_stile(
            p, testo, run_modello_spettle, bold=is_ragione
        )

    # 3) Email/PEC accanto a "Mail: "
    _aggiungi_run_ereditando_stile(paragraphs[7], mail, run_modello_mail, bold=True)


# ============================================================
# PATCH DATA (Tabella 0 - cella in basso a sinistra)
# ============================================================
def patch_data(doc, data_str):
    """
    T0/R1/C0: 'Villafranca di Verona, lì <DATA>'.
    Sostituisco la data esistente con regex sul testo dei run.
    """
    t0 = doc.tables[0]
    cell = t0.rows[1].cells[0]
    p = cell.paragraphs[0]

    pattern_data = re.compile(r"\d{1,2}/\d{1,2}/\d{4}|gg/mm/aaaa")
    for run in p.runs:
        if pattern_data.search(run.text):
            run.text = pattern_data.sub(data_str, run.text)
            return
    raise RuntimeError(f"Template inatteso: data segnaposto non trovata in {p.text!r}")


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
# GRAFICA DI PRIMA PAGINA
# ============================================================
def patch_grafiche(doc, ids):
    """
    Sostituisce l'immagine della sezione "Descrizione del servizio" (la prima
    immagine del corpo) con le grafiche scelte, una sotto l'altra, a tutta
    larghezza disponibile mantenendo le proporzioni.
    """
    from PIL import Image

    ids = ids or ["infografica"]
    ignoti = [g for g in ids if g not in GRAFICHE]
    if ignoti:
        raise ValueError(f"Grafiche ignote {ignoti}; disponibili: {list(GRAFICHE)}")

    blip = doc.element.body.find(".//" + qn("a:blip"))
    if blip is None:
        raise RuntimeError("Template inatteso: nessuna immagine nel corpo del documento")
    drawing = blip
    while drawing.tag != qn("w:drawing"):
        drawing = drawing.getparent()
    run_el = drawing.getparent()
    para_el = run_el
    while para_el.tag != qn("w:p"):
        para_el = para_el.getparent()

    extent = drawing.find(".//" + qn("wp:extent"))
    larghezza = int(extent.get("cx"))          # EMU: larghezza del riquadro originale
    sezione = doc.sections[0]
    utile = int(sezione.page_width - sezione.left_margin - sezione.right_margin)
    larghezza = min(max(larghezza, utile), utile)

    ultimo = para_el
    for i, gid in enumerate(ids):
        path = ASSETS / "grafiche" / GRAFICHE[gid]["file"]
        w, h = Image.open(path).size
        altezza = int(larghezza * h / w)
        p_el = para_el if i == 0 else copy.deepcopy(para_el)
        if i:
            ultimo.addnext(p_el)
            ultimo = p_el
        rid, _ = doc.part.get_or_add_image(str(path))
        p_el.find(".//" + qn("a:blip")).set(qn("r:embed"), rid)
        for tag in ("wp:extent", "a:ext"):
            for el in p_el.iter(qn(tag)):
                el.set("cx", str(larghezza))
                el.set("cy", str(altezza))
        for el in p_el.iter(qn("wp:docPr")):
            el.set("id", str(9000 + i))
            el.set("name", f"Grafica {gid}")


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


def patch_prezzo(doc, info, nota=None):
    """Sostituisce la Tabella 1 del template con la tabella prezzi (+ nota facoltativa)."""
    t1 = doc.tables[1]
    parent = t1._element.getparent()
    indice = list(parent).index(t1._element)
    nuova = _tabella_prezzi(doc, info)
    parent.insert(indice, nuova._element)
    dopo = OxmlElement("w:p")  # spaziatura prima dell'art. A2
    parent.insert(indice + 1, dopo)
    if nota:
        from docx.text.paragraph import Paragraph
        p_el = OxmlElement("w:p")
        dopo.addprevious(p_el)
        p = Paragraph(p_el, t1._parent)
        run = p.add_run(nota)
        run.font.name = "Gill Sans MT"
        run.font.size = Pt(8)
    parent.remove(t1._element)


# ============================================================
# ORCHESTRAZIONE
# ============================================================
def genera_offerta(template_path, out_path, ragione, indirizzo, piva, mail, spec, data_str=None):
    data_str = data_str or date.today().strftime("%d/%m/%Y")
    info = calcola_righe(spec)
    doc = Document(template_path)
    patch_intestazione(doc, ragione=ragione, indirizzo=indirizzo, piva=piva, mail=mail)
    patch_data(doc, data_str=data_str)
    patch_grafiche(doc, spec.get("grafiche"))
    patch_prezzo(doc, info, nota=spec.get("nota"))
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return {**info, "data": data_str, "out_path": out_path}


def main():
    ap = argparse.ArgumentParser()
    for a in ("--template", "--out", "--ragione", "--indirizzo", "--piva", "--mail", "--spec"):
        ap.add_argument(a, required=True)
    ap.add_argument("--data", default=None)
    args = ap.parse_args()
    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    info = genera_offerta(args.template, args.out, args.ragione, args.indirizzo.replace("\\n", "\n"),
                          args.piva, args.mail, spec, args.data)
    print("✅ Offerta generata")
    for r in info["righe"]:
        print(f"  {r['voce'][:60]:<60} {r['qta']:>3}  {formatta_eur(r['importo']):>12}")
    print(f"  {'Totale annuo (IVA esclusa)':<60}      {formatta_eur(info['totale']):>12}")
    print(f"  Grafiche: {', '.join(spec.get('grafiche') or ['infografica'])}")
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
