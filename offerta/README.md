# offerta/ — sorgenti della skill kitformasubito-offerta

La skill li scarica all'avvio (`git clone --depth 1` di questo repo) e usa:

- `scripts/genera_offerta.py` — genera l'offerta Word: intestazione cliente, data,
  grafica di prima pagina, tabella prezzi con KIT, sconti e totale.
  Input: dati cliente da riga di comando + `--spec offerta.json` (formato nel docstring).
  `python3 scripts/genera_offerta.py --test` esegue l'autoverifica dei calcoli.
- `assets/template_offerta.docx` — fac-simile uniformato (Gill Sans MT, titoli su
  striscia verde FORMASUBITO #0E3A27, data e piè di pagina come testo fisso).
  Ancore usate dallo script: Tabella 0 ("Spett.le", 6 paragrafi vuoti, "Mail:",
  data "Villafranca di Verona, lì 00/00/0000"), Tabella 1 (sostituita dalla
  tabella prezzi), prima immagine del corpo (sostituita dalla grafica scelta).
- `assets/grafiche/` — grafiche selezionabili per la prima pagina
  (id in `GRAFICHE` dello script): infografica, tu-formi, cosa-ricevi,
  scegli-kit, tre-kit. Le ultime due riportano prezzi stampati.

Listino nel catalogo dello script (IVA esclusa, canoni annui): KIT Lavoratori
249,00 € fino a 3 mansioni + 19,99 € per mansione oltre la 3ª; KIT Preposti
149,00 €; KIT Aggiornamenti 149,00 €; i 3 KIT insieme: sconto 100,00 € (447,00 €).
