# offerta/ — sorgenti della skill kitformasubito-offerta

La skill li scarica all'avvio (`git clone --depth 1` di questo repo) e usa:

- `scripts/genera_offerta.py` — compone l'offerta (3 pagine) su carta intestata Overall:
  1. lettera (impaginazione B): data/riferimenti e destinatario, oggetto, riquadro
     "La vostra proposta in sintesi" (KIT, sconti, totale), grafica ridotta facoltativa
     e prossimi passi;
  2. condizioni di vendita (`assets/testi/condizioni.txt`);
  3. approvazione specifica, informativa privacy e consenso (`assets/testi/informativa.txt`).
  Input: dati cliente da riga di comando + `--spec offerta.json` (formato nel docstring).
  `python3 scripts/genera_offerta.py --test` esegue l'autoverifica dei calcoli.
- `assets/carta_intestata.docx` — intestazione (loghi Overall e FORMASUBITO, dati societari, filetto verde), piè di pagina (barra verde, contatti, Pag. X/Y) e stili (Gill Sans MT).
- `assets/testi/*.txt` — testi di pag. 2 e 3, modificabili senza toccare lo script:
  `## Titolo` = striscia verde; `**grassetto**`, `*corsivo*`; `|a|b|c|` = tabella;
  `[FIRMA]`, `[FIRMA_CLIENTE]`, `[FIRMA_CONSENSO]` = blocchi firma.
- `assets/grafiche/` — grafiche per la pag. 1 (id in `GRAFICHE`): infografica,
  tu-formi (default), cosa-ricevi, scegli-kit, tre-kit. Le ultime due hanno prezzi stampati.
- `assets/logo_formasubito.png`, `assets/firma_overall.jpeg` — logo e timbro/firma Overall.

Listino (IVA esclusa, canoni annui): KIT Lavoratori 249,00 € fino a 3 mansioni
+ 19,99 € per mansione oltre la 3ª; KIT Preposti 149,00 €; KIT Aggiornamenti 149,00 €;
i 3 KIT insieme: sconto 100,00 € (447,00 €), ripartito 45,52 / 27,24 / 27,24 €.

Testi di pag. 2 e 3: bozza novembre 2026 da far rivedere a un avvocato prima dell'uso.
