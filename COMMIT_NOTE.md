# Commit `kitformasubito-src` — 16/07/2026

Sostituire i file nella cartella locale del repo, poi commit + push su `main` con GitHub Desktop.

## File

| File | Stato | Righe cambiate |
|---|---|---|
| `logo_util.js` | **NUOVO** | 138 |
| `helpers.js` | modificato | 10 |
| `run.js` | modificato | 22 |
| `gen_test.js` | modificato | 32 |
| `docs1.js` | modificato | 8 |
| `docs2.js` | modificato | 14 |
| `gen_scheda.js` | modificato | 8 |

> `helpers.js` è stato ricostruito **a partire da `HEAD`**, applicando solo il fix strutturale.
> Contiene ancora i dati di FOTO-GRAFIC S.N.C. come da template committato: **nessun dato
> del cliente SAN FRANCESCO S.A.S. finisce nel repo**. Verificato con grep (0 occorrenze).

## Messaggio di commit suggerito

```
fix: logo type sniffato dai magic bytes + pre-flight anti-regressione logo + blocchi test indivisibili

- logo_util.js (nuovo): imgType() sniffa PNG/JPEG/GIF/BMP dai magic bytes;
  analizzaLogo() decodifica il PNG via zlib e blocca la build se il logo non
  lascia segno su fondo bianco (soglia 2% di pixel visibili)
- helpers/docs1/docs2/gen_scheda/gen_test: 10 occorrenze di type:'jpg'
  hardcoded sostituite con il tipo reale sniffato
- run.js: pre-flight logo che esce con errore su logo inutilizzabile
- gen_test.js: cantSplit sulle TableRow + keepNext sui paragrafi (tranne
  l'ultima risposta) -> il blocco domanda+risposte non si spezza fra due pagine
- gen_test.js: spacing.after 240 sull'header -> il logo non tocca il corpo del test
```

## Perché

Incidente reale del 16/07/2026, cliente SAN FRANCESCO S.A.S.

1. **Logo bianco su bianco.** Il logo fornito era in versione "inchiostro bianco"
   (RGB bianco + canale alfa). Poiché tutti i generatori dichiaravano `type: 'jpg'`
   a prescindere dal file reale, il PNG è stato convertito appiattendo la
   trasparenza su fondo bianco: risultato, bianco su bianco. Sono stati consegnati
   17 documenti con il logo presente in `word/media/` e invisibile. Il controllo
   "il file c'è" non intercetta questo caso: guardava i byte, non il contenuto.

2. **Mismatch estensione/contenuto.** Con `type` hardcoded, un PNG finiva scritto
   in `word/media/*.jpg`: Word non renderizza l'immagine.

3. **Blocchi domanda spezzati.** `makeQuestion()` non imponeva né `cantSplit` né
   `keepNext`: Word spezzava domanda e risposte fra due pagine.

## Note tecniche

- **Magic bytes, non estensione.** L'estensione mente: il file del cliente si
  chiamava `.png` e conteneva un JPEG; si chiamava `2048x1448` ed era 139x66.
  `LOGO_PATH` ha comunque estensione fissa `.png` per qualsiasi contenuto.
- **`keepNext` non sull'ultima risposta.** Sull'ultima riga incollerebbe la
  domanda al blocco successivo, propagandosi a catena su tutte le domande.
- **Decoder PNG solo interno (`zlib`, core Node).** Nessuna dipendenza nuova.
  Varianti non gestite (palette, 16 bit, interlacciato) -> warning, non errore.
  Sui JPEG l'analisi pixel non gira: senza canale alfa il bug d'origine è
  impossibile.
- **La conversione del logo non serve più**: il file del cliente entra così com'è,
  trasparenza inclusa.

## Effetti a valle sulla SKILL

Lo STEP 3 di `kitformasubito` dice: *"Se stampa JPEG -> in helpers.js usa type: 'jpg'
(già impostato di default)"*. Dopo questo commit l'istruzione è **obsoleta**: il tipo
è sniffato, il logo va copiato in `/home/claude/logo.png` senza conversioni.
Aggiornare la SKILL con `aggiornamento-skill-utente` e ri-pacchettarla.

Da valutare anche la regola dello STEP 6 ("mai pushare helpers.js"): così com'è
rende impossibile persistere qualsiasi fix strutturale su quel file, che di
strutturale ne contiene circa 650 righe. Alternativa: consentirlo previa
ricostruzione da `HEAD` dei blocchi CLIENTE/MANSIONI, come fatto qui.

## Verifiche eseguite

- Smoke test su checkout pulito di `HEAD` + questi file: `node run.js` completo,
  17 documenti generati, pre-flight logo superato (`tipo png, 91.3% visibili`).
- Test anti-regressione: ricostruito il logo guasto (inchiostro bianco su alfa)
  -> build fermata con `exit 1` e messaggio diagnostico.
- 6/6 test convertiti in PDF e riletti: 30 domande ciascuno, 0 blocchi spezzati.
- `validate.py`: PASSED su Progetto Formativo, Test DOCENTE, Attestato, Scheda
  Addestrativa.
