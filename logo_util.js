'use strict';
// ─────────────────────────────────────────────────────────────────────────────
//  logo_util.js — utilità per il logo cliente (modulo strutturale, pushabile)
//
//  Nasce da due incidenti reali del 16/07/2026 (cliente SAN FRANCESCO S.A.S.):
//
//  1. TIPO IMMAGINE HARDCODED. Tutti i generatori dichiaravano `type: 'jpg'`
//     in ImageRun a prescindere dal file reale, costringendo a convertire a
//     monte ogni logo PNG. imgType() sniffa i MAGIC BYTES: è più robusto sia
//     dell'hardcoding sia del controllo sull'estensione, perché l'estensione
//     mente (il file del cliente si chiamava .png e conteneva già un JPEG, e
//     si chiamava "2048x1448" pur essendo 139x66).
//
//  2. LOGO BIANCO SU BIANCO. Il logo fornito era in versione "inchiostro
//     bianco" (RGB bianco + canale alfa). Appiattito su fondo bianco è
//     diventato un rettangolo invisibile: 17 documenti consegnati con il logo
//     formalmente presente in word/media/ e sostanzialmente inesistente.
//     analizzaLogo() misura i pixel VISIBILI e blocca la build.
// ─────────────────────────────────────────────────────────────────────────────

const zlib = require('zlib');

// ── 1. Sniffing del tipo reale dai magic bytes ──────────────────────────────
// Restituisce la stringa attesa da ImageRun di docx: 'png' | 'jpg' | 'gif' | 'bmp'.
function imgType(bytes) {
  const b = bytes;
  if (b.length < 12) throw new Error('[logo] file troppo corto per essere un\'immagine valida');
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'png';
  if (b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) return 'jpg';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif';
  if (b[0] === 0x42 && b[1] === 0x4D) return 'bmp';
  throw new Error('[logo] formato non riconosciuto: attesi PNG, JPEG, GIF o BMP');
}

// ── 2. Decodifica PNG minimale (senza dipendenze esterne) ───────────────────
// Copre i casi reali dei loghi: bit depth 8, non interlacciato,
// colorType 0 (gray), 2 (RGB), 4 (gray+alpha), 6 (RGBA).
// Sui casi non coperti (palette, 16 bit, interlacciato) restituisce null:
// l'analisi viene saltata con un warning, non con un errore.
function _decodePng(bytes) {
  let off = 8, w = 0, hgt = 0, depth = 0, colorType = -1;
  const idat = [];
  while (off < bytes.length) {
    const len = bytes.readUInt32BE(off);
    const tipo = bytes.toString('ascii', off + 4, off + 8);
    const dati = bytes.slice(off + 8, off + 8 + len);
    if (tipo === 'IHDR') {
      w = dati.readUInt32BE(0); hgt = dati.readUInt32BE(4);
      depth = dati[8]; colorType = dati[9];
      if (dati[12] !== 0) return null;        // interlacciato → non gestito
    } else if (tipo === 'IDAT') idat.push(dati);
    else if (tipo === 'IEND') break;
    off += 12 + len;
  }
  if (depth !== 8) return null;
  const canali = { 0: 1, 2: 3, 4: 2, 6: 4 }[colorType];
  if (!canali) return null;                    // palette → non gestito

  const raw = zlib.inflateSync(Buffer.concat(idat));
  const bpp = canali, stride = w * bpp;
  const out = Buffer.alloc(hgt * stride);
  let pos = 0;
  for (let y = 0; y < hgt; y++) {
    const filtro = raw[pos++];
    const riga = raw.slice(pos, pos + stride); pos += stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? out[y * stride + x - bpp] : 0;
      const b = y > 0 ? out[(y - 1) * stride + x] : 0;
      const c = (x >= bpp && y > 0) ? out[(y - 1) * stride + x - bpp] : 0;
      let v = riga[x];
      if (filtro === 1) v += a;
      else if (filtro === 2) v += b;
      else if (filtro === 3) v += (a + b) >> 1;
      else if (filtro === 4) {                 // Paeth
        const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
        v += (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
      }
      out[y * stride + x] = v & 0xFF;
    }
  }
  return { w, h: hgt, canali, px: out };
}

// ── 3. Check anti-regressione ───────────────────────────────────────────────
// Un logo è INUTILIZZABILE se, una volta stampato su carta bianca, non lascia
// segno: cioè se i pixel opachi e non-bianchi sono una frazione trascurabile.
// È esattamente il caso "inchiostro bianco appiattito su fondo bianco".
//
// L'analisi pixel è possibile solo sui PNG (decoder interno). Sui JPEG il tipo
// è comunque validato, ma il contenuto non viene ispezionato: nessuna
// dipendenza esterna è ammessa nel repo. Il fatto che il canale alfa non
// esista nel JPEG rende comunque impossibile il bug specifico all'origine.
const SOGLIA_VISIBILE = 0.02;   // almeno il 2% di pixel che lasciano segno

function analizzaLogo(bytes) {
  const tipo = imgType(bytes);
  const esito = { tipo, warnings: [], analizzato: false, fraseVisibile: null };

  if (tipo !== 'png') {
    esito.warnings.push(`tipo ${tipo}: analisi dei pixel non eseguita (decoder interno solo per PNG)`);
    return esito;
  }
  const img = _decodePng(bytes);
  if (!img) {
    esito.warnings.push('PNG in variante non gestita dal decoder interno (palette/16bit/interlacciato): analisi saltata');
    return esito;
  }

  const { w, h, canali, px } = img;
  let visibili = 0;
  const tot = w * h;
  for (let i = 0; i < tot; i++) {
    const o = i * canali;
    let r, g, b, a;
    if (canali === 1)      { r = g = b = px[o];     a = 255; }
    else if (canali === 2) { r = g = b = px[o];     a = px[o + 1]; }
    else if (canali === 3) { r = px[o]; g = px[o+1]; b = px[o+2]; a = 255; }
    else                   { r = px[o]; g = px[o+1]; b = px[o+2]; a = px[o+3]; }
    // "lascia segno su carta bianca" = sufficientemente opaco E non quasi-bianco
    if (a > 32 && (r < 240 || g < 240 || b < 240)) visibili++;
  }
  const frazione = visibili / tot;
  esito.analizzato = true;
  esito.frazioneVisibile = frazione;

  if (frazione < SOGLIA_VISIBILE) {
    throw new Error(
      `[logo] LOGO INUTILIZZABILE: solo il ${(frazione * 100).toFixed(1)}% dei pixel lascia segno su fondo bianco ` +
      `(soglia ${(SOGLIA_VISIBILE * 100).toFixed(0)}%). Causa tipica: logo in versione "inchiostro bianco" ` +
      `(RGB bianco su canale alfa) destinato a fondi scuri. Richiedere al cliente la versione a colori ` +
      `oppure appiattire la trasparenza sul colore di sfondo del marchio, MAI sul bianco.`
    );
  }
  if (canali === 4 || canali === 2) esito.warnings.push('il PNG ha canale alfa: la trasparenza viene preservata nel .docx');
  return esito;
}

module.exports = { imgType, analizzaLogo, SOGLIA_VISIBILE };
