'use strict';
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, TabStopPosition, SectionType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak, SimpleField, LineRuleType,
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── PALETTE ────────────────────────────────────────────────────────────────
const C = {
  BLU_DARK   : '1F3864',
  BLU_HEADER : '1F4E79',
  BLU_MED    : '2E75B6',
  BLU_LIGHT  : 'D5E8F0',
  GRIGIO_ALT : 'F2F2F2',
  BLU_ALT    : 'EBF3FB',
  GRIGIO     : '595959',
  BIANCO     : 'FFFFFF',
  ROSSO      : 'C00000',
  ARANCIO    : 'E07000',
  VERDE      : '538135',
  GIALLO     : 'FFFF00',
};

const FONT = 'Gill Sans MT';
const LOGO_PATH = '/home/claude/logo.png';

// ─── DEFAULT STYLES ──────────────────────────────────────────────────────────
const docStyles = {
  default: {
    document: { run: { font: FONT, size: 22 } }
  }
};

// ─── LOGO BYTES ─────────────────────────────────────────────────────────────
const logoBytes = fs.readFileSync(LOGO_PATH);

// ─── HEADER ─────────────────────────────────────────────────────────────────
function makeHeader(ragioneSociale, titoloDoc, atecoCodice, atecoDesc) {
  const tbl = new Table({
    width: { size: 9638, type: WidthType.DXA },
    columnWidths: [2977, 6661],
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 2977, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            borders: {
              top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
              left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
            },
            children: [new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [new ImageRun({
                data: logoBytes, type: 'png',
                transformation: { width: 70, height: 70 },
              })],
            })],
          }),
          new TableCell({
            width: { size: 6661, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            borders: {
              top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
              left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: ragioneSociale, bold: true, color: C.BLU_HEADER, font: FONT, size: 22 })],
              }),
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: titoloDoc, bold: true, color: C.BLU_HEADER, font: FONT, size: 22 })],
              }),
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [new TextRun({ text: `ATECO: ${atecoCodice} – ${atecoDesc}`, color: C.BLU_HEADER, font: FONT, size: 18 })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
  return new Header({ children: [tbl] });
}

// ─── FOOTER ─────────────────────────────────────────────────────────────────
function makeFooter(docTitle = '') {
  return new Footer({
    children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: 9638 }],
      children: [
        new TextRun({ text: docTitle, size: 18, font: FONT, color: C.GRIGIO }),
        new TextRun({ text: '\tPag. ', size: 18, font: FONT }),
        new SimpleField('PAGE'),
        new TextRun({ text: ' a ', size: 18, font: FONT }),
        new SimpleField('NUMPAGES'),
      ],
    })],
  });
}

// ─── TITOLO SEZIONE (riga con bordo inferiore blu) ───────────────────────────
function titoloSezione(testo, sz = 26) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.BLU_MED } },
    children: [new TextRun({ text: testo, bold: true, color: C.BLU_HEADER, font: FONT, size: sz })],
  });
}

// ─── CORPO TESTO ─────────────────────────────────────────────────────────────
function corpo(testo, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { before: opts.before || 30, after: opts.after || 30 },
    children: [new TextRun({
      text: testo, font: FONT, size: opts.sz || 20,
      bold: opts.bold || false, color: opts.color || undefined,
    })],
  });
}

// ─── RIGA DATI (etichetta: valore) ───────────────────────────────────────────
function rigaDati(etichetta, valore, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 30, after: 30 },
    children: [
      new TextRun({ text: etichetta + ': ', bold: true, font: FONT, size: opts.sz || 20 }),
      new TextRun({ text: valore, font: FONT, size: opts.sz || 20 }),
    ],
  });
}

// ─── PARAGRAFO VUOTO ─────────────────────────────────────────────────────────
function vuoto(sp = 60) {
  return new Paragraph({ spacing: { after: sp }, children: [new TextRun({ text: '' })] });
}

// ─── CELLA TABELLA helper ────────────────────────────────────────────────────
function cella(content, opts = {}) {
  const children = typeof content === 'string'
    ? [new Paragraph({
        alignment: opts.align === 'center' ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
        children: [new TextRun({ text: content, font: FONT, size: opts.sz || 20, bold: opts.bold || false, color: opts.color || undefined })],
      })]
    : content;
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    verticalAlign: opts.vAlign || VerticalAlign.TOP,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    columnSpan: opts.span,
    borders: opts.borders || undefined,
    children,
  });
}

// ─── SALVA DOCUMENTO ─────────────────────────────────────────────────────────
async function salvaDoc(doc, filePath) {
  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buf);
  console.log('✓', path.relative('/home/claude/kit/OUT', filePath));
}

// ─── COSTANTI DOCUMENTO ──────────────────────────────────────────────────────
const A4_P = { width: 11906, height: 16838 }; // A4 portrait (DXA)
const A4_L = { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }; // A4 landscape
const MARGIN_STD = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
const MARGIN_REG = { top: 876, right: 1134, bottom: 1134, left: 1134 }; // registro

// ─── DATI CLIENTE ────────────────────────────────────────────────────────────
const CLIENTE = {
  ragioneSociale      : 'RAGIONE SOCIALE CLIENTE',
  ragioneSocialeBreve : 'NOME BREVE CLIENTE',
  indirizzo           : 'VIA, CAP CITTÀ (PROV)',
  piva                : '00000000000',
  atecoCodice         : '00.00.00',
  atecoDesc           : 'DESCRIZIONE ATTIVITÀ',
  datoreLavoro        : 'NOME COGNOME',
  anno                : '2026',
};

// ─── MANSIONI ────────────────────────────────────────────────────────────────
// Compilato da Claude ad ogni generazione in base al DVR del cliente.
// Struttura di ogni mansione:
// {
//   id: 'NomeBreve',          // es. 'ImpiegataAmm' – senza spazi
//   nome: 'Nome esteso',
//   reparto: 'Reparto',
//   livello: 'BASSO',         // BASSO / MEDIO / ALTO
//   oreSpec: 4,               // 4 / 8 / 12
//   rischi: [
//     {
//       nome: 'Nome rischio',
//       misure: ['Misura 1', 'Misura 2'],
//       dpi: ['DPI specifico'],
//     },
//   ],
//   dpi: ['DPI completi della mansione'],
// }
const MANSIONI = [];

module.exports = {
  C, FONT, CLIENTE, MANSIONI, logoBytes,
  docStyles, A4_P, A4_L, MARGIN_STD, MARGIN_REG,
  makeHeader, makeFooter,
  titoloSezione, corpo, rigaDati, vuoto, cella, salvaDoc,
  // re-export docx classes
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, TabStopPosition, SectionType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak, SimpleField, LineRuleType,
};
