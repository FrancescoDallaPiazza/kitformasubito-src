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
                data: logoBytes, type: 'jpg',
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
  ragioneSociale : 'Calor Energy Verona Società Cooperativa',
  ragioneSocialeBreve : 'Calor Energy Verona Soc. Coop.',
  indirizzo      : 'Via M.G. Deledda, 6 - 37059 Campagnola di Zevio (VR)',
  piva           : '04905340230',
  atecoCodice    : '43.22.07',
  atecoDesc      : 'Installazione impianti riscaldamento e condizionamento aria',
  datoreLavoro   : 'Giulia Furlani',
  anno           : '2026',
};

// ─── MANSIONI ────────────────────────────────────────────────────────────────
const MANSIONI = [
  {
    id: 'ImpiegataAmm',
    nome: 'Impiegata/o Amministrativa/o',
    reparto: 'Amministrazione',
    livello: 'BASSO',
    oreSpec: 4,
    rischi: [
      {
        nome: 'Utilizzo di videoterminali (VDT)',
        misure: [
          'Corretto posizionamento del monitor per evitare riflessi e abbagliamenti',
          'Pause regolari di 15 minuti ogni 2 ore di lavoro continuativo al VDT',
          'Regolazione di luminosità e contrasto dello schermo',
          'Uso di seduta ergonomica con schienale regolabile',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
      {
        nome: 'Posture incongrue / sovraccarico muscolo-scheletrico',
        misure: [
          'Strutturazione ergonomica della postazione di lavoro (seduta, schermo, tastiera)',
          'Alternanza tra posizione seduta e movimenti durante la giornata',
          'Esercizi di stretching e rilassamento muscolare',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
      {
        nome: 'Stress lavoro-correlato',
        misure: [
          'Organizzazione dei carichi di lavoro con scadenze realistiche',
          'Comunicazione trasparente tra colleghi e responsabili',
          'Gestione dei conflitti interpersonali e supporto psicologico',
          'Pause regolari e rispetto degli orari di lavoro',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
      {
        nome: 'Scivolamenti e inciampi in ambiente domestico/smart working',
        misure: [
          'Mantenimento dell\'ordine e pulizia della postazione di lavoro',
          'Cavi elettrici ordinati e protetti per evitare inciampi',
          'Chiusura dei cassetti dopo l\'uso',
          'Uso di scalette idonee per accedere a ripiani alti',
        ],
        dpi: ['Calzature con suola antiscivolo'],
      },
      {
        nome: 'Rischio incendio (ambiente domestico)',
        misure: [
          'Disponibilità di estintore e piano di evacuazione',
          'Non utilizzare fiamme libere in prossimità di materiale infiammabile',
          'Segnalazione immediata di guasti agli impianti elettrici',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
    ],
    dpi: ['Calzature con suola antiscivolo'],
  },
  {
    id: 'Commerciale',
    nome: 'Commerciale',
    reparto: 'Commerciale',
    livello: 'BASSO',
    oreSpec: 4,
    rischi: [
      {
        nome: 'Guida di autoveicolo (rischio stradale)',
        misure: [
          'Divieto assoluto di utilizzo del telefono durante la guida (salvo vivavoce)',
          'Rispetto dei limiti di velocità e del codice della strada',
          'Verifica preliminare del funzionamento del mezzo (freni, pneumatici, luci)',
          'Divieto di guida in condizioni di stanchezza estrema o dopo assunzione di alcol',
          'Manutenzione programmata del veicolo con revisioni periodiche',
        ],
        dpi: ['Gilet alta visibilità (in caso di emergenza su strada)'],
      },
      {
        nome: 'Stress lavoro-correlato',
        misure: [
          'Pianificazione realistica delle visite e degli spostamenti',
          'Limitazione delle ore di guida consecutive con pause regolari',
          'Supporto da parte del responsabile in caso di difficoltà relazionali con clienti',
          'Gestione assertiva dei conflitti con clienti',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
      {
        nome: 'Utilizzo di videoterminali (VDT)',
        misure: [
          'Corretto posizionamento del monitor/laptop per evitare riflessi',
          'Pause regolari di 15 minuti ogni 2 ore di utilizzo del VDT',
          'Uso di seduta ergonomica nelle postazioni fisse',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
      {
        nome: 'Posture incongrue durante guida e lavoro',
        misure: [
          'Corretta regolazione del sedile e degli specchietti dell\'auto',
          'Pause di deambulazione durante i trasferimenti lunghi',
          'Esercizi di stretching durante le soste',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
      {
        nome: 'Aggressioni / comportamenti ostili da parte di clienti',
        misure: [
          'Formazione sulle tecniche di comunicazione assertiva e de-escalation',
          'Non recarsi mai da soli in contesti percepiti come rischiosi',
          'Segnalazione immediata al datore di lavoro di episodi critici',
        ],
        dpi: ['Nessun DPI specifico richiesto'],
      },
    ],
    dpi: ['Calzature con suola antiscivolo', 'Gilet alta visibilità'],
  },
  {
    id: 'Manutentore',
    nome: 'Manutentore',
    reparto: 'Manutenzione / Installazione',
    livello: 'ALTO',
    oreSpec: 12,
    rischi: [
      {
        nome: 'Cadute dall\'alto (uso scale portatili e lavori in quota)',
        misure: [
          'Utilizzo esclusivo di scale portatili conformi alla norma EN 131',
          'Ispezione della scala prima dell\'uso (pioli, piedini antiscivolo, vincoli)',
          'Posizionamento della scala su superficie stabile e piana',
          'Mantenimento del centro di gravità all\'interno degli stili durante la salita',
          'Divieto di lavori in quota per tirocinanti non formati',
          'Uso di imbragatura anticaduta per lavori su tetti o dislivelli > 2 m',
        ],
        dpi: ['Calzature antinfortunistiche', 'Imbragatura anticaduta (per quota > 2 m)'],
      },
      {
        nome: 'Tagli e lesioni meccaniche',
        misure: [
          'Uso di attrezzature da taglio con protezioni attive (carter, schermi)',
          'Verifica dello stato di usura degli utensili prima dell\'uso',
          'Divieto di rimuovere le protezioni fisse dalle attrezzature',
          'Riporre gli utensili taglienti in apposite custodie dopo l\'uso',
        ],
        dpi: ['Guanti antitaglio', 'Guanti in cuoio', 'Occhiali di sicurezza'],
      },
      {
        nome: 'Agenti chimici – gas refrigeranti e combustibili',
        misure: [
          'Verifica della disponibilità e lettura delle Schede Dati di Sicurezza (SDS)',
          'Aerazione del locale prima e durante interventi su impianti a gas',
          'Utilizzo di rilevatore di perdite gas (analizzatore) certificato',
          'Divieto di accensione di fiamme libere in locali con sospetta presenza di gas',
          'Smaltimento dei refrigeranti secondo normativa F-Gas (Reg. UE 517/2014)',
        ],
        dpi: ['Guanti in gomma/PVC resistenti agli agenti chimici', 'Mascherina facciale filtrante (FFP2)', 'Occhiali di sicurezza'],
      },
      {
        nome: 'Rischio elettrico',
        misure: [
          'Verifica assenza di tensione prima di interventi su impianti elettrici (multimetro)',
          'Sezionamento e blocco dell\'impianto prima di ogni intervento (LOTO)',
          'Utilizzo di attrezzature con doppio isolamento (Classe II)',
          'Verifica periodica del collegamento a terra degli impianti',
        ],
        dpi: ['Guanti isolanti per lavori elettrici (quando necessari)', 'Calzature antinfortunistiche con suola isolante'],
      },
      {
        nome: 'Esposizione a rumore',
        misure: [
          'Preferire attrezzature con minore emissione sonora',
          'Limitazione dei tempi di esposizione alle fasi più rumorose',
          'Manutenzione periodica delle attrezzature per ridurre le emissioni',
        ],
        dpi: ['Otoprotettori (inserti o cuffie) nelle fasi con livello > 80 dB(A)'],
      },
      {
        nome: 'Esposizione a calore (caldaie, bruciatori, impianti termici)',
        misure: [
          'Attesa del raffreddamento dell\'impianto prima di interventi su parti calde',
          'Segnalazione delle superfici calde con apposita segnaletica',
          'Ventilazione adeguata dei locali caldaia durante l\'intervento',
        ],
        dpi: ['Guanti in cuoio resistenti al calore'],
      },
      {
        nome: 'Ambienti confinati / rischio asfissia (locali tecnici)',
        misure: [
          'Verifica della qualità dell\'aria con idoneo strumento prima dell\'ingresso',
          'Presenza obbligatoria di un addetto esterno durante il lavoro in confinato',
          'Divieto di lavorare da soli in spazi confinati',
          'Piano di emergenza e procedure di evacuazione predisposte prima dell\'intervento',
        ],
        dpi: ['Maschera con filtro adeguato (ABEK) se necessario'],
      },
    ],
    dpi: [
      'Calzature antinfortunistiche',
      'Occhiali di sicurezza',
      'Guanti antitaglio',
      'Guanti in cuoio',
      'Guanti in gomma/PVC',
      'Imbragatura anticaduta',
      'Elmetto (in cantieri edili)',
      'Mascherina facciale filtrante FFP2',
      'Gilet alta visibilità',
    ],
  },
];

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
