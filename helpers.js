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
  ragioneSociale      : 'FOTO-GRAFIC S.N.C. di Rossi E. & Ronca M.',
  ragioneSocialeBreve : 'FOTO-GRAFIC S.N.C.',
  indirizzo           : 'C/o Centro Commerciale Verona Uno – Via Cesare Battisti, 266 – 37067 San Giovanni Lupatoto (VR)',
  piva                : '03348080239',
  atecoCodice         : '74.20.19',
  atecoDesc           : 'Altre attività di riprese fotografiche',
  datoreLavoro        : 'Mara Ronca',
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
const MANSIONI = [
  {
    id: 'AddCommMinuto',
    nome: 'Addetto al Commercio al Minuto',
    reparto: 'Negozio / Magazzino',
    livello: 'BASSO',
    oreSpec: 4,
    rischi: [
      {
        nome: 'Caduta in piano (pavimenti e passaggi)',
        livello: 'ALTO',
        misure: [
          'Mantenere pavimenti e passaggi sgombri da materiali e oggetti',
          'Segnalare immediatamente pavimento bagnato o scivoloso con apposita segnaletica',
          'Indossare calzature con suola antiscivolo durante tutta la giornata lavorativa',
        ],
        dpi: ['Calzature antiscivolo'],
      },
      {
        nome: 'Aggressione fisica (contatto con il pubblico)',
        livello: 'MEDIO',
        misure: [
          'Non restare mai soli durante l\'orario di apertura al pubblico',
          'Utilizzare i sistemi di videosorveglianza attivi nel negozio',
          'Segnalare immediatamente al datore di lavoro episodi di comportamento aggressivo',
        ],
        dpi: [],
      },
      {
        nome: 'Agenti chimici per la pulizia (alcool etilico denaturato)',
        livello: 'BASSO',
        misure: [
          'Consultare la Scheda di Sicurezza (SdS) prima dell\'utilizzo del prodotto',
          'Utilizzare il prodotto in ambienti aerati e in quantità minime necessarie',
          'Conservare il contenitore chiuso lontano da fonti di calore e fiamme libere',
        ],
        dpi: ['Guanti in nitrile monouso'],
      },
      {
        nome: 'Stress lavoro-correlato',
        livello: 'BASSO',
        misure: [
          'Organizzare il lavoro in modo da evitare carichi eccessivi e ritmi insostenibili',
          'Garantire pause regolari durante il turno lavorativo',
          'Favorire la comunicazione tra lavoratori e datore di lavoro per segnalare situazioni di disagio',
        ],
        dpi: [],
      },
      {
        nome: 'Posture incongrue (stazione eretta prolungata)',
        livello: 'BASSO',
        misure: [
          'Alternare la posizione eretta con brevi pause di seduta o movimento',
          'Utilizzare tappeti antifatica nelle postazioni dove si staziona a lungo in piedi',
          'Eseguire esercizi di stretching durante le pause per ridurre la tensione muscolare',
        ],
        dpi: ['Calzature ergonomiche con supporto plantare'],
      },
    ],
    dpi: ['Calzature antiscivolo/ergonomiche', 'Guanti in nitrile monouso (per pulizie)'],
    quizExtra: [
      {d:'Cosa si deve fare immediatamente se si spilla un liquido sul pavimento del negozio?', r:[
        {lettera:'A', testo:'Pulire subito e posizionare il cartello di pavimento bagnato', corretta:true},
        {lettera:'B', testo:'Continuare il lavoro e rimandare la pulizia a fine turno', corretta:false},
        {lettera:'C', testo:'Segnalarlo solo verbalmente ai colleghi', corretta:false},
        {lettera:'D', testo:'Chiudere temporaneamente il negozio', corretta:false},
      ]},
      {d:'Qual è il comportamento corretto se un cliente diventa verbalmente aggressivo?', r:[
        {lettera:'A', testo:'Rispondere con tono duro per ristabilire l\'autorità', corretta:false},
        {lettera:'B', testo:'Mantenere la calma, non restare soli e avvisare il responsabile', corretta:true},
        {lettera:'C', testo:'Ignorare il cliente e continuare a lavorare', corretta:false},
        {lettera:'D', testo:'Allontanarsi immediatamente senza avvisare nessuno', corretta:false},
      ]},
      {d:'Quale DPI è indicato per la manipolazione dell\'alcool etilico denaturato durante la pulizia?', r:[
        {lettera:'A', testo:'Occhiali di protezione a visiera intera', corretta:false},
        {lettera:'B', testo:'Elmetto di sicurezza', corretta:false},
        {lettera:'C', testo:'Guanti in nitrile monouso', corretta:true},
        {lettera:'D', testo:'Scarpe con puntale in acciaio', corretta:false},
      ]},
      {d:'Dopo quante ore di stazione eretta continuativa è consigliabile effettuare una pausa posturale?', r:[
        {lettera:'A', testo:'Dopo 4 ore senza interruzione', corretta:false},
        {lettera:'B', testo:'Solo a fine turno', corretta:false},
        {lettera:'C', testo:'Mai, il corpo si adatta automaticamente', corretta:false},
        {lettera:'D', testo:'Ogni 60-90 minuti con breve pausa di movimento o seduta', corretta:true},
      ]},
      {d:'L\'alcool etilico denaturato utilizzato per le pulizie è classificato come:', r:[
        {lettera:'A', testo:'Sostanza infiammabile che richiede precauzioni nella conservazione', corretta:true},
        {lettera:'B', testo:'Sostanza inerte e priva di rischi', corretta:false},
        {lettera:'C', testo:'Agente cancerogeno di prima categoria', corretta:false},
        {lettera:'D', testo:'Prodotto alimentare diluito', corretta:false},
      ]},
      {d:'Quale segnale di sicurezza si usa per avvisare del pericolo pavimento bagnato?', r:[
        {lettera:'A', testo:'Segnale di divieto rosso', corretta:false},
        {lettera:'B', testo:'Segnale di avvertimento giallo triangolare con simbolo pavimento scivoloso', corretta:true},
        {lettera:'C', testo:'Segnale di salvataggio verde', corretta:false},
        {lettera:'D', testo:'Nessuna segnaletica è necessaria', corretta:false},
      ]},
      {d:'Cosa prevede la normativa per i lavoratori a contatto con il pubblico rispetto al rischio aggressione?', r:[
        {lettera:'A', testo:'Nessun obbligo specifico poiché il rischio è minimo', corretta:false},
        {lettera:'B', testo:'Solo formazione antincendio', corretta:false},
        {lettera:'C', testo:'Valutazione del rischio e adozione di misure preventive e protettive specifiche', corretta:true},
        {lettera:'D', testo:'Assunzione di personale di sicurezza armato sempre', corretta:false},
      ]},
      {d:'Lo stress lavoro-correlato, secondo l\'Accordo Europeo del 2004, si verifica quando:', r:[
        {lettera:'A', testo:'Il lavoratore è in vacanza', corretta:false},
        {lettera:'B', testo:'Le richieste lavorative superano le capacità del lavoratore di farvi fronte', corretta:false},
        {lettera:'C', testo:'Il lavoratore è in formazione obbligatoria', corretta:false},
        {lettera:'D', testo:'C\'è uno squilibrio tra le richieste dell\'ambiente lavorativo e le risorse del lavoratore', corretta:true},
      ]},
      {d:'Quali sono i principali effetti fisici dell\'assunzione di posture incongrue prolungate?', r:[
        {lettera:'A', testo:'Lombalgie, tendiniti, disturbi muscoloscheletrici', corretta:true},
        {lettera:'B', testo:'Miglioramento della circolazione periferica', corretta:false},
        {lettera:'C', testo:'Aumento della concentrazione', corretta:false},
        {lettera:'D', testo:'Nessun effetto rilevante', corretta:false},
      ]},
      {d:'Come deve essere conservato l\'alcool etilico denaturato in negozio?', r:[
        {lettera:'A', testo:'Vicino al registratore di cassa per averlo a portata di mano', corretta:false},
        {lettera:'B', testo:'In contenitore chiuso, lontano da fonti di calore e in quantità limitata', corretta:true},
        {lettera:'C', testo:'In frigorifero per mantenerne l\'efficacia', corretta:false},
        {lettera:'D', testo:'In qualsiasi posto purché non sia visibile ai clienti', corretta:false},
      ]},
      {d:'Cosa deve fare il lavoratore prima di utilizzare un detergente per la pulizia dei locali?', r:[
        {lettera:'A', testo:'Attendere l\'arrivo dei clienti e poi usarlo', corretta:false},
        {lettera:'B', testo:'Diluirlo sempre con acqua calda senza leggere le istruzioni', corretta:false},
        {lettera:'C', testo:'Consultare la Scheda di Sicurezza e indossare i DPI previsti', corretta:true},
        {lettera:'D', testo:'Nessuna precauzione, i detergenti domestici sono sempre sicuri', corretta:false},
      ]},
      {d:'In caso di contatto cutaneo accidentale con alcool etilico denaturato cosa si deve fare?', r:[
        {lettera:'A', testo:'Applicare ghiaccio sulla zona interessata', corretta:false},
        {lettera:'B', testo:'Ignorare l\'episodio se non compare bruciore', corretta:false},
        {lettera:'C', testo:'Sciacquare abbondantemente con acqua per almeno 15 minuti', corretta:false},
        {lettera:'D', testo:'Lavare la parte con abbondante acqua corrente e consultare le indicazioni della SdS', corretta:true},
      ]},
      {d:'Il mancato utilizzo delle calzature di sicurezza da parte del lavoratore configura:', r:[
        {lettera:'A', testo:'Una violazione degli obblighi del lavoratore sanzionabile', corretta:true},
        {lettera:'B', testo:'Una scelta discrezionale del lavoratore', corretta:false},
        {lettera:'C', testo:'Un obbligo solo per il settore manifatturiero', corretta:false},
        {lettera:'D', testo:'Una responsabilità esclusiva del datore di lavoro', corretta:false},
      ]},
      {d:'Quale misura organizzativa riduce efficacemente il rischio di caduta in piano nel negozio?', r:[
        {lettera:'A', testo:'Aumentare il numero di prodotti esposti sul pavimento', corretta:false},
        {lettera:'B', testo:'Effettuare la pulizia solo dopo la chiusura al pubblico e mantenere i percorsi liberi', corretta:true},
        {lettera:'C', testo:'Illuminare solo le vetrine esterne', corretta:false},
        {lettera:'D', testo:'Vietare l\'accesso ai clienti anziani', corretta:false},
      ]},
      {d:'Quale parte del corpo è maggiormente sollecitata dal lavoro in stazione eretta prolungata?', r:[
        {lettera:'A', testo:'Polsi e mani', corretta:false},
        {lettera:'B', testo:'Solo le spalle', corretta:false},
        {lettera:'C', testo:'Esclusivamente il collo', corretta:false},
        {lettera:'D', testo:'Colonna vertebrale, arti inferiori e piedi', corretta:true},
      ]},
      {d:'Quando si manipolano prodotti chimici per la pulizia, la ventilazione del locale deve essere:', r:[
        {lettera:'A', testo:'Adeguata per disperdere eventuali vapori e garantire ricambio d\'aria', corretta:true},
        {lettera:'B', testo:'Ridotta al minimo per non raffreddare l\'ambiente', corretta:false},
        {lettera:'C', testo:'Irrilevante per prodotti di uso domestico', corretta:false},
        {lettera:'D', testo:'Garantita solo in estate', corretta:false},
      ]},
      {d:'Quale tra le seguenti è una misura preventiva efficace contro il rischio di stress lavoro-correlato?', r:[
        {lettera:'A', testo:'Aumentare il numero di clienti da servire contemporaneamente', corretta:false},
        {lettera:'B', testo:'Eliminare tutte le pause durante il turno', corretta:false},
        {lettera:'C', testo:'Favorire la comunicazione aperta tra lavoratori e responsabili', corretta:false},
        {lettera:'D', testo:'Definire chiaramente i ruoli, garantire supporto e riconoscere il lavoro svolto', corretta:true},
      ]},
    ],
  },
  {
    id: 'AddLabFoto',
    nome: 'Addetto al Laboratorio Fotografico',
    reparto: 'Laboratorio / Magazzino',
    livello: 'MEDIO',
    oreSpec: 8,
    rischi: [
      {
        nome: 'Incidenti meccanici (taglierine manuali, forbici, taglierini)',
        livello: 'MEDIO',
        misure: [
          'Utilizzare le taglierine solo dopo specifica formazione e con la protezione lama abbassata a riposo',
          'Riporre forbici e taglierini in contenitori appositi con la lama protetta',
          'Non distrarsi durante le operazioni di taglio e non eseguire movimenti bruschi',
        ],
        dpi: ['Guanti antitaglio (livello EN 388 adeguato)'],
      },
      {
        nome: 'Caduta dall\'alto (scale semplici metalliche)',
        livello: 'MEDIO',
        misure: [
          'Verificare la stabilità e la conformità della scala prima dell\'uso (certificazione UNI EN 131)',
          'Assicurare la base della scala su superficie piana e antiscivolo durante l\'uso',
          'Non superare il penultimo gradino e non usare la scala da soli per operazioni in quota',
        ],
        dpi: ['Calzature antiscivolo con suola antifatica'],
      },
      {
        nome: 'Esposizione a superfici calde (pressa a caldo, plastificatrice)',
        livello: 'MEDIO',
        misure: [
          'Attendere il completo raffreddamento prima di toccare le superfici di lavoro delle macchine a caldo',
          'Non avvicinarsi alle zone calde con indumenti sintetici o infiammabili',
          'Segnalare con apposita segnaletica le macchine in fase di riscaldamento',
        ],
        dpi: ['Guanti termoresistenti per l\'uso della pressa a caldo'],
      },
      {
        nome: 'Agenti chimici (toner di stampanti, alcool etilico denaturato)',
        livello: 'MEDIO',
        misure: [
          'Sostituire le cartucce toner esaurite seguendo le istruzioni del produttore e smaltirle correttamente',
          'In caso di fuoriuscita accidentale di toner rimuoverlo con panno umido, mai con aria compressa',
          'Lavare le mani dopo la rimozione dei guanti e dopo ogni contatto con agenti chimici',
        ],
        dpi: ['Guanti in nitrile monouso', 'Mascherina FFP1 in caso di sostituzione toner'],
      },
      {
        nome: 'Lavoro al videoterminale (PC, stampanti, scanner)',
        livello: 'BASSO',
        misure: [
          'Posizionare il monitor a circa 50-70 cm dagli occhi e regolare luminosità e contrasto',
          'Effettuare pause di almeno 15 minuti ogni 2 ore di utilizzo continuativo del VDT',
          'Regolare la sedia e il piano di lavoro per mantenere una postura corretta',
        ],
        dpi: [],
      },
    ],
    dpi: [
      'Guanti antitaglio EN 388 (per taglierine)',
      'Guanti termoresistenti (per pressa a caldo)',
      'Guanti in nitrile monouso (per agenti chimici)',
      'Mascherina FFP1 (per sostituzione toner)',
      'Calzature antiscivolo',
    ],
    quizExtra: [
      {d:'Quale norma tecnica disciplina la conformità e il collaudo delle scale portatili?', r:[
        {lettera:'A', testo:'UNI EN 131 parte 1 e 2', corretta:true},
        {lettera:'B', testo:'ISO 9001', corretta:false},
        {lettera:'C', testo:'EN 13501', corretta:false},
        {lettera:'D', testo:'UNI 8457', corretta:false},
      ]},
      {d:'Durante l\'uso della taglierina manuale, dove devono essere posizionate le dita dell\'operatore?', r:[
        {lettera:'A', testo:'Sulla lama per guidare il taglio con maggiore precisione', corretta:false},
        {lettera:'B', testo:'Lontano dalla traiettoria della lama, sul bordo opposto al taglio', corretta:true},
        {lettera:'C', testo:'Sopra il materiale da tagliare per tenerlo fermo', corretta:false},
        {lettera:'D', testo:'La posizione delle dita è irrilevante se si usa la protezione', corretta:false},
      ]},
      {d:'Dopo l\'utilizzo della plastificatrice a caldo, quando è sicuro toccare il rullo?', r:[
        {lettera:'A', testo:'Immediatamente dopo lo spegnimento', corretta:false},
        {lettera:'B', testo:'Dopo 1 minuto dal termine dell\'ultima plastificazione', corretta:false},
        {lettera:'C', testo:'Solo dopo che la macchina ha raggiunto la temperatura ambiente', corretta:true},
        {lettera:'D', testo:'Il rullo non raggiunge mai temperature pericolose', corretta:false},
      ]},
      {d:'Le cartucce toner esaurite delle stampanti devono essere:', r:[
        {lettera:'A', testo:'Smaltite nel cestino dei rifiuti indifferenziati', corretta:false},
        {lettera:'B', testo:'Aperte e svuotate prima dello smaltimento', corretta:false},
        {lettera:'C', testo:'Conservate all\'infinito in magazzino', corretta:false},
        {lettera:'D', testo:'Smaltite secondo le indicazioni del produttore e la normativa sui rifiuti speciali', corretta:true},
      ]},
      {d:'Qual è il principale rischio nell\'uso della pressa a caldo per sublimazione?', r:[
        {lettera:'A', testo:'Ustioni per contatto con superfici a elevata temperatura', corretta:true},
        {lettera:'B', testo:'Elettrocuzione per tensione eccessiva', corretta:false},
        {lettera:'C', testo:'Caduta della pressa per peso eccessivo', corretta:false},
        {lettera:'D', testo:'Inalazione di gas tossici ad alta concentrazione', corretta:false},
      ]},
      {d:'Quante persone devono essere presenti durante l\'uso di una scala portatile per operazioni in quota?', r:[
        {lettera:'A', testo:'Una sola persona, per evitare ingombri', corretta:false},
        {lettera:'B', testo:'Almeno due: una che sale e una che stabilizza la base', corretta:true},
        {lettera:'C', testo:'Il numero non è rilevante se la scala è certificata', corretta:false},
        {lettera:'D', testo:'Tre persone obbligatoriamente', corretta:false},
      ]},
      {d:'Cosa si intende per "zona calda" nella plastificatrice a caldo?', r:[
        {lettera:'A', testo:'Il cavo di alimentazione durante il funzionamento', corretta:false},
        {lettera:'B', testo:'Il pannello di controllo con i tasti', corretta:false},
        {lettera:'C', testo:'Il vassoio di raccolta dei fogli plastificati', corretta:false},
        {lettera:'D', testo:'I rulli di laminazione che raggiungono temperature superiori ai 100°C', corretta:true},
      ]},
      {d:'Come si devono riporre le forbici dopo l\'uso in laboratorio?', r:[
        {lettera:'A', testo:'Appoggiate sulla superficie del banco con la lama aperta', corretta:false},
        {lettera:'B', testo:'In un contenitore apposito con la punta rivolta verso il basso', corretta:false},
        {lettera:'C', testo:'Chiuse e riposte in un contenitore o portautensili dedicato', corretta:true},
        {lettera:'D', testo:'Lasciate sul bordo del banco per averle sempre a portata di mano', corretta:false},
      ]},
      {d:'In caso di taglio accidentale con una taglierina, qual è la prima azione da compiere?', r:[
        {lettera:'A', testo:'Continuare il lavoro applicando un cerotto in autonomia senza avvisare', corretta:false},
        {lettera:'B', testo:'Applicare pressione sulla ferita, fermare l\'emorragia e segnalare l\'infortunio al responsabile', corretta:true},
        {lettera:'C', testo:'Lavare la ferita con alcool etilico denaturato e riprendere subito il lavoro', corretta:false},
        {lettera:'D', testo:'Togliere i guanti e valutare l\'entità del danno senza interrompere la produzione', corretta:false},
      ]},
      {d:'Qual è la distanza raccomandata tra gli occhi e lo schermo del monitor per ridurre l\'affaticamento visivo?', r:[
        {lettera:'A', testo:'10-20 cm', corretta:false},
        {lettera:'B', testo:'50-70 cm', corretta:true},
        {lettera:'C', testo:'Oltre 150 cm', corretta:false},
        {lettera:'D', testo:'La distanza è irrilevante se lo schermo è di ultima generazione', corretta:false},
      ]},
      {d:'L\'esposizione cronica a polvere di toner di stampanti può provocare principalmente:', r:[
        {lettera:'A', testo:'Ustioni superficiali alle mani', corretta:false},
        {lettera:'B', testo:'Sordità professionale', corretta:false},
        {lettera:'C', testo:'Irritazione delle vie respiratorie e potenziale rischio per la salute dei polmoni', corretta:true},
        {lettera:'D', testo:'Danni esclusivamente alle unghie', corretta:false},
      ]},
      {d:'Quale DPI è specificamente indicato per l\'utilizzo della pressa a caldo?', r:[
        {lettera:'A', testo:'Guanti antitaglio in fibra di acciaio', corretta:false},
        {lettera:'B', testo:'Cuffie antirumore', corretta:false},
        {lettera:'C', testo:'Elmetto di protezione', corretta:false},
        {lettera:'D', testo:'Guanti termoresistenti omologati per il calore', corretta:true},
      ]},
      {d:'Ogni quanto tempo è raccomandata una pausa durante l\'uso continuativo del videoterminale?', r:[
        {lettera:'A', testo:'15 minuti ogni 2 ore di lavoro continuativo al VDT', corretta:true},
        {lettera:'B', testo:'Una pausa giornaliera di 5 minuti a metà turno', corretta:false},
        {lettera:'C', testo:'Solo a fine giornata lavorativa', corretta:false},
        {lettera:'D', testo:'Non è prevista alcuna pausa specifica per il VDT', corretta:false},
      ]},
      {d:'In caso di inalazione accidentale di vapori di alcool etilico denaturato, cosa si deve fare?', r:[
        {lettera:'A', testo:'Continuare il lavoro aprendo una finestra', corretta:false},
        {lettera:'B', testo:'Allontanarsi immediatamente nell\'aria fresca e consultare la SdS; se i sintomi persistono, richiedere assistenza medica', corretta:true},
        {lettera:'C', testo:'Bere acqua abbondante per neutralizzare l\'effetto', corretta:false},
        {lettera:'D', testo:'Inspirare profondamente per accelerare lo smaltimento', corretta:false},
      ]},
      {d:'Come si trasporta correttamente una scala portatile metallica in laboratorio?', r:[
        {lettera:'A', testo:'Verticalmente con un braccio solo per velocizzare gli spostamenti', corretta:false},
        {lettera:'B', testo:'Orizzontalmente con entrambe le mani, prestando attenzione agli ingombri laterali', corretta:true},
        {lettera:'C', testo:'Trascinandola sul pavimento per non affaticarsi', corretta:false},
        {lettera:'D', testo:'Il trasporto della scala non richiede precauzioni particolari', corretta:false},
      ]},
      {d:'Qual è la caratteristica di sicurezza principale di un taglierino (cutter) idoneo all\'uso lavorativo?', r:[
        {lettera:'A', testo:'Lama estraibile a scatto con meccanismo di blocco nella posizione di lavoro', corretta:false},
        {lettera:'B', testo:'Lama fissa senza alcun sistema di protezione', corretta:false},
        {lettera:'C', testo:'Lama con sistema di ritiro automatico o blocco sicuro e impugnatura antiscivolo', corretta:true},
        {lettera:'D', testo:'Colore brillante del manico per una facile visibilità', corretta:false},
      ]},
      {d:'Il "rischio meccanico" legato alle taglierine manuali include principalmente:', r:[
        {lettera:'A', testo:'Solo il rischio di rottura della lama per materiali troppo duri', corretta:false},
        {lettera:'B', testo:'Esclusivamente il rischio di rumore durante il funzionamento', corretta:false},
        {lettera:'C', testo:'Solo il rischio di danneggiamento del materiale lavorato', corretta:false},
        {lettera:'D', testo:'Tagli, abrasioni e schiacciamenti alle mani durante l\'operazione di taglio', corretta:true},
      ]},
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
