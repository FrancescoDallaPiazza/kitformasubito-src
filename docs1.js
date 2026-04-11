'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  TabStopType, SimpleField, LevelFormat,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, A4_L, MARGIN_STD,
  makeHeader, makeFooter, vuoto, cella, salvaDoc, logoBytes,
} = h;

// Import aggiuntivi
const { PageBreak, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType, TextWrappingSide } = require('docx');

const OUT = `/home/claude/kit/OUT/KIT FORMASUBITO - ${CLIENTE.ragioneSocialeBreve}`;

// ── helpers locali ─────────────────────────────────────────────────────────
const BD = {
  top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
  left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
};
const NO = {
  top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
  left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
};

// paragrafo normale giustificato
function N(testo, opts = {}) {
  return new Paragraph({
    alignment: opts.align !== undefined ? opts.align : AlignmentType.JUSTIFIED,
    spacing: { before: opts.spB !== undefined ? opts.spB*20 : 0, after: opts.spA !== undefined ? opts.spA*20 : 0 },
    children: [new TextRun({
      text: testo, font: FONT,
      size: opts.sz ? opts.sz*2 : undefined,
      bold: opts.bold || false,
      color: opts.col || undefined,
    })],
  });
}

// paragrafo con page break incorporato (come nel master)
function NBrk(testo, opts = {}) {
  const runs = [];
  if (testo) runs.push(new TextRun({ text: testo, font: FONT, size: opts.sz ? opts.sz*2 : undefined, bold: opts.bold, color: opts.col }));
  runs.push(new PageBreak());
  return new Paragraph({
    style: opts.style || 'Normal',
    alignment: opts.align !== undefined ? opts.align : AlignmentType.JUSTIFIED,
    spacing: { before: opts.spB !== undefined ? opts.spB*20 : 0, after: opts.spA !== undefined ? opts.spA*20 : 0 },
    children: runs,
  });
}

// paragrafo lista con bullet vero (numbering)
function LP(testo, opts = {}) {
  const runs = [new TextRun({ text: testo, font: FONT, size: 20 })];
  if (opts.pageBreak) runs.push(new PageBreak());
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 4, after: 4 },
    children: runs,
  });
}

// titolo sezione: bold sz=13 col=1F3864 spB=14 spA=6
function SEC(n, testo, opts = {}) {
  const runs = [new TextRun({ text: `${n}. ${testo}`, bold: true, font: FONT, size: 26, color: C.BLU_DARK })];
  if (opts.pageBreak) runs.push(new PageBreak());
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 14*20, after: 6*20 },
    children: runs,
  });
}

// sottotitolo: bold col=1F4E79 sz=default
function SUB(testo) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: testo, bold: true, font: FONT, color: C.BLU_HEADER })],
  });
}

// subsection 4.x: bold sz=12 col=1F4E79
function SUBSEC(testo) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: testo, bold: true, font: FONT, size: 24, color: C.BLU_HEADER })],
  });
}

// "Durata:..." in rosso bold sz=10
function DUR(testo) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: testo, bold: true, font: FONT, size: 20, color: C.ROSSO })],
  });
}

// "Modulo Specifico" bold sz=10.5 col=1F3864 spB=2 spA=2
function MOD(testo) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 4, after: 4 },
    children: [new TextRun({ text: testo, bold: true, font: FONT, size: 21, color: C.BLU_DARK })],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// PROGETTO FORMATIVO
// ─────────────────────────────────────────────────────────────────────────────
async function genProgettoFormativo() {
  // Margine T=3.75cm dal master
  const MARGIN = { top: 2126, right: 1134, bottom: 1134, left: 1134 };
  const W = 9638;
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'Progetto Formativo Aziendale', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('');

  // Livelli distinti tra le mansioni
  const livelli = [...new Set(MANSIONI.map(m => m.livello))];
  const livFill = { BASSO: C.VERDE, MEDIO: C.ARANCIO, ALTO: C.ROSSO };

  // ── SUMMARY TABLE 2×2 ───────────────────────────────────────────────────
  // Cella [1,0]: "Livello di rischio:" (1F4E79) + livello(i) + ore (C00000)
  function livelloCellChildren() {
    const children = [
      new Paragraph({ children: [new TextRun({ text: 'Livello di rischio:', bold: true, font: FONT, color: C.BLU_HEADER })] }),
    ];
    livelli.forEach(liv => {
      const m = MANSIONI.find(m2 => m2.livello === liv);
      children.push(new Paragraph({ children: [new TextRun({ text: liv, bold: true, font: FONT, color: C.ROSSO })] }));
      children.push(new Paragraph({ children: [new TextRun({ text: `(4 ore formazione generale + ${m.oreSpec} ore formazione specifica)`, bold: true, font: FONT, size: 18, color: C.ROSSO })] }));
    });
    return children;
  }

  function cellKV(label, value, fill) {
    return new TableCell({
      width: { size: Math.floor(W/2), type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      borders: BD,
      children: [
        new Paragraph({ children: [new TextRun({ text: label, bold: true, font: FONT, color: C.BLU_HEADER })] }),
        new Paragraph({ children: [new TextRun({ text: value, font: FONT, size: 20 })] }),
      ],
    });
  }

  const summaryTable = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [Math.floor(W/2), W - Math.floor(W/2)],
    borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: BD.top, insideV: BD.top },
    rows: [
      new TableRow({ children: [
        cellKV('Datore di Lavoro / RSPP:', CLIENTE.datoreLavoro, C.BLU_LIGHT),
        cellKV('Codice ATECO:', `${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`, C.BLU_LIGHT),
      ]}),
      new TableRow({ children: [
        new TableCell({
          width: { size: Math.floor(W/2), type: WidthType.DXA },
          shading: { fill: C.BIANCO, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          borders: BD,
          children: livelloCellChildren(),
        }),
        cellKV('Anno:', CLIENTE.anno, C.BIANCO),
      ]}),
    ],
  });

  // ── CORSO INFO TABLE 11×2 ─────────────────────────────────────────────
  // Righe alternate: key=D5E8F0/EBF3FB, value=FFFFFF/F2F2F2
  const wKey = 3000; const wVal = W - wKey;
  const corsoRows = [
    ['Soggetto formatore', CLIENTE.ragioneSociale],
    ['Codice ATECO', `${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`],
    ['P.IVA / C.F.', CLIENTE.piva],
    ['Sede', CLIENTE.indirizzo],
    ['Datore di Lavoro / RSPP', CLIENTE.datoreLavoro],
    ['Livello di rischio', livelli.map(l => { const m = MANSIONI.find(m2=>m2.livello===l); return `${l} (ATECO ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc})`; }).join(' / ')],
    ['Ore formazione specifica', livelli.map(l => { const m = MANSIONI.find(m2=>m2.livello===l); return `${m.oreSpec} ore (Rischio ${l})`; }).join(' / ')],
    ['Numero max partecipanti per sessione', '30'],
    ['Soglia di presenza minima', '90% delle ore previste'],
    ['Metodologia didattica', 'Lezioni frontali sul campo ed esempi pratici'],
    ['Verifica finale', 'Test a risposta multipla – superamento con almeno 70% di risposte corrette'],
  ];
  const keyFills = ['D5E8F0','EBF3FB'];
  const valFills = ['FFFFFF','F2F2F2'];
  const corsoTable = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wKey, wVal],
    borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: BD.top, insideV: BD.top },
    rows: corsoRows.map(([k, v], i) => new TableRow({ children: [
      new TableCell({
        width: { size: wKey, type: WidthType.DXA },
        shading: { fill: keyFills[i%2], type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        borders: BD,
        children: [new Paragraph({ children: [new TextRun({ text: k, bold: true, font: FONT, color: C.BLU_HEADER })] })],
      }),
      new TableCell({
        width: { size: wVal, type: WidthType.DXA },
        shading: { fill: valFills[i%2], type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        borders: BD,
        children: [new Paragraph({ children: [new TextRun({ text: v, font: FONT, size: 20 })] })],
      }),
    ]})),
  });

  // ── MANSIONI TABLE 3×4 ───────────────────────────────────────────────
  const wM = 3680; const wR = 2980; const wL2 = 1840; const wO = W - wM - wR - wL2;
  const mansioniTable = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wM, wR, wL2, wO],
    borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: BD.top, insideV: BD.top },
    rows: [
      new TableRow({ tableHeader: true, children: [
        cella('MANSIONE', { width: wM, bold: true, fill: C.BLU_HEADER, color: C.BIANCO }),
        cella('REPARTO', { width: wR, bold: true, fill: C.BLU_HEADER, color: C.BIANCO }),
        cella('LIVELLO RISCHIO', { width: wL2, bold: true, fill: C.BLU_HEADER, color: C.BIANCO, align: 'center' }),
        cella('ORE TOT.', { width: wO, bold: true, fill: C.BLU_HEADER, color: C.BIANCO, align: 'center' }),
      ]}),
      ...MANSIONI.map(m => new TableRow({ children: [
        cella(m.nome, { width: wM }),
        cella(m.reparto, { width: wR }),
        cella(m.livello, { width: wL2, bold: true, fill: livFill[m.livello] || C.VERDE, color: C.BIANCO, align: 'center' }),
        cella(`${m.oreSpec + 4}h`, { width: wO, align: 'center' }),
      ]})),
    ],
  });

  // ── SOGGETTI TABLES ─────────────────────────────────────────────────────
  function tKV2(righe) {
    return new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [wKey, wVal],
      borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: BD.top, insideV: BD.top },
      rows: righe.map(([k, v], i) => new TableRow({ children: [
        new TableCell({
          width: { size: wKey, type: WidthType.DXA },
          shading: { fill: C.BLU_LIGHT, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 80, right: 80 }, borders: BD,
          children: [new Paragraph({ children: [new TextRun({ text: k, bold: true, font: FONT, color: C.BLU_HEADER })] })],
        }),
        new TableCell({
          width: { size: wVal, type: WidthType.DXA },
          shading: { fill: C.BIANCO, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 80, right: 80 }, borders: BD,
          children: [new Paragraph({ children: [new TextRun({ text: v, font: FONT, size: 20 })] })],
        }),
      ]})),
    });
  }

  // ── NOTA BENE TABLE ──────────────────────────────────────────────────────
  const notaBeneTable = new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: NO.top, insideV: NO.top },
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: W, type: WidthType.DXA },
        shading: { fill: C.BLU_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 }, borders: BD,
        children: [
          new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: 'NOTA BENE', bold: true, font: FONT, color: C.BLU_DARK })] }),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 20 }, children: [new TextRun({ text: 'In coerenza con le previsioni di cui all\'articolo 37, comma 12, del D.Lgs. n. 81/08, i corsi di formazione vanno realizzati previa richiesta di collaborazione agli organismi paritetici di cui al Decreto del Ministro del Lavoro e delle Politiche Sociali dell\'11 ottobre 2022, n. 171, ove presenti nel settore e nel territorio in cui si svolge l\'attività del datore.', font: FONT, size: 20 })] }),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 20 }, children: [new TextRun({ text: 'In mancanza, il datore di lavoro procede alla pianificazione e realizzazione delle attività di formazione. Ove la richiesta riceva riscontro da parte dell\'organismo paritetico, delle relative indicazioni occorre tener conto nella pianificazione e realizzazione delle attività di formazione.', font: FONT, size: 20 })] }),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 0 }, children: [new TextRun({ text: 'Ove la richiesta di cui al precedente periodo non riceva riscontro dall\'organismo paritetico entro quindici giorni dal suo invio, il datore di lavoro procede autonomamente alla pianificazione e realizzazione delle attività di formazione.', font: FONT, size: 20 })] }),
        ],
      }),
    ]})]
  });

  // ── CORPO COMPLETO ────────────────────────────────────────────────────────
  // Struttura esatta del master (con page break agli indici corretti)
  const children = [
    // Info azienda nel corpo (come nel master)
    N(CLIENTE.ragioneSociale, { bold: true, sz: 18, col: C.BLU_DARK, spA: 4 }),
    N(CLIENTE.indirizzo, { col: C.GRIGIO, spA: 4 }),
    N(`P.IVA: ${CLIENTE.piva}`, { col: C.GRIGIO, spA: 4 }),
    N(`Codice Fiscale: ${CLIENTE.piva}`, { col: C.GRIGIO, spA: 4 }),
    new Paragraph({ children: [] }), // empty [4]

    // Title table 1×1
    new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [W],
      borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: NO.top, insideV: NO.top },
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: W, type: WidthType.DXA },
          shading: { fill: C.BLU_DARK, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 }, borders: BD,
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: 'PROGETTO FORMATIVO AZIENDALE', bold: true, font: FONT, size: 26, color: C.BIANCO })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: 'Formazione sulla Salute e Sicurezza sul Lavoro', font: FONT, size: 22, color: 'D5E8F0' })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: 'D.Lgs. 81/2008 e s.m.i. – Accordo Stato-Regioni 17/04/2025', font: FONT, size: 20, color: 'D5E8F0' })] }),
          ],
        }),
      ]})]
    }),
    new Paragraph({ children: [] }), // empty [6]

    // Summary table 2×2
    summaryTable,
    new Paragraph({ children: [] }), // empty [8]

    // PAGE BREAK [9] — tra copertina e sezioni
    new Paragraph({ children: [new PageBreak()] }),

    // ── SEZIONE 1 ─────────────────────────────────────────────────────────
    N('1. PREFAZIONE AL PROGETTO FORMATIVO', { bold: true, sz: 13, col: C.BLU_DARK, spB: 14, spA: 6 }),
    SUB('Formazione del Datore di Lavoro che svolge il ruolo di RSPP'),
    new Paragraph({ children: [] }),
    N('Il presente Progetto Formativo Aziendale definisce in modo strutturato e coerente il percorso di formazione in materia di salute e sicurezza sul lavoro adottato dal Datore di Lavoro che svolge direttamente il ruolo di Responsabile del Servizio di Prevenzione e Protezione (RSPP).', { sz: 10 }),
    new Paragraph({ children: [] }),
    N('La progettazione della formazione è stata sviluppata in conformità al D.Lgs. 81/2008 e s.m.i., nonché agli indirizzi introdotti dall\'Accordo Stato-Regioni del 17 aprile 2025 (entrato in vigore il 24 maggio 2025), che rafforzano il principio secondo cui la formazione non deve essere considerata un adempimento formale, ma uno strumento operativo e funzionale alla gestione reale dei rischi aziendali.', { sz: 10 }),
    new Paragraph({ children: [] }),
    N('In tale contesto, il Datore di Lavoro assume un ruolo centrale non solo come destinatario della formazione, ma anche come soggetto responsabile della pianificazione, organizzazione e verifica del processo formativo, garantendo che i contenuti siano coerenti con:', { sz: 10 }),
    LP('la specifica attività aziendale;'),
    LP('i rischi effettivamente presenti nei luoghi di lavoro;'),
    LP('l\'organizzazione interna e le modalità operative adottate.'),
    new Paragraph({ children: [] }),
    N('Il presente progetto ha lo scopo di:', { sz: 10 }),
    LP('collegare la formazione alla Valutazione dei Rischi, rendendo il percorso uno strumento integrato al sistema di prevenzione aziendale;'),
    LP('garantire la tracciabilità delle attività formative, delle verifiche di apprendimento e degli aggiornamenti periodici;'),
    LP('dimostrare l\'effettiva coerenza tra formazione erogata, competenze acquisite e gestione operativa della sicurezza.'),
    new Paragraph({ children: [] }),
    N('La formazione viene progettata con un approccio pratico e applicativo, privilegiando esempi concreti, analisi di casi reali aziendali e collegamenti diretti sul campo, con le procedure operative adottate.', { sz: 10 }),
    new Paragraph({ children: [] }),
    N('Il presente Progetto Formativo Aziendale rappresenta pertanto uno strumento dinamico, soggetto ad aggiornamento periodico in relazione:', { sz: 10 }),
    LP('all\'evoluzione normativa;'),
    LP('alle modifiche organizzative o produttive dell\'azienda;'),
    LP('all\'esito delle verifiche di efficacia della formazione e degli eventi infortunistici o dei mancati infortuni.'),
    new Paragraph({ children: [] }),
    N('Attraverso questo progetto, l\'azienda intende dimostrare una gestione consapevole e responsabile della formazione, orientata alla prevenzione reale dei rischi e alla tutela della salute e sicurezza dei lavoratori, nel rispetto del principio di miglioramento continuo e una costante presenza del Datore di lavoro RSPP negli interventi formativi e addestrativi del proprio personale.', { sz: 10 }),
    new Paragraph({ children: [] }),

    // PAGE BREAK [36] — paragrafo vuoto con pagebreak e spaziatura 14/6
    new Paragraph({
      spacing: { before: 14*20, after: 6*20 },
      children: [new PageBreak()],
    }),

    // ── SEZIONE 2 ─────────────────────────────────────────────────────────
    N('2. RIFERIMENTO NORMATIVO', { bold: true, sz: 13, col: C.BLU_DARK, spB: 14, spA: 6 }),
    SUB('Accordo Stato-Regioni del 17 aprile 2025 – Parte II, Punto 2 e Parte IV, Punto 1'),
    new Paragraph({ children: [] }),
    // Paragrafo 1: "Parte II dell'Accordo – Punto 2" sottolineato + testo in corsivo
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: 'Parte II dell\'Accordo – Punto 2', font: FONT, size: 20, color: '000000', underline: { type: 'single' } }),
        new TextRun({ text: ': ', font: FONT, size: 20, color: '000000' }),
        new TextRun({ text: 'i datori di lavoro possono organizzare direttamente i corsi di formazione ex art. 37, comma 2, del D.Lgs. n. 81/2008 nei confronti dei propri lavoratori, preposti e dirigenti, a condizione che venga rispettato quanto previsto dal presente Accordo. In questo caso il datore di lavoro riveste il ruolo di soggetto formatore cui spettano gli adempimenti del presente accordo', font: FONT, size: 20, color: '000000', italics: true }),
        new TextRun({ text: '.', font: FONT, size: 20, color: '000000' }),
      ],
    }),
    new Paragraph({ children: [] }),
    // Paragrafo 2: tutto in corsivo
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: 'Il datore di lavoro in possesso dei requisiti per lo svolgimento diretto dei compiti del servizio di prevenzione e protezione di cui all\'articolo 34 del D.Lgs. n. 81/2008, può svolgere anche in qualità di docente, esclusivamente nei riguardi dei propri lavoratori, preposti e dirigenti, la formazione di cui ai paragrafi: 2.1, 2.2 e 2.3.', font: FONT, size: 20, color: '000000', italics: true }),
      ],
    }),
    new Paragraph({ children: [] }),
    // Paragrafo 3: "Parte IV dell'Accordo" sottolineato + " – " normale + testo in corsivo
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 0 },
      children: [
        new TextRun({ text: 'Parte IV dell\'Accordo', font: FONT, size: 20, color: '000000', underline: { type: 'single' } }),
        new TextRun({ text: ' – ', font: FONT, size: 20, color: '000000' }),
        new TextRun({ text: 'Punto 1: le indicazioni metodologiche per l\'organizzazione e la gestione dei corsi, fatta eccezione dei punti 3.2, 3.3, 3.4, 3.5, 6.3 e 7, non si applicano ai Datori di Lavoro che organizzano ed erogano autonomamente, all\'interno delle proprie aziende, la formazione sulla salute e sicurezza sul lavoro, ma esse possono trovare indicazioni utili per la gestione dei percorsi formativi di cui al presente accordo.', font: FONT, size: 20, color: '000000', italics: true }),
      ],
    }),
    new Paragraph({ children: [] }),

    // ── SEZIONE 3 ─────────────────────────────────────────────────────────
    N('3. DATI IDENTIFICATIVI DEL CORSO', { bold: true, sz: 13, col: C.BLU_DARK, spB: 14, spA: 6 }),
    new Paragraph({ children: [] }),
    corsoTable,
    new Paragraph({ children: [] }),
    mansioniTable,
    new Paragraph({ children: [] }),
    new Paragraph({ children: [] }),
    new Paragraph({ children: [] }),

    // ── SEZIONE 4 ─────────────────────────────────────────────────────────
    N('4. STRUTTURA DEL CORSO', { bold: true, sz: 13, col: C.BLU_DARK, spB: 14, spA: 6 }),
    new Paragraph({ children: [] }),

    SUBSEC('4.1 – Formazione Generale'),
    new Paragraph({ children: [] }),
    DUR('Durata: 4 ore – Costituisce credito formativo permanente'),
    new Paragraph({ children: [] }),
    N('Sono trattati i seguenti argomenti:', { sz: 10 }),
    new Paragraph({ children: [] }),
    LP('Concetti di pericolo, rischio e danno (1 ora)'),
    LP('Prevenzione e protezione (1 ora)'),
    LP('Organizzazione della prevenzione aziendale e sistema di partecipazione dei lavoratori (1 ora)'),
    LP('Diritti, doveri e sanzioni dei vari soggetti aziendali; organi di vigilanza, controllo e assistenza (1 ora)'),
    new Paragraph({ children: [] }),
    new Paragraph({ children: [] }),

    SUBSEC('4.2 – Formazione Specifica dei Rischi'),
    new Paragraph({ children: [] }),
    // Per ogni livello di rischio distinto tra le mansioni
    ...(() => {
      const result = [];
      livelli.forEach(livello => {
        const mansioniLiv = MANSIONI.filter(m => m.livello === livello);
        const oreSpec = mansioniLiv[0].oreSpec;
        const rischiUnici = [...new Set(mansioniLiv.flatMap(m => m.rischi.map(r => r.nome)))];
        // Riga mansione/i prima della durata (solo se ci sono più livelli distinti)
        if (livelli.length > 1) {
          const nomiMansioni = mansioniLiv.map(m => m.nome.toUpperCase()).join(' / ');
          result.push(new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { before: 6*20, after: 0 },
            children: [new TextRun({ text: `Mansione: ${nomiMansioni}`, bold: true, font: FONT, size: 20, color: C.ROSSO })],
          }));
        }
        result.push(DUR(`Durata: ${oreSpec} ore – Settore a rischio ${livello} (ATECO ${CLIENTE.atecoCodice})`));
        result.push(new Paragraph({ children: [] }));
        result.push(N('La formazione specifica è mirata ai rischi effettivamente presenti nel luogo di lavoro, identificati dalla valutazione dei rischi aziendale (DVR). Per ogni mansione sono trattati i rischi specifici della postazione lavorativa:', { sz: 10 }));
        result.push(new Paragraph({ children: [] }));
        result.push(MOD(`Modulo Specifico – Rischio ${livello} (${oreSpec} ore)`));
        rischiUnici.forEach(r => result.push(LP(r)));
        result.push(LP('DPI: tipologie, scelta, uso e manutenzione'));
        result.push(LP('Segnaletica di sicurezza'));
        result.push(LP('Prevenzione incendi ed evacuazione'));
        result.push(LP('Segnaletica'));
        result.push(LP('Procedure organizzative per il primo soccorso'));
        result.push(LP('Stress lavoro correlato'));
        result.push(LP('Incidenti, mancati infortuni'));
        result.push(new Paragraph({ children: [] }));
      });
      return result;
    })(),

    SUBSEC('4.3 – Aggiornamento della formazione specifica'),
    new Paragraph({ children: [] }),
    DUR('Durata: 6 ore ogni 5 anni (Accordo Stato-Regioni 17/04/2025, Parte III).'),
    new Paragraph({ children: [] }),
    N('Modalità: colloquio individuale o test scritto a risposta multipla.', { sz: 10 }),
    new Paragraph({ children: [] }),
    N('Contenuti:', { sz: 10 }),
    LP('aggiornamento sui rischi specifici'),
    LP('nuove normative'),
    // PAGE BREAK [97] — nell'ultimo elemento della lista 4.3
    new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 4, after: 4 },
      children: [
        new TextRun({ text: 'cambiamenti organizzativi/produttivi.', font: FONT, size: 20 }),
        new PageBreak(),
      ],
    }),

    // ── SEZIONE 5 ─────────────────────────────────────────────────────────
    N('5. SOGGETTI FORMATORI E DOCENTI', { bold: true, sz: 13, col: C.BLU_DARK, spB: 14, spA: 6 }),
    new Paragraph({ children: [] }),
    tKV2([
      ['Soggetto formatore', `${CLIENTE.datoreLavoro} – Datore di Lavoro e RSPP`],
      ['Base normativa', 'ASR 17/04/2025, Punto 2 – Parte II'],
    ]),
    new Paragraph({ children: [] }),
    tKV2([
      ['Soggetto relatore / docente', `${CLIENTE.datoreLavoro} – Datore di Lavoro e RSPP`],
      ['Base normativa docente', 'ASR 17/04/2025, Punto 2 – Parte II (deroga per Datore di Lavoro RSPP)'],
    ]),
    new Paragraph({ children: [] }),

    // ── SEZIONE 6 ─────────────────────────────────────────────────────────
    N('6. NOTA BENE', { bold: true, sz: 13, col: C.BLU_DARK, spB: 14, spA: 6 }),
    new Paragraph({ children: [] }),
    notaBeneTable,
    new Paragraph({ children: [] }),
  ];

  const doc = new Document({
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: '●',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      }],
    },
    styles: docStyles,
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: MARGIN } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });
  await salvaDoc(doc, `${OUT}/00 - PROGETTO FORMATIVO/ProgettoFormativo.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// REGISTRO PRESENZE FORMAZIONE INIZIALE
// ─────────────────────────────────────────────────────────────────────────────
async function genRegistroFormIniziale(mansione) {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 709, footer: 709 };
  const W = 14570;

  // ── Bordi AAAAAA (come nel master) ────────────────────────────────────────
  const BAA = {
    top:    { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
    bottom: { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
    left:   { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
    right:  { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
  };
  // Bordi intestazione presenze (1F4E79)
  const BHDR = {
    top:    { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
    bottom: { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
    left:   { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
    right:  { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
  };

  // ── Header inline: logo sinistra + ragione sociale destra ────────────────
  const NO_HDR = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} };
  const header = new Header({
    children: [new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [2400, W - 2400],
      borders: NO_HDR,
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 2400, type: WidthType.DXA },
          borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 140, height: 32 } })] })],
        }),
        new TableCell({
          width: { size: W - 2400, type: WidthType.DXA },
          borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: CLIENTE.ragioneSociale, bold: true, font: FONT, size: 20, color: C.BLU_DARK })] })],
        }),
      ]})],
    })],
  });

  // ── Footer con bordo superiore blu ────────────────────────────────────────
  const footer = new Footer({
    children: [new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
      children: [
        new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   Pag. `, size: 16, font: FONT, color: C.GRIGIO }),
        new SimpleField('PAGE'),
      ],
    })],
  });

  // ── Helper cella con bordi AAAAAA e margini precisi ───────────────────────
  function cellaAA(children, opts = {}) {
    return new TableCell({
      width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
      borders: BAA,
      margins: { top: opts.mTB || 100, bottom: opts.mTB || 100, left: opts.mLR || 150, right: opts.mLR || 150 },
      verticalAlign: VerticalAlign.CENTER,
      children: Array.isArray(children) ? children : [children],
    });
  }

  // ── Helper cella dati presenze (bordi AAAAAA, margini ridotti) ────────────
  function cellaDato(w) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: BAA,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({ children: [new TextRun({ text: ' ' })] })],
    });
  }

  // ── Helper cella intestazione presenze (fill 1F4E79, testo bianco) ────────
  function cellaHdr(w, testo) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: BHDR,
      shading: { fill: '1F4E79', type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: testo, bold: true, font: FONT, size: 18, color: 'FFFFFF' })],
      })],
    });
  }

  const colW = [2025, 2004, 3196, 708, 3402, 709, 2526];
  const wL = 6964; const wR = 7606;

  // ── Paragrafo spacer (come nel master: sz=8, after=100) ───────────────────
  const spacer = new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: ' ', font: FONT, size: 8 })],
  });

  const children = [
    // Titolo
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'REGISTRO PRESENZE', bold: true, font: FONT, size: 34, color: C.BLU_DARK })],
    }),
    // Ragione sociale
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 20 },
      children: [new TextRun({ text: CLIENTE.ragioneSociale, bold: true, font: FONT, size: 24 })],
    }),
    // Indirizzo
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: CLIENTE.indirizzo, font: FONT, size: 20 })],
    }),

    // ── Tabella info corso (bordi AAAAAA, label bold + valore normale) ────
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [wL, wR],
      borders: { top: BAA.top, bottom: BAA.bottom, left: BAA.left, right: BAA.right, insideH: BAA.top, insideV: BAA.left },
      rows: [
        new TableRow({ children: [
          cellaAA(new Paragraph({ children: [
            new TextRun({ text: 'Corso: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: `Formazione Generale + Specifica (D.Lgs. 81/08 – ASR 17/04/2025)`, font: FONT, size: 20 }),
          ]}), { width: wL }),
          cellaAA(new Paragraph({ children: [
            new TextRun({ text: 'Mansione: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: mansione.nome, font: FONT, size: 20 }),
          ]}), { width: wR }),
        ]}),
        new TableRow({ children: [
          cellaAA(new Paragraph({ children: [
            new TextRun({ text: 'Relatore / Docente: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: CLIENTE.datoreLavoro, font: FONT, size: 20 }),
          ]}), { width: wL }),
          cellaAA(new Paragraph({ children: [new TextRun({ text: ' ', font: FONT, size: 20 })] }), { width: wR }),
        ]}),
      ],
    }),

    spacer,

    // ── Tabella presenze (intestazione 1F4E79, celle dati AAAAAA) ──────────
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: colW,
      borders: { top: BAA.top, bottom: BAA.bottom, left: BAA.left, right: BAA.right, insideH: BAA.top, insideV: BAA.left },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cellaHdr(colW[0], 'COGNOME'),
          cellaHdr(colW[1], 'NOME'),
          cellaHdr(colW[2], 'FIRMA ENTRATA'),
          cellaHdr(colW[3], 'ORA'),
          cellaHdr(colW[4], 'FIRMA USCITA'),
          cellaHdr(colW[5], 'ORA'),
          cellaHdr(colW[6], 'DATA INTERVENTO'),
        ]}),
        ...Array.from({ length: 15 }, () => new TableRow({
          height: { value: 500, rule: 'atLeast' },
          children: colW.map(w => cellaDato(w)),
        })),
      ],
    }),

    spacer,

    // ── Tabella argomenti (bordi AAAAAA, label bold + valore normale) ──────
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [W],
      borders: { top: BAA.top, bottom: BAA.bottom, left: BAA.left, right: BAA.right, insideH: BAA.top, insideV: BAA.left },
      rows: [new TableRow({ children: [
        cellaAA(new Paragraph({ children: [
          new TextRun({ text: 'Argomenti trattati:  ', bold: true, font: FONT, size: 20 }),
          new TextRun({ text: 'Vedasi progetto formativo', font: FONT, size: 20 }),
        ]}), { width: W }),
      ]})],
    }),
  ];

  const doc = new Document({ styles: docStyles, sections: [{ 
    properties: { page: { size: { width: 16838, height: 11906 }, margin: MARGIN } },
    headers: { default: header },
    footers: { default: footer },
    children,
  }]});
  await salvaDoc(doc, `${OUT}/02 - REGISTRO PRESENZE/00.Registro_FormIniziale_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// REGISTRO AGGIORNAMENTO
// ─────────────────────────────────────────────────────────────────────────────
async function genRegistroAggiornamento() {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 709, footer: 709 };
  const W = 14570;

  const BAA = {
    top:    { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
    bottom: { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
    left:   { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
    right:  { style: BorderStyle.SINGLE, size: 4, space: 0, color: 'AAAAAA' },
  };
  const BHDR = {
    top:    { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
    bottom: { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
    left:   { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
    right:  { style: BorderStyle.SINGLE, size: 4, space: 0, color: '1F4E79' },
  };

  // ── Header inline: logo sinistra + ragione sociale destra ────────────────
  const NO_HDR = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE} };
  const header = new Header({
    children: [new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [2400, W - 2400],
      borders: NO_HDR,
      rows: [new TableRow({ children: [
        new TableCell({
          width: { size: 2400, type: WidthType.DXA },
          borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 140, height: 32 } })] })],
        }),
        new TableCell({
          width: { size: W - 2400, type: WidthType.DXA },
          borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: CLIENTE.ragioneSociale, bold: true, font: FONT, size: 20, color: C.BLU_DARK })] })],
        }),
      ]})],
    })],
  });

  const footer = new Footer({
    children: [new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
      children: [
        new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   Pag. `, size: 16, font: FONT, color: C.GRIGIO }),
        new SimpleField('PAGE'),
      ],
    })],
  });

  function cellaAA(children, opts = {}) {
    return new TableCell({
      width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
      borders: BAA,
      margins: { top: opts.mTB || 100, bottom: opts.mTB || 100, left: opts.mLR || 150, right: opts.mLR || 150 },
      verticalAlign: VerticalAlign.CENTER,
      children: Array.isArray(children) ? children : [children],
    });
  }
  function cellaDato(w) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: BAA,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({ children: [new TextRun({ text: ' ' })] })],
    });
  }
  function cellaHdr(w, testo) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: BHDR,
      shading: { fill: '1F4E79', type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 100, right: 100 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: testo, bold: true, font: FONT, size: 18, color: 'FFFFFF' })],
      })],
    });
  }

  const colW = [2025, 2004, 3196, 708, 3402, 709, 2526];
  const wL = 6964; const wR = 7606;
  const spacer = new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: ' ', font: FONT, size: 8 })],
  });

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'REGISTRO PRESENZE', bold: true, font: FONT, size: 34, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 20 },
      children: [new TextRun({ text: CLIENTE.ragioneSociale, bold: true, font: FONT, size: 24 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: CLIENTE.indirizzo, font: FONT, size: 20 })],
    }),
    new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: [wL, wR],
      borders: { top: BAA.top, bottom: BAA.bottom, left: BAA.left, right: BAA.right, insideH: BAA.top, insideV: BAA.left },
      rows: [
        new TableRow({ children: [
          cellaAA(new Paragraph({ children: [
            new TextRun({ text: 'Corso: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: 'Aggiornamento della formazione specifica (6 ore)', font: FONT, size: 20 }),
          ]}), { width: wL }),
          cellaAA(new Paragraph({ children: [
            new TextRun({ text: 'Mansione: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: '_________________________________', font: FONT, size: 20 }),
          ]}), { width: wR }),
        ]}),
        new TableRow({ children: [
          cellaAA(new Paragraph({ children: [
            new TextRun({ text: 'Relatore / Docente: ', bold: true, font: FONT, size: 20 }),
            new TextRun({ text: CLIENTE.datoreLavoro, font: FONT, size: 20 }),
          ]}), { width: wL }),
          cellaAA(new Paragraph({ children: [new TextRun({ text: ' ', font: FONT, size: 20 })] }), { width: wR }),
        ]}),
      ],
    }),
    spacer,
    new Table({
      width: { size: W, type: WidthType.DXA }, columnWidths: colW,
      borders: { top: BAA.top, bottom: BAA.bottom, left: BAA.left, right: BAA.right, insideH: BAA.top, insideV: BAA.left },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cellaHdr(colW[0], 'COGNOME'),
          cellaHdr(colW[1], 'NOME'),
          cellaHdr(colW[2], 'FIRMA ENTRATA'),
          cellaHdr(colW[3], 'ORA'),
          cellaHdr(colW[4], 'FIRMA USCITA'),
          cellaHdr(colW[5], 'ORA'),
          cellaHdr(colW[6], 'DATA INTERVENTO'),
        ]}),
        ...Array.from({ length: 14 }, () => new TableRow({
          height: { value: 500, rule: 'atLeast' },
          children: colW.map(w => cellaDato(w)),
        })),
      ],
    }),
    spacer,
    // ── Argomenti trattati: label + 3 righe libere con underscore ──────────
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: 'Argomenti trattati:', bold: true, font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { after: 40, line: 276, lineRule: 'auto' },
      children: [new TextRun({ text: '_________________________________________________________________________________________________________________________________________________', bold: true, font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { after: 40, line: 276, lineRule: 'auto' },
      children: [new TextRun({ text: '_________________________________________________________________________________________________________________________________________________', bold: true, font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { after: 40, line: 276, lineRule: 'auto' },
      children: [new TextRun({ text: '_________________________________________________________________________________________________________________________________________________', bold: true, font: FONT, size: 20 })],
    }),
  ];

  const doc = new Document({ styles: docStyles, sections: [{
    properties: { page: { size: { width: 16838, height: 11906 }, margin: MARGIN } },
    headers: { default: header },
    footers: { default: footer },
    children,
  }]});
  await salvaDoc(doc, `${OUT}/02 - REGISTRO PRESENZE/01.Registro_Aggiornamento.docx`);
}

module.exports = { genProgettoFormativo, genRegistroFormIniziale, genRegistroAggiornamento };
