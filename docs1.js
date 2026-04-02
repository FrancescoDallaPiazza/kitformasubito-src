'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, SimpleField, LineRuleType,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, A4_L, MARGIN_STD, MARGIN_REG,
  makeHeader, makeFooter, titoloSezione, corpo, rigaDati, vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

// ─────────────────────────────────────────────────────────────────────────────
// PROGETTO FORMATIVO
// ─────────────────────────────────────────────────────────────────────────────
async function genProgettoFormativo() {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'PROGETTO FORMATIVO', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Progetto Formativo – D.Lgs. 81/2008 – ASR 17/04/2025');

  const mansioniList = MANSIONI.map(m => `- ${m.nome} (Rischio ${m.livello} – ${m.oreSpec + 4} ore totali)`).join('\n');

  const children = [
    // Titolo
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: 'PROGETTO FORMATIVO', bold: true, font: FONT, size: 36, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'D.Lgs. 81/2008 – Accordo Stato-Regioni 17/04/2025', font: FONT, size: 24, color: C.BLU_HEADER })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 280 },
      children: [new TextRun({ text: 'Formazione Generale e Specifica per i Lavoratori', font: FONT, size: 22, color: C.GRIGIO })],
    }),

    // 1. DATI AZIENDALI
    titoloSezione('1. DATI AZIENDALI'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [3200, 6438],
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        insideH: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        insideV: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      },
      rows: [
        ['Ragione Sociale', CLIENTE.ragioneSociale],
        ['Sede Operativa', CLIENTE.indirizzo],
        ['Partita IVA', CLIENTE.piva],
        ['Codice ATECO', `${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`],
        ['Datore di Lavoro / RSPP', CLIENTE.datoreLavoro],
        ['Medico Competente', 'Studio S.M.A.L – Dr. Lorenzo Adami'],
        ['Consulenza esterna SPP', 'Overall Group Srl'],
      ].map((r, i) => new TableRow({
        children: [
          cella(r[0], { width: 3200, bold: true, fill: i % 2 === 0 ? C.GRIGIO_ALT : C.BIANCO }),
          cella(r[1], { width: 6438, fill: i % 2 === 0 ? C.GRIGIO_ALT : C.BIANCO }),
        ],
      })),
    }),
    vuoto(120),

    // 2. RIFERIMENTI NORMATIVI
    titoloSezione('2. RIFERIMENTI NORMATIVI'),
    corpo('Il presente Progetto Formativo è redatto ai sensi di:', { before: 60 }),
    ...[
      'D.Lgs. 9 aprile 2008, n. 81 – Testo Unico sulla Salute e Sicurezza sul Lavoro, artt. 36 e 37;',
      'Accordo Stato-Regioni del 17/04/2025 (in sostituzione dell\'ASR 21/12/2011) – Formazione dei lavoratori, dirigenti e preposti;',
      'D.Lgs. 106/2009 – Modifiche integrative al D.Lgs. 81/08.',
    ].map(t => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 30, after: 30 },
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text: t, font: FONT, size: 20 })],
    })),
    vuoto(60),

    // 3. OBIETTIVI
    titoloSezione('3. OBIETTIVI DELLA FORMAZIONE'),
    corpo('La formazione ha lo scopo di fornire ai lavoratori adeguate conoscenze in materia di salute e sicurezza sul lavoro, con particolare riferimento a:', { before: 60 }),
    ...[
      'Concetti di rischio, danno, prevenzione e protezione nell\'ambiente lavorativo;',
      'Organizzazione della prevenzione aziendale e principali soggetti del sistema prevenzionistico;',
      'Diritti e doveri dei lavoratori, sanzioni per i lavoratori inadempienti;',
      'Rischi specifici propri della mansione svolta e relative misure di prevenzione e protezione;',
      'Corretto utilizzo dei Dispositivi di Protezione Individuale (DPI);',
      'Segnaletica di sicurezza e procedure per il primo soccorso e la gestione delle emergenze;',
      'Stress lavoro-correlato, incidenti e mancati infortuni (near miss).',
    ].map(t => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 30, after: 30 },
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text: t, font: FONT, size: 20 })],
    })),
    vuoto(60),

    // 4. STRUTTURA DEL CORSO
    titoloSezione('4. STRUTTURA DEL CORSO'),
    corpo('Il percorso formativo si articola in due moduli distinti, come previsto dall\'ASR 17/04/2025:', { before: 60 }),
    vuoto(60),

    // Tabella struttura
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [2200, 4038, 1700, 1700],
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        insideH: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        insideV: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            cella('Modulo', { width: 2200, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
            cella('Contenuti', { width: 4038, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
            cella('Durata', { width: 1700, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
            cella('Riferimento normativo', { width: 1700, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          ],
        }),
        new TableRow({
          children: [
            cella('FORMAZIONE GENERALE', { width: 2200, bold: true }),
            cella('Concetti di rischio, organizzazione prevenzione, diritti e doveri, segnaletica, primo soccorso, stress lavoro-correlato, near miss', { width: 4038 }),
            cella('4 ore', { width: 1700, align: 'center' }),
            cella('ASR 17/04/2025 – Parte II, Punto 2', { width: 1700 }),
          ],
        }),
        new TableRow({
          children: [
            cella('FORMAZIONE SPECIFICA – Rischio BASSO', { width: 2200, bold: true, fill: C.GRIGIO_ALT }),
            cella('Rischi specifici per Impiegata/o Amministrativa/o e Commerciale – DPI, procedure operative', { width: 4038, fill: C.GRIGIO_ALT }),
            cella('4 ore', { width: 1700, align: 'center', fill: C.GRIGIO_ALT }),
            cella('ASR 17/04/2025 – Parte IV, Punto 1', { width: 1700, fill: C.GRIGIO_ALT }),
          ],
        }),
        new TableRow({
          children: [
            cella('FORMAZIONE SPECIFICA – Rischio MEDIO', { width: 2200, bold: true }),
            cella('Rischi specifici per Manutentore – cadute dall\'alto, agenti chimici, rischio elettrico, DPI, ambienti confinati', { width: 4038 }),
            cella('8 ore', { width: 1700, align: 'center' }),
            cella('ASR 17/04/2025 – Parte IV, Punto 1', { width: 1700 }),
          ],
        }),
      ],
    }),
    vuoto(120),

    // 5. MANSIONI E LAVORATORI
    titoloSezione('5. MANSIONI E LAVORATORI INTERESSATI'),
    corpo('Le seguenti mansioni aziendali sono soggette a formazione obbligatoria ai sensi del D.Lgs. 81/2008:', { before: 60 }),
    ...MANSIONI.map(m => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 30, after: 30 },
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text: `${m.nome} – Rischio ${m.livello}: 4 ore generali + ${m.oreSpec} ore specifiche = ${4 + m.oreSpec} ore totali`, font: FONT, size: 20 })],
    })),
    corpo('I Tirocinanti seguono il percorso formativo della mansione di riferimento con supervisione obbligatoria del tutor aziendale.', { before: 60, after: 60 }),

    // 6. METODOLOGIA
    titoloSezione('6. METODOLOGIA DIDATTICA'),
    corpo('La formazione si svolge con le seguenti modalità, nel rispetto dell\'ASR 17/04/2025:', { before: 60 }),
    ...[
      'Presenza sul campo (formazione diretta con il Datore di Lavoro/RSPP);',
      'Utilizzo di materiale didattico specifico per mansione (schede mansione, test, attestati);',
      'Verifica dell\'apprendimento tramite test scritto a risposta multipla (30 domande);',
      'Possibilità di videoconferenza sincrona per i moduli di aggiornamento.',
    ].map(t => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 30, after: 30 },
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text: t, font: FONT, size: 20 })],
    })),
    vuoto(60),

    // 7. PROGRAMMA DETTAGLIATO FORMAZIONE GENERALE
    titoloSezione('7. PROGRAMMA DETTAGLIATO – FORMAZIONE GENERALE (4 ore)'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [1200, 6438, 2000],
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        insideH: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        insideV: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            cella('U.D.', { width: 1200, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
            cella('Argomento', { width: 6438, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
            cella('Durata', { width: 2000, bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          ],
        }),
        ...[
          ['U.D. 1', 'Il sistema legislativo: D.Lgs. 81/08, figure della sicurezza, diritti e doveri dei lavoratori', '45 min'],
          ['U.D. 2', 'Concetti di rischio, danno, prevenzione e protezione – metodologia di valutazione dei rischi', '45 min'],
          ['U.D. 3', 'Organizzazione della prevenzione aziendale – SPP, RLS, MC, addetti emergenza', '30 min'],
          ['U.D. 4', 'Segnaletica di sicurezza – tipologie, colori, significati (D.Lgs. 81/08 All. XXV)', '30 min'],
          ['U.D. 5', 'Procedure organizzative per il primo soccorso e la gestione delle emergenze', '30 min'],
          ['U.D. 6', 'Stress lavoro-correlato – definizione, fattori di rischio, misure di prevenzione', '20 min'],
          ['U.D. 7', 'Incidenti e mancati infortuni (near miss) – riconoscimento e segnalazione', '20 min'],
        ].map(([ud, arg, dur], i) => new TableRow({
          children: [
            cella(ud, { width: 1200, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
            cella(arg, { width: 6438, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
            cella(dur, { width: 2000, align: 'center', fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
          ],
        })),
      ],
    }),
    vuoto(120),

    // 8. PROGRAMMA FORMAZIONE SPECIFICA
    titoloSezione('8. PROGRAMMA DETTAGLIATO – FORMAZIONE SPECIFICA'),
    ...MANSIONI.map(m => [
      new Paragraph({
        spacing: { before: 60, after: 30 },
        children: [new TextRun({ text: `${m.nome} – Rischio ${m.livello} (${m.oreSpec} ore)`, bold: true, font: FONT, size: 22, color: C.BLU_MED })],
      }),
      new Table({
        width: { size: 9638, type: WidthType.DXA },
        columnWidths: [5638, 4000],
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          insideH: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
          insideV: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
        },
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              cella('Argomento specifico', { width: 5638, bold: true, fill: C.BLU_LIGHT }),
              cella('DPI associati', { width: 4000, bold: true, fill: C.BLU_LIGHT }),
            ],
          }),
          ...m.rischi.map((r, i) => new TableRow({
            children: [
              cella(r.nome, { width: 5638, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
              cella(r.dpi.join(', '), { width: 4000, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
            ],
          })),
        ],
      }),
      vuoto(80),
    ]).flat(),

    // 9. AGGIORNAMENTO
    titoloSezione('9. AGGIORNAMENTO PERIODICO'),
    corpo('L\'aggiornamento della formazione è obbligatorio ogni 5 anni per tutti i lavoratori, per una durata minima di 6 ore. L\'aggiornamento può essere svolto in modalità e-learning o in presenza, anche in un\'unica soluzione.', { before: 60 }),
    vuoto(60),

    // 10. VERIFICA APPRENDIMENTO
    titoloSezione('10. VERIFICA DELL\'APPRENDIMENTO'),
    corpo('Al termine di ciascun modulo (formazione generale e formazione specifica) viene effettuata una verifica dell\'apprendimento mediante:', { before: 60 }),
    ...[
      'Test scritto a risposta multipla (30 domande con 4 opzioni ciascuna);',
      'Soglia minima di superamento: 60% delle risposte corrette (18/30);',
      'In caso di mancato superamento: colloquio individuale orale con il Docente/RSPP.',
    ].map(t => new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 30, after: 30 },
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text: t, font: FONT, size: 20 })],
    })),
    vuoto(60),

    // 11. FIRME
    titoloSezione('11. APPROVAZIONE E FIRME'),
    vuoto(80),
    rigaDati('Datore di Lavoro / RSPP', `${CLIENTE.datoreLavoro}`, { sz: 20 }),
    vuoto(60),
    corpo('Luogo e data: Campagnola di Zevio, ___/___/_____', { before: 60 }),
    vuoto(60),
    new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: 'Firma: ___________________________________', font: FONT, size: 20 })],
    }),
  ];

  const doc = new Document({
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      }],
    },
    styles: docStyles,
    sections: [{
      properties: {
        page: { size: A4_P, margin: MARGIN_STD },
      },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/00 - PROGETTO FORMATIVO/ProgettoFormativo.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// REGISTRO PRESENZE FORMAZIONE INIZIALE (per mansione)
// ─────────────────────────────────────────────────────────────────────────────
async function genRegistroFormIniziale(mansione) {
  // Struttura corretta: NO header standard, footer con azienda+pag, corpo con titolo centrato
  const W = 16838 - 1134 - 2267; // landscape content: margini L2.0 R1.3 circa → uso valori corretti
  const WL = 16838; // page width landscape

  const footer = new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: 16000 }],
    children: [
      new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   Pag. `, size: 16, font: FONT, color: C.GRIGIO }),
      new SimpleField('PAGE'),
    ],
  })]});

  const nRighe = 15;
  const W9 = 13438; // landscape content ~ (29.7cm - 2cm - 2cm margins = 25.7cm = 14566 DXA, using approx)
  const colW = [2600, 2000, 2200, 900, 2200, 900, 2638]; // sum=13438

  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 30 },
      children: [new TextRun({ text: 'REGISTRO PRESENZE', bold: true, font: FONT, size: 34, color: C.BLU_DARK })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 10 },
      children: [new TextRun({ text: CLIENTE.ragioneSociale, bold: true, font: FONT, size: 24 })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: CLIENTE.indirizzo, font: FONT, size: 20 })],
    }),
    // Info corso table
    new Table({
      width: { size: W9, type: WidthType.DXA }, columnWidths: [Math.floor(W9/2), W9 - Math.floor(W9/2)],
      borders: { top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'} },
      rows: [
        new TableRow({ children: [
          cella(`Corso: Formazione Generale + Specifica (D.Lgs. 81/08 – ASR 17/04/2025)  |  ${mansione.oreSpec + 4} ore totali`, { width: Math.floor(W9/2) }),
          cella(`Mansione: ${mansione.nome}`, { width: W9 - Math.floor(W9/2) }),
        ]}),
        new TableRow({ children: [
          cella(`Relatore / Docente: ${CLIENTE.datoreLavoro}`, { width: Math.floor(W9/2) }),
          cella('', { width: W9 - Math.floor(W9/2) }),
        ]}),
      ],
    }),
    vuoto(60),
    // Firma table
    new Table({
      width: { size: W9, type: WidthType.DXA }, columnWidths: colW,
      borders: { top:{style:BorderStyle.SINGLE,size:1,color:'999999'},bottom:{style:BorderStyle.SINGLE,size:1,color:'999999'},left:{style:BorderStyle.SINGLE,size:1,color:'999999'},right:{style:BorderStyle.SINGLE,size:1,color:'999999'},insideH:{style:BorderStyle.SINGLE,size:1,color:'999999'},insideV:{style:BorderStyle.SINGLE,size:1,color:'999999'} },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cella('COGNOME', { width: colW[0], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('NOME', { width: colW[1], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('FIRMA ENTRATA', { width: colW[2], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('ORA', { width: colW[3], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('FIRMA USCITA', { width: colW[4], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('ORA', { width: colW[5], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('DATA INTERVENTO', { width: colW[6], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
        ]}),
        ...Array.from({ length: nRighe }, (_, i) => new TableRow({
          height: { value: 550, rule: 'atLeast' },
          children: colW.map(w => cella('', { width: w, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT })),
        })),
      ],
    }),
    vuoto(60),
    new Table({
      width: { size: W9, type: WidthType.DXA }, columnWidths: [W9],
      borders: { top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE} },
      rows: [new TableRow({ children: [
        cella('Argomenti trattati:  Vedasi progetto formativo', { width: W9 }),
      ]})],
    }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_L, margin: { top: 850, right: 1134, bottom: 1134, left: 1134 } } },
      footers: { default: footer },
      children,
    }],
  });
  await salvaDoc(doc, `${OUT}/02 - REGISTRO PRESENZE/Registro_FormIniziale_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// REGISTRO AGGIORNAMENTO
// ─────────────────────────────────────────────────────────────────────────────
async function genRegistroAggiornamento() {
  const footer = new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: 16000 }],
    children: [
      new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   Pag. `, size: 16, font: FONT, color: C.GRIGIO }),
      new SimpleField('PAGE'),
    ],
  })]});

  const nRighe = 14;
  const W9 = 13438;
  const colW = [2600, 2000, 2200, 900, 2200, 900, 2638];

  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 30 },
      children: [new TextRun({ text: 'REGISTRO PRESENZE', bold: true, font: FONT, size: 34, color: C.BLU_DARK })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 10 },
      children: [new TextRun({ text: CLIENTE.ragioneSociale, bold: true, font: FONT, size: 24 })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: CLIENTE.indirizzo, font: FONT, size: 20 })],
    }),
    new Table({
      width: { size: W9, type: WidthType.DXA }, columnWidths: [Math.floor(W9/2), W9 - Math.floor(W9/2)],
      borders: { top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'} },
      rows: [
        new TableRow({ children: [
          cella('Corso: Aggiornamento della formazione specifica (6 ore)', { width: Math.floor(W9/2) }),
          cella('Mansione: _________________________________', { width: W9 - Math.floor(W9/2) }),
        ]}),
        new TableRow({ children: [
          cella(`Relatore / Docente: ${CLIENTE.datoreLavoro}`, { width: Math.floor(W9/2) }),
          cella('', { width: W9 - Math.floor(W9/2) }),
        ]}),
      ],
    }),
    vuoto(60),
    new Table({
      width: { size: W9, type: WidthType.DXA }, columnWidths: colW,
      borders: { top:{style:BorderStyle.SINGLE,size:1,color:'999999'},bottom:{style:BorderStyle.SINGLE,size:1,color:'999999'},left:{style:BorderStyle.SINGLE,size:1,color:'999999'},right:{style:BorderStyle.SINGLE,size:1,color:'999999'},insideH:{style:BorderStyle.SINGLE,size:1,color:'999999'},insideV:{style:BorderStyle.SINGLE,size:1,color:'999999'} },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cella('COGNOME', { width: colW[0], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('NOME', { width: colW[1], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('FIRMA ENTRATA', { width: colW[2], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('ORA', { width: colW[3], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('FIRMA USCITA', { width: colW[4], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('ORA', { width: colW[5], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
          cella('DATA INTERVENTO', { width: colW[6], bold: true, fill: C.BLU_LIGHT, align: 'center' }),
        ]}),
        ...Array.from({ length: nRighe }, (_, i) => new TableRow({
          height: { value: 550, rule: 'atLeast' },
          children: colW.map(w => cella('', { width: w, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT })),
        })),
      ],
    }),
    vuoto(60),
    new Paragraph({ children: [new TextRun({ text: 'Argomenti trattati:', bold: true, font: FONT, size: 20 })] }),
    new Paragraph({ children: [new TextRun({ text: '____________________________________________________________________________________________________', font: FONT, size: 20 })] }),
    new Paragraph({ children: [new TextRun({ text: '____________________________________________________________________________________________________', font: FONT, size: 20 })] }),
    new Paragraph({ children: [new TextRun({ text: '____________________________________________________________________________________________________', font: FONT, size: 20 })] }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_L, margin: { top: 453, right: 1134, bottom: 1134, left: 1134 } } },
      footers: { default: footer },
      children,
    }],
  });
  await salvaDoc(doc, `${OUT}/02 - REGISTRO PRESENZE/Registro_Aggiornamento.docx`);
}

module.exports = { genProgettoFormativo, genRegistroFormIniziale, genRegistroAggiornamento };
