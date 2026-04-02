'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  TabStopType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, SimpleField, LineRuleType,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, A4_L, MARGIN_STD,
  makeHeader, makeFooter, titoloSezione, corpo, rigaDati, vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

// ─────────────────────────────────────────────────────────────────────────────
// COLLOQUIO INDIVIDUALE AGGIORNAMENTO (per mansione)
// ─────────────────────────────────────────────────────────────────────────────
async function genColloquio(mansione) {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'VERBALE COLLOQUIO INDIVIDUALE', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Verbale Colloquio Individuale – Aggiornamento');

  const checkbox = (testo, checked = false) => new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: `${checked ? '☒' : '☐'}  ${testo}`, font: FONT, size: 20 })],
  });

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'VERBALE DI COLLOQUIO INDIVIDUALE', bold: true, font: FONT, size: 30, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      children: [new TextRun({ text: `Aggiornamento Formazione – D.Lgs. 81/2008 – Mansione: ${mansione.nome}`, font: FONT, size: 20, color: C.BLU_HEADER })],
    }),

    // SEZIONE 1 – SOGGETTO FORMATORE
    titoloSezione('1. DATI DEL SOGGETTO FORMATORE'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [3200, 6438],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        ['Denominazione soggetto formatore', CLIENTE.ragioneSociale],
        ['Docente / Relatore', CLIENTE.datoreLavoro],
        ['Qualifica', 'Datore di Lavoro / RSPP interno'],
      ].map(([et, va], i) => new TableRow({
        children: [
          cella(et, { width: 3200, bold: true, fill: C.BLU_LIGHT }),
          cella(va, { width: 6438, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
        ],
      })),
    }),
    vuoto(120),

    // SEZIONE 2 – DATI CORSO
    titoloSezione('2. DATI DEL CORSO DI AGGIORNAMENTO'),
    new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [
        new TextRun({ text: 'Modalità: ', bold: true, font: FONT, size: 20 }),
        new TextRun({ text: '☐  In presenza     ☐  Videoconferenza sincrona', font: FONT, size: 20 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: 'Durata: _______ ore     Date: ___/___/_____  –  ___/___/_____', font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { before: 60, after: 120 },
      children: [new TextRun({ text: 'Sede / Luogo: ' + CLIENTE.indirizzo, font: FONT, size: 20 })],
    }),

    // SEZIONE 3 – PARTECIPANTE
    titoloSezione('3. DATI DEL PARTECIPANTE'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [3200, 6438],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        ['Cognome e Nome', ''],
        ['Data di nascita', ''],
        ['Mansione', mansione.nome],
        ['Reparto', mansione.reparto],
      ].map(([et, va], i) => new TableRow({
        children: [
          cella(et, { width: 3200, bold: true, fill: C.BLU_LIGHT }),
          cella(va, { width: 6438, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
        ],
      })),
    }),
    vuoto(120),

    // SEZIONE 4 – FINALITÀ
    titoloSezione('4. FINALITÀ E CONTENUTI DEL COLLOQUIO'),
    corpo('Argomenti trattati nel colloquio individuale (barrare le voci pertinenti):', { before: 60 }),
    checkbox('Rischi specifici della mansione e aggiornamenti normativi'),
    checkbox('Procedure aziendali di sicurezza e modifiche organizzative'),
    checkbox('Gestione emergenze e procedure di evacuazione'),
    checkbox('Uso corretto e manutenzione dei DPI'),
    checkbox('Segnalazione pericoli e mancati infortuni (near miss)'),
    checkbox('Addestramento specifico su attrezzature/procedure'),
    new Paragraph({
      spacing: { before: 60, after: 120 },
      children: [new TextRun({ text: '☐  Altro: _______________________________________________', font: FONT, size: 20 })],
    }),

    // SEZIONE 5 – VERIFICA
    titoloSezione('5. MODALITÀ DI VERIFICA DELL\'APPRENDIMENTO'),
    checkbox('Domande aperte a risposta libera'),
    checkbox('Caso pratico situazionale'),
    checkbox('Simulazione di una procedura operativa'),
    vuoto(120),

    corpo('Esito del colloquio:', { before: 60 }),
    new Paragraph({
      spacing: { before: 60, after: 80 },
      children: [new TextRun({ text: '☐  POSITIVO     ☐  NEGATIVO     ☐  Da ripetere entro: ___/___/_____', font: FONT, size: 20 })],
    }),
    vuoto(80),

    new Paragraph({
      spacing: { after: 80, line: 360, lineRule: LineRuleType.AUTO },
      children: [new TextRun({ text: 'Luogo e data: Campagnola di Zevio, ___/___/_____', font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { after: 80, line: 360, lineRule: LineRuleType.AUTO },
      children: [new TextRun({ text: 'Firma del Docente / RSPP: ___________________________________', font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { after: 80, line: 360, lineRule: LineRuleType.AUTO },
      children: [new TextRun({ text: 'Firma del Lavoratore: ___________________________________', font: FONT, size: 20 })],
    }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/01. AGGIORNAMENTO/Colloquio_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SCHEDA GRADIMENTO
// ─────────────────────────────────────────────────────────────────────────────
async function genGradimento() {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'SCHEDA DI GRADIMENTO', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Scheda di Gradimento – Formazione Sicurezza');

  const scala = (etichetta) => new TableRow({
    children: [
      cella(etichetta, { width: 5438 }),
      ...[1,2,3,4,5].map(n => cella(`${n}`, { width: 841, align: 'center' })),
    ],
  });

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'SCHEDA DI GRADIMENTO', bold: true, font: FONT, size: 30, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: 'Formazione Sicurezza sul Lavoro – D.Lgs. 81/2008', font: FONT, size: 20, color: C.BLU_HEADER })],
    }),
    corpo('Il presente questionario è anonimo e ha lo scopo di migliorare la qualità dei percorsi formativi. Ti chiediamo di esprimere il tuo giudizio su una scala da 1 (per niente) a 5 (moltissimo).', { before: 30, after: 60 }),

    titoloSezione('A – ORGANIZZAZIONE DEL CORSO'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [5438, 841, 841, 841, 841, 836],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cella('Domanda', { width: 5438, bold: true, fill: C.BLU_LIGHT }),
          ...[1,2,3,4,5].map(n => cella(`${n}`, { width: 841, bold: true, fill: C.BLU_LIGHT, align: 'center' })),
        ]}),
        scala('Le informazioni organizzative (orari, sede, durata) erano chiare?'),
        scala('Il materiale didattico era comprensibile e ben strutturato?'),
        scala('I tempi e la durata del corso erano adeguati?'),
      ],
    }),
    vuoto(60),

    titoloSezione('B – CONTENUTI E DOCENTE'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [5438, 841, 841, 841, 841, 836],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cella('Domanda', { width: 5438, bold: true, fill: C.BLU_LIGHT }),
          ...[1,2,3,4,5].map(n => cella(`${n}`, { width: 841, bold: true, fill: C.BLU_LIGHT, align: 'center' })),
        ]}),
        scala('I contenuti erano pertinenti alla tua mansione?'),
        scala('Il docente ha esposto gli argomenti in modo chiaro?'),
        scala('Il docente era disponibile a rispondere alle domande?'),
        scala('Il corso ha aumentato la tua consapevolezza sui rischi?'),
      ],
    }),
    vuoto(60),

    titoloSezione('C – UTILITÀ E SUGGERIMENTI'),
    corpo('Il corso è stato utile per la tua attività lavorativa?', { before: 60, after: 60 }),
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: '☐  Per niente   ☐  Poco   ☐  Abbastanza   ☐  Molto   ☐  Moltissimo', font: FONT, size: 20 })],
    }),
    corpo('Suggerimenti o commenti liberi:', { before: 60, after: 60 }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: '___________________________________________________________________________', font: FONT, size: 20 })] }),
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: '___________________________________________________________________________', font: FONT, size: 20 })] }),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: '___________________________________________________________________________', font: FONT, size: 20 })] }),

    corpo('Grazie per la tua collaborazione.', { before: 0, after: 60 }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/04 - GRADIMENTO/Gradimento.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTESTATO (per mansione)
// ─────────────────────────────────────────────────────────────────────────────
async function genAttestato(mansione) {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'ATTESTATO DI FORMAZIONE', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Attestato di Formazione – D.Lgs. 81/2008');

  const livColor = mansione.livello === 'ALTO' ? C.ROSSO : mansione.livello === 'MEDIO' ? C.ARANCIO : C.VERDE;
  const rigaFirma = (testo) => new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text: testo, font: FONT, size: 20 })],
  });

  const sezione = (tipo, rif, durata) => [
    new Paragraph({
      spacing: { after: 40 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED } },
      children: [new TextRun({ text: `── ${tipo} ──────────────────────────────────────────`, font: FONT, size: 18, color: C.BLU_MED })],
    }),
    rigaFirma(`Tipologia:           ${tipo}`),
    rigaFirma(`Rif. normativo:      ${rif}`),
    rigaFirma(`Durata:              ${durata}`),
    rigaFirma('Modalità:            Presenza sul campo'),
    rigaFirma('Data inizio/fine:    ___/___/_____'),
    rigaFirma(`Sede:                ${CLIENTE.indirizzo}`),
    vuoto(60),
  ];

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 60 },
      children: [new TextRun({ text: 'ATTESTATO DI FORMAZIONE', bold: true, font: FONT, size: 36, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: 'Formazione Generale e Specifica – D.Lgs. 81/2008 – ASR 17/04/2025', font: FONT, size: 22, color: C.BLU_HEADER })],
    }),

    corpo(`Il/La sottoscritto/a ${CLIENTE.datoreLavoro} in qualità di Datore di Lavoro e Soggetto Formatore della Società ${CLIENTE.ragioneSociale} con Codice ATECO ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc},`, { before: 40, after: 40 }),
    corpo(`nei confronti del/la lavoratore/rice avente Mansione di ${mansione.nome}:`, { before: 0, after: 80 }),

    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({ text: 'Sig./Sig.ra ________________________________', font: FONT, size: 22, bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({ text: 'Nato/a a ______________________ il ___/___/_____', font: FONT, size: 20 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: 'Codice Fiscale: ______________________________', font: FONT, size: 20 })],
    }),

    corpo('ha frequentato con esito positivo il corso di formazione in materia di Salute e Sicurezza sul Lavoro:', { before: 40, after: 80 }),

    ...sezione(
      'FORMAZIONE GENERALE',
      'ASR 17/04/2025 (D.Lgs. 81/2008, art. 37) – Parte II, Punto 2',
      '4 ore'
    ),
    ...sezione(
      `FORMAZIONE SPECIFICA – RISCHIO ${mansione.livello}`,
      'ASR 17/04/2025 – Parte IV, Punto 1',
      `${mansione.oreSpec} ore`
    ),

    corpo('Superando con esito positivo la verifica finale dell\'apprendimento, effettuata secondo quanto previsto dall\'Accordo Stato-Regioni 17/04/2025. Il presente attestato costituisce credito formativo permanente e ha validità su tutto il territorio nazionale.', { before: 40, after: 80 }),

    new Paragraph({
      spacing: { after: 40 },
      children: [new TextRun({ text: 'Luogo e data: Campagnola di Zevio, ___/___/_____', font: FONT, size: 20 })],
    }),
    vuoto(80),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [4819, 4819],
      borders: {
        top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
        left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
        insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE},
      },
      rows: [new TableRow({ children: [
        new TableCell({ width:{size:4819,type:WidthType.DXA}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children: [
          new Paragraph({ children: [new TextRun({ text: 'Firma Soggetto Formatore / DL', font: FONT, size: 20 })] }),
          new Paragraph({ spacing:{after:40}, children: [new TextRun({ text: '___________________________________', font: FONT, size: 20 })] }),
        ]}),
        new TableCell({ width:{size:4819,type:WidthType.DXA}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children: [
          new Paragraph({ children: [new TextRun({ text: 'Firma Relatore / DL / RSPP', font: FONT, size: 20 })] }),
          new Paragraph({ spacing:{after:40}, children: [new TextRun({ text: '___________________________________', font: FONT, size: 20 })] }),
        ]}),
      ]})],
    }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/05 - ATTESTATI/Attestato_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTESTATO AGGIORNAMENTO
// ─────────────────────────────────────────────────────────────────────────────
async function genAttestatiAggiornamento() {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'ATTESTATO DI AGGIORNAMENTO', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Attestato Aggiornamento – D.Lgs. 81/2008');

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 60 },
      children: [new TextRun({ text: 'ATTESTATO DI AGGIORNAMENTO', bold: true, font: FONT, size: 36, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 100 },
      children: [new TextRun({ text: 'Aggiornamento Formazione – D.Lgs. 81/2008 – ASR 17/04/2025', font: FONT, size: 22, color: C.BLU_HEADER })],
    }),
    corpo(`Il/La sottoscritto/a ${CLIENTE.datoreLavoro} in qualità di Datore di Lavoro e Soggetto Formatore della Società ${CLIENTE.ragioneSociale}`, { before: 40, after: 40 }),
    corpo('nei confronti del/la lavoratore/rice:', { before: 0, after: 80 }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({ text: 'Sig./Sig.ra ________________________________', font: FONT, size: 22, bold: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [new TextRun({ text: 'Mansione: ______________________________', font: FONT, size: 20 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: 'Nato/a a ______________________ il ___/___/_____', font: FONT, size: 20 })],
    }),
    corpo('ha frequentato con esito positivo il corso di AGGIORNAMENTO in materia di Salute e Sicurezza sul Lavoro:', { before: 40, after: 80 }),
    new Paragraph({
      spacing: { after: 40 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED } },
      children: [new TextRun({ text: '── AGGIORNAMENTO QUINQUENNALE ─────────────────────────', font: FONT, size: 18, color: C.BLU_MED })],
    }),
    ...[
      `Rif. normativo:      ASR 17/04/2025 – Aggiornamento art. 37, D.Lgs. 81/2008`,
      'Durata:              6 ore',
      'Modalità:            ☐ Presenza   ☐ Videoconferenza sincrona',
      'Data inizio/fine:    ___/___/_____',
      `Sede:                ${CLIENTE.indirizzo}`,
    ].map(t => new Paragraph({ spacing:{after:40}, children:[new TextRun({text:t,font:FONT,size:20})] })),
    vuoto(80),
    corpo('Il presente attestato costituisce credito formativo quinquennale ai sensi dell\'ASR 17/04/2025.', { before: 40, after: 80 }),
    new Paragraph({ spacing:{after:40}, children:[new TextRun({text:'Luogo e data: Campagnola di Zevio, ___/___/_____',font:FONT,size:20})] }),
    vuoto(80),
    new Paragraph({ spacing:{after:40}, children:[new TextRun({text:'Firma Soggetto Formatore / DL: ___________________________________',font:FONT,size:20})] }),
    new Paragraph({ spacing:{after:40}, children:[new TextRun({text:'Firma Lavoratore/rice: ___________________________________',font:FONT,size:20})] }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/05 - ATTESTATI/Attestato_Aggiorn.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERBALE VERIFICA FINALE
// ─────────────────────────────────────────────────────────────────────────────
async function genVerbaleVerifica() {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'VERBALE VERIFICA FINALE', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Verbale Verifica Finale – D.Lgs. 81/2008');

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'VERBALE DI VERIFICA FINALE DELL\'APPRENDIMENTO', bold: true, font: FONT, size: 28, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: 'D.Lgs. 81/2008 – ASR 17/04/2025', font: FONT, size: 20, color: C.BLU_HEADER })],
    }),

    titoloSezione('1. DATI DEL CORSO'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [3200, 6438],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        ['Soggetto formatore', CLIENTE.ragioneSociale],
        ['Docente / Esaminatore', CLIENTE.datoreLavoro],
        ['Data verifica', '___/___/_____'],
        ['Modalità verifica', '☐ Test scritto   ☐ Colloquio orale   ☐ Caso pratico'],
        ['Soglia superamento', '18/30 (60%)'],
      ].map(([et, va], i) => new TableRow({
        children: [
          cella(et, { width: 3200, bold: true, fill: C.BLU_LIGHT }),
          cella(va, { width: 6438, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
        ],
      })),
    }),
    vuoto(120),

    titoloSezione('2. RISULTATI INDIVIDUALI'),
    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [2700, 1800, 1200, 1200, 1200, 1538],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'999999'},bottom:{style:BorderStyle.SINGLE,size:1,color:'999999'},
        left:{style:BorderStyle.SINGLE,size:1,color:'999999'},right:{style:BorderStyle.SINGLE,size:1,color:'999999'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'999999'},insideV:{style:BorderStyle.SINGLE,size:1,color:'999999'},
      },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cella('Cognome e Nome', { width:2700, bold:true, fill:C.BLU_LIGHT, align:'center' }),
          cella('Mansione', { width:1800, bold:true, fill:C.BLU_LIGHT, align:'center' }),
          cella('Punteggio', { width:1200, bold:true, fill:C.BLU_LIGHT, align:'center' }),
          cella('Max', { width:1200, bold:true, fill:C.BLU_LIGHT, align:'center' }),
          cella('Esito', { width:1200, bold:true, fill:C.BLU_LIGHT, align:'center' }),
          cella('Firma', { width:1538, bold:true, fill:C.BLU_LIGHT, align:'center' }),
        ]}),
        ...Array.from({length:10}, (_,i) => new TableRow({
          height: { value: 550, rule: 'atLeast' },
          children: [
            cella('', {width:2700, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
            cella('', {width:1800, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
            cella('', {width:1200, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
            cella('30', {width:1200, align:'center', fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
            cella('', {width:1200, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
            cella('', {width:1538, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
          ],
        })),
      ],
    }),
    vuoto(120),

    new Paragraph({ spacing:{ before:200, after:60 }, children:[new TextRun({text:'Note del commissario:',bold:true,font:FONT,size:20})] }),
    new Paragraph({ spacing:{after:60}, children:[new TextRun({text:'_________________________________________________________________',font:FONT,size:20})] }),
    new Paragraph({ spacing:{after:60}, children:[new TextRun({text:'_________________________________________________________________',font:FONT,size:20})] }),
    vuoto(80),
    new Paragraph({ spacing:{ after:80, line:360, lineRule:LineRuleType.AUTO }, children:[new TextRun({text:'Luogo e data: Campagnola di Zevio, ___/___/_____',font:FONT,size:20})] }),
    new Paragraph({ spacing:{ after:80, line:360, lineRule:LineRuleType.AUTO }, children:[new TextRun({text:'Firma del Docente / Esaminatore: ___________________________________',font:FONT,size:20})] }),
    new Paragraph({ spacing:{ after:80, line:360, lineRule:LineRuleType.AUTO }, children:[new TextRun({text:'Firma del Datore di Lavoro: ___________________________________',font:FONT,size:20})] }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/06 - VERBALE VERIFICA/Verbale_Verifica_Finale.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERIFICA EFFICACIA FORMAZIONE
// ─────────────────────────────────────────────────────────────────────────────
async function genVerificaEfficacia() {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'VERIFICA EFFICACIA FORMAZIONE', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('Verifica Efficacia Formazione');

  const check = (testo) => new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text: `☐  ${testo}`, font: FONT, size: 20 })],
  });

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'VERIFICA EFFICACIA DELLA FORMAZIONE', bold: true, font: FONT, size: 28, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: 'D.Lgs. 81/2008 – ASR 17/04/2025 – Checklist di osservazione sul campo', font: FONT, size: 20, color: C.BLU_HEADER })],
    }),

    new Table({
      width: { size: 9638, type: WidthType.DXA },
      columnWidths: [3200, 6438],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        ['Lavoratore/rice', ''],
        ['Mansione', ''],
        ['Data formazione', ''],
        ['Data osservazione', '___/___/_____'],
        ['Osservatore (RSPP/DL)', CLIENTE.datoreLavoro],
      ].map(([et, va], i) => new TableRow({
        children: [
          cella(et, { width: 3200, bold: true, fill: C.BLU_LIGHT }),
          cella(va, { width: 6438, fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT }),
        ],
      })),
    }),
    vuoto(120),

    titoloSezione('A – CONOSCENZE (comprovate da osservazione)'),
    check('Il lavoratore conosce i rischi specifici della propria mansione'),
    check('Il lavoratore conosce le procedure di emergenza e le vie di esodo'),
    check('Il lavoratore conosce i numeri di emergenza e gli addetti al primo soccorso'),
    check('Il lavoratore è in grado di identificare la segnaletica di sicurezza'),
    vuoto(60),

    titoloSezione('B – COMPORTAMENTI E USO DPI'),
    check('Il lavoratore utilizza correttamente i DPI previsti per la propria mansione'),
    check('Il lavoratore mantiene in ordine il proprio spazio di lavoro'),
    check('Il lavoratore segnala tempestivamente condizioni di rischio o near miss'),
    check('Il lavoratore rispetta le procedure operative di sicurezza'),
    vuoto(60),

    titoloSezione('C – ESITO DELLA VERIFICA'),
    new Paragraph({
      spacing: { before: 60, after: 80 },
      children: [new TextRun({ text: '☐  EFFICACE – La formazione ha prodotto i comportamenti sicuri attesi', font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { before: 60, after: 80 },
      children: [new TextRun({ text: '☐  PARZIALMENTE EFFICACE – Necessario rinforzo su: _________________________', font: FONT, size: 20 })],
    }),
    new Paragraph({
      spacing: { before: 60, after: 80 },
      children: [new TextRun({ text: '☐  NON EFFICACE – Prevista nuova sessione formativa entro: ___/___/_____', font: FONT, size: 20 })],
    }),
    vuoto(80),
    new Paragraph({ spacing:{after:80,line:360,lineRule:LineRuleType.AUTO}, children:[new TextRun({text:'Luogo e data: Campagnola di Zevio, ___/___/_____',font:FONT,size:20})] }),
    new Paragraph({ spacing:{after:80,line:360,lineRule:LineRuleType.AUTO}, children:[new TextRun({text:'Firma Osservatore (RSPP/DL): ___________________________________',font:FONT,size:20})] }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/07 - VERIFICA EFFICACIA/Verifica_Efficacia.docx`);
}

module.exports = { genColloquio, genGradimento, genAttestato, genAttestatiAggiornamento, genVerbaleVerifica, genVerificaEfficacia };
