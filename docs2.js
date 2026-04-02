'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  TabStopType, SimpleField, LineRuleType,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, A4_L, MARGIN_STD, logoBytes,
  makeHeader, makeFooter, titoloSezione, corpo, vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';
const W = 9638; // portrait content width

// ── helpers locali ────────────────────────────────────────────────────────────
const BD = { top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'}, bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'}, left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'}, right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'} };
const NO = { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} };

function footerAzienda(conPag = true) {
  const children = [
    new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   `, size: 16, font: FONT, color: C.GRIGIO }),
  ];
  if (conPag) {
    children.push(new TextRun({ text: 'Pag. ', size: 16, font: FONT }));
    children.push(new SimpleField('PAGE'));
  }
  return new Footer({ children: [new Paragraph({ children })] });
}

function tabellaKV(rows, wL = 3200) {
  const wR = W - wL;
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [wL, wR],
    borders: { top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top },
    rows: rows.map(([et, va], i) => new TableRow({ children: [
      cella(et, { width: wL, bold: true, fill: i%2===0?C.BLU_LIGHT:C.BIANCO }),
      cella(va, { width: wR, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
    ]})),
  });
}

function centrato(txt, opts = {}) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: opts.after || 30 },
    children: [new TextRun({ text: txt, bold: opts.bold, font: FONT, size: opts.sz || 20, color: opts.color })],
  });
}

function tabellaFirme2col(lab1, lab2) {
  const half = Math.floor(W/2);
  return new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[half,W-half],
    borders:{top:NO.top,bottom:NO.bottom,left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[new TableRow({ children:[
      new TableCell({ width:{size:half,type:WidthType.DXA}, borders:NO, margins:{top:60,bottom:60,left:0,right:60}, children:[
        new Paragraph({children:[new TextRun({text:lab1,font:FONT,size:20})]}),
        new Paragraph({spacing:{after:40},children:[new TextRun({text:'_________________________________',font:FONT,size:20})]}),
      ]}),
      new TableCell({ width:{size:W-half,type:WidthType.DXA}, borders:NO, margins:{top:60,bottom:60,left:60,right:0}, children:[
        new Paragraph({children:[new TextRun({text:lab2,font:FONT,size:20})]}),
        new Paragraph({spacing:{after:40},children:[new TextRun({text:'_________________________________',font:FONT,size:20})]}),
      ]}),
    ]})]
  });
}

// logo + azienda header (usato in gradimento e test)
function topLogoAzienda(sub) {
  const wL = 1300; const wR = W - wL;
  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wL,wR],
    borders:{top:NO.top,bottom:NO.bottom,left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[
      new TableRow({ children:[
        new TableCell({ width:{size:wL,type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER, borders:NO, margins:{top:40,bottom:40,left:0,right:60},
          children:[new Paragraph({alignment:AlignmentType.LEFT, children:[new ImageRun({data:logoBytes,type:'jpg',transformation:{width:55,height:55}})]})],
        }),
        new TableCell({ width:{size:wR,type:WidthType.DXA}, verticalAlign:VerticalAlign.CENTER, borders:NO, margins:{top:40,bottom:40,left:60,right:0},
          children:[
            new Paragraph({alignment:AlignmentType.RIGHT, children:[new TextRun({text:CLIENTE.ragioneSociale,bold:true,font:FONT,size:20,color:C.BLU_HEADER})]}),
            ...(sub ? [new Paragraph({alignment:AlignmentType.RIGHT, children:[new TextRun({text:sub,font:FONT,size:18,color:C.GRIGIO})]})] : []),
          ],
        }),
      ]}),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// COLLOQUIO INDIVIDUALE AGGIORNAMENTO
// Struttura corretta: header standard (logo | azienda | titolo | ATECO), footer "Pag. X a Y"
// ─────────────────────────────────────────────────────────────────────────────
async function genColloquio(mansione) {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'Colloquio Individuale – Aggiornamento', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter('');

  const children = [
    centrato('VERBALE DI COLLOQUIO INDIVIDUALE', {bold:true, sz:32, color:C.BLU_DARK, after:20}),
    centrato('Accordo Stato-Regioni 17/04/2025 – Parte IV, Punto 6 e 6.3', {sz:18, color:C.GRIGIO, after:80}),

    new Paragraph({ spacing:{after:20}, children:[new TextRun({text:'1. DATI DEL SOGGETTO FORMATORE',bold:true,font:FONT,size:22,color:C.BLU_MED})] }),
    tabellaKV([
      ['Denominazione', CLIENTE.ragioneSociale],
      ['Docente', CLIENTE.datoreLavoro],
      ['Qualifica', 'Datore di Lavoro / RSPP'],
    ]),
    vuoto(80),

    new Paragraph({ spacing:{after:20}, children:[new TextRun({text:'2. DATI DEL CORSO DI AGGIORNAMENTO',bold:true,font:FONT,size:22,color:C.BLU_MED})] }),
    tabellaKV([
      ['Modalità', '☐ In presenza    ☐ Videoconferenza sincrona'],
      ['Durata', '___ ore / minuti'],
      ['Data/e svolgimento', '___________________________'],
    ]),
    vuoto(80),

    new Paragraph({ spacing:{after:20}, children:[new TextRun({text:'3. DATI DEL PARTECIPANTE',bold:true,font:FONT,size:22,color:C.BLU_MED})] }),
    tabellaKV([
      ['Cognome e Nome', '___________________________'],
      ['Data di nascita', '___________________________'],
      ['Mansione', mansione.nome],
      ['Reparto / Area', mansione.reparto],
    ]),
    vuoto(80),

    new Paragraph({ spacing:{after:20}, children:[new TextRun({text:'4. FINALITÀ E CONTENUTI',bold:true,font:FONT,size:22,color:C.BLU_MED})] }),
    ...['Rischi specifici della mansione','Procedure aziendali di sicurezza','Gestione emergenze e evacuazione','Uso corretto dei DPI','Segnalazione pericoli / near miss','Addestramento specifico','Altro: _______________________________'].map(t =>
      new Paragraph({ spacing:{before:40,after:40}, children:[new TextRun({text:`☐  ${t}`,font:FONT,size:20})] })
    ),
    vuoto(80),

    new Paragraph({ spacing:{after:20}, children:[new TextRun({text:'5. MODALITÀ DI VERIFICA',bold:true,font:FONT,size:22,color:C.BLU_MED})] }),
    new Paragraph({ spacing:{after:80}, children:[new TextRun({text:'☐ Domande aperte\t\t☐ Caso pratico\t\t☐ Simulazione\t\t☐ Questionario scritto',font:FONT,size:20})] }),

    new Paragraph({ spacing:{after:20}, children:[new TextRun({text:'6. ESITO',bold:true,font:FONT,size:22,color:C.BLU_MED})] }),
    new Paragraph({ spacing:{after:80}, children:[new TextRun({text:'☐  IDONEO          ☐  NON IDONEO',bold:true,font:FONT,size:20})] }),

    new Paragraph({ spacing:{after:30}, children:[new TextRun({text:`Luogo: ${CLIENTE.indirizzo}`,font:FONT,size:20})] }),
    new Paragraph({ spacing:{after:80}, children:[new TextRun({text:'Data: ___/___/______',font:FONT,size:20})] }),

    tabellaFirme2col('Firma del Datore di Lavoro / RSPP','Firma del Lavoratore'),
  ];

  const doc = new Document({ styles:docStyles, sections:[{
    properties:{page:{size:A4_P,margin:MARGIN_STD}},
    headers:{default:header}, footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/01. AGGIORNAMENTO/Colloquio_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SCHEDA GRADIMENTO
// Struttura corretta: NO header standard, footer azienda+pag,
// corpo: topLogoAzienda table + titolo centrato + legenda + tabella 20x5 (scala 1-4)
// ─────────────────────────────────────────────────────────────────────────────
async function genGradimento() {
  const footer = footerAzienda(true);
  const MARGIN = { top: 737, right: 1134, bottom: 1134, left: 1134 };

  // Tabella rating 1-4
  const wDom = 5638; const wN = Math.floor((W-wDom)/4);
  const wNLast = W - wDom - wN*3;

  function domandaRow(testo, isHeader = false, isSection = false) {
    const fill = isHeader ? C.BLU_LIGHT : isSection ? C.BLU_ALT : C.BIANCO;
    const colSpan = isSection ? 4 : 0;
    if (isSection) {
      return new TableRow({ children:[
        cella(testo, {width:wDom, bold:true, fill}),
        new TableCell({ columnSpan:4, width:{size:W-wDom,type:WidthType.DXA}, shading:{fill,type:ShadingType.CLEAR},
          borders:BD, margins:{top:60,bottom:60,left:80,right:80},
          children:[new Paragraph({children:[new TextRun({text:'',font:FONT,size:20})]})],
        }),
      ]});
    }
    return new TableRow({ children:[
      cella(testo, {width:wDom, bold:isHeader, fill}),
      cella(isHeader?'1':'☐', {width:wN, bold:isHeader, fill, align:'center'}),
      cella(isHeader?'2':'☐', {width:wN, bold:isHeader, fill, align:'center'}),
      cella(isHeader?'3':'☐', {width:wN, bold:isHeader, fill, align:'center'}),
      cella(isHeader?'4':'☐', {width:wNLast, bold:isHeader, fill, align:'center'}),
    ]});
  }

  const domande = [
    {testo:'DOMANDA', isHeader:true},
    {testo:'CONTENUTI', isSection:true},
    {testo:'I contenuti del corso erano pertinenti alle mie mansioni lavorative'},
    {testo:'I contenuti erano aggiornati e conformi alla normativa vigente'},
    {testo:'La quantità di informazioni fornita era adeguata'},
    {testo:'DOCENTE', isSection:true},
    {testo:'Il docente ha esposto gli argomenti in modo chiaro e comprensibile'},
    {testo:'Il docente era disponibile a rispondere alle domande'},
    {testo:'Gli esempi pratici erano pertinenti alla realtà lavorativa'},
    {testo:'ORGANIZZAZIONE', isSection:true},
    {testo:'La durata del corso era adeguata agli argomenti trattati'},
    {testo:'Il luogo/modalità di svolgimento era adeguato'},
    {testo:'UTILITÀ', isSection:true},
    {testo:'Il corso ha aumentato la mia consapevolezza sui rischi'},
    {testo:'Le informazioni ricevute sono applicabili nella mia attività quotidiana'},
    {testo:'La verifica finale era coerente con quanto trattato nel corso'},
    {testo:'VALUTAZIONE COMPLESSIVA', isSection:true},
    {testo:'Valutazione complessiva del corso'},
    {testo:'Consiglierei questo corso ad altri colleghi'},
    {testo:'Mi ritengo più preparato/a sulla sicurezza dopo il corso'},
  ];

  const children = [
    topLogoAzienda('QUESTIONARIO DI GRADIMENTO'),
    vuoto(40),
    centrato('QUESTIONARIO DI VALUTAZIONE DEL GRADIMENTO', {bold:true, sz:28, color:C.BLU_DARK, after:20}),
    centrato('Formazione Obbligatoria – D.Lgs. 81/08', {sz:20, color:C.GRIGIO, after:40}),
    new Paragraph({spacing:{after:40}, children:[new TextRun({text:'Il presente questionario è anonimo e ha lo scopo di migliorare la qualità dei futuri interventi formativi. La preghiamo di compilarlo con sincerità.',font:FONT,size:20})]}),
    new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[W],
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:NO.top,insideV:NO.top},
      rows:[new TableRow({children:[cella('LEGENDA: 1 = Insufficiente   2 = Sufficiente   3 = Buono   4 = Ottimo',{width:W,bold:true,fill:C.BLU_ALT})]})]
    }),
    vuoto(20),
    new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[wDom,wN,wN,wN,wNLast],
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
      rows: domande.map(d => domandaRow(d.testo, d.isHeader, d.isSection)),
    }),
    vuoto(40),
    new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[W],
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:NO.top,insideV:NO.top},
      rows:[new TableRow({children:[cella('OSSERVAZIONI E SUGGERIMENTI',{width:W,bold:true,fill:C.BLU_LIGHT})]})]
    }),
    new Paragraph({spacing:{before:40,after:20}, children:[new TextRun({text:'Cosa ha apprezzato maggiormente?',font:FONT,size:20})]}),
    new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right},rows:[new TableRow({height:{value:800,rule:'atLeast'},children:[cella('',{width:W})]})]}),
    new Paragraph({spacing:{before:40,after:20}, children:[new TextRun({text:'Cosa migliorereste?',font:FONT,size:20})]}),
    new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right},rows:[new TableRow({height:{value:800,rule:'atLeast'},children:[cella('',{width:W})]})]}),
  ];

  const doc = new Document({styles:docStyles, sections:[{
    properties:{page:{size:A4_P,margin:MARGIN}},
    footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/04 - GRADIMENTO/Gradimento.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTESTATO (per mansione)
// Struttura corretta: NO header, footer azienda (no pag), 2 attestati per file
// ─────────────────────────────────────────────────────────────────────────────
async function genAttestato(mansione) {
  const MARGIN_ATT = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
  const footer = new Footer({ children: [new Paragraph({
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});

  const wL = 3800; const wR = W - wL; const half = Math.floor(W/2);

  function buildSection(tipoLabel, durata, nota) {
    return [
      centrato('ATTESTATO DI FORMAZIONE', {bold:true, sz:40, color:C.BLU_DARK, after:20}),
      centrato('Formazione Generale e Specifica – D.Lgs. 81/2008 – ASR 17/04/2025', {sz:22, color:C.BLU_MED, after:40}),
      centrato(`Il/La sottoscritto/a ${CLIENTE.datoreLavoro} in qualità di Datore di Lavoro e Soggetto Formatore`),
      centrato(`della Società ${CLIENTE.ragioneSociale}`),
      centrato(`con Codice ATECO ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`),
      centrato('nei confronti del/la lavoratore/rice avente'),
      centrato(`Mansione di ${mansione.nome}`, {bold:true, sz:24, color:C.BLU_DARK, after:40}),
      centrato('Sig./Sig.ra  ____________________________________________'),
      centrato('Nato/a a ________________________ il ___/___/_____'),
      centrato('Codice Fiscale: ___________________________________________', {after:60}),
      centrato('ha frequentato il corso di formazione in materia di Salute e Sicurezza sul Lavoro:', {sz:20, after:40}),
      new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[wL,wR],
        borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
        rows:[
          ['Tipologia corso:', tipoLabel],
          ['Riferimento normativo:', 'ASR 17/04/2025 (D.Lgs. 81/2008, art. 37)\nRiferimento normativo e contenuti minimi – Parte II'],
          ['Durata:', durata],
          ['Modalità:', 'Presenza sul campo'],
          ['Data inizio:', '___/___/_____'],
          ['Data fine:', '___/___/_____'],
        ].map(([et,va],i) => new TableRow({children:[
          cella(et,{width:wL,bold:true,fill:i%2===0?C.GRIGIO_ALT:C.BIANCO}),
          cella(va,{width:wR,fill:i%2===0?C.GRIGIO_ALT:C.BIANCO}),
        ]})),
      }),
      vuoto(40),
      centrato('Superando con esito positivo la verifica finale dell\'apprendimento,'),
      centrato('effettuata in data ___/___/_____ secondo quanto previsto dall\'Accordo.', {after:20}),
      centrato(nota, {sz:19, color:C.GRIGIO, after:40}),
      new Paragraph({spacing:{after:40}, children:[new TextRun({text:'Luogo e data: ______________________',font:FONT,size:20})]}),
      vuoto(40),
      tabellaFirme2col('Firma del Soggetto Formatore / Datore di Lavoro','Firma del Relatore / Datore di Lavoro / RSPP'),
    ];
  }

  const sec1 = buildSection('FORMAZIONE GENERALE','4 ore generali','Il presente attestato costituisce credito formativo permanente e ha validità su tutto il territorio nazionale.');
  const sec2 = buildSection(`FORMAZIONE SPECIFICA RISCHIO ${mansione.livello}`,`${mansione.oreSpec} ore specifiche`,'Il presente attestato ha validità di 5 anni su tutto il territorio nazionale.');

  const doc = new Document({styles:docStyles, sections:[
    {properties:{page:{size:A4_P,margin:MARGIN_ATT}},footers:{default:footer},children:sec1},
    {properties:{page:{size:A4_P,margin:MARGIN_ATT}},footers:{default:footer},children:sec2},
  ]});
  await salvaDoc(doc, `${OUT}/05 - ATTESTATI/Attestato_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTESTATO AGGIORNAMENTO
// ─────────────────────────────────────────────────────────────────────────────
async function genAttestatiAggiornamento() {
  const MARGIN_ATT = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
  const footer = new Footer({ children: [new Paragraph({
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});
  const wL = 3800; const wR = W - wL; const half = Math.floor(W/2);

  const children = [
    centrato('ATTESTATO DI FORMAZIONE', {bold:true, sz:40, color:C.BLU_DARK, after:20}),
    centrato('Aggiornamento – D.Lgs. 81/2008 – ASR 17/04/2025', {sz:22, color:C.BLU_MED, after:40}),
    centrato(`Il/La sottoscritto/a ${CLIENTE.datoreLavoro} in qualità di Datore di Lavoro e Soggetto Formatore`),
    centrato(`della Società ${CLIENTE.ragioneSociale}`),
    centrato(`con Codice ATECO ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`),
    centrato('nei confronti del/la lavoratore/rice avente'),
    centrato('Mansione di _________________________________', {bold:true, sz:24, color:C.BLU_DARK, after:40}),
    centrato('Sig./Sig.ra  ____________________________________________'),
    centrato('Nato/a a ________________________ il ___/___/_____'),
    centrato('Codice Fiscale: ___________________________________________', {after:60}),
    centrato('ha frequentato il corso di formazione in materia di Salute e Sicurezza sul Lavoro:', {sz:20, after:40}),
    new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[wL,wR],
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
      rows:[
        ['Tipologia corso:', 'AGGIORNAMENTO FORMAZIONE LAVORATORI'],
        ['Riferimento normativo:', 'ASR 17/04/2025 (D.Lgs. 81/2008, art. 37)\nParte II dell\'Accordo - Punto 2 /2.1 e Parte III'],
        ['Durata:', '6 ore'],
        ['Modalità:', 'Presenza sul campo'],
        ['Data inizio:', '___/___/_____'],
        ['Data fine:', '___/___/_____'],
      ].map(([et,va],i) => new TableRow({children:[
        cella(et,{width:wL,bold:true,fill:i%2===0?C.GRIGIO_ALT:C.BIANCO}),
        cella(va,{width:wR,fill:i%2===0?C.GRIGIO_ALT:C.BIANCO}),
      ]})),
    }),
    vuoto(40),
    centrato('Superando con esito positivo la verifica finale dell\'apprendimento,'),
    centrato('effettuata in data ___/___/_____ secondo quanto previsto dall\'Accordo.', {after:20}),
    centrato('Aggiornamento effettuato ai sensi dell\'art. 37 del D.Lgs. 81/2008.', {sz:19, color:C.GRIGIO}),
    centrato('Il presente attestato ha validità di 5 anni su tutto il territorio nazionale.', {sz:19, color:C.GRIGIO, after:40}),
    new Paragraph({spacing:{after:40}, children:[new TextRun({text:'Luogo e data: ______________________',font:FONT,size:20})]}),
    vuoto(40),
    tabellaFirme2col('Firma del Soggetto Formatore / Datore di Lavoro','Firma del Relatore / Datore di Lavoro / RSPP'),
  ];

  const doc = new Document({styles:docStyles, sections:[{
    properties:{page:{size:A4_P,margin:MARGIN_ATT}},
    footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/05 - ATTESTATI/Attestato_Aggiorn.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERBALE VERIFICA FINALE
// Struttura corretta: NO header, footer azienda+pag, sezioni numerate
// ─────────────────────────────────────────────────────────────────────────────
async function genVerbaleVerifica() {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
  const footer = footerAzienda(true);

  function sez(n, tit) {
    return new Paragraph({ spacing:{before:80,after:20}, children:[new TextRun({text:`${n}. ${tit}`,bold:true,font:FONT,size:22,color:C.BLU_MED})] });
  }
  function riga(txt) {
    return new Paragraph({ spacing:{before:40,after:40}, children:[new TextRun({text:txt,font:FONT,size:20})] });
  }
  function check(txt) {
    return new Paragraph({ spacing:{before:20,after:20}, children:[new TextRun({text:`☐  ${txt}`,font:FONT,size:20})] });
  }

  // Tabella presenti
  const wN=600;const wCF=2800;const wNome=3238;const wAm=1500;const wEs=1500;
  const colW=[wN,wNome,wCF,wAm,wEs];
  const nRighe = 8;

  const children = [
    centrato('VERBALE DI VERIFICA FINALE', {bold:true, sz:32, color:C.BLU_DARK, after:20}),
    centrato('Accordo Stato-Regioni 17/04/2025 (in G.U. il 24/05/2025) – Punto 5', {sz:18, color:C.GRIGIO, after:60}),

    sez(1,'DATI DEL SOGGETTO FORMATORE'),
    tabellaKV([
      ['Denominazione', CLIENTE.ragioneSociale],
      ['Codice Fiscale / P.IVA', CLIENTE.piva],
      ['Responsabile progetto formativo', CLIENTE.datoreLavoro],
      ['Soggetto Formatore / Docente', `${CLIENTE.datoreLavoro} – Datore di Lavoro / RSPP`],
    ]),
    vuoto(60),

    sez(2,'DATI DEL CORSO'),
    new Paragraph({ spacing:{before:20,after:10}, children:[new TextRun({text:'Tipologia e durata:',bold:true,font:FONT,size:20})] }),
    check('Formazione Generale – Durata: 4 ore'),
    ...MANSIONI.map(m => check(`Formazione Specifica – Durata: ${m.oreSpec} ore (Settore rischio ${m.livello}) – Mansione: ${m.nome}`)),
    check('Aggiornamento – Durata: ___ ore'),
    vuoto(20),
    riga('Modalità di erogazione:  ☐ Aula / Sul campo'),
    vuoto(40),

    sez(3,'ELENCO AMMESSI ALLA VERIFICA FINALE ED ESITO'),
    new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:colW,
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
      rows:[
        new TableRow({tableHeader:true, children:[
          cella('N.',{width:wN,bold:true,fill:C.BLU_LIGHT,align:'center'}),
          cella('Cognome e Nome',{width:wNome,bold:true,fill:C.BLU_LIGHT}),
          cella('Codice Fiscale',{width:wCF,bold:true,fill:C.BLU_LIGHT}),
          cella('Ammesso (SI/NO)',{width:wAm,bold:true,fill:C.BLU_LIGHT,align:'center'}),
          cella('Esito (Idoneo / Non idoneo)',{width:wEs,bold:true,fill:C.BLU_LIGHT,align:'center'}),
        ]}),
        ...Array.from({length:nRighe},(_,i) => new TableRow({
          height:{value:500,rule:'atLeast'},
          children:[
            cella(`${i+1}`,{width:wN,align:'center',fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
            cella('',{width:wNome,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
            cella('',{width:wCF,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
            cella('',{width:wAm,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
            cella('',{width:wEs,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
          ],
        })),
      ],
    }),
    vuoto(60),

    sez(4,'MODALITÀ DI VERIFICA'),
    check('Test scritto a risposte multiple'),
    check('Colloquio individuale'),
    check('Prova pratica'),
    vuoto(40),

    sez(5,'LUOGO E DATA DELLA VERIFICA FINALE'),
    riga('Luogo: _____________________________________________________________'),
    riga('Data: _____________________________ - Orario: dalle ______ alle ______'),
    vuoto(40),

    sez(6,'DICHIARAZIONE FINALE'),
    new Paragraph({ alignment:AlignmentType.JUSTIFIED, spacing:{after:40}, children:[new TextRun({text:'Il Responsabile del progetto formativo attesta che la verifica finale è stata svolta nel rispetto delle modalità e dei criteri previsti dall\'Accordo Stato-Regioni 17/04/2025 e che i risultati sono stati comunicati ai partecipanti.',font:FONT,size:20})] }),
    riga('Data: ____ / ____ / ______'),
    vuoto(60),
    new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[W],
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:NO.top,insideV:NO.top},
      rows:[new TableRow({children:[
        new TableCell({width:{size:W,type:WidthType.DXA},borders:NO,margins:{top:60,bottom:60,left:80,right:80},children:[
          new Paragraph({children:[new TextRun({text:'Responsabile del Progetto Formativo',font:FONT,size:20})]}),
          new Paragraph({spacing:{after:40},children:[new TextRun({text:'_________________________________',font:FONT,size:20})]}),
        ]}),
      ]})]
    }),
  ];

  const doc = new Document({styles:docStyles, sections:[{
    properties:{page:{size:A4_P,margin:MARGIN}},
    footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/06 - VERBALE VERIFICA/Verbale_Verifica_Finale.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERIFICA EFFICACIA
// Struttura corretta: NO header, footer azienda (no pag), una sezione per mansione
// ─────────────────────────────────────────────────────────────────────────────
async function genVerificaEfficacia() {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134 };
  const footer = new Footer({ children: [new Paragraph({
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});

  const wVoce=2000;const wCrit=4238;const wEs=1700;const wNote=1700;

  function buildSezione(mansione) {
    // Criteri di osservazione specifici per mansione
    const voci = mansione.rischi.slice(0,4).map(r => ({
      voce: r.nome.length > 30 ? r.nome.substring(0,30)+'…' : r.nome,
      criterio: `Il lavoratore applica correttamente le misure di prevenzione per: ${r.nome.toLowerCase()}.`,
    }));
    if (voci.length < 4) voci.push({voce:'Uso DPI', criterio:`Il lavoratore indossa correttamente i DPI richiesti per la mansione di ${mansione.nome}.`});

    return [
      centrato('VERIFICA DELL\'EFFICACIA DELLA FORMAZIONE', {bold:true, sz:32, color:C.BLU_DARK, after:20}),
      centrato('D.Lgs. 81/2008 – Accordo Stato-Regioni 17/04/2025', {sz:18, color:C.GRIGIO, after:60}),
      tabellaKV([
        ['Corso di riferimento:', `Formazione Generale e Specifica – ${mansione.nome}`],
        ['Data verifica:', '___/___/______'],
        ['Lavoratore verificato:', ''],
        ['Reparto / Area:', mansione.reparto],
        ['Osservatore (nome e qualifica):', ''],
      ], 3200),
      vuoto(40),
      new Paragraph({ spacing:{before:0,after:20}, children:[new TextRun({text:`VALUTAZIONE COMPORTAMENTALE – MANSIONE: ${mansione.nome.toUpperCase()}`,bold:true,font:FONT,size:22,color:C.BLU_DARK})] }),
      new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[wVoce,wCrit,wEs,wNote],
        borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
        rows:[
          new TableRow({tableHeader:true, children:[
            cella('Voce',{width:wVoce,bold:true,fill:C.BLU_LIGHT}),
            cella('Criterio di osservazione',{width:wCrit,bold:true,fill:C.BLU_LIGHT}),
            cella('Esito',{width:wEs,bold:true,fill:C.BLU_LIGHT,align:'center'}),
            cella('Note',{width:wNote,bold:true,fill:C.BLU_LIGHT}),
          ]}),
          ...voci.map((v,i) => new TableRow({
            height:{value:700,rule:'atLeast'},
            children:[
              cella(v.voce,{width:wVoce,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
              cella(v.criterio,{width:wCrit,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
              new TableCell({width:{size:wEs,type:WidthType.DXA},shading:{fill:i%2===0?C.BIANCO:C.GRIGIO_ALT,type:ShadingType.CLEAR},borders:BD,margins:{top:60,bottom:60,left:80,right:80},children:[
                new Paragraph({children:[new TextRun({text:'☐ Adeguato',font:FONT,size:19})]}),
                new Paragraph({children:[new TextRun({text:'☐ Non adeguato',font:FONT,size:19})]}),
              ]}),
              cella('',{width:wNote,fill:i%2===0?C.BIANCO:C.GRIGIO_ALT}),
            ],
          })),
        ],
      }),
      vuoto(40),
      new Paragraph({spacing:{after:20},children:[new TextRun({text:'Esito complessivo:',bold:true,font:FONT,size:20})]}),
      new Paragraph({spacing:{after:40},children:[new TextRun({text:'☐ ADEGUATO     ☐ PARZIALMENTE ADEGUATO     ☐ NON ADEGUATO',bold:true,font:FONT,size:20})]}),
      new Paragraph({spacing:{after:20},children:[new TextRun({text:'Azioni correttive proposte:',font:FONT,size:20})]}),
      new Paragraph({spacing:{after:60},children:[new TextRun({text:'____________________________________________________________',font:FONT,size:20})]}),
      tabellaFirme2col('Firma Osservatore','Firma Datore di Lavoro / RSPP'),
    ];
  }

  const sections = MANSIONI.map((m, i) => ({
    properties: { page:{ size:A4_P, margin:MARGIN }, ...(i>0?{type:'nextPage'}:{}) },
    footers: { default: footer },
    children: buildSezione(m),
  }));

  const doc = new Document({styles:docStyles, sections});
  await salvaDoc(doc, `${OUT}/07 - VERIFICA EFFICACIA/Verifica_Efficacia.docx`);
}

module.exports = { genColloquio, genGradimento, genAttestato, genAttestatiAggiornamento, genVerbaleVerifica, genVerificaEfficacia };
