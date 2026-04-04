'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  TabStopType, SimpleField,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, MARGIN_STD, logoBytes,
  makeHeader, makeFooter, vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';
const W = 9638;

// ── helpers ──────────────────────────────────────────────────────────────────
const BD = {top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'}};
const NO = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};

function footerAziendaPag() {
  return new Footer({ children: [new Paragraph({
    children: [
      new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   Pag. `, size: 16, font: FONT, color: C.GRIGIO }),
      new SimpleField('PAGE'),
    ],
  })]});
}
function footerAziendaNoPag() {
  return new Footer({ children: [new Paragraph({
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});
}

function PAR(txt, opts = {}) {
  return new Paragraph({
    alignment: opts.align !== undefined ? opts.align : AlignmentType.CENTER,
    spacing: { before: 0, after: opts.spA !== undefined ? opts.spA*20 : 0 },
    children: [new TextRun({ text: txt, bold: opts.bold, font: FONT, size: opts.sz ? opts.sz*2 : 20, color: opts.col })],
  });
}

function tabellaKV(righe, wL, wR) {
  return new Table({
    width:{size:wL+wR,type:WidthType.DXA}, columnWidths:[wL,wR],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
    rows: righe.map(([k,v],i) => new TableRow({children:[
      cella(k,{width:wL,bold:true,fill:C.BLU_LIGHT}),
      cella(v,{width:wR}),
    ]})),
  });
}

function tabellaFirme(lab1, lab2) {
  const half = Math.floor(W/2);
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:[half,W-half],
    borders:{top:NO.top,bottom:NO.bottom,left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[new TableRow({children:[
      new TableCell({width:{size:half,type:WidthType.DXA},borders:NO,margins:{top:60,bottom:60,left:0,right:60},children:[
        new Paragraph({children:[new TextRun({text:lab1,font:FONT,size:20})]}),
        new Paragraph({spacing:{after:40},children:[new TextRun({text:'_________________________________',font:FONT,size:20})]}),
      ]}),
      new TableCell({width:{size:W-half,type:WidthType.DXA},borders:NO,margins:{top:60,bottom:60,left:60,right:0},children:[
        new Paragraph({children:[new TextRun({text:lab2,font:FONT,size:20})]}),
        new Paragraph({spacing:{after:40},children:[new TextRun({text:'_________________________________',font:FONT,size:20})]}),
      ]}),
    ]})]
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// COLLOQUIO INDIVIDUALE AGGIORNAMENTO
// Header standard, footer "Pag. X a Y"
// Sezioni bold col=2E75B6 spB=14 spA=6, tabelle fill=D5E8F0 col0 169pt / 313pt col1
// ─────────────────────────────────────────────────────────────────────────────
async function genColloquio(mansione) {

  // ── Header: tabella testo 2 col (5301+4337) – NO logo ──────────────────
  const NO_HDR = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
  const header = new Header({ children: [
    new Table({
      width:{size:9638,type:WidthType.DXA}, columnWidths:[5301,4337],
      borders:{top:NO_HDR.top,bottom:NO_HDR.bottom,left:NO_HDR.left,right:NO_HDR.right,insideH:NO_HDR.top,insideV:NO_HDR.top},
      rows:[new TableRow({children:[
        new TableCell({width:{size:5301,type:WidthType.DXA},borders:NO_HDR,
          margins:{top:80,bottom:80,left:120,right:120},
          children:[new Paragraph({children:[new TextRun({text:CLIENTE.ragioneSociale,bold:true,font:FONT,size:22,color:C.BLU_DARK})]})],
        }),
        new TableCell({width:{size:4337,type:WidthType.DXA},borders:NO_HDR,
          margins:{top:80,bottom:80,left:120,right:120},
          children:[
            new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:'Colloquio Individuale – Aggiornamento',bold:true,font:FONT,size:24,color:C.BLU_HEADER})]}),
            new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:`ATECO: ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`,font:FONT,size:17,color:C.GRIGIO})]}),
          ],
        }),
      ]})]
    }),
  ]});

  // ── Footer: bordo blu top + "Pag. X a Y" ───────────────────────────────
  const footer = new Footer({ children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
    spacing: { before: 80 },
    tabStops: [{ type: TabStopType.RIGHT, position: 9638 }],
    children: [
      new TextRun({ text: 'Pag. ', font: FONT, size: 18 }),
      new SimpleField('PAGE'),
      new TextRun({ text: ' a ', font: FONT, size: 18 }),
      new SimpleField('NUMPAGES'),
    ],
  })]});

  const wL = 3373; const wR = W - wL; // 6265

  // ── Sezione: titolo con bordo inferiore blu ─────────────────────────────
  function SEZ(n, txt) {
    return new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, space: 2, color: '2E75B6' } },
      spacing: { before: 280, after: 120 },
      children: [new TextRun({ text: `${n}. ${txt}`, bold: true, font: FONT, size: 22, color: C.BLU_MED })],
    });
  }

  // ── Check item: spacing corretto ────────────────────────────────────────
  function CHECK(txt) {
    return new Paragraph({ spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: `☐  ${txt}`, font: FONT, size: 20, color: '000000' })],
    });
  }

  // ── Paragrafo vuoto inter-sezione (after SEZ e after tabella) ──────────
  const gapSez = new Paragraph({ spacing: { before: 60 }, children: [] });
  const gapTbl = new Paragraph({ spacing: { before: 120 }, children: [] });

  // ── Tabella KV con colori master: label=1F4E79, value=000000 ──────────
  function kvTable(righe) {
    const BD_KV = {top:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'}};
    return new Table({
      width:{size:wL+wR,type:WidthType.DXA}, columnWidths:[wL,wR],
      borders:{top:BD_KV.top,bottom:BD_KV.bottom,left:BD_KV.left,right:BD_KV.right,insideH:BD_KV.top,insideV:BD_KV.top},
      rows: righe.map(([k,v]) => new TableRow({ children: [
        cella(k, { width:wL, bold:true, fill:C.BLU_LIGHT, color:C.BLU_HEADER }),
        cella(v, { width:wR, color:'000000' }),
      ]})),
    });
  }

  // ── Firme: 2 col 4819+4819, NO borders, bordo inferiore CCCCCC ────────
  const wF = 4819;
  const NO_F = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
  function firmeCol(label) {
    return new TableCell({ width:{size:wF,type:WidthType.DXA}, borders:NO_F,
      margins:{top:80,bottom:80,left:120,right:120},
      children:[
        new Paragraph({children:[]}),
        new Paragraph({children:[]}),
        new Paragraph({children:[]}),
        new Paragraph({children:[]}),
        new Paragraph({children:[new TextRun({text:label,bold:true,font:FONT,color:'000000'})]}),
      ],
    });
  }
  function firmeLinea() {
    return new TableCell({ width:{size:wF,type:WidthType.DXA}, borders:NO_F,
      margins:{top:80,bottom:80,left:120,right:120},
      children:[
        new Paragraph({spacing:{before:200},children:[]}),
        new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:4,space:0,color:'CCCCCC'}},children:[]}),
      ],
    });
  }
  const tableFirme = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wF,wF],
    borders:{top:NO_F.top,bottom:NO_F.bottom,left:NO_F.left,right:NO_F.right,insideH:NO_F.top,insideV:NO_F.top},
    rows:[
      new TableRow({children:[firmeCol('Firma del Datore di Lavoro / RSPP'), firmeCol('Firma del Lavoratore')]}),
      new TableRow({children:[firmeLinea(), firmeLinea()]}),
    ],
  });

  const children = [
    PAR('VERBALE DI COLLOQUIO INDIVIDUALE',{bold:true,sz:16,col:C.BLU_DARK,spA:4}),
    PAR('Accordo Stato-Regioni 17/04/2025 – Parte IV, Punto 6 e 6.3',{col:C.GRIGIO,sz:10,spA:10}),

    SEZ(1,'DATI DEL SOGGETTO FORMATORE'),
    gapSez,
    kvTable([
      ['Denominazione', CLIENTE.ragioneSociale],
      ['Docente', CLIENTE.datoreLavoro],
      ['Qualifica', 'Datore di Lavoro / RSPP'],
    ]),
    gapTbl,

    SEZ(2,'DATI DEL CORSO DI AGGIORNAMENTO'),
    gapSez,
    kvTable([
      ['Modalità', '☐ In presenza    ☐ Videoconferenza sincrona'],
      ['Durata', '___ ore / minuti'],
      ['Data/e svolgimento', '___________________________'],
    ]),
    gapTbl,

    SEZ(3,'DATI DEL PARTECIPANTE'),
    gapSez,
    kvTable([
      ['Cognome e Nome', '___________________________'],
      ['Data di nascita', '___________________________'],
      ['Mansione', mansione.nome],
      ['Reparto / Area', mansione.reparto],
    ]),
    gapTbl,

    SEZ(4,'FINALITÀ E CONTENUTI'),
    gapSez,
    CHECK(`Rischi specifici della mansione (${mansione.nome})`),
    CHECK('Procedure aziendali di sicurezza'),
    CHECK('Gestione emergenze e evacuazione'),
    CHECK('Uso corretto dei DPI'),
    CHECK('Segnalazione pericoli / near miss'),
    CHECK('Addestramento specifico'),
    new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'☐  Altro: _______________________________',font:FONT,size:20,color:'000000'})]}),
    gapTbl,

    SEZ(5,'MODALITÀ DI VERIFICA'),
    gapSez,
    new Paragraph({spacing:{after:0},children:[new TextRun({text:'☐ Domande aperte\t\t☐ Caso pratico\t\t☐ Simulazione\t\t☐ Questionario scritto',font:FONT,size:20,color:'000000'})]}),
    gapTbl,

    SEZ(6,'ESITO'),
    gapSez,
    new Paragraph({spacing:{after:0},children:[new TextRun({text:'☐  IDONEO          ☐  NON IDONEO',bold:true,font:FONT,size:22,color:'000000'})]}),
    gapTbl,

    new Paragraph({spacing:{after:0},children:[new TextRun({text:`Luogo: ${CLIENTE.indirizzo}`,font:FONT,size:20,color:'000000'})]}),
    new Paragraph({spacing:{after:0},children:[new TextRun({text:'Data: ___/___/______',font:FONT,size:20,color:'000000'})]}),
    new Paragraph({children:[]}),
    tableFirme,
  ];

  const doc = new Document({styles:docStyles,sections:[{
    properties:{page:{size:A4_P,margin:MARGIN_STD}},
    headers:{default:header},footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/01. AGGIORNAMENTO/Colloquio_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SCHEDA GRADIMENTO
// Top table 2x2: [vuoto|fill=1F3864 azienda] [vuoto|fill=2E75B6 "QUESTIONARIO"]
// Rating table 20x5: col0=5590dxa, col1-4=1020dxa each
// ─────────────────────────────────────────────────────────────────────────────
async function genGradimento() {
  const MARGIN = { top: 710, right: 1134, bottom: 1134, left: 1134 };
  const footer = footerAziendaPag();

  const wLogo = 3380; const wAz = W - wLogo;
  const topTable = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wLogo,wAz],
    borders:{top:NO.top,bottom:NO.bottom,left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[
      new TableRow({children:[
        new TableCell({width:{size:wLogo,type:WidthType.DXA},borders:NO,children:[new Paragraph({children:[]})]}),
        new TableCell({width:{size:wAz,type:WidthType.DXA},borders:NO,shading:{fill:C.BLU_DARK,type:ShadingType.CLEAR},margins:{top:40,bottom:40,left:120,right:120},children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:CLIENTE.ragioneSociale,bold:true,font:FONT,size:20,color:C.BIANCO})]}),
        ]}),
      ]}),
      new TableRow({children:[
        new TableCell({width:{size:wLogo,type:WidthType.DXA},borders:NO,children:[new Paragraph({children:[]})]}),
        new TableCell({width:{size:wAz,type:WidthType.DXA},borders:NO,shading:{fill:C.BLU_MED,type:ShadingType.CLEAR},margins:{top:40,bottom:40,left:120,right:120},children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'QUESTIONARIO DI GRADIMENTO',bold:true,font:FONT,size:20,color:C.BIANCO})]}),
        ]}),
      ]}),
    ],
  });

  // Rating table: 5 colonne
  const wDom = 5590; const wN = 1012; // 5590+4*1012=9638
  function domRow(testo, isHeader=false, isSection=false) {
    const fill = isHeader?C.BLU_HEADER:(isSection?C.BLU_LIGHT:undefined);
    const colr = (isHeader||isSection)?C.BIANCO:undefined;
    if (isSection) {
      return new TableRow({children:[
        new TableCell({columnSpan:5,width:{size:W,type:WidthType.DXA},shading:{fill:C.BLU_LIGHT,type:ShadingType.CLEAR},borders:BD,margins:{top:40,bottom:40,left:80,right:80},children:[
          new Paragraph({children:[new TextRun({text:testo,bold:true,font:FONT,size:20,color:C.BLU_DARK})]}),
        ]}),
      ]});
    }
    return new TableRow({children:[
      cella(testo,{width:wDom,bold:isHeader,fill:fill||undefined,color:colr}),
      cella(isHeader?'1':'☐',{width:wN,bold:isHeader,fill:fill||undefined,color:colr,align:'center'}),
      cella(isHeader?'2':'☐',{width:wN,bold:isHeader,fill:fill||undefined,color:colr,align:'center'}),
      cella(isHeader?'3':'☐',{width:wN,bold:isHeader,fill:fill||undefined,color:colr,align:'center'}),
      cella(isHeader?'4':'☐',{width:W-wDom-3*wN,bold:isHeader,fill:fill||undefined,color:colr,align:'center'}),
    ]});
  }

  const ratingTable = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wDom,wN,wN,wN,W-wDom-3*wN],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
    rows:[
      domRow('DOMANDA',true),
      domRow('CONTENUTI',false,true),
      domRow('I contenuti del corso erano pertinenti alle mie mansioni lavorative'),
      domRow('I contenuti erano aggiornati e conformi alla normativa vigente'),
      domRow('La quantità di informazioni fornita era adeguata'),
      domRow('DOCENTE',false,true),
      domRow('Il docente ha esposto gli argomenti in modo chiaro e comprensibile'),
      domRow('Il docente era disponibile a rispondere alle domande'),
      domRow('Gli esempi pratici erano pertinenti alla realtà lavorativa'),
      domRow('ORGANIZZAZIONE',false,true),
      domRow('La durata del corso era adeguata agli argomenti trattati'),
      domRow('Il luogo/modalità di svolgimento era adeguato'),
      domRow('UTILITÀ',false,true),
      domRow('Il corso ha aumentato la mia consapevolezza sui rischi'),
      domRow('Le informazioni ricevute sono applicabili nella mia attività quotidiana'),
      domRow('La verifica finale era coerente con quanto trattato nel corso'),
      domRow('VALUTAZIONE COMPLESSIVA',false,true),
      domRow('Valutazione complessiva del corso'),
      domRow('Consiglierei questo corso ad altri colleghi'),
      domRow('Mi ritengo più preparato/a sulla sicurezza dopo il corso'),
    ],
  });

  const ossTbl = new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:NO.top,insideV:NO.top},
    rows:[new TableRow({children:[cella('OSSERVAZIONI E SUGGERIMENTI',{width:W,bold:true,fill:C.BLU_HEADER,color:C.BIANCO})]})]
  });
  const emptyTbl = () => new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right},
    rows:[new TableRow({height:{value:800,rule:'atLeast'},children:[cella('',{width:W})]})]
  });

  const children = [
    topTable, vuoto(20),
    PAR('QUESTIONARIO DI VALUTAZIONE DEL GRADIMENTO',{bold:true,sz:14,col:C.BLU_DARK,spA:3}),
    PAR('Formazione Obbligatoria – D.Lgs. 81/08',{sz:10,col:C.GRIGIO,spA:10}),
    new Paragraph({spacing:{after:20},children:[new TextRun({text:'Il presente questionario è anonimo e ha lo scopo di migliorare la qualità dei futuri interventi formativi. La preghiamo di compilarlo con sincerità.',font:FONT,size:20})]}),
    new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right},rows:[new TableRow({children:[cella('LEGENDA: 1 = Insufficiente   2 = Sufficiente   3 = Buono   4 = Ottimo',{width:W,bold:true,fill:C.BLU_HEADER,color:C.BIANCO})]})]
    }),
    vuoto(10),
    ratingTable,
    vuoto(20),
    ossTbl,
    new Paragraph({spacing:{before:10,after:6},children:[new TextRun({text:'Cosa ha apprezzato maggiormente?',font:FONT,size:20})]}),
    emptyTbl(),
    new Paragraph({spacing:{before:10,after:6},children:[new TextRun({text:'Cosa migliorereste?',font:FONT,size:20})]}),
    emptyTbl(),
  ];

  const doc = new Document({styles:docStyles,sections:[{
    properties:{page:{size:A4_P,margin:MARGIN}},
    footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/04 - GRADIMENTO/Gradimento.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTESTATO (per mansione)
// NO header, footer azienda (no pag), testo CENTER, tabella 7x2 (col0=2700 col1=6938)
// Due attestati in file unico (generale + specifica)
// ─────────────────────────────────────────────────────────────────────────────
async function genAttestato(mansione) {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 708, footer: 708 };
  // Master: header con logo inline 164x36, footer indirizzo NO pag
  const header = new Header({ children: [new Paragraph({
    children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 164, height: 36 } })],
  })]});
  const footer = new Footer({ children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});
  const wL = 2698; const wR = W - wL; // 6940
  const BD_A = {top:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'}};
  const NO_F = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
  const wFirma = 4819;

  const BF = {top:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'}};
  function firmaCol(label) {
    return new TableCell({ width:{size:wFirma,type:WidthType.DXA}, borders:BF,
      margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:label,bold:true,font:FONT,size:20,color:'000000'})]})],
    });
  }
  function firmaVuota() {
    return new TableCell({ width:{size:wFirma,type:WidthType.DXA}, borders:BF,
      margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:40})]})],
    });
  }
  const tableFirme = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wFirma,wFirma],
    borders:{top:BF.top,bottom:BF.bottom,left:BF.left,right:BF.right,insideH:BF.top,insideV:BF.left},
    rows:[
      new TableRow({children:[firmaCol('Firma del Soggetto Formatore / Datore di Lavoro'), firmaCol('Firma del Relatore / Datore di Lavoro / RSPP')]}),
      new TableRow({children:[firmaVuota(), firmaVuota()]}),
    ],
  });

  // Cella "Riferimento normativo" con 3 paragrafi: riga 1 normale sz20, righe 2-3 italic sz18
  function cellaRifNorm(wR, riga2, riga3) {
    const BD_A2 = {top:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'}};
    return new TableCell({ width:{size:wR,type:WidthType.DXA}, borders:BD_A2,
      shading:{fill:C.BLU_LIGHT,type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:120,right:120},
      children:[
        new Paragraph({children:[new TextRun({text:"ASR 17/04/2025 (D.Lgs. 81/2008, art. 37)",font:FONT,size:20,color:'000000'})]}),
        new Paragraph({children:[new TextRun({text:riga2,font:FONT,size:18,italic:true,color:'000000'})]}),
        new Paragraph({children:[new TextRun({text:riga3,font:FONT,size:18,italic:true,color:'000000'})]}),
      ],
    });
  }

  function buildAttest(tipoLabel, durata, nota) {
    return [
      PAR('ATTESTATO DI FORMAZIONE',{bold:true,sz:20,col:C.BLU_DARK,spA:4}),
      PAR(`Il/La sottoscritto/a ${CLIENTE.datoreLavoro} in qualità di Datore di Lavoro e Soggetto Formatore`,{spA:2}),
      PAR(`della Società ${CLIENTE.ragioneSociale}`,{spA:2}),
      PAR(`con Codice ATECO ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`,{spA:2}),
      PAR('nei confronti del/la lavoratore/rice avente',{spA:2}),
      PAR(`Mansione di ${mansione.nome}`,{bold:true,sz:12,col:C.BLU_DARK,spA:3}),
      PAR('Sig./Sig.ra  ____________________________________________',{spA:2}),
      PAR('Nato/a a ________________________ il ___/___/_____',{spA:2}),
      PAR('Codice Fiscale: ___________________________________________',{spA:3}),
      PAR('ha frequentato il corso di formazione in materia di Salute e Sicurezza sul Lavoro:',{sz:10,spA:5}),
      new Table({width:{size:W,type:WidthType.DXA},columnWidths:[wL,wR],
        borders:{top:BD_A.top,bottom:BD_A.bottom,left:BD_A.left,right:BD_A.right,insideH:BD_A.top,insideV:BD_A.top},
        rows:[
          new TableRow({children:[cella('Tipologia corso:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella(tipoLabel,{width:wR,color:'000000'})]}),
          new TableRow({children:[cella('Riferimento normativo:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cellaRifNorm(wR,"Riferimento normativo e contenuti minimi secondo l'Accordo Stato-Regioni del 17 aprile 2025","Parte II dell'Accordo - Punto 2 e Parte IV dell'Accordo - Punto 1")]}),
          new TableRow({children:[cella('Durata:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella(durata,{width:wR,color:'000000'})]}),
          new TableRow({children:[cella('Modalità:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('Presenza sul campo',{width:wR,color:'000000'})]}),
          new TableRow({children:[cella('Data inizio:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('___/___/_____',{width:wR,color:'000000'})]}),
          new TableRow({children:[cella('Data fine:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('___/___/_____',{width:wR,color:'000000'})]}),
          new TableRow({children:[cella('Sede:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella(CLIENTE.indirizzo,{width:wR,color:'000000'})]}),
        ],
      }),
      vuoto(20),
      PAR('Superando con esito positivo la verifica finale dell\'apprendimento,',{sz:10,spA:2}),
      PAR('effettuata in data ___/___/_____ secondo quanto previsto dall\'Accordo.',{sz:10,spA:5}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100},children:[new TextRun({text:nota,bold:true,italic:true,font:FONT,size:19,color:C.GRIGIO})]}),
      new Paragraph({spacing:{after:100},children:[new TextRun({text:'Luogo e data: ______________________',font:FONT,size:20})]}),
      vuoto(30),
      tableFirme,
    ];
  }

  const sec1 = buildAttest('FORMAZIONE GENERALE','4 ore generali','Il presente attestato costituisce credito formativo permanente e ha validità su tutto il territorio nazionale.');
  const sec2 = buildAttest(`FORMAZIONE SPECIFICA RISCHIO ${mansione.livello}`,`${mansione.oreSpec} ore specifiche`,'Il presente attestato ha validità di 5 anni su tutto il territorio nazionale.');

  const doc = new Document({styles:docStyles,sections:[
    {properties:{page:{size:A4_P,margin:MARGIN}},headers:{default:header},footers:{default:footer},children:sec1},
    {properties:{page:{size:A4_P,margin:MARGIN}},headers:{default:header},footers:{default:footer},children:sec2},
  ]});
  await salvaDoc(doc, `${OUT}/05 - ATTESTATI/Attestato_${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// ATTESTATO AGGIORNAMENTO
// ─────────────────────────────────────────────────────────────────────────────
async function genAttestatiAggiornamento() {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 708, footer: 708 };
  const wL = 2698; const wR = W - wL;
  const header = new Header({ children: [new Paragraph({
    children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 164, height: 36 } })],
  })]});
  const footer = new Footer({ children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});
  const BD_A = {top:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'}};
  const NO_F = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
  const wFirma = 4819;

  const BF = {top:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:4,space:0,color:'AAAAAA'}};
  function firmaCol(label) {
    return new TableCell({ width:{size:wFirma,type:WidthType.DXA}, borders:BF,
      margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:label,bold:true,font:FONT,size:20,color:'000000'})]})],
    });
  }
  function firmaVuota() {
    return new TableCell({ width:{size:wFirma,type:WidthType.DXA}, borders:BF,
      margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:40})]})],
    });
  }
  const tableFirme = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wFirma,wFirma],
    borders:{top:BF.top,bottom:BF.bottom,left:BF.left,right:BF.right,insideH:BF.top,insideV:BF.left},
    rows:[
      new TableRow({children:[firmaCol('Firma del Soggetto Formatore / Datore di Lavoro'), firmaCol('Firma del Relatore / Datore di Lavoro / RSPP')]}),
      new TableRow({children:[firmaVuota(), firmaVuota()]}),
    ],
  });

  const children = [
    PAR('ATTESTATO DI FORMAZIONE',{bold:true,sz:20,col:C.BLU_DARK,spA:4}),
    PAR(`Il/La sottoscritto/a ${CLIENTE.datoreLavoro} in qualità di Datore di Lavoro e Soggetto Formatore`,{spA:2}),
    PAR(`della Società ${CLIENTE.ragioneSociale}`,{spA:2}),
    PAR(`con Codice ATECO ${CLIENTE.atecoCodice} – ${CLIENTE.atecoDesc}`,{spA:2}),
    PAR('nei confronti del/la lavoratore/rice avente',{spA:2}),
    PAR('Mansione di _________________________________',{bold:true,sz:12,col:C.BLU_DARK,spA:3}),
    PAR('Sig./Sig.ra  ____________________________________________',{spA:2}),
    PAR('Nato/a a ________________________ il ___/___/_____',{spA:2}),
    PAR('Codice Fiscale: ___________________________________________',{spA:3}),
    PAR('ha frequentato il corso di formazione in materia di Salute e Sicurezza sul Lavoro:',{sz:10,spA:5}),
    new Table({width:{size:W,type:WidthType.DXA},columnWidths:[wL,wR],
      borders:{top:BD_A.top,bottom:BD_A.bottom,left:BD_A.left,right:BD_A.right,insideH:BD_A.top,insideV:BD_A.top},
      rows:[
        new TableRow({children:[cella('Tipologia corso:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('AGGIORNAMENTO FORMAZIONE LAVORATORI',{width:wR,color:'000000'})]}),
        new TableRow({children:[cella('Riferimento normativo:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}),
          new TableCell({width:{size:wR,type:WidthType.DXA},borders:{top:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:4,color:'AAAAAA'}},shading:{fill:C.BLU_LIGHT,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},children:[
            new Paragraph({children:[new TextRun({text:"ASR 17/04/2025 (D.Lgs. 81/2008, art. 37)",font:FONT,size:20,color:'000000'})]}),
            new Paragraph({children:[new TextRun({text:"Parte II dell'Accordo - Punto 2 /2.1 e Parte IV dell'Accordo - Punto 6 e 6.3",font:FONT,size:18,italic:true,color:'000000'})]}),
          ]})
        ]}),
        new TableRow({children:[cella('Durata:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('6 ore',{width:wR,color:'000000'})]}),
        new TableRow({children:[cella('Modalità:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('Presenza sul campo',{width:wR,color:'000000'})]}),
        new TableRow({children:[cella('Data inizio:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('___/___/_____',{width:wR,color:'000000'})]}),
        new TableRow({children:[cella('Data fine:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella('___/___/_____',{width:wR,color:'000000'})]}),
        new TableRow({children:[cella('Sede:',{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}), cella(CLIENTE.indirizzo,{width:wR,color:'000000'})]}),
      ],
    }),
    vuoto(20),
    PAR('Superando con esito positivo la verifica finale dell\'apprendimento,',{sz:10,spA:2}),
    PAR('effettuata in data ___/___/_____ secondo quanto previsto dall\'Accordo.',{sz:10,spA:5}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100},children:[new TextRun({text:'Aggiornamento effettuato ai sensi dell\'art. 37 del D.Lgs. 81/2008.',bold:true,italic:true,font:FONT,size:19,color:C.GRIGIO})]}),
    new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:100},children:[new TextRun({text:'Il presente attestato ha validità di 5 anni su tutto il territorio nazionale.',bold:true,italic:true,font:FONT,size:19,color:C.GRIGIO})]}),
    new Paragraph({spacing:{after:100},children:[new TextRun({text:'Luogo e data: ______________________',font:FONT,size:20})]}),
    vuoto(30),
    tableFirme,
  ];

  const doc = new Document({styles:docStyles,sections:[{
    properties:{page:{size:A4_P,margin:MARGIN}},headers:{default:header},footers:{default:footer},children,
  }]});
  await salvaDoc(doc, `${OUT}/05 - ATTESTATI/Attestato_Aggiorn.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERBALE VERIFICA FINALE
// NO header, footer azienda+pag, titolo CENTER, sezioni bold col=2E75B6 spB=10 spA=5
// ─────────────────────────────────────────────────────────────────────────────
async function genVerbaleVerifica() {
  const MARGIN = { top: 1134, right: 1134, bottom: 1134, left: 1134, header: 426 };
  const header = new Header({ children: [new Paragraph({
    children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 164, height: 36 } })],
  })]});
  const footer = footerAziendaPag();
  const wL = 3539; const wR = W - wL; // 6099

  function SEZ(n, txt) {
    return new Paragraph({spacing:{before:200,after:100},children:[new TextRun({text:`${n}. ${txt}`,bold:true,font:FONT,color:C.BLU_MED})]});
  }
  function CHECK(txt) {
    return new Paragraph({spacing:{before:0,after:60},children:[new TextRun({text:`☐  ${txt}`,font:FONT,size:20})]});
  }
  function RIGA(txt) {
    return new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:txt,font:FONT,size:20})]});
  }

  // Elenco: col widths dal master [464, 2384, 1986, 2403, 2401]
  const wN=464;const wNome=2384;const wCF=3243;const wAm=1417;const wEs=2130;
  const elencoTable = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wN,wNome,wCF,wAm,wEs],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
    rows:[
      new TableRow({tableHeader:true,children:[
        cella('N.',{width:wN,bold:true,fill:C.BLU_HEADER,color:C.BIANCO,align:'center'}),
        cella('Cognome e Nome',{width:wNome,bold:true,fill:C.BLU_HEADER,color:C.BIANCO}),
        cella('Codice Fiscale',{width:wCF,bold:true,fill:C.BLU_HEADER,color:C.BIANCO}),
        cella('Ammesso (SI/NO)',{width:wAm,bold:true,fill:C.BLU_HEADER,color:C.BIANCO,align:'center'}),
        new TableCell({width:{size:wEs,type:WidthType.DXA},
          borders:{top:{style:BorderStyle.SINGLE,size:4,color:'1F4E79'},bottom:{style:BorderStyle.SINGLE,size:4,color:'1F4E79'},left:{style:BorderStyle.SINGLE,size:4,color:'1F4E79'},right:{style:BorderStyle.SINGLE,size:4,color:'1F4E79'}},
          shading:{fill:'1F4E79',type:ShadingType.CLEAR},
          margins:{top:80,bottom:80,left:100,right:100},
          verticalAlign:VerticalAlign.CENTER,
          children:[
            new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'Esito ',bold:true,font:FONT,size:18,color:'FFFFFF'})]}),
            new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'(Idoneo / Non idoneo)',bold:true,font:FONT,size:18,color:'FFFFFF'})]}),
          ],
        }),
      ]}),
      ...Array.from({length:8},(_,i) => new TableRow({height:{value:454,rule:'atLeast'},children:[
        cella(`${i+1}`,{width:wN,align:'center'}),
        cella('',{width:wNome}),
        cella('',{width:wCF}),
        cella('',{width:wAm}),
        cella('',{width:wEs}),
      ]})),
    ],
  });

  // Firma: single column 7132 DXA dal master
  const wFirma = 7132;
  const firmaTable = new Table({width:{size:wFirma,type:WidthType.DXA},columnWidths:[wFirma],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:NO.top,insideV:NO.top},
    rows:[
      new TableRow({children:[new TableCell({width:{size:wFirma,type:WidthType.DXA},
        borders:{top:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'}},
        margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Responsabile del Progetto Formativo',bold:true,font:FONT,size:20})]})],
      })]}),
      new TableRow({children:[new TableCell({width:{size:wFirma,type:WidthType.DXA},
        borders:{top:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'}},
        margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:50})]})],
      })]}),
    ],
  });

  const children = [
    PAR('VERBALE DI VERIFICA FINALE',{bold:true,sz:16,col:C.BLU_DARK,spA:3}),
    PAR('Accordo Stato-Regioni 17/04/2025 (in G.U. il 24/05/2025) – Punto 5',{sz:10,col:C.GRIGIO,spA:10}),

    SEZ(1,'DATI DEL SOGGETTO FORMATORE'),
    new Table({width:{size:W,type:WidthType.DXA},columnWidths:[wL,wR],
      borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
      rows:[
        ['Denominazione',CLIENTE.ragioneSociale],
        ['Codice Fiscale / P.IVA',CLIENTE.piva],
        ['Responsabile progetto formativo',CLIENTE.datoreLavoro],
        ['Soggetto Formatore / Docente',`${CLIENTE.datoreLavoro} – Datore di Lavoro / RSPP`],
      ].map(([k,v]) => new TableRow({children:[cella(k,{width:wL,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}),cella(v,{width:wR,color:'000000'})]})),
    }),
    vuoto(30),

    SEZ(2,'DATI DEL CORSO'),
    new Paragraph({spacing:{after:80},children:[new TextRun({text:'Tipologia e durata:',bold:true,font:FONT,size:20})]}),
    CHECK('Formazione Generale – Durata: 4 ore'),
    ...MANSIONI.map(m => CHECK(`Formazione Specifica – Durata: ${m.oreSpec} ore (Settore rischio ${m.livello}) – Mansione: ${m.nome}`)),
    CHECK('Aggiornamento – Durata: ___ ore'),
    CHECK('Formazione Aggiuntiva Preposti – Durata: ___ ore'),
    CHECK('Aggiornamento Preposti – Durata: ___ ore'),
    new Paragraph({spacing:{before:80,after:100},children:[new TextRun({text:'Modalità di erogazione:  ☐ Aula / Sul campo',font:FONT,size:20})]}),

    SEZ(3,'ELENCO AMMESSI ALLA VERIFICA FINALE ED ESITO'),
    elencoTable,
    vuoto(30),

    SEZ(4,'MODALITÀ DI VERIFICA'),
    CHECK('Test scritto a risposte multiple'),
    CHECK('Colloquio individuale'),
    CHECK('Prova pratica'),
    vuoto(30),

    SEZ(5,'LUOGO E DATA DELLA VERIFICA FINALE'),
    RIGA('Luogo: _____________________________________________________________'),
    RIGA('Data: _____________________________ - Orario: dalle ______ alle ______'),
    vuoto(30),

    SEZ(6,'DICHIARAZIONE FINALE'),
    new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:20},children:[new TextRun({text:'Il Responsabile del progetto formativo attesta che la verifica finale è stata svolta nel rispetto dell\'Accordo Stato-Regioni 17/04/2025 e che gli esiti sono stati documentati e archiviati.',font:FONT,size:20})]}),
    RIGA('Data: ____ / ____ / ______'),
    vuoto(40),
    firmaTable,
  ];

  const doc = new Document({styles:docStyles,sections:[{
    properties:{page:{size:A4_P,margin:MARGIN}},
    headers:{default:header},footers:{default:footer},
    children,
  }]});
  await salvaDoc(doc, `${OUT}/06 - VERBALE VERIFICA/Verbale_Verifica_Finale.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERIFICA EFFICACIA
// NO header, footer azienda (no pag), una sezione per mansione nello stesso doc
// Tabella valutazione 11x4, col widths precisi dal modello
// ─────────────────────────────────────────────────────────────────────────────
async function genVerificaEfficacia() {
  const MARGIN = { top: 1134, right: 1133, bottom: 1134, left: 1134, header: 426, footer: 708 };
  // Master: header logo inline 164×36, footer indirizzo NO pag
  const header = new Header({ children: [new Paragraph({
    children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 164, height: 36 } })],
  })]});
  const footer = new Footer({ children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
    children: [new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}`, size: 16, font: FONT, color: C.GRIGIO })],
  })]});
  const NO_F = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
  const wFirma = 4819;

  const BF_VE = {top:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'}};

  function buildSezione(mansione) {
    const wDati = 3256; const wVal = W - wDati; // 6382
    // col widths valutazione: dal master [2405, 4253, 1742, 1238]


    // 10 voci fisse dai rischi della mansione (+ trasversali)
    const voceFromRischio = (r) => ({
      voce: r.nome.length>22 ? r.nome.substring(0,22)+'…' : r.nome,
      criterio: r.misure && r.misure[0]
        ? `Il lavoratore ${r.misure[0].charAt(0).toLowerCase()}${r.misure[0].slice(1)}?`
        : `Il lavoratore applica correttamente le misure di prevenzione per ${r.nome.toLowerCase()}?`,
    });
    const voceBase = mansione.rischi.map(voceFromRischio);
    const voceTraversali = [
      {voce:'Movimentazione carichi', criterio:`Il lavoratore utilizza la tecnica corretta per sollevare e spostare carichi, piegando le ginocchia e mantenendo la schiena dritta?`},
      {voce:'Postura',                criterio:`Il lavoratore mantiene una postura corretta durante le lavorazioni prolungate e sfrutta le pause previste?`},
      {voce:'Procedure emergenza',    criterio:`Il lavoratore conosce ed applica le procedure di emergenza ed evacuazione e sa usare gli estintori?`},
      {voce:'Segnalazione rischi',    criterio:`Il lavoratore segnala tempestivamente situazioni di pericolo, anomalie o near miss al responsabile?`},
      {voce:'Ordine e pulizia',       criterio:`Il lavoratore mantiene l'area di lavoro in ordine e pulisce regolarmente le attrezzature?`},
    ];
    // Combina fino a 10 voci
    const vociAll = [...voceBase, ...voceTraversali];
    const voci = vociAll.slice(0, 10);
    // Col widths valutazione calibrate sul numero voci e sulla mansione
    const wVoce = 2405;
    const wNote = 1238;
    const wEs_m = 1600;
    const wCrit = W - wVoce - wEs_m - wNote; // adatta criterio al layout

    const tableFirme = new Table({
      width:{size:W,type:WidthType.DXA}, columnWidths:[wFirma,wFirma],
      borders:{top:BF_VE.top,bottom:BF_VE.bottom,left:BF_VE.left,right:BF_VE.right,insideH:BF_VE.top,insideV:BF_VE.left},
      rows:[
        new TableRow({children:[
          new TableCell({width:{size:wFirma,type:WidthType.DXA},borders:BF_VE,margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:'Firma Osservatore',bold:true,font:FONT,size:20})]})]}) ,
          new TableCell({width:{size:wFirma,type:WidthType.DXA},borders:BF_VE,margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:20})]})]})
        ]}),
        new TableRow({children:[
          new TableCell({width:{size:wFirma,type:WidthType.DXA},borders:BF_VE,margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:50})]})]}) ,
          new TableCell({width:{size:wFirma,type:WidthType.DXA},borders:BF_VE,margins:{top:80,bottom:80,left:120,right:120},children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:50})]})]})
        ]}),
      ],
    });

    return [
      PAR('VERIFICA DELL\'EFFICACIA DELLA FORMAZIONE',{bold:true,sz:16,col:C.BLU_DARK,spA:3}),
      PAR('D.Lgs. 81/2008 – Accordo Stato-Regioni 17/04/2025',{sz:10,col:C.GRIGIO,spA:10}),
      new Table({width:{size:W,type:WidthType.DXA},columnWidths:[wDati,wVal],
        borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
        rows:[
          ['Corso di riferimento:',`Formazione Generale e Specifica – ${mansione.nome}`],
          ['Data verifica:','___/___/______'],
          ['Lavoratore verificato:',''],
          ['Reparto / Area:',mansione.reparto],
          ['Osservatore (nome e qualifica):',''],
        ].map(([k,v]) => new TableRow({children:[cella(k,{width:wDati,bold:true,fill:C.BLU_LIGHT,color:C.BLU_HEADER}),cella(v,{width:wVal,color:'000000'})]})),
      }),
      vuoto(20),
      new Paragraph({spacing:{before:160,after:80},children:[new TextRun({text:`VALUTAZIONE COMPORTAMENTALE – MANSIONE: ${mansione.nome.toUpperCase()}`,bold:true,font:FONT,color:C.BLU_DARK})]}),
      new Table({width:{size:W,type:WidthType.DXA},columnWidths:[wVoce,wCrit,wEs_m,wNote],
        borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
        rows:[
          new TableRow({tableHeader:true,children:[
            cella('Voce',{width:wVoce,bold:true,fill:C.BLU_HEADER,color:C.BIANCO}),
            cella('Criterio di osservazione',{width:wCrit,bold:true,fill:C.BLU_HEADER,color:C.BIANCO}),
            cella('Esito',{width:wEs,bold:true,fill:C.BLU_HEADER,color:C.BIANCO,align:'center'}),
            cella('Note',{width:wNote,bold:true,fill:C.BLU_HEADER,color:C.BIANCO}),
          ]}),
          ...voci.map((v,i) => new TableRow({height:{value:650,rule:'atLeast'},children:[
            cella(v.voce,{width:wVoce}),
            cella(v.criterio,{width:wCrit}),
            new TableCell({width:{size:wEs_m,type:WidthType.DXA},borders:BD,margins:{top:60,bottom:60,left:80,right:80},children:[
              new Paragraph({children:[new TextRun({text:'☐ Adeguato',font:FONT,size:18})]}),
              new Paragraph({children:[new TextRun({text:'☐ Non adeguato',font:FONT,size:18})]}),
            ]}),
            cella('',{width:wNote}),
          ]})),
        ],
      }),
      vuoto(20),
      new Paragraph({spacing:{before:160,after:80},children:[new TextRun({text:'Esito complessivo:',bold:true,font:FONT,size:20})]}),
      new Paragraph({spacing:{after:100},children:[new TextRun({text:'☐ ADEGUATO     ☐ PARZIALMENTE ADEGUATO     ☐ NON ADEGUATO',bold:true,font:FONT,size:20})]}),
      new Paragraph({spacing:{after:60},children:[new TextRun({text:'Azioni correttive proposte:',font:FONT,size:20})]}),
      new Paragraph({spacing:{after:200},children:[new TextRun({text:'____________________________________________________________',font:FONT,size:20})]}),
      tableFirme,
    ];
  }

  const sections = MANSIONI.map((m) => ({
    properties:{page:{size:A4_P,margin:MARGIN}},
    headers:{default:header},
    footers:{default:footer},
    children: buildSezione(m),
  }));

  const doc = new Document({styles:docStyles, sections});
  await salvaDoc(doc, `${OUT}/07 - VERIFICA EFFICACIA/Verifica_Efficacia.docx`);
}

module.exports = { genColloquio, genGradimento, genAttestato, genAttestatiAggiornamento, genVerbaleVerifica, genVerificaEfficacia };
