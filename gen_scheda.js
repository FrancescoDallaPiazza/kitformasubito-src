'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageOrientation,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, A4_L, logoBytes,
  vuoto, cella, salvaDoc,
} = h;

const OUT = `/home/claude/kit/OUT/KIT FORMASUBITO - ${CLIENTE.ragioneSocialeBreve}`;

const BD = {top:{style:BorderStyle.SINGLE,size:1,color:'BBBBBB'},bottom:{style:BorderStyle.SINGLE,size:1,color:'BBBBBB'},left:{style:BorderStyle.SINGLE,size:1,color:'BBBBBB'},right:{style:BorderStyle.SINGLE,size:1,color:'BBBBBB'}};
const NO = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};

// ─────────────────────────────────────────────────────────────────────────────
// SCHEDA MANSIONE
// Landscape A4, margini 1.27cm (720 DXA), NO header, NO footer
// Struttura dal modello:
//   [top bar: logo (no fill) | fill=1F3864 titolo | fill=1F4E79 info]
//   [rischi table: header fill=1F4E79, righe alt FFFFFF/F2F2F2, 4 colonne]
//   [DPI bar: fill=1F3864 label | fill=D5E8F0 content]
//   [revisione: no fill]
// ─────────────────────────────────────────────────────────────────────────────
async function genSchedaMansione(mansione) {
  // Landscape A4
  const { PageOrientation: PO } = require('docx');
  const PAGE = { width: 11906, height: 16838, orientation: PO.LANDSCAPE };
  const MARGIN = { top: 540, right: 540, bottom: 540, left: 540 };
  const W = 15758; // 16838 - 540*2

  const LIV_COL = { ALTO: 'C00000', MEDIO: 'E07000', BASSO: '538135' };
  const LIV_FILL= { ALTO: 'FFE8E8', MEDIO: 'FFF3E0', BASSO: 'E8F5E9' };

  const BDt = {
    top:    { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    left:   { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    right:  { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    insideH:{ style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    insideV:{ style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  };
  const BDc = {
    top:    { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    left:   { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
    right:  { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' },
  };
  const NO_BDR = {
    top:    { style: BorderStyle.NONE },
    bottom: { style: BorderStyle.NONE },
    left:   { style: BorderStyle.NONE },
    right:  { style: BorderStyle.NONE },
  };

  function tc(children, { w, fill, span, vAlign, borders } = {}) {
    return new TableCell({
      width:         w ? { size: w, type: WidthType.DXA } : undefined,
      columnSpan:    span,
      verticalAlign: vAlign || VerticalAlign.CENTER,
      shading:       fill ? { fill, type: ShadingType.CLEAR } : undefined,
      borders:       borders || BDc,
      margins:       { top: 40, bottom: 40, left: 60, right: 60 },
      children:      Array.isArray(children) ? children : [children],
    });
  }

  function p(text, { sz = 20, bold = false, color, italics = false, align = AlignmentType.LEFT, spB = 0, spA = 0 } = {}) {
    return new Paragraph({
      alignment: align,
      spacing:   { before: spB, after: spA },
      children:  [new TextRun({ text, font: FONT, size: sz, bold, italics, color })],
    });
  }

  // ═══ TAB1 – HEADER (logo | SCHEDA MANSIONE + nome | reparto/rischio/norma) ═══
  const wL = 1900; const wC = 7700; const wR = W - wL - wC;
  const tblHeader = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wL, wC, wR],
    borders: BDt,
    rows: [new TableRow({ children: [
      tc(new Paragraph({
        children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 100, height: 45 } })],
      }), { w: wL, borders: NO_BDR }),
      tc([
        p('SCHEDA MANSIONE', { sz: 18, bold: true, color: C.BLU_HEADER, align: AlignmentType.CENTER }),
        p(mansione.nome.toUpperCase(), { sz: 28, bold: true, color: C.BLU_DARK, align: AlignmentType.CENTER }),
      ], { w: wC, borders: NO_BDR }),
      tc([
        p(`Reparto: ${mansione.reparto}`, { sz: 16, align: AlignmentType.RIGHT }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({ text: 'Rischio: ', font: FONT, size: 16 }),
            new TextRun({ text: mansione.livello, font: FONT, size: 16, bold: true,
              color: LIV_COL[mansione.livello] || '000000' }),
          ],
        }),
        p('D.Lgs. 81/08 \u2013 Art. 36-37', { sz: 14, align: AlignmentType.RIGHT, color: C.GRIGIO }),
      ], { w: wR, borders: NO_BDR }),
    ]})],
  });

  // ═══ TAB2 – RISCHI (4 colonne: rischio | livello | misure | DPI) ═══
  const Cr = 3500; const Cl = 2100; const Cm = 6400; const Cd = W - Cr - Cl - Cm;

  const hdrRow = new TableRow({ children: [
    tc(p('RISCHIO IDENTIFICATO', { sz: 16, bold: true, color: C.BIANCO, align: AlignmentType.CENTER }),
       { w: Cr, fill: C.BLU_DARK }),
    tc(p('LIVELLO RISCHIO',       { sz: 16, bold: true, color: C.BIANCO, align: AlignmentType.CENTER }),
       { w: Cl, fill: C.BLU_DARK }),
    tc(p('MISURE DI PREVENZIONE', { sz: 16, bold: true, color: C.BIANCO, align: AlignmentType.CENTER }),
       { w: Cm, fill: C.BLU_DARK }),
    tc(p('DPI RICHIESTI',         { sz: 16, bold: true, color: C.BIANCO, align: AlignmentType.CENTER }),
       { w: Cd, fill: C.BLU_DARK }),
  ]});

  const rischioRows = mansione.rischi.map((r) => {
    const lv   = r.livello || mansione.livello;
    const lCol = LIV_COL[lv]  || '000000';
    const lFill= LIV_FILL[lv] || 'FFFFFF';

    const misurePars = r.misure.map(m => new Paragraph({
      spacing: { before: 0, after: 20 },
      children: [
        new TextRun({ text: '\u2022 ', font: FONT, size: 16 }),
        new TextRun({ text: m, font: FONT, size: 16 }),
      ],
    }));

    const dpiPars = r.dpi.length
      ? r.dpi.map(d => new Paragraph({
          spacing: { before: 0, after: 20 },
          children: [
            new TextRun({ text: '\u2022 ', font: FONT, size: 16 }),
            new TextRun({ text: d, font: FONT, size: 16 }),
          ],
        }))
      : [p('\u2014', { color: 'AAAAAA', align: AlignmentType.CENTER })];

    return new TableRow({ children: [
      // C1 – Rischio bold
      tc(p(r.nome, { sz: 16, bold: true }), { w: Cr }),
      // C2 – Livello con icona triangolo + testo colorato
      tc([
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 20 },
          children: [new TextRun({ text: '\u26a0', font: FONT, size: 22, bold: true, color: lCol })],
        }),
        p(lv, { sz: 14, bold: true, color: lCol, align: AlignmentType.CENTER }),
      ], { w: Cl }),
      // C3 – Misure
      tc(misurePars, { w: Cm }),
      // C4 – DPI
      tc(dpiPars, { w: Cd }),
    ]});
  });

  const tblRischi = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [Cr, Cl, Cm, Cd],
    borders: BDt,
    rows: [hdrRow, ...rischioRows],
  });

  // ═══ TAB3 – DPI OBBLIGATORI ═══
  const wDpiL = 3500; const wDpiR = W - wDpiL;
  const dpiTxt = mansione.dpi.join('   |   ');
  const tblDpi = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wDpiL, wDpiR],
    borders: BDt,
    rows: [new TableRow({ children: [
      tc(p('DPI OBBLIGATORI', { sz: 16, bold: true, color: C.BIANCO }),
         { w: wDpiL, fill: C.BLU_MED }),
      tc(p(dpiTxt, { sz: 16 }), { w: wDpiR }),
    ]})],
  });

  // ═══ TAB4 – FORMAZIONE + FIRMA ═══
  const wF1 = 2800; const wF2 = 2800; const wF3 = W - wF1 - wF2;
  const tblForm = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wF1, wF2, wF3],
    borders: BDt,
    rows: [new TableRow({ children: [
      tc([
        p('FORMAZIONE SPECIFICA', { sz: 16, color: C.BIANCO }),
        p(`${mansione.oreSpec} ORE`, { sz: 22, bold: true, color: C.BIANCO }),
        p(`Rischio ${mansione.livello}`, { sz: 14, color: 'BDD7EE' }),
      ], { w: wF1, fill: C.BLU_MED }),
      tc([
        p('AGGIORNAMENTO', { sz: 16, color: C.BLU_DARK }),
        p('6 ORE ogni 5 anni', { sz: 18, bold: true, color: C.BLU_DARK }),
      ], { w: wF2, fill: C.BLU_LIGHT }),
      tc([
        p('Firma Datore di Lavoro / RSPP', { sz: 16, bold: true, spB: 20 }),
        p(CLIENTE.datoreLavoro, { sz: 16, spB: 120 }),
        p('_______________________________', { sz: 16 }),
      ], { w: wF3 }),
    ]})],
  });

  // ═══ TAB5 – REVISIONE ═══
  const wRev1 = 11000; const wRev2 = W - wRev1;
  const revTxt = 'Revisione, aggiornamento e consegna al lavoratore: il presente documento è soggetto a revisione periodica o in seguito a variazioni delle mansioni, dell\'organizzazione del lavoro o dei rischi presenti.';
  const tblRev = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wRev1, wRev2],
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED },
      bottom: { style: BorderStyle.NONE },
      left:   { style: BorderStyle.NONE },
      right:  { style: BorderStyle.NONE },
      insideH:{ style: BorderStyle.NONE },
      insideV:{ style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED },
    },
    rows: [new TableRow({ children: [
      tc(p(revTxt, { sz: 14, color: C.GRIGIO }), {
        w: wRev1,
        borders: { top:{ style:BorderStyle.NONE }, bottom:{ style:BorderStyle.NONE }, left:{ style:BorderStyle.NONE }, right:{ style:BorderStyle.NONE } },
      }),
      tc(p(`${CLIENTE.ragioneSocialeBreve} \u2013 Rev. ${CLIENTE.anno}`, {
        sz: 16, bold: true, color: C.BLU_HEADER, align: AlignmentType.RIGHT,
      }), {
        w: wRev2,
        borders: { top:{ style:BorderStyle.NONE }, bottom:{ style:BorderStyle.NONE }, left:{ style:BorderStyle.NONE }, right:{ style:BorderStyle.NONE } },
      }),
    ]})],
  });

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: PAGE, margin: MARGIN } },
      children: [tblHeader, vuoto(10), tblRischi, vuoto(8), tblDpi, vuoto(8), tblForm, vuoto(8), tblRev],
    }],
  });

  await salvaDoc(doc, `${OUT}/01 - SCHEDE MANSIONI/${mansione.id}.docx`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SCHEDA ADDESTRATIVA
// Portrait A4, margini T1.04 R2.0 B0.98 L2.0 (DXA: T590 R1134 B554 L1134)
// NO header, NO footer
// Title stile "Title" (sz=11pt)
// Main table: header fill=F7CAAC, righe con contenuto specifico
// Firma table: header fill=F7CAAC, 4 colonne, 11 righe
// ─────────────────────────────────────────────────────────────────────────────
async function genSchedaAddestrativa(mansione) {
  // Margini corretti: header=567 (~1cm dal bordo), top=1417 (~2.5cm corpo), footer=400
  const MARGIN = { top: 1417, right: 1134, bottom: 556, left: 1134, header: 567, footer: 400 };
  const W = 9638;
  // Header: logo inline 131×29 px
  const { Header: HdrCls, Footer: FtrCls, SimpleField } = require('docx');
  const header = new HdrCls({ children: [new Paragraph({
    children: [new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 131, height: 29 } })],
  })]});
  // Footer: Pag. X right-aligned, no border
  const footer = new FtrCls({ children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [
      new TextRun({ text: 'Pag. ', font: FONT }),
      new SimpleField('PAGE'),
    ],
  })]});

  const BD_A = {top:{style:BorderStyle.SINGLE,size:1,color:'AAAAAA'},bottom:{style:BorderStyle.SINGLE,size:1,color:'AAAAAA'},left:{style:BorderStyle.SINGLE,size:1,color:'AAAAAA'},right:{style:BorderStyle.SINGLE,size:1,color:'AAAAAA'}};
  const SALMON = 'F7CAAC';
  const PURPLE = 'CCC0D9';

  // Colonne dal master: [1413, 3347, 1883, 2985]
  const COLS = [1413, 3347, 1883, 2985];

  function hdrFull(txt, fill) {
    return new TableRow({children:[
      new TableCell({columnSpan:4, width:{size:W,type:WidthType.DXA},
        shading:{fill,type:ShadingType.CLEAR}, borders:BD_A,
        margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
          new TextRun({text:txt,bold:true,italics:true,font:FONT,size:20,color:'000000'})
        ]})],
      }),
    ]});
  }

  function twoCell(txt1, txt2) {
    return new TableRow({children:[
      new TableCell({columnSpan:2,width:{size:COLS[0]+COLS[1],type:WidthType.DXA},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:txt1,font:FONT,size:20})]})]
      }),
      new TableCell({columnSpan:2,width:{size:COLS[2]+COLS[3],type:WidthType.DXA},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:txt2,font:FONT,size:20})]})]
      }),
    ]});
  }

  function fullRow(children, fill) {
    return new TableRow({children:[
      new TableCell({columnSpan:4,width:{size:W,type:WidthType.DXA},
        shading:fill?{fill,type:ShadingType.CLEAR}:undefined,
        borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children,
      }),
    ]});
  }

  function istrRow(txt, fill) {
    return new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,
        margins:{top:40,bottom:40,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'',font:FONT,size:20})]})]
      }),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        shading:fill?{fill,type:ShadingType.CLEAR}:undefined,
        borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:txt,font:FONT,size:20})]})]
      }),
    ]});
  }

  const attPrincipale = mansione.rischi.length > 0
    ? `UTILIZZO CORRETTO ED IN SICUREZZA DELLE ATTREZZATURE – ${mansione.rischi[0].nome.toUpperCase()}`
    : 'UTILIZZO CORRETTO ED IN SICUREZZA DELLE ATTREZZATURE DI LAVORO';

  const dpiTxt = mansione.dpi.slice(0,5).join(', ');

  const mainRows = [
    hdrFull('MANSIONE COINVOLTA', SALMON),
    twoCell('☐ '+mansione.nome, '☐ ___________________________________________'),
    fullRow([new Paragraph({children:[
      new TextRun({text:'Reparto/Area: ',bold:true,italics:true,font:FONT,size:20}),
      new TextRun({text:`${mansione.reparto} ___________________________ - _____________________`,font:FONT,size:20}),
    ]})]),
    fullRow([
      new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'Motivazioni addestramento:',bold:true,italics:true,font:FONT,size:20})]}),
      new Paragraph({children:[new TextRun({text:'☐ Nuova assunzione   ☐ Cambio mansione   ☐ Interinale   ☐ Altra attività di addestramento',font:FONT,size:20})]}),
    ]),
    hdrFull('ATTIVITÀ DI ADDESTRAMENTO DEI LAVORATORI', SALMON),
    fullRow([
      new Paragraph({spacing:{after:6},children:[
        new TextRun({text:'Affiancamento avvenuto con (Cognome/Nome): ',bold:true,font:FONT,size:20}),
        new TextRun({text:CLIENTE.datoreLavoro,font:FONT,size:20}),
      ]}),
      new Paragraph({spacing:{after:6},children:[new TextRun({
        text:`Ruolo: (☐ Datore di Lavoro / ☐ Preposto / ☐ Lavoratore / ☐ Resp. Produz. / ☐ RSPP / ☐ Altro____________________), il quale ha provveduto a fornire adeguato addestramento teorico-pratico, specifico e con riferimenti alla sicurezza e salute sul lavoro all'operatore di cui sopra, rispetto a all'attività specifica di:`,
        font:FONT,size:20,
      })]}),
      new Paragraph({spacing:{after:6},children:[new TextRun({text:attPrincipale,bold:true,font:FONT,size:20})]}),
      new Paragraph({spacing:{after:6},children:[
        new TextRun({text:'Utilizzo della ☐  macchina / ☐ ',font:FONT,size:20}),
        new TextRun({text:'attrezzatura',bold:true,font:FONT,size:20}),
        new TextRun({text:' / ☐ impianto / ☐ procedura di lavoro / ☐ altro _______________',font:FONT,size:20}),
      ]}),
      new Paragraph({spacing:{after:6},children:[
        new TextRun({text:'Durata addestramento ____ ☐ mesi - ☐ ____settimana/e - ☐ ____giorno/i – per un totale di ________ore - ',font:FONT,size:20}),
        new TextRun({text:'10 min',bold:true,font:FONT,size:20}),
      ]}),
      new Paragraph({children:[new TextRun({text:'Al termine dell\'attività si rilascia copia della presente a comprova dell\'attività svolta.',font:FONT,size:20})]}),
    ]),
    // ── R7: header "Al lavoratore..." — C1 vuota, C2 span3 SALMON italic bold ──
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'',font:FONT,size:18})]})]}),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        shading:{fill:SALMON,type:ShadingType.CLEAR},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'Al lavoratore sono state illustrate e consegnate le seguenti informazioni - istruzioni di lavoro:',bold:true,italics:true,font:FONT,size:18})]})],
      }),
    ]}),
    // ── R8: C1=CCC0D9 "Istruzioni di lavoro in sicurezza" | C2 span3 bianca ──
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,
        shading:{fill:'CCC0D9',type:ShadingType.CLEAR},
        verticalAlign:VerticalAlign.CENTER,
        margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'Istruzioni di lavoro in sicurezza',bold:true,font:FONT,size:18})]})]}),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ Utilizzo corretto ed in sicurezza delle attrezzature in dotazione',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ Sicurezze presenti sulle attrezzature in uso (emergenze, microinterruttori, allarmi)',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ Segnaletica di sicurezza, salute ed emergenza in reparto.',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ Istruzioni specifiche di reparto (specificare di seguito se presenti)',font:FONT,size:18})]}),
        ],
      }),
    ]}),
    // ── R9: C1=CCC0D9 "DPI da utilizzare" | C2 span3 bianca ──
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,
        shading:{fill:'CCC0D9',type:ShadingType.CLEAR},
        verticalAlign:VerticalAlign.CENTER,
        margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'DPI da utilizzare',bold:true,font:FONT,size:18})]})]}),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ DPI necessari alla lavorazione (specificare di seguito se necessari):',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:0,after:40},children:[new TextRun({text:dpiTxt,font:FONT,size:18})]}),
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ Rischi per i quali sono necessari i DPI.',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:"☐ Utilizzo dei DPI (modalità d'impiego, verifica della necessità di utilizzo).",font:FONT,size:18})]}),
          new Paragraph({spacing:{before:40,after:40},children:[new TextRun({text:'☐ Modalità di conservazione e richiesta di sostituzione/integrazione dei DPI.',font:FONT,size:18})]}),
        ],
      }),
    ]}),
    // ── R10: C1=CCC0D9 "Istruttori e Preposto" | C2=span2 bianca | C3=SALMON con GIUDIZIO+opzioni ──
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,
        shading:{fill:'CCC0D9',type:ShadingType.CLEAR},
        verticalAlign:VerticalAlign.CENTER,
        margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'Istruttori e Preposto',bold:true,font:FONT,size:18})]})]}),
      new TableCell({columnSpan:2,width:{size:COLS[1]+COLS[2],type:WidthType.DXA},borders:BD_A,
        margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:"Al termine dell'addestramento, effettuato secondo quanto sopra esposto, l'Istruttore e il Preposto valutando in campo le modalità operative e le conoscenze ricevute, ritengono il lavoratore:",font:FONT,size:18})]})]}),
      new TableCell({width:{size:COLS[3],type:WidthType.DXA},borders:BD_A,
        shading:{fill:SALMON,type:ShadingType.CLEAR},
        verticalAlign:VerticalAlign.CENTER,
        margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:'GIUDIZIO',bold:true,font:FONT,size:18})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'☐ Adeguato',font:FONT,size:18})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:0},children:[new TextRun({text:'☐ Non adeguato',font:FONT,size:18})]}),
        ]}),
    ]}),
    // ── R13: Note: fullspan ──
    new TableRow({children:[
      new TableCell({columnSpan:4,width:{size:W,type:WidthType.DXA},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'Note:',bold:true,font:FONT,size:20})]})]}),
    ]}),
  ];

  const mainTable = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:COLS,
    borders:{top:BD_A.top,bottom:BD_A.bottom,left:BD_A.left,right:BD_A.right,insideH:BD_A.top,insideV:BD_A.top},
    rows:mainRows,
  });

  // FIRMA TABLE – colonne dal master: [3399, 1632, 2482, 2126]
  const FCOLS = [3399, 1632, 2482, 2126];
  function fCell(txt, fill, span) {
    return new TableCell({
      ...(span?{columnSpan:span}:{}),
      width:{size:span?FCOLS.slice(0,span).reduce((a,b)=>a+b,0):FCOLS[0],type:WidthType.DXA},
      shading:fill?{fill,type:ShadingType.CLEAR}:undefined,
      borders:BD_A, margins:{top:40,bottom:40,left:80,right:80},
      children:[new Paragraph({children:[new TextRun({text:txt,font:FONT,size:18,bold:!!fill&&fill===SALMON})]})],
    });
  }

  function fCellW(txt, opts={}) {
    const w = opts.w || FCOLS[0];
    const span = opts.span || 1;
    const realW = span === 2 ? FCOLS[0]+FCOLS[1] : w;
    return new TableCell({
      ...(span>1?{columnSpan:span}:{}),
      width:{size:realW,type:WidthType.DXA},
      shading:opts.fill?{fill:opts.fill,type:ShadingType.CLEAR}:undefined,
      borders:BD_A, margins:{top:40,bottom:40,left:80,right:80},
      children:[new Paragraph({alignment:opts.center?AlignmentType.CENTER:AlignmentType.LEFT,
        children:[new TextRun({text:txt,font:FONT,size:opts.sz||20,bold:opts.bold||false,color:opts.col||undefined})],
      })],
    });
  }
  const firmaTable = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:FCOLS,
    borders:{top:BD_A.top,bottom:BD_A.bottom,left:BD_A.left,right:BD_A.right,insideH:BD_A.top,insideV:BD_A.top},
    rows:[
      // Row 1: FIRME header SALMON full width
      new TableRow({children:[new TableCell({columnSpan:4,width:{size:W,type:WidthType.DXA},shading:{fill:SALMON,type:ShadingType.CLEAR},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'FIRME ADDESTRAMENTO SUL CAMPO',bold:true,font:FONT,size:20})]})]})]},),
      // Row 2: Nome lavoratore (colspan 2) + Firma lavoratore (colspan 2)
      new TableRow({height:{value:270,rule:'exact'},children:[
        fCellW('Nome lavoratore:',{span:2,bold:true,col:'000000'}),
        fCellW('Firma lavoratore:',{span:2,bold:true,col:'000000'}),
      ]}),
      // Rows 3-8: empty (6 righe per la firma)
      ...[...Array(6)].map(()=>new TableRow({height:{value:270,rule:'exact'},children:[
        fCellW('',{span:2}), fCellW('',{span:2}),
      ]})),
      // Row 9: separatore SALMON full width
      new TableRow({height:{value:212,rule:'exact'},children:[new TableCell({columnSpan:4,width:{size:W,type:WidthType.DXA},shading:{fill:SALMON,type:ShadingType.CLEAR},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},children:[new Paragraph({children:[new TextRun({text:' ',font:FONT,size:18})]})]})]},),
      // Row 10: Istruttore | Firma Istruttore | Data
      new TableRow({height:{value:420,rule:'exact'},children:[
        new TableCell({columnSpan:2,width:{size:FCOLS[0]+FCOLS[1],type:WidthType.DXA},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},children:[new Paragraph({children:[new TextRun({text:`Istruttore: `,font:FONT,size:20})]})]}),
        fCellW('Firma Istruttore:',{}),
        fCellW('Data:',{w:FCOLS[3]}),
      ]}),
      // Row 11: DDL/RSPP | Firma DDL/RSPP | Data
      new TableRow({height:{value:420,rule:'exact'},children:[
        new TableCell({columnSpan:2,width:{size:FCOLS[0]+FCOLS[1],type:WidthType.DXA},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},children:[new Paragraph({children:[
          new TextRun({text:' DDL/RSPP: ',font:FONT,size:20}),
          new TextRun({text:CLIENTE.datoreLavoro,font:FONT,size:18}),
        ]})]}),
        fCellW('Firma DDL/RSPP: ',{}),
        fCellW('Data:',{w:FCOLS[3]}),
      ]}),
    ],
  });

  const doc = new Document({styles:docStyles,sections:[{
    properties:{page:{size:A4_P,margin:MARGIN}},
    headers:{default:header},
    footers:{default:footer},
    children:[
      new Paragraph({spacing:{after:60},children:[new TextRun({text:'SCHEDA ADDESTRAMENTO SUL CAMPO',bold:true,font:FONT,size:22})]}),
      mainTable,
      new Paragraph({spacing:{after:40},children:[]}),
      firmaTable,
    ],
  }]});
  await salvaDoc(doc, `${OUT}/BONUS - SCHEDA ADDESTRATIVA/Scheda_Addestramento_${mansione.id}.docx`);
}


module.exports = { genSchedaMansione, genSchedaAddestrativa };
