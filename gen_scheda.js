'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
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
  const MARGIN = { top: 720, right: 720, bottom: 720, left: 720 };
  // Content width: 16838 - 2*720 = 15398 DXA
  const W = 15398;

  const livColor = {ALTO:C.ROSSO,MEDIO:C.ARANCIO,BASSO:C.VERDE}[mansione.livello]||C.GRIGIO;

  // Misure colonne da modello (pt → DXA): RISCHIO=169pt=3380, MISURE=293pt=5860, DPI_REQ=169pt=3380, DPI_AZ=W-resto
  const wRis=3380; const wMis=5860; const wDpi=3380; const wAz=W-wRis-wMis-wDpi;

  // TOP BAR 1x3
  const wLogo=1400; const wTit=Math.floor(W*0.53); const wInf=W-wLogo-wTit;
  const topBar = new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wLogo,wTit,wInf],
    borders:{top:{style:BorderStyle.SINGLE,size:4,color:C.BLU_DARK},bottom:{style:BorderStyle.SINGLE,size:4,color:C.BLU_DARK},left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[new TableRow({children:[
      // Logo cell: no fill
      new TableCell({width:{size:wLogo,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,borders:NO,margins:{top:80,bottom:80,left:80,right:120},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:logoBytes,type:'jpg',transformation:{width:60,height:60}})]})],
      }),
      // Title cell: fill=1F3864
      new TableCell({width:{size:wTit,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
        shading:{fill:C.BLU_DARK,type:ShadingType.CLEAR},
        borders:{top:NO.top,bottom:NO.bottom,left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
        margins:{top:80,bottom:80,left:160,right:160},
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:20},children:[new TextRun({text:'SCHEDA MANSIONE',bold:true,font:FONT,size:20,color:C.BIANCO})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:0},children:[new TextRun({text:mansione.nome.toUpperCase(),bold:true,font:FONT,size:26,color:C.BIANCO})]}),
        ],
      }),
      // Info cell: fill=1F4E79
      new TableCell({width:{size:wInf,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
        shading:{fill:C.BLU_HEADER,type:ShadingType.CLEAR},
        borders:NO,margins:{top:80,bottom:80,left:160,right:80},
        children:[
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{after:20},children:[new TextRun({text:`Reparto: ${mansione.reparto}`,font:FONT,size:18,color:C.BIANCO})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{after:20},children:[new TextRun({text:'Rischio: ',font:FONT,size:18,color:C.BIANCO}),new TextRun({text:mansione.livello,bold:true,font:FONT,size:18,color:C.BIANCO})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{after:0},children:[new TextRun({text:'D.Lgs. 81/08 – Art. 36-37',font:FONT,size:16,color:C.BIANCO})]}),
        ],
      }),
    ]})]
  });

  // RISCHI TABLE (header fill=1F4E79, righe alt FFFFFF/F2F2F2)
  function hdCell(txt, w) {
    return new TableCell({width:{size:w,type:WidthType.DXA},shading:{fill:C.BLU_HEADER,type:ShadingType.CLEAR},
      margins:{top:80,bottom:80,left:100,right:80},borders:BD,
      children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:txt,bold:true,font:FONT,size:18,color:C.BIANCO})]})],
    });
  }
  function blt(txt,col) {
    return new Paragraph({spacing:{before:0,after:24},children:[
      new TextRun({text:'• ',font:FONT,size:18,color:col||C.BLU_MED}),
      new TextRun({text:txt,font:FONT,size:18}),
    ]});
  }
  function dc(w,fill,children) {
    return new TableCell({width:{size:w,type:WidthType.DXA},shading:{fill,type:ShadingType.CLEAR},
      margins:{top:60,bottom:60,left:100,right:80},verticalAlign:VerticalAlign.TOP,borders:BD,children});
  }

  const rischioTable = new Table({
    width:{size:W,type:WidthType.DXA},columnWidths:[wRis,wMis,wDpi,wAz],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
    rows:[
      new TableRow({tableHeader:true,children:[hdCell('RISCHIO',wRis),hdCell('MISURE DI PREVENZIONE',wMis),hdCell('DPI RICHIESTI',wDpi),hdCell('DPI AZIENDA',wAz)]}),
      ...mansione.rischi.map((r,i) => {
        const fill = i%2===0 ? C.BIANCO : C.GRIGIO_ALT;
        return new TableRow({children:[
          dc(wRis,fill,[new Paragraph({children:[new TextRun({text:r.nome,bold:true,font:FONT,size:18,color:C.BLU_DARK})]})]),
          dc(wMis,fill,r.misure.map(m=>blt(m,C.BLU_MED))),
          dc(wDpi,fill,r.dpi.length?r.dpi.map(d=>blt(d,C.ARANCIO)):[new Paragraph({children:[new TextRun({text:'—',font:FONT,size:18,color:C.GRIGIO})]})]),
          dc(wAz,fill,[new Paragraph({children:[new TextRun({text:'',font:FONT,size:18})]})]),
        ]});
      }),
    ],
  });

  // DPI BAR: fill=1F3864 label, fill=D5E8F0 content
  const wDpiL=3380;
  const dpiBar = new Table({
    width:{size:W,type:WidthType.DXA},columnWidths:[wDpiL,W-wDpiL],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:NO.top,insideV:BD.top},
    rows:[new TableRow({children:[
      new TableCell({width:{size:wDpiL,type:WidthType.DXA},shading:{fill:C.BLU_DARK,type:ShadingType.CLEAR},borders:BD,margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'DPI OBBLIGATORI',bold:true,font:FONT,size:20,color:C.BIANCO})]})],
      }),
      new TableCell({width:{size:W-wDpiL,type:WidthType.DXA},shading:{fill:C.BLU_LIGHT,type:ShadingType.CLEAR},borders:BD,margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:mansione.dpi.join('   |   '),font:FONT,size:18})]})],
      }),
    ]})]
  });

  // REVISIONE BAR (no fill)
  const wRev=Math.floor(W*0.7);
  const reviBar = new Table({
    width:{size:W,type:WidthType.DXA},columnWidths:[wRev,W-wRev],
    borders:{top:NO.top,bottom:NO.bottom,left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[new TableRow({children:[
      new TableCell({width:{size:wRev,type:WidthType.DXA},borders:NO,margins:{top:60,bottom:40,left:0,right:80},
        children:[new Paragraph({children:[new TextRun({text:'Revisione, aggiornamento e consegna al lavoratore: il presente documento è soggetto a revisione periodica o in seguito a variazioni delle mansioni, dell\'organizzazione del lavoro o dei rischi presenti.',font:FONT,size:16,color:C.GRIGIO})]})],
      }),
      new TableCell({width:{size:W-wRev,type:WidthType.DXA},borders:NO,verticalAlign:VerticalAlign.CENTER,margins:{top:60,bottom:40,left:80,right:0},
        children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:`${CLIENTE.ragioneSocialeBreve} – Rev. ${CLIENTE.anno}`,bold:true,font:FONT,size:16,color:C.BLU_HEADER})]})],
      }),
    ]})]
  });

  const doc = new Document({styles:docStyles,sections:[{
    properties:{page:{size:A4_L,margin:MARGIN}},
    children:[topBar,vuoto(40),rischioTable,vuoto(40),dpiBar,vuoto(20),reviBar],
  }]});
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
          new TextRun({text:txt,bold:true,font:FONT,size:20,color:'000000'})
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
    twoCell(mansione.nome, '___________________________________________'),
    fullRow([new Paragraph({children:[
      new TextRun({text:'Reparto/Area: ',bold:true,font:FONT,size:20}),
      new TextRun({text:`${mansione.reparto} ___________________________ - _____________________`,font:FONT,size:20}),
    ]})]),
    fullRow([
      new Paragraph({children:[new TextRun({text:'Motivazioni addestramento:',bold:true,font:FONT,size:20})]}),
      new Paragraph({children:[new TextRun({text:'☐ Nuova assunzione   ☐ Cambio mansione   ☐ Interinale   ☐ Altra attività di addestramento',font:FONT,size:20})]}),
    ]),
    hdrFull('ATTIVITÀ DI ADDESTRAMENTO DEI LAVORATORI', SALMON),
    fullRow([
      new Paragraph({spacing:{after:6},children:[
        new TextRun({text:'Affiancamento avvenuto con (Cognome/Nome): ',bold:true,font:FONT,size:20}),
        new TextRun({text:CLIENTE.datoreLavoro,font:FONT,size:20}),
      ]}),
      new Paragraph({spacing:{after:6},children:[new TextRun({
        text:`Ruolo: ( Datore di Lavoro /  Preposto /  Lavoratore /  Resp. Produz. /  RSPP /  Altro____________________), il quale ha provveduto a fornire adeguato addestramento teorico-pratico, specifico e con riferimenti alla sicurezza e salute sul lavoro all'operatore di cui sopra, rispetto a all'attività specifica di:`,
        font:FONT,size:20,
      })]}),
      new Paragraph({spacing:{after:6},children:[new TextRun({text:attPrincipale,bold:true,font:FONT,size:20})]}),
      new Paragraph({spacing:{after:6},children:[new TextRun({text:'Utilizzo della   macchina /  attrezzatura /  impianto /  procedura di lavoro /  altro _______________',font:FONT,size:20})]}),
      new Paragraph({spacing:{after:6},children:[new TextRun({text:'Durata addestramento _____ mesi -  _____settimana/e -  _____giorno/i – per un totale di ________ore – 10 min',font:FONT,size:20})]}),
      new Paragraph({children:[new TextRun({text:'Al termine dell\'attività si rilascia copia della presente a comprova dell\'attività svolta.',font:FONT,size:20})]}),
    ]),
    // ── "Al lavoratore..." header + tutti i sottopunti istruzioni in UNA sola cella ──
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'',font:FONT,size:18})]})]}),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        shading:{fill:SALMON,type:ShadingType.CLEAR},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'Al lavoratore sono state illustrate e consegnate le seguenti informazioni - istruzioni di lavoro:',bold:true,font:FONT,size:18})]})],
      }),
    ]}),
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'',font:FONT,size:18})]})]}),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        shading:{fill:PURPLE,type:ShadingType.CLEAR},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Istruzioni di lavoro in sicurezza',bold:true,font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Utilizzo corretto ed in sicurezza delle attrezzature in dotazione',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Sicurezze presenti sulle attrezzature in uso (emergenze, microinterruttori, allarmi)',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Segnaletica di sicurezza, salute ed emergenza in reparto.',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Istruzioni specifiche di reparto (specificare di seguito se presenti)',font:FONT,size:18})]}),
        ],
      }),
    ]}),
    // ── DPI: col[0] vuota + cols[1-3] merged CCC0D9 con "DPI da utilizzare" header ──
    new TableRow({children:[
      new TableCell({width:{size:COLS[0],type:WidthType.DXA},borders:BD_A,margins:{top:40,bottom:40,left:80,right:80},
        children:[new Paragraph({children:[new TextRun({text:'',font:FONT,size:18})]})]}),
      new TableCell({columnSpan:3,width:{size:COLS[1]+COLS[2]+COLS[3],type:WidthType.DXA},
        shading:{fill:PURPLE,type:ShadingType.CLEAR},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'DPI da utilizzare',bold:true,font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'DPI necessari alla lavorazione (specificare di seguito se necessari):',bold:true,font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:dpiTxt,font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Rischi per i quali sono necessari i DPI.',font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:"Utilizzo dei DPI (modalità d'impiego, verifica della necessità di utilizzo).",font:FONT,size:18})]}),
          new Paragraph({spacing:{before:60,after:60},children:[new TextRun({text:'Modalità di conservazione e richiesta di sostituzione/integrazione dei DPI.',font:FONT,size:18})]}),
        ],
      }),
    ]}),
    new TableRow({children:[
      new TableCell({columnSpan:2,width:{size:COLS[0]+COLS[1],type:WidthType.DXA},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({spacing:{after:4},children:[new TextRun({text:"Al termine dell'addestramento, effettuato secondo quanto sopra esposto, l'Istruttore e il Preposto valutando in campo le modalità operative e le conoscenze ricevute, ritengono il lavoratore:",font:FONT,size:18})]}),
        ]
      }),
      new TableCell({width:{size:COLS[2],type:WidthType.DXA},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'GIUDIZIO',bold:true,font:FONT,size:20})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'☐ Adeguato',font:FONT,size:20})]}),
          new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'☐ Non adeguato',font:FONT,size:20})]}),
        ]
      }),
      new TableCell({width:{size:COLS[3],type:WidthType.DXA},borders:BD_A,margins:{top:60,bottom:60,left:80,right:80},
        children:[
          new Paragraph({children:[new TextRun({text:'Note:',bold:true,font:FONT,size:20})]}),
          new Paragraph({children:[new TextRun({text:'',font:FONT,size:20})]}),
        ]
      }),
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
