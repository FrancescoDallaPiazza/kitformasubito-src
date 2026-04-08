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
  // Landscape A4 corretto: width=lungo, height=corto
  const PAGE_L = { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE };
  const MARGIN = { top: 720, right: 720, bottom: 720, left: 720 };
  const W = 15398; // 16838 - 720*2

  const fs_mod = require('fs');
  const path_mod = require('path');
  const { PageOrientation: PO } = require('docx');

  // Colori categoria
  const CAT_FILL = {
    'Sicurezza': 'FFE6E6', 'Chimico / Cancerogeno': 'FFF2CC',
    'Fisico': 'DAE3F3',    'Ergonomico': 'E2EFDA',
    'Elettrico': 'DAE3F3', 'Biologico': 'F5E6FF',
  };
  const ROW_FILL = { ALTO:'FFF5F5', MEDIO:'FFFBF0', BASSO:'F5FFF5' };
  const LIV_COL  = { ALTO:'C00000', MEDIO:'E07000', BASSO:'538135' };

  function catRischio(nome) {
    const n = nome.toLowerCase();
    if (/elettr/.test(n)) return 'Elettrico';
    if (/chimico|toner|alcool|cancerogeno|polvere|solvente|detergente|agente chim|inchiostro/.test(n)) return 'Chimico / Cancerogeno';
    if (/rumore|vibrazion|microclima|calore|freddo|radiazion|illuminaz|campi elettro/.test(n)) return 'Fisico';
    if (/postura|movimentazione|ergono|vdt|videoterminale|sollevamento|ripetitiv|stress/.test(n)) return 'Ergonomico';
    if (/biologico|virus|batterio/.test(n)) return 'Biologico';
    return 'Sicurezza';
  }

  const BDt = {
    top:    {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    bottom: {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    left:   {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    right:  {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    insideH:{style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    insideV:{style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
  };
  const BDc = {
    top:    {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    bottom: {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    left:   {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
    right:  {style:BorderStyle.SINGLE, size:4, color:'AAAAAA'},
  };

  function tc(children, {w, fill, span, vAlign}={}) {
    return new TableCell({
      width: w ? {size:w, type:WidthType.DXA} : undefined,
      columnSpan: span,
      verticalAlign: vAlign || VerticalAlign.CENTER,
      shading: fill ? {fill, type:ShadingType.CLEAR} : undefined,
      borders: BDc,
      margins: {top:60, bottom:60, left:80, right:80},
      children: Array.isArray(children) ? children : [children],
    });
  }

  function p(text, {sz=20, bold=false, color, align=AlignmentType.LEFT, spB=0, italics=false}={}) {
    return new Paragraph({
      alignment: align,
      spacing: {before:0, after:spB},
      children: [new TextRun({text, font:FONT, size:sz, bold, italics, color})],
    });
  }

  // ═══ TABELLA 1 – HEADER AZIENDALE (2 righe) ══════════════════════════════
  const tblHeader = new Table({
    width: {size:W, type:WidthType.DXA},
    columnWidths: [2800, 7200, 5398],
    borders: BDt,
    rows: [
      // Riga 1: P.IVA+ragione breve | ragione completa+indirizzo+ATECO | box SCHEDA MANSIONE
      new TableRow({children:[
        tc([
          p(CLIENTE.ragioneSocialeBreve, {sz:16, bold:true, color:'1F3864'}),
          p(`P.IVA: ${CLIENTE.piva}`, {sz:14, color:'595959'}),
        ], {w:2800, fill:'FFFFFF'}),
        tc([
          p(CLIENTE.ragioneSociale, {sz:18, bold:true, color:'1F3864', spB:20}),
          p(CLIENTE.indirizzo, {sz:15, color:'595959'}),
          p(`ATECO ${CLIENTE.atecoCodice} \u2013 ${CLIENTE.atecoDesc}`, {sz:15, color:'595959'}),
        ], {w:7200, fill:'FFFFFF'}),
        tc([
          p('SCHEDA MANSIONE', {sz:14, color:'BDD7EE', align:AlignmentType.RIGHT}),
          p(mansione.nome.toUpperCase(), {sz:26, bold:true, color:'FFFFFF', align:AlignmentType.RIGHT}),
          p(mansione.reparto, {sz:16, color:'BDD7EE', align:AlignmentType.RIGHT}),
        ], {w:5398, fill:'1F3864'}),
      ]}),
      // Riga 2: RSPP | ATECO | Livello | Anno revisione
      new TableRow({children:[
        tc([
          p('RSPP / DATORE DI LAVORO', {sz:13, color:'BDD7EE'}),
          p(CLIENTE.datoreLavoro, {sz:16, bold:true, color:'FFFFFF'}),
        ], {w:3849, fill:'2E75B6'}),
        tc([
          p('CODICE ATECO', {sz:13, color:'1F4E79'}),
          p(`${CLIENTE.atecoCodice} \u2013 ${CLIENTE.atecoDesc}`, {sz:15, bold:true, color:'1F3864'}),
        ], {w:3849, fill:'D5E8F0'}),
        tc([
          p('LIVELLO RISCHIO SETTORE', {sz:13, color:'FFD7D7'}),
          p(`\u25cf ${mansione.livello} \u2013 ${mansione.oreSpec} ore specifica`, {sz:16, bold:true, color:'FFFFFF'}),
        ], {w:3849, fill:'C00000'}),
        tc([
          p('DATA REVISIONE', {sz:13, color:'595959'}),
          p(CLIENTE.anno, {sz:16, bold:true, color:'1F3864'}),
        ], {w:3851, fill:'F2F2F2'}),
      ]}),
    ],
  });

  // ═══ TABELLA 2 – RISCHI ══════════════════════════════════════════════════
  const Cn=400, Ccat=1300, Cris=3400, Cliv=1100, Cmis=5498, Cdpi=3700;

  const rHdr = new TableRow({children:[
    tc(p('N.',        {sz:16,bold:true,color:'FFFFFF',align:AlignmentType.CENTER}), {w:Cn,   fill:'1F3864'}),
    tc(p('Categoria', {sz:16,bold:true,color:'FFFFFF'}),                            {w:Ccat, fill:'1F3864'}),
    tc(p('Rischio identificato',{sz:16,bold:true,color:'FFFFFF'}),                  {w:Cris, fill:'1F3864'}),
    tc(p('Livello',   {sz:16,bold:true,color:'FFFFFF',align:AlignmentType.CENTER}), {w:Cliv, fill:'1F3864'}),
    tc(p('Misure di prevenzione e protezione',{sz:16,bold:true,color:'FFFFFF'}),    {w:Cmis, fill:'1F3864'}),
    tc(p('DPI richiesti',{sz:16,bold:true,color:'FFFFFF'}),                         {w:Cdpi, fill:'1F3864'}),
  ]});

  const rischioRows = mansione.rischi.map((r, i) => {
    const cat     = catRischio(r.nome);
    const catFill = CAT_FILL[cat] || 'FFFFFF';
    const livello = r.livello || mansione.livello;
    const rowFill = ROW_FILL[livello] || 'FFFBF0';
    const livCol  = LIV_COL[livello]  || '000000';

    const misurePars = r.misure.map(m => new Paragraph({
      spacing:{before:0, after:40},
      children:[
        new TextRun({text:'\u25cf ', font:FONT, size:14, color:'538135', bold:true}),
        new TextRun({text:m, font:FONT, size:14}),
      ],
    }));

    const dpiPars = r.dpi.map(d => new Paragraph({
      spacing:{before:0, after:40},
      children:[new TextRun({text:'  '+d, font:FONT, size:14, color:'1F4E79'})],
    }));

    return new TableRow({children:[
      tc(p(String(i+1), {sz:16,bold:true,color:'C00000',align:AlignmentType.CENTER}), {w:Cn,   fill:rowFill}),
      tc(p(cat,          {sz:14,bold:true,color:'1F3864'}),                             {w:Ccat, fill:catFill, vAlign:VerticalAlign.CENTER}),
      tc(p(r.nome,       {sz:15}),                                                       {w:Cris, fill:rowFill}),
      tc([
        p('\u26a0', {sz:20, bold:true, color:livCol, align:AlignmentType.CENTER}),
        p(livello,  {sz:14, bold:true, color:livCol, align:AlignmentType.CENTER}),
      ], {w:Cliv, fill:rowFill}),
      tc(misurePars.length ? misurePars : [p('')], {w:Cmis, fill:rowFill}),
      tc(dpiPars.length   ? dpiPars   : [p('\u2014', {color:'AAAAAA'})], {w:Cdpi, fill:rowFill}),
    ]});
  });

  const tblRischi = new Table({
    width: {size:W, type:WidthType.DXA},
    columnWidths: [Cn, Ccat, Cris, Cliv, Cmis, Cdpi],
    borders: BDt,
    rows: [rHdr, ...rischioRows],
  });

  // ═══ TABELLA 3 – DPI PREVISTI (header + icone) ═══════════════════════════
  const ICON_MAP = [
    {re:/occhiali/i,                                  file:'Immagine10.png', label:'Occhiali protettivi'},
    {re:/calzatur|scarpe|antiscivolo|ergonomic/i,     file:'Immagine9.png',  label:'Calzature di sicurezza'},
    {re:/antitaglio|nitrile/i,                        file:'Immagine7.png',  label:'Guanti chimico-res.'},
    {re:/termoresist|da lavoro/i,                     file:'Immagine8.png',  label:'Guanti da lavoro'},
    {re:/antivibranti/i,                              file:'Immagine3.png',  label:'Guanti antivibranti'},
    {re:/otoprotett|auricolar/i,                      file:'Immagine4.png',  label:'Otoprotettori'},
    {re:/ffp3/i,                                      file:'Immagine5.png',  label:'Maschera FFP3'},
    {re:/ffp1|ffp2|mascherina|maschera/i,             file:'Immagine6.png',  label:'Maschera FFP2'},
    {re:/tuta/i,                                      file:'Immagine2.png',  label:'Tuta monouso'},
    {re:/alta visibilit|hv|gilet/i,                   file:'Immagine1.png',  label:'Alta visibilit\u00e0'},
  ];
  const ICONS_DIR = path_mod.join(__dirname, 'icons');

  const usedFiles = new Set();
  const dpiIconData = [];
  for (const dpiName of mansione.dpi) {
    for (const m of ICON_MAP) {
      if (m.re.test(dpiName) && !usedFiles.has(m.file)) {
        usedFiles.add(m.file);
        const iconPath = path_mod.join(ICONS_DIR, m.file);
        const iconBytes = fs_mod.existsSync(iconPath) ? fs_mod.readFileSync(iconPath) : null;
        dpiIconData.push({iconBytes, label: dpiName.split('(')[0].trim()});
        break;
      }
    }
  }
  if (!dpiIconData.length) dpiIconData.push({iconBytes:null, label:'\u2014'});

  const nDpi = dpiIconData.length;
  const dpiCellW  = Math.floor(W / nDpi);
  const dpiCellWL = W - dpiCellW*(nDpi-1);
  const dpiColW   = Array.from({length:nDpi}, (_,i)=> i===nDpi-1 ? dpiCellWL : dpiCellW);

  const rDpiHdr = new TableRow({children:[
    new TableCell({
      width:{size:W, type:WidthType.DXA},
      columnSpan: nDpi,
      verticalAlign: VerticalAlign.CENTER,
      shading:{fill:'1F3864', type:ShadingType.CLEAR},
      borders: BDc,
      margins:{top:60,bottom:60,left:80,right:80},
      children:[p('DPI PREVISTI PER LA MANSIONE',{sz:16,bold:true,color:'FFFFFF',align:AlignmentType.CENTER})],
    }),
  ]});

  const rDpiIcons = new TableRow({children: dpiIconData.map(({iconBytes, label}, i) =>
    new TableCell({
      width:{size:dpiColW[i], type:WidthType.DXA},
      verticalAlign: VerticalAlign.CENTER,
      shading:{fill:'EBF3FB', type:ShadingType.CLEAR},
      borders: BDc,
      margins:{top:40,bottom:40,left:40,right:40},
      children:[
        new Paragraph({
          alignment:AlignmentType.CENTER,
          spacing:{after:20},
          children: iconBytes
            ? [new ImageRun({data:iconBytes, type:'png', transformation:{width:36,height:36}})]
            : [new TextRun({text:'\u25a1', font:FONT, size:24})],
        }),
        p(label, {sz:12, align:AlignmentType.CENTER}),
      ],
    })
  )});

  const tblDpi = new Table({
    width:{size:W, type:WidthType.DXA},
    columnWidths: dpiColW,
    borders: BDt,
    rows: [rDpiHdr, rDpiIcons],
  });

  // ═══ TABELLA 4 – FORMAZIONE + FIRME ══════════════════════════════════════
  const tblForm = new Table({
    width:{size:W, type:WidthType.DXA},
    columnWidths:[2308,2308,2308,4237,4237],
    borders: BDt,
    rows:[
      new TableRow({children:[
        tc([
          p('FORMAZIONE GENERALE', {sz:14,bold:true,color:'FFFFFF',align:AlignmentType.CENTER}),
          p('4 ORE',               {sz:20,bold:true,color:'FFFFFF',align:AlignmentType.CENTER}),
          p('Credito permanente',  {sz:12,color:'BDD7EE',align:AlignmentType.CENTER}),
        ], {w:2308, fill:'1F3864'}),
        tc([
          p('FORMAZIONE SPECIFICA',    {sz:14,bold:true,color:'FFFFFF',align:AlignmentType.CENTER}),
          p(`${mansione.oreSpec} ORE`, {sz:20,bold:true,color:'FFFFFF',align:AlignmentType.CENTER}),
          p(`Rischio ${mansione.livello}`, {sz:12,color:'BDD7EE',align:AlignmentType.CENTER}),
        ], {w:2308, fill:'2E75B6'}),
        tc([
          p('AGGIORNAMENTO',  {sz:14,bold:true,color:'1F4E79',align:AlignmentType.CENTER}),
          p('6 ORE',          {sz:20,bold:true,color:'1F4E79',align:AlignmentType.CENTER}),
          p('ogni 5 anni',    {sz:12,color:'1F4E79',align:AlignmentType.CENTER}),
        ], {w:2308, fill:'D5E8F0'}),
        tc([
          p('Firma Datore di Lavoro / RSPP', {sz:13,bold:true,spB:20}),
          p(CLIENTE.datoreLavoro,            {sz:14,color:'595959',spB:60}),
          p('_______________________________',{sz:14}),
        ], {w:4237, fill:'FFFFFF'}),
        tc([
          p('Firma Medico Competente', {sz:13,bold:true,spB:20}),
          p('',                       {sz:14,spB:60}),
          p('_______________________________',{sz:14}),
        ], {w:4237, fill:'FFFFFF'}),
      ]}),
    ],
  });

  // ═══ DOCUMENTO ═══════════════════════════════════════════════════════════
  const doc = new Document({
    styles: docStyles,
    sections:[{
      properties:{page:{size:PAGE_L, margin:MARGIN}},
      children:[tblHeader, tblRischi, tblDpi, tblForm],
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
