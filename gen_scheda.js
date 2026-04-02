'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType, VerticalAlign,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_L, MARGIN_STD,
  makeHeader, makeFooter, titoloSezione, corpo, vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

// A4 landscape content width: 16838 - 1134*2 = 14570 DXA
const W = 14570;

async function genSchedaMansione(mansione) {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'SCHEDA MANSIONE', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter(`Scheda Mansione – ${mansione.nome}`);

  const livColor = mansione.livello === 'ALTO' ? C.ROSSO : mansione.livello === 'MEDIO' ? C.ARANCIO : C.VERDE;
  const livLabel = { 'BASSO': '🟢 BASSO', 'MEDIO': '🟡 MEDIO', 'ALTO': '🔴 ALTO' }[mansione.livello] || mansione.livello;

  // Tabella principale: left panel (DPI + level) + right panel (rischi tabella)
  const nRischi = mansione.rischi.length;

  // Panel sinistro: 4000 DXA, Panel destro: 10570 DXA
  const panelL = 4000;
  const panelR = W - panelL;

  // DPI list
  const dpiRows = mansione.dpi.map((d, i) => new TableRow({
    children: [
      new TableCell({
        width: { size: panelL, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? C.BIANCO : C.BLU_ALT, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        borders: {
          top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
          left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
        },
        children: [new Paragraph({
          children: [
            new TextRun({ text: '▪ ', font: FONT, size: 18, color: C.BLU_MED }),
            new TextRun({ text: d, font: FONT, size: 18 }),
          ],
        })],
      }),
    ],
  }));

  // Tabella rischi destra
  const rischioRows = [
    new TableRow({
      tableHeader: true,
      children: [
        cella('RISCHIO', { width: Math.floor(panelR * 0.25), bold: true, fill: C.BLU_LIGHT, align: 'center' }),
        cella('MISURE DI PREVENZIONE', { width: Math.floor(panelR * 0.45), bold: true, fill: C.BLU_LIGHT, align: 'center' }),
        cella('DPI SPECIFICI', { width: panelR - Math.floor(panelR * 0.25) - Math.floor(panelR * 0.45), bold: true, fill: C.BLU_LIGHT, align: 'center' }),
      ],
    }),
    ...mansione.rischi.map((r, i) => new TableRow({
      children: [
        new TableCell({
          width: { size: Math.floor(panelR * 0.25), type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          verticalAlign: VerticalAlign.TOP,
          borders: {
            top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
          },
          children: [new Paragraph({
            children: [new TextRun({ text: r.nome, bold: true, font: FONT, size: 18, color: C.BLU_DARK })],
          })],
        }),
        new TableCell({
          width: { size: Math.floor(panelR * 0.45), type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          verticalAlign: VerticalAlign.TOP,
          borders: {
            top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
          },
          children: r.misure.map(m => new Paragraph({
            spacing: { before: 0, after: 20 },
            children: [
              new TextRun({ text: '→ ', font: FONT, size: 18, color: C.BLU_MED }),
              new TextRun({ text: m, font: FONT, size: 18 }),
            ],
          })),
        }),
        new TableCell({
          width: { size: panelR - Math.floor(panelR * 0.25) - Math.floor(panelR * 0.45), type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          verticalAlign: VerticalAlign.TOP,
          borders: {
            top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
            right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
          },
          children: r.dpi.map(d => new Paragraph({
            spacing: { before: 0, after: 20 },
            children: [
              new TextRun({ text: '◆ ', font: FONT, size: 18, color: C.ARANCIO }),
              new TextRun({ text: d, font: FONT, size: 18 }),
            ],
          })),
        }),
      ],
    })),
  ];

  const children = [
    // Titolo mansione
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: `SCHEDA MANSIONE: ${mansione.nome.toUpperCase()}`, bold: true, font: FONT, size: 30, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [
        new TextRun({ text: `Reparto: ${mansione.reparto}   `, font: FONT, size: 20, color: C.GRIGIO }),
        new TextRun({ text: `   Livello di rischio: ${livLabel}   `, font: FONT, size: 20, bold: true, color: livColor }),
        new TextRun({ text: `   Formazione: 4h generale + ${mansione.oreSpec}h specifica = ${4+mansione.oreSpec}h`, font: FONT, size: 20, color: C.GRIGIO }),
      ],
    }),

    // Tabella principale (2 colonne: DPI panel + rischi table)
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [panelL, panelR],
      borders: {
        top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
        left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
        insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.SINGLE,size:4,color:C.BLU_MED},
      },
      rows: [
        new TableRow({
          children: [
            // Panel sinistro: DPI riassuntivo
            new TableCell({
              width: { size: panelL, type: WidthType.DXA },
              verticalAlign: VerticalAlign.TOP,
              borders: {
                top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
                left:{style:BorderStyle.NONE},right:{style:BorderStyle.SINGLE,size:4,color:C.BLU_MED},
              },
              margins: { top: 80, bottom: 80, left: 100, right: 200 },
              children: [
                new Paragraph({
                  spacing: { before: 0, after: 80 },
                  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED } },
                  children: [new TextRun({ text: 'DPI PREVISTI', bold: true, font: FONT, size: 22, color: C.BLU_HEADER })],
                }),
                ...mansione.dpi.map((d, i) => new Paragraph({
                  spacing: { before: 0, after: 40 },
                  children: [
                    new TextRun({ text: '▪ ', font: FONT, size: 18, color: C.BLU_MED }),
                    new TextRun({ text: d, font: FONT, size: 18 }),
                  ],
                })),
                vuoto(120),
                new Paragraph({
                  spacing: { before: 0, after: 80 },
                  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED } },
                  children: [new TextRun({ text: 'ARGOMENTI TRASVERSALI', bold: true, font: FONT, size: 22, color: C.BLU_HEADER })],
                }),
                ...['DPI e loro corretto utilizzo', 'Segnaletica di sicurezza', 'Primo soccorso', 'Stress lavoro-correlato', 'Near miss – mancati infortuni'].map(t => new Paragraph({
                  spacing: { before: 0, after: 40 },
                  children: [
                    new TextRun({ text: '✓ ', font: FONT, size: 18, color: C.VERDE }),
                    new TextRun({ text: t, font: FONT, size: 18 }),
                  ],
                })),
                vuoto(120),
                new Paragraph({
                  spacing: { before: 0, after: 60 },
                  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED } },
                  children: [new TextRun({ text: 'DATORE DI LAVORO / RSPP', bold: true, font: FONT, size: 22, color: C.BLU_HEADER })],
                }),
                new Paragraph({
                  spacing: { before: 0, after: 40 },
                  children: [new TextRun({ text: CLIENTE.datoreLavoro, font: FONT, size: 20 })],
                }),
                new Paragraph({
                  spacing: { before: 0, after: 40 },
                  children: [new TextRun({ text: CLIENTE.ragioneSociale, font: FONT, size: 18, color: C.GRIGIO })],
                }),
                new Paragraph({
                  spacing: { before: 60, after: 40 },
                  children: [new TextRun({ text: 'Data/Firma: ___________________', font: FONT, size: 18 })],
                }),
              ],
            }),
            // Panel destro: tabella rischi
            new TableCell({
              width: { size: panelR, type: WidthType.DXA },
              verticalAlign: VerticalAlign.TOP,
              borders: {
                top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
                left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
              },
              margins: { top: 80, bottom: 80, left: 200, right: 80 },
              children: [
                new Paragraph({
                  spacing: { before: 0, after: 80 },
                  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.BLU_MED } },
                  children: [new TextRun({ text: 'RISCHI SPECIFICI, MISURE DI PREVENZIONE E DPI', bold: true, font: FONT, size: 22, color: C.BLU_HEADER })],
                }),
                new Table({
                  width: { size: panelR - 280, type: WidthType.DXA },
                  columnWidths: [
                    Math.floor((panelR - 280) * 0.25),
                    Math.floor((panelR - 280) * 0.45),
                    (panelR - 280) - Math.floor((panelR - 280) * 0.25) - Math.floor((panelR - 280) * 0.45),
                  ],
                  borders: {
                    top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                    bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                    left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                    right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                    insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                    insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                  },
                  rows: [
                    new TableRow({
                      tableHeader: true,
                      children: [
                        cella('RISCHIO', { width: Math.floor((panelR-280)*0.25), bold:true, fill:C.BLU_LIGHT, align:'center', sz:18 }),
                        cella('MISURE DI PREVENZIONE', { width: Math.floor((panelR-280)*0.45), bold:true, fill:C.BLU_LIGHT, align:'center', sz:18 }),
                        cella('DPI SPECIFICI', { width: (panelR-280)-Math.floor((panelR-280)*0.25)-Math.floor((panelR-280)*0.45), bold:true, fill:C.BLU_LIGHT, align:'center', sz:18 }),
                      ],
                    }),
                    ...mansione.rischi.map((r, i) => new TableRow({
                      children: [
                        new TableCell({
                          width: { size: Math.floor((panelR-280)*0.25), type: WidthType.DXA },
                          shading: { fill: i%2===0?C.BIANCO:C.GRIGIO_ALT, type: ShadingType.CLEAR },
                          margins: { top:40, bottom:40, left:80, right:80 },
                          verticalAlign: VerticalAlign.TOP,
                          borders: {
                            top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                            left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                          },
                          children: [new Paragraph({ children: [new TextRun({ text: r.nome, bold:true, font:FONT, size:16, color:C.BLU_DARK })] })],
                        }),
                        new TableCell({
                          width: { size: Math.floor((panelR-280)*0.45), type: WidthType.DXA },
                          shading: { fill: i%2===0?C.BIANCO:C.GRIGIO_ALT, type: ShadingType.CLEAR },
                          margins: { top:40, bottom:40, left:80, right:80 },
                          verticalAlign: VerticalAlign.TOP,
                          borders: {
                            top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                            left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                          },
                          children: r.misure.slice(0,3).map(m => new Paragraph({
                            spacing:{before:0,after:20},
                            children:[
                              new TextRun({text:'→ ',font:FONT,size:16,color:C.BLU_MED}),
                              new TextRun({text:m,font:FONT,size:16}),
                            ],
                          })),
                        }),
                        new TableCell({
                          width: { size: (panelR-280)-Math.floor((panelR-280)*0.25)-Math.floor((panelR-280)*0.45), type: WidthType.DXA },
                          shading: { fill: i%2===0?C.BIANCO:C.GRIGIO_ALT, type: ShadingType.CLEAR },
                          margins: { top:40, bottom:40, left:80, right:80 },
                          verticalAlign: VerticalAlign.TOP,
                          borders: {
                            top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                            left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
                          },
                          children: r.dpi.map(d => new Paragraph({
                            spacing:{before:0,after:20},
                            children:[
                              new TextRun({text:'◆ ',font:FONT,size:16,color:C.ARANCIO}),
                              new TextRun({text:d,font:FONT,size:16}),
                            ],
                          })),
                        }),
                      ],
                    })),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_L, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/01 - SCHEDE MANSIONI/${mansione.id}.docx`);
}

async function genSchedaAddestrativa(mansione) {
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, 'SCHEDA ADDESTRATIVA', CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter(`Scheda Addestrativa – ${mansione.nome}`);

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: `SCHEDA ADDESTRATIVA – ${mansione.nome.toUpperCase()}`, bold: true, font: FONT, size: 28, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: 'D.Lgs. 81/2008 – Addestramento specifico ai sensi dell\'art. 37 co. 5', font: FONT, size: 20, color: C.BLU_HEADER })],
    }),

    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [3000, W - 3000],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        ['Lavoratore/rice', ''], ['Mansione', mansione.nome],
        ['Tutor / Addestratore', CLIENTE.datoreLavoro], ['Data addestramento', '___/___/_____'],
        ['Durata', '_____ ore'],
      ].map(([et, va], i) => new TableRow({ children: [
        cella(et, { width:3000, bold:true, fill:C.BLU_LIGHT }),
        cella(va, { width: W-3000, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
      ]})),
    }),
    vuoto(120),

    titoloSezione('ATTIVITÀ DI ADDESTRAMENTO SVOLTE'),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [Math.floor(W*0.4), Math.floor(W*0.35), W - Math.floor(W*0.4) - Math.floor(W*0.35)],
      borders: {
        top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
        insideH:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},insideV:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},
      },
      rows: [
        new TableRow({ tableHeader: true, children: [
          cella('Attività addestrativa', { width:Math.floor(W*0.4), bold:true, fill:C.BLU_LIGHT }),
          cella('Competenza acquisita', { width:Math.floor(W*0.35), bold:true, fill:C.BLU_LIGHT }),
          cella('Valutazione (1-5)', { width: W-Math.floor(W*0.4)-Math.floor(W*0.35), bold:true, fill:C.BLU_LIGHT, align:'center' }),
        ]}),
        ...mansione.rischi.map((r, i) => new TableRow({ children: [
          cella(`Gestione sicura: ${r.nome}`, { width:Math.floor(W*0.4), fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
          cella(r.misure[0] || '', { width:Math.floor(W*0.35), fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
          cella('', { width: W-Math.floor(W*0.4)-Math.floor(W*0.35), fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
        ]})),
        new TableRow({ children: [
          cella('Corretto utilizzo dei DPI in dotazione', { width:Math.floor(W*0.4) }),
          cella('Uso, manutenzione e sostituzione DPI', { width:Math.floor(W*0.35) }),
          cella('', { width: W-Math.floor(W*0.4)-Math.floor(W*0.35) }),
        ]}),
        new TableRow({ children: [
          cella('Gestione emergenze e primo soccorso', { width:Math.floor(W*0.4), fill:C.GRIGIO_ALT }),
          cella('Conoscenza procedure di evacuazione', { width:Math.floor(W*0.35), fill:C.GRIGIO_ALT }),
          cella('', { width: W-Math.floor(W*0.4)-Math.floor(W*0.35), fill:C.GRIGIO_ALT }),
        ]}),
      ],
    }),
    vuoto(120),

    corpo('Note e osservazioni del tutor:', { before: 60 }),
    new Paragraph({ spacing:{after:60}, children:[new TextRun({text:'_________________________________________________________________________',font:FONT,size:20})] }),
    new Paragraph({ spacing:{after:120}, children:[new TextRun({text:'_________________________________________________________________________',font:FONT,size:20})] }),

    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [Math.floor(W/2), W - Math.floor(W/2)],
      borders: {
        top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},
        left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE},
        insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE},
      },
      rows: [new TableRow({ children: [
        new TableCell({ width:{size:Math.floor(W/2),type:WidthType.DXA}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children:[
          new Paragraph({children:[new TextRun({text:'Firma Tutor / Addestratore',font:FONT,size:20})]}),
          new Paragraph({spacing:{after:40},children:[new TextRun({text:'___________________________________',font:FONT,size:20})]}),
        ]}),
        new TableCell({ width:{size:W-Math.floor(W/2),type:WidthType.DXA}, borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children:[
          new Paragraph({children:[new TextRun({text:'Firma Lavoratore/rice',font:FONT,size:20})]}),
          new Paragraph({spacing:{after:40},children:[new TextRun({text:'___________________________________',font:FONT,size:20})]}),
        ]}),
      ]})],
    }),
  ];

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_L, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  await salvaDoc(doc, `${OUT}/BONUS - SCHEDA ADDESTRATIVA/Scheda_Addestramento_${mansione.id}.docx`);
}

module.exports = { genSchedaMansione, genSchedaAddestrativa };
