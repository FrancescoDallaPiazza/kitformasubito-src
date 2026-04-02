'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  C, FONT, CLIENTE, docStyles, A4_P, A4_L, logoBytes,
  vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

const BRD = {
  top:    { style: BorderStyle.SINGLE, size: 1, color: 'BBBBBB' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'BBBBBB' },
  left:   { style: BorderStyle.SINGLE, size: 1, color: 'BBBBBB' },
  right:  { style: BorderStyle.SINGLE, size: 1, color: 'BBBBBB' },
};
const NO_BRD = {
  top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
};

// ─── SCHEDA MANSIONE ─────────────────────────────────────────────────────────
async function genSchedaMansione(mansione) {
  const MARGIN = { top: 737, right: 737, bottom: 737, left: 737 };
  const W = 15364; // 16838 - 737*2

  const livColor = { ALTO: C.ROSSO, MEDIO: C.ARANCIO, BASSO: C.VERDE }[mansione.livello] || C.GRIGIO;

  // TOP BAR: logo | titolo mansione | info
  const wLogo = 1400; const wTitle = 8000; const wInfo = W - wLogo - wTitle;
  const topBar = new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [wLogo, wTitle, wInfo],
    borders: { top: { style: BorderStyle.SINGLE, size: 6, color: C.BLU_HEADER },
               bottom: { style: BorderStyle.SINGLE, size: 6, color: C.BLU_HEADER },
               left: NO_BRD.left, right: NO_BRD.right,
               insideH: NO_BRD.top, insideV: NO_BRD.top },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: wLogo, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        borders: NO_BRD, margins: { top: 80, bottom: 80, left: 80, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
          new ImageRun({ data: logoBytes, type: 'jpg', transformation: { width: 60, height: 60 } }),
        ]})],
      }),
      new TableCell({ width: { size: wTitle, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        borders: { top: NO_BRD.top, bottom: NO_BRD.bottom,
                   left: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' },
                   right: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' } },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [
            new TextRun({ text: 'SCHEDA MANSIONE', bold: true, font: FONT, size: 20, color: C.BLU_HEADER }),
          ]}),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [
            new TextRun({ text: mansione.nome.toUpperCase(), bold: true, font: FONT, size: 26, color: C.BLU_DARK }),
          ]}),
        ],
      }),
      new TableCell({ width: { size: wInfo, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        borders: NO_BRD, margins: { top: 80, bottom: 80, left: 160, right: 80 },
        children: [
          new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 20 }, children: [
            new TextRun({ text: `Reparto: ${mansione.reparto}`, font: FONT, size: 18, color: C.GRIGIO }),
          ]}),
          new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 20 }, children: [
            new TextRun({ text: 'Rischio: ', font: FONT, size: 18 }),
            new TextRun({ text: mansione.livello, bold: true, font: FONT, size: 18, color: livColor }),
          ]}),
          new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [
            new TextRun({ text: 'D.Lgs. 81/08 – Art. 36-37', font: FONT, size: 16, color: C.GRIGIO }),
          ]}),
        ],
      }),
    ]})]
  });

  // RISCHI TABLE: 4 colonne
  const wR = 3200; const wM = 6000; const wD1 = 3164; const wD2 = W - wR - wM - wD1;
  function hdCell(txt, w) {
    return new TableCell({ width: { size: w, type: WidthType.DXA },
      shading: { fill: C.BLU_LIGHT, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 80 }, borders: BRD,
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: txt, bold: true, font: FONT, size: 18, color: C.BLU_DARK }),
      ]})],
    });
  }
  function blt(txt, col) {
    return new Paragraph({ spacing: { before: 0, after: 24 }, children: [
      new TextRun({ text: '• ', font: FONT, size: 18, color: col || C.BLU_MED }),
      new TextRun({ text: txt, font: FONT, size: 18 }),
    ]});
  }
  function dc(w, fill, children) {
    return new TableCell({ width: { size: w, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 100, right: 80 },
      verticalAlign: VerticalAlign.TOP, borders: BRD, children });
  }

  const rischioTable = new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [wR, wM, wD1, wD2],
    borders: { top: BRD.top, bottom: BRD.bottom, left: BRD.left, right: BRD.right, insideH: BRD.top, insideV: BRD.top },
    rows: [
      new TableRow({ tableHeader: true, children: [hdCell('RISCHIO', wR), hdCell('MISURE DI PREVENZIONE', wM), hdCell('DPI RICHIESTI', wD1), hdCell('DPI AZIENDA', wD2)] }),
      ...mansione.rischi.map((r, i) => {
        const fill = i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT;
        return new TableRow({ children: [
          dc(wR, fill, [new Paragraph({ children: [new TextRun({ text: r.nome, bold: true, font: FONT, size: 18, color: C.BLU_DARK })] })]),
          dc(wM, fill, r.misure.map(m => blt(m, C.BLU_MED))),
          dc(wD1, fill, r.dpi.length ? r.dpi.map(d => blt(d, C.ARANCIO)) : [new Paragraph({ children: [new TextRun({ text: '—', font: FONT, size: 18, color: C.GRIGIO })] })]),
          dc(wD2, fill, [new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: 18 })] })]),
        ]});
      }),
    ],
  });

  // DPI OBBLIGATORI BAR
  const wDL = 2600;
  const dpiBar = new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [wDL, W - wDL],
    borders: { top: BRD.top, bottom: BRD.bottom, left: BRD.left, right: BRD.right, insideH: NO_BRD.top, insideV: BRD.top },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: wDL, type: WidthType.DXA },
        shading: { fill: C.BLU_HEADER, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 }, borders: BRD,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
          new TextRun({ text: 'DPI OBBLIGATORI', bold: true, font: FONT, size: 20, color: C.BIANCO }),
        ]})],
      }),
      new TableCell({ width: { size: W - wDL, type: WidthType.DXA },
        shading: { fill: C.BLU_ALT, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 }, borders: BRD,
        children: [new Paragraph({ children: [
          new TextRun({ text: mansione.dpi.join('   |   '), font: FONT, size: 18 }),
        ]})],
      }),
    ]})]
  });

  // REVISIONE BAR
  const wRev = Math.floor(W * 0.7);
  const reviBar = new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [wRev, W - wRev],
    borders: { top: NO_BRD.top, bottom: NO_BRD.bottom, left: NO_BRD.left, right: NO_BRD.right, insideH: NO_BRD.top, insideV: NO_BRD.top },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: wRev, type: WidthType.DXA }, borders: NO_BRD,
        margins: { top: 60, bottom: 40, left: 0, right: 80 },
        children: [new Paragraph({ children: [
          new TextRun({ text: 'Revisione, aggiornamento e consegna al lavoratore: il presente documento è soggetto a revisione periodica o in seguito a variazioni delle mansioni, dell\'organizzazione del lavoro o dei rischi presenti.', font: FONT, size: 16, color: C.GRIGIO }),
        ]})],
      }),
      new TableCell({ width: { size: W - wRev, type: WidthType.DXA }, borders: NO_BRD, verticalAlign: VerticalAlign.CENTER,
        margins: { top: 60, bottom: 40, left: 80, right: 0 },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
          new TextRun({ text: `${CLIENTE.ragioneSocialeBreve} – Rev. ${CLIENTE.anno}`, bold: true, font: FONT, size: 16, color: C.BLU_HEADER }),
        ]})],
      }),
    ]})]
  });

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_L, margin: MARGIN } },
      children: [topBar, vuoto(40), rischioTable, vuoto(40), dpiBar, vuoto(20), reviBar],
    }],
  });
  await salvaDoc(doc, `${OUT}/01 - SCHEDE MANSIONI/${mansione.id}.docx`);
}

// ─── SCHEDA ADDESTRATIVA ─────────────────────────────────────────────────────
async function genSchedaAddestrativa(mansione) {
  const MARGIN = { top: 567, right: 1134, bottom: 567, left: 1134 };
  const W = 9638;
  const H2 = Math.floor(W / 2);
  const Q = Math.floor(H2 / 2);

  const BD = BRD;

  function fullSpan(content, fill) {
    return new TableRow({ children: [
      new TableCell({ columnSpan: 4, width: { size: W, type: WidthType.DXA },
        shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
        margins: { top: 60, bottom: 60, left: 120, right: 120 }, borders: BD,
        children: content,
      }),
    ]});
  }

  function span2(txt, w, fill) {
    return new TableCell({ columnSpan: 2, width: { size: w, type: WidthType.DXA },
      shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
      margins: { top: 60, bottom: 60, left: 120, right: 120 }, borders: BD,
      children: [new Paragraph({ children: [new TextRun({ text: txt, font: FONT, size: 19 })] })],
    });
  }

  const attivita = [
    'Utilizzo corretto delle attrezzature di lavoro',
    'Uso e manutenzione dei DPI in dotazione',
    'Procedure di emergenza e primo soccorso',
    'Gestione rischi specifici della mansione',
    'Segnaletica di sicurezza e comportamenti corretti',
    'Prevenzione infortuni e segnalazione near miss',
  ];

  function emptyCell(w, fill) {
    return new TableCell({ width: { size: w, type: WidthType.DXA },
      shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
      margins: { top: 60, bottom: 60, left: 120, right: 120 }, borders: BD,
      children: [new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: 19 })] })],
    });
  }

  const mainRows = [
    fullSpan([new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: 'MANSIONE COINVOLTA', bold: true, font: FONT, size: 20, color: C.BIANCO }),
    ]})], C.BLU_HEADER),
    new TableRow({ children: [span2(mansione.nome, H2, C.BLU_ALT), span2('___________________________________________', H2)] }),
    fullSpan([new Paragraph({ children: [
      new TextRun({ text: `Reparto/Area:  ${mansione.reparto}`, font: FONT, size: 19 }),
      new TextRun({ text: '    ___________________________    -    ___/___/______', font: FONT, size: 19 }),
    ]})]),
    fullSpan([new Paragraph({ children: [
      new TextRun({ text: 'Motivazioni addestramento:  ', bold: true, font: FONT, size: 19 }),
      new TextRun({ text: '☐ Nuova assunzione   ☐ Cambio mansione   ☐ Interinale   ☐ Altra: _________________', font: FONT, size: 19 }),
    ]})]),
    fullSpan([new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: 'ATTIVITÀ DI ADDESTRAMENTO DEI LAVORATORI', bold: true, font: FONT, size: 19, color: C.BIANCO }),
    ]})], C.BLU_MED),
    ...attivita.map((att, i) => {
      const fill = i % 2 === 0 ? C.BIANCO : C.GRIGIO_ALT;
      return new TableRow({ children: [span2(att, H2, fill), span2('', H2, fill)] });
    }),
    fullSpan([
      new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: 'Note del Tutor:', bold: true, font: FONT, size: 19 })] }),
      new Paragraph({ children: [new TextRun({ text: '_________________________________________________________________________________', font: FONT, size: 19 })] }),
    ]),
  ];

  const mainTable = new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [Q, Q, Q, Q],
    borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: BD.top, insideV: BD.top },
    rows: mainRows,
  });

  function qCell(txt, fill) {
    return new TableCell({ width: { size: Q, type: WidthType.DXA },
      shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
      margins: { top: 40, bottom: 40, left: 120, right: 80 }, borders: BD,
      children: [new Paragraph({ children: [new TextRun({ text: txt, font: FONT, size: 18 })] })],
    });
  }

  const firmaTable = new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [Q, Q, Q, Q],
    borders: { top: BD.top, bottom: BD.bottom, left: BD.left, right: BD.right, insideH: BD.top, insideV: BD.top },
    rows: [
      fullSpan([new Paragraph({ alignment: AlignmentType.CENTER, children: [
        new TextRun({ text: 'FIRME ADDESTRAMENTO SUL CAMPO', bold: true, font: FONT, size: 19, color: C.BIANCO }),
      ]})], C.BLU_HEADER),
      new TableRow({ children: [qCell('Nome lavoratore:'), qCell('Nome lavoratore:'), qCell('Firma lavoratore:'), qCell('Firma lavoratore:')] }),
      ...[...Array(4)].map(() => new TableRow({ children: [qCell(''), qCell(''), qCell(''), qCell('')] })),
      new TableRow({ children: [
        new TableCell({ columnSpan: 2, width: { size: H2, type: WidthType.DXA }, borders: BD, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [
          new Paragraph({ children: [new TextRun({ text: `Tutor / Addestratore: ${CLIENTE.datoreLavoro}`, font: FONT, size: 18 })] }),
          new Paragraph({ children: [new TextRun({ text: 'Firma: _________________________', font: FONT, size: 18 })] }),
        ]}),
        new TableCell({ columnSpan: 2, width: { size: H2, type: WidthType.DXA }, borders: BD, margins: { top: 60, bottom: 60, left: 120, right: 80 }, children: [
          new Paragraph({ children: [new TextRun({ text: 'Data addestramento: ___/___/______', font: FONT, size: 18 })] }),
        ]}),
      ]}),
    ],
  });

  const doc = new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN } },
      children: [
        new Paragraph({ children: [new TextRun({ text: 'SCHEDA ADDESTRAMENTO SUL CAMPO', bold: true, font: FONT, size: 22, color: C.BLU_DARK })] }),
        vuoto(40),
        mainTable,
        vuoto(40),
        firmaTable,
      ],
    }],
  });
  await salvaDoc(doc, `${OUT}/BONUS - SCHEDA ADDESTRATIVA/Scheda_Addestramento_${mansione.id}.docx`);
}

module.exports = { genSchedaMansione, genSchedaAddestrativa };
