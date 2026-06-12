'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  SimpleField,
  C, FONT, CLIENTE, MANSIONI, kitOutDir, docStyles, A4_P, logoBytes,
  vuoto, cella, salvaDoc,
} = h;

const OUT = kitOutDir();

// W = 9638 DXA portrait content width

// Header con logo inline (in alto a sinistra)
function makeHeaderTest() {
  return new Header({
    children: [new Paragraph({
      children: [new ImageRun({
        data: logoBytes,
        type: 'jpg',
        transformation: { width: 70, height: 70 },
      })],
    })],
  });
}

// Footer con bordo superiore blu + azienda + pag
function makeFooterTest() {
  return new Footer({ children: [new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 6, space: 1, color: '2E75B6' } },
    children: [
      new TextRun({ text: `${CLIENTE.ragioneSociale} – ${CLIENTE.indirizzo}   |   Pag. `, size: 16, font: FONT, color: C.GRIGIO }),
      new SimpleField('PAGE'),
    ],
  })]});
}

// Top table: [vuoto] [fill=1F3864, ragione sociale, sz=18, right]
// Col widths dal master: 3373 + 6265 = 9638
function makeTopTable() {
  const wL = 3373; const wR = 9638 - wL;
  const NO = {top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
  return new Table({
    width:{size:9638,type:WidthType.DXA}, columnWidths:[wL,wR],
    borders:{top:NO.top,bottom:NO.bottom,left:NO.left,right:NO.right,insideH:NO.top,insideV:NO.top},
    rows:[new TableRow({children:[
      new TableCell({width:{size:wL,type:WidthType.DXA},borders:NO,
        margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[]})]}),
      new TableCell({width:{size:wR,type:WidthType.DXA},borders:NO,
        shading:{fill:C.BLU_DARK,type:ShadingType.CLEAR},
        margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:CLIENTE.ragioneSociale,bold:true,font:FONT,size:18,color:C.BIANCO})]})],
      }),
    ]})]
  });
}

// Tabella dati discente 2x2
function makeDiscente() {
  const W = 9638; const half = Math.floor(W/2);
  const BD = {top:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:1,color:'CCCCCC'}};
  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[half,W-half],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
    rows:[
      new TableRow({children:[cella('Cognome e Nome: _________________________________',{width:half}),cella('Mansione: _________________________________',{width:W-half})]}),
      new TableRow({children:[cella('Data: ____/____/__________',{width:half}),cella('',{width:W-half})]}),
    ],
  });
}

// Tabella domanda: riga 0 = cella merged (colspan=2), righe 1-4 = lettera + risposta
// Col widths dal master: col0=415dxa (lettera), col1=9223dxa (risposta)
function makeQuestion(domanda, risposte, isDocente) {
  const W = 9638;
  const wL = 415;   // colonna lettera – dal master
  const wR = W - wL; // 9223
  const BD = {top:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},bottom:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},left:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'},right:{style:BorderStyle.SINGLE,size:4,color:'CCCCCC'}};

  const rows = [
    // Riga 0: domanda full-width, fill D5E8F0, bold=true, sz=20, color=000000, margins 80/120
    new TableRow({children:[
      new TableCell({
        columnSpan: 2,
        width:{size:W,type:WidthType.DXA},
        shading:{fill:C.BLU_LIGHT,type:ShadingType.CLEAR},
        margins:{top:80,bottom:80,left:120,right:120},
        borders:BD,
        children:[new Paragraph({children:[new TextRun({text:domanda,font:FONT,size:20,bold:true,color:'000000'})]})]
      }),
    ]}),
    // Righe risposte: lettera sz=18 center color=000000, testo sz=18 color=000000, margins 80/120
    ...risposte.map(r => {
      const corrFill = isDocente && r.corretta ? C.VERDE : undefined;
      return new TableRow({children:[
        new TableCell({
          width:{size:wL,type:WidthType.DXA},
          shading:corrFill?{fill:corrFill,type:ShadingType.CLEAR}:undefined,
          margins:{top:80,bottom:80,left:120,right:120},
          borders:BD,
          children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:`${r.lettera}.`,font:FONT,size:18,color:'000000',bold:isDocente&&r.corretta})]})]
        }),
        new TableCell({
          width:{size:wR,type:WidthType.DXA},
          shading:corrFill?{fill:corrFill,type:ShadingType.CLEAR}:undefined,
          margins:{top:80,bottom:80,left:120,right:120},
          borders:BD,
          children:[new Paragraph({children:[new TextRun({text:isDocente&&r.corretta?`${r.testo}  ✓`:r.testo,font:FONT,size:18,color:'000000',bold:isDocente&&r.corretta})]})]
        }),
      ]});
    }),
  ];

  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[wL,wR],
    borders:{top:BD.top,bottom:BD.bottom,left:BD.left,right:BD.right,insideH:BD.top,insideV:BD.top},
    rows,
  });
}

// Firme finali
function firme() {
  return [
    new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:'Firma del Datore di Lavoro – RSPP: _________________________________',bold:true,font:FONT,size:18,color:'000000'})]}),
    new Paragraph({spacing:{after:0},children:[new TextRun({text:'Firma del discente: _______________________________________',bold:true,font:FONT,size:18,color:'000000'})]}),
  ];
}

// 30 domande generali standardizzate (fisse – non modificare)
function domandeGenerali() {
  return [
    {d:"1. Che cosa si intende per 'pericolo' in ambito lavorativo?", r:[
      {lettera:'A', testo:"Una proprietà o condizione che può causare un danno", corretta:true},
      {lettera:'B', testo:"Un evento senza conseguenze", corretta:false},
      {lettera:'C', testo:"Una persona incaricata della sicurezza", corretta:false},
      {lettera:'D', testo:"Una procedura amministrativa", corretta:false},
    ]},
    {d:"2. Come si definisce 'rischio'?", r:[
      {lettera:'A', testo:"La sola gravità del danno", corretta:false},
      {lettera:'B', testo:"Probabilità e gravità di un danno", corretta:true},
      {lettera:'C', testo:"Il costo della prevenzione", corretta:false},
      {lettera:'D', testo:"Un comportamento sicuro", corretta:false},
    ]},
    {d:"3. Cosa si intende per 'danno'?", r:[
      {lettera:'A', testo:"Una misura preventiva", corretta:false},
      {lettera:'B', testo:"Una segnalazione", corretta:false},
      {lettera:'C', testo:"Conseguenza negativa su persone o cose", corretta:true},
      {lettera:'D', testo:"Un obbligo formale", corretta:false},
    ]},
    {d:"4. Quale delle seguenti situazioni rappresenta un rischio elevato?", r:[
      {lettera:'A', testo:"Ambiente ordinato", corretta:false},
      {lettera:'B', testo:"Usare DPI adeguati", corretta:false},
      {lettera:'C', testo:"Lavorare in quota senza protezioni", corretta:true},
      {lettera:'D', testo:"Seguire la procedura", corretta:false},
    ]},
    {d:"5. Qual è l'obiettivo della valutazione dei rischi?", r:[
      {lettera:'A', testo:"Valutare la produttività", corretta:false},
      {lettera:'B', testo:"Individuare pericoli e definire misure di prevenzione", corretta:true},
      {lettera:'C', testo:"Pianificare le ferie", corretta:false},
      {lettera:'D', testo:"Organizzare i turni", corretta:false},
    ]},
    {d:"6. Quale intervento riduce maggiormente il rischio?", r:[
      {lettera:'A', testo:"Solo segnaletica", corretta:false},
      {lettera:'B', testo:"Uso sporadico dei DPI", corretta:false},
      {lettera:'C', testo:"Eliminazione del pericolo alla fonte", corretta:true},
      {lettera:'D', testo:"Controlli saltuari", corretta:false},
    ]},
    {d:"7. Che cosa significa 'prevenzione' sul luogo di lavoro?", r:[
      {lettera:'A', testo:"Limitare i danni dopo l'evento", corretta:false},
      {lettera:'B', testo:"Evitare che accadano eventi dannosi", corretta:true},
      {lettera:'C', testo:"Curare gli infortunati", corretta:false},
      {lettera:'D', testo:"Archiviare documenti", corretta:false},
    ]},
    {d:"8. Che cosa significa 'protezione'?", r:[
      {lettera:'A', testo:"Eliminare ogni rischio", corretta:false},
      {lettera:'B', testo:"Migliorare la produttività", corretta:false},
      {lettera:'C', testo:"Ridurre le conseguenze di un evento pericoloso", corretta:true},
      {lettera:'D', testo:"Trasferire il rischio ai fornitori", corretta:false},
    ]},
    {d:"9. La formazione dei lavoratori è una misura di:", r:[
      {lettera:'A', testo:"Protezione", corretta:false},
      {lettera:'B', testo:"Sanzione", corretta:false},
      {lettera:'C', testo:"Sorveglianza sanitaria", corretta:false},
      {lettera:'D', testo:"Prevenzione", corretta:true},
    ]},
    {d:"10. Un estintore è un esempio di misura di:", r:[
      {lettera:'A', testo:"Prevenzione", corretta:false},
      {lettera:'B', testo:"Partecipazione", corretta:false},
      {lettera:'C', testo:"Protezione", corretta:true},
      {lettera:'D', testo:"Organizzazione", corretta:false},
    ]},
    {d:"11. Nell'ordine delle priorità della prevenzione qual è la prima scelta?", r:[
      {lettera:'A', testo:"Fornire DPI", corretta:false},
      {lettera:'B', testo:"Eliminare il pericolo alla fonte", corretta:false},
      {lettera:'C', testo:"Mettere cartelli", corretta:false},
      {lettera:'D', testo:"Formare i lavoratori", corretta:true},
    ]},
    {d:"12. Che cos'è una procedura di lavoro sicuro?", r:[
      {lettera:'A', testo:"Una nota informale tra colleghi", corretta:false},
      {lettera:'B', testo:"Un documento contabile", corretta:false},
      {lettera:'C', testo:"Istruzione scritta che standardizza attività in sicurezza", corretta:true},
      {lettera:'D', testo:"Un suggerimento non vincolante", corretta:false},
    ]},
    {d:"13. Qual è il compito principale dell'RSPP?", r:[
      {lettera:'A', testo:"Redigere cartelle cliniche", corretta:false},
      {lettera:'B', testo:"Coordinare attività di prevenzione e protezione", corretta:true},
      {lettera:'C', testo:"Sostituire il datore di lavoro", corretta:false},
      {lettera:'D', testo:"Rappresentare i lavoratori", corretta:false},
    ]},
    {d:"14. Chi è l'RLS?", r:[
      {lettera:'A', testo:"Un revisore legale", corretta:false},
      {lettera:'B', testo:"Rappresentante dei Lavoratori per la Sicurezza", corretta:true},
      {lettera:'C', testo:"Responsabile delle risorse umane", corretta:false},
      {lettera:'D', testo:"Medico competente", corretta:false},
    ]},
    {d:"15. Chi nomina il medico competente in azienda?", r:[
      {lettera:'A', testo:"Il datore di lavoro", corretta:true},
      {lettera:'B', testo:"L'INAIL", corretta:false},
      {lettera:'C', testo:"L'RSPP", corretta:false},
      {lettera:'D', testo:"Il lavoratore", corretta:false},
    ]},
    {d:"16. Chi partecipa alla riunione periodica di prevenzione?", r:[
      {lettera:'A', testo:"Solo i lavoratori", corretta:false},
      {lettera:'B', testo:"Solo i dirigenti", corretta:false},
      {lettera:'C', testo:"Datore di lavoro, RSPP, medico competente, RLS", corretta:true},
      {lettera:'D', testo:"Solo il preposto", corretta:false},
    ]},
    {d:"17. Se un lavoratore individua un pericolo, cosa deve fare immediatamente?", r:[
      {lettera:'A', testo:"Risolverlo da solo", corretta:false},
      {lettera:'B', testo:"Aspettare la riunione periodica", corretta:false},
      {lettera:'C', testo:"Segnalare al preposto o datore di lavoro", corretta:true},
      {lettera:'D', testo:"Ignorarlo", corretta:false},
    ]},
    {d:"18. Qual è il ruolo del preposto?", r:[
      {lettera:'A', testo:"Emettere sanzioni disciplinari", corretta:false},
      {lettera:'B', testo:"Sovrintendere e vigilare sul rispetto delle misure di sicurezza", corretta:true},
      {lettera:'C', testo:"Redigere la valutazione dei rischi", corretta:false},
      {lettera:'D', testo:"Svolgere funzioni mediche", corretta:false},
    ]},
    {d:"19. Quale diritto fondamentale hanno i lavoratori in materia di sicurezza?", r:[
      {lettera:'A', testo:"Decidere le procedure", corretta:false},
      {lettera:'B', testo:"Rifiutare ogni mansione", corretta:false},
      {lettera:'C', testo:"Ricevere formazione, informazione e addestramento", corretta:true},
      {lettera:'D', testo:"Non usare DPI", corretta:false},
    ]},
    {d:"20. Qual è un obbligo del lavoratore?", r:[
      {lettera:'A', testo:"Nominare l'RSPP", corretta:false},
      {lettera:'B', testo:"Usare correttamente i DPI forniti", corretta:true},
      {lettera:'C', testo:"Gestire la sorveglianza sanitaria", corretta:false},
      {lettera:'D', testo:"Redigere la valutazione dei rischi", corretta:false},
    ]},
    {d:"21. Quale conseguenza rischia il datore di lavoro che non rispetta le norme?", r:[
      {lettera:'A', testo:"Nessuna conseguenza", corretta:false},
      {lettera:'B', testo:"Sanzioni amministrative e/o penali", corretta:true},
      {lettera:'C', testo:"Solo un richiamo verbale", corretta:false},
      {lettera:'D', testo:"Un bonus fiscale", corretta:false},
    ]},
    {d:"22. In quali casi il lavoratore può rifiutare una mansione?", r:[
      {lettera:'A', testo:"Se il compito è noioso", corretta:false},
      {lettera:'B', testo:"Se esiste un pericolo grave e immediato non controllato", corretta:true},
      {lettera:'C', testo:"Se manca un collega", corretta:false},
      {lettera:'D', testo:"Se non è il suo giorno", corretta:false},
    ]},
    {d:"23. Cosa deve garantire il datore di lavoro ai lavoratori?", r:[
      {lettera:'A', testo:"Solo produttività", corretta:false},
      {lettera:'B', testo:"Turni più lunghi", corretta:false},
      {lettera:'C', testo:"Tutela della salute e sicurezza sul lavoro", corretta:true},
      {lettera:'D', testo:"Premi di rendimento", corretta:false},
    ]},
    {d:"24. Le sanzioni per i lavoratori sono previste quando:", r:[
      {lettera:'A', testo:"Propongono miglioramenti", corretta:false},
      {lettera:'B', testo:"Violano consapevolmente le misure di sicurezza", corretta:true},
      {lettera:'C', testo:"Chiedono DPI", corretta:false},
      {lettera:'D', testo:"Raggiungono gli obiettivi", corretta:false},
    ]},
    {d:"25. Quali sono gli organi di vigilanza sulla sicurezza sul lavoro?", r:[
      {lettera:'A', testo:"Agenzie interinali", corretta:false},
      {lettera:'B', testo:"ASL/ATS, Ispettorato Nazionale del Lavoro, Vigili del Fuoco", corretta:true},
      {lettera:'C', testo:"RSPP e RLS", corretta:false},
      {lettera:'D', testo:"INPS e CAF", corretta:false},
    ]},
    {d:"26. Quale ente assicura contro gli infortuni sul lavoro?", r:[
      {lettera:'A', testo:"Prefettura", corretta:false},
      {lettera:'B', testo:"INPS", corretta:false},
      {lettera:'C', testo:"Ministero dell'Interno", corretta:false},
      {lettera:'D', testo:"INAIL", corretta:true},
    ]},
    {d:"27. Chi può svolgere ispezioni in azienda per la sicurezza?", r:[
      {lettera:'A', testo:"Il medico di base", corretta:false},
      {lettera:'B', testo:"I fornitori", corretta:false},
      {lettera:'C', testo:"Ispettorato Nazionale del Lavoro e ASL/ATS", corretta:true},
      {lettera:'D', testo:"Solo il datore di lavoro", corretta:false},
    ]},
    {d:"28. Un compito dell'ASL/ATS in materia di sicurezza è:", r:[
      {lettera:'A', testo:"Pagare stipendi", corretta:false},
      {lettera:'B', testo:"Vigilanza igienico-sanitaria e prevenzione", corretta:true},
      {lettera:'C', testo:"Vendere estintori", corretta:false},
      {lettera:'D', testo:"Gestire ferie", corretta:false},
    ]},
    {d:"29. A cosa serve la sorveglianza sanitaria?", r:[
      {lettera:'A', testo:"Aumentare i turni", corretta:false},
      {lettera:'B', testo:"Punire i lavoratori", corretta:false},
      {lettera:'C', testo:"Tutelare l'idoneità alla mansione e prevenire malattie lavoro-correlate", corretta:true},
      {lettera:'D', testo:"Stabilire le sanzioni", corretta:false},
    ]},
    {d:"30. Chi fornisce supporto tecnico e linee guida alle aziende in materia di prevenzione?", r:[
      {lettera:'A', testo:"Solo i sindacati", corretta:false},
      {lettera:'B', testo:"Camere di commercio", corretta:false},
      {lettera:'C', testo:"INAIL e organismi regionali di prevenzione", corretta:true},
      {lettera:'D', testo:"Agenzie di viaggio", corretta:false},
    ]},
  ];
}

// Domande specifiche per mansione
// ── Distribuzione risposta corretta: sequenza pseudo-casuale bilanciata ─
// Sequenza fissa (seed 11): 8×A 8×B 8×C 8×D, max 2 consecutive uguali,
// nessun pattern sequenziale ovvio (non A→B→C→D in loop).
const POS_CORRETTA = [1,3,0,1,0,0,3,3,2,0,2,3,0,1,2,1,1,3,1,3,3,1,2,2,0,2,2,0,2,1,3,0];
// 0=A 1=B 2=C 3=D → B,D,A,B,A,A,D,D,C,A,C,D,A,B,C,B,B,D,B,D,D,B,C,C,A,C,C,A,C,B,D,A

function ruotaRisposte(risposte, qIdx) {
  const LETTERE = ['A','B','C','D'];
  const posTarget = POS_CORRETTA[qIdx % POS_CORRETTA.length];
  const idxCorretta = risposte.findIndex(r => r.corretta);
  // Difensivo: se la domanda non ha (o ha più di) una risposta corretta non
  // rimescolare; ci penserà validaDomande() a far fallire la build con un
  // messaggio chiaro, invece di produrre uno swap su indice -1.
  if (idxCorretta < 0) return risposte.map((r, i) => ({ ...r, lettera: LETTERE[i] }));
  if (idxCorretta === posTarget) return risposte;
  const res = [...risposte];
  [res[idxCorretta], res[posTarget]] = [res[posTarget], res[idxCorretta]];
  return res.map((r, i) => ({ ...r, lettera: LETTERE[i] }));
}

// ── Validatore integrità domande (anti doppia-risposta / opzioni duplicate) ──
// Fa fallire la build se una domanda ha != 1 risposta corretta oppure opzioni
// con testo identico. Intercetta a monte il bug segnalato dal cliente
// ("la risposta ha due opzioni giuste uguali") su QUALSIASI test.
function validaDomande(domande, etichetta) {
  domande.forEach((q, i) => {
    const r = q.r || [];
    const corrette = r.filter(x => x.corretta).length;
    if (corrette !== 1) {
      throw new Error(`[validaDomande] ${etichetta}: domanda ${i + 1} con ${corrette} risposte corrette (ne serve esattamente 1) → "${q.d}"`);
    }
    const testi = r.map(x => (x.testo || '').trim().toLowerCase());
    const idxDup = testi.findIndex((t, idx) => t && testi.indexOf(t) !== idx);
    if (idxDup !== -1) {
      throw new Error(`[validaDomande] ${etichetta}: domanda ${i + 1} con opzioni di risposta duplicate ("${r[idxDup].testo}") → "${q.d}"`);
    }
  });
  return domande;
}

// ── Pool di distrattori plausibili (per variare le opzioni tra domande) ──────
// I distrattori "preferiti" sono i DPI/misure reali delle ALTRE voci della
// stessa mansione: sbagliati per quel rischio ma verosimili nello stesso
// contesto. Il pool generico serve solo da riempimento quando le voci reali
// non bastano. Evita il difetto segnalato dal cliente: stesse 3 risposte in
// tutte le domande, con sola variazione della corretta.
const DPI_POOL = [
  'Otoprotettori (cuffie o inserti auricolari)',
  'Maschera a pieno facciale con filtri ABEK',
  'Imbracatura anticaduta con cordino e assorbitore',
  'Occhiali di protezione a mascherina EN 166',
  'Guanti antitaglio EN 388',
  'Elmetto di protezione EN 397',
  'Calzature di sicurezza S3 con puntale e lamina',
  'Semimaschera filtrante FFP3 antipolvere',
  'Schermo facciale e grembiule per saldatura',
  'Guanti dielettrici EN 60903',
];
const MISURA_POOL = [
  'Ignorare il rischio se l\u2019attivit\u00e0 \u00e8 di breve durata',
  'Proseguire senza interruzione fino a fine turno',
  'Attendere disposizioni del datore prima di qualunque azione',
  'Aumentare il ritmo per ridurre il tempo di esposizione',
  'Rimuovere le protezioni della macchina per lavorare pi\u00f9 comodamente',
  'Affidarsi alla sola esperienza personale, senza seguire procedure',
  'Ritenere sufficiente la sola segnaletica, senza altre misure',
];

// Sceglie k distrattori distinti dal corretto, in modo deterministico ma
// "sfasato" tra domande (seed), così domande diverse mostrano opzioni diverse.
function _distrattori(pool, corretto, k, seed) {
  const cand = [];
  for (const t of pool) {
    if (t && t !== corretto && !cand.includes(t)) cand.push(t);
  }
  const res = [];
  if (cand.length) {
    let i = Math.abs(seed) % cand.length, guard = 0;
    while (res.length < k && guard < cand.length) {
      const t = cand[i % cand.length];
      if (!res.includes(t)) res.push(t);
      i++; guard++;
    }
  }
  const FALLBACK = ['Nessuna delle altre opzioni', 'Non \u00e8 prevista alcuna misura specifica', 'Tutte le opzioni indicate'];
  let f = 0;
  while (res.length < k) res.push(FALLBACK[f++ % FALLBACK.length]);
  return res.slice(0, k);
}

// Costruisce 4 opzioni (1 corretta + 3 distrattori) con lettere provvisorie.
function _quattroOpzioni(corretto, distrattori) {
  const opts = [{ testo: corretto, corretta: true },
    ...distrattori.slice(0, 3).map(t => ({ testo: t, corretta: false }))];
  return opts.map((o, i) => ({ lettera: ['A', 'B', 'C', 'D'][i], ...o }));
}

// Fonde due liste alternando in proporzione alle lunghezze (deterministico):
// distribuisce i quizExtra dentro le domande automatiche invece di accodarli.
function _interleave(a, b) {
  if (!b.length) return a.slice();
  if (!a.length) return b.slice();
  const out = []; let ia = 0, ib = 0;
  for (let k = 0; k < a.length + b.length; k++) {
    const wantA = (ia / a.length) <= (ib / b.length);
    if (wantA && ia < a.length) out.push(a[ia++]);
    else if (ib < b.length) out.push(b[ib++]);
    else if (ia < a.length) out.push(a[ia++]);
  }
  return out;
}

// Evita che due domande consecutive provengano dallo stesso rischio (_rk):
// così la domanda DPI del rischio A non è mai seguita subito dalla domanda
// "misura" dello stesso rischio A (difetto a schema fisso segnalato).
function _spezzaAdiacenze(list) {
  for (let i = 1; i < list.length; i++) {
    if (list[i]._rk != null && list[i]._rk === list[i - 1]._rk) {
      for (let j = i + 1; j < list.length; j++) {
        if (list[j]._rk !== list[i - 1]._rk) {
          [list[i], list[j]] = [list[j], list[i]];
          break;
        }
      }
    }
  }
  return list;
}

function domandeSpecifiche(mansione) {
  const rischi = mansione.rischi || [];
  const tuttiDPI = rischi.map(r => (r.dpi && r.dpi[0]) || null).filter(Boolean);
  const tutteMisure = rischi.map(r => (r.misure && r.misure[0]) || null).filter(Boolean);

  const dpiQ = [];   // una domanda DPI per rischio
  const misQ = [];   // una domanda "misura" per rischio (se presente)
  rischi.forEach((r, i) => {
    const dpiCorretto = (r.dpi && r.dpi[0]) || 'Nessun DPI specifico previsto dal DVR';
    const distrDpi = _distrattori([...tuttiDPI.filter(d => d !== dpiCorretto), ...DPI_POOL], dpiCorretto, 3, i * 7 + 1);
    dpiQ.push({
      _rk: i,
      // Dicitura ancorata al DVR aziendale: il DPI è quello previsto dal DVR,
      // non un obbligo generico (richiesta cliente: niente DPI non in DVR).
      d: `Secondo il DVR aziendale, quale DPI \u00e8 previsto come misura di protezione per il rischio "${r.nome}"?`,
      r: _quattroOpzioni(dpiCorretto, distrDpi),
    });
    if (r.misure && r.misure.length > 0) {
      const misCorretta = r.misure[0];
      const distrMis = _distrattori([...tutteMisure.filter(m => m !== misCorretta), ...MISURA_POOL], misCorretta, 3, i * 5 + 3);
      misQ.push({
        _rk: i,
        d: `Qual \u00e8 la principale misura di prevenzione prevista per il rischio "${r.nome}"?`,
        r: _quattroOpzioni(misCorretta, distrMis),
      });
    }
  });

  // ── Trasversali (3) — contenuto invariato ───────────────────────────────
  const tras = [
    {d:`In caso di infortunio durante la mansione di ${mansione.nome}, il lavoratore deve:`, r:[
      {lettera:'A',testo:'Continuare a lavorare e segnalare a fine turno',corretta:false},
      {lettera:'B',testo:'Informare immediatamente il responsabile e ricevere le cure necessarie',corretta:true},
      {lettera:'C',testo:'Recarsi autonomamente in ospedale senza avvisare nessuno',corretta:false},
      {lettera:'D',testo:'Compilare il registro presenze e proseguire',corretta:false},
    ]},
    {d:`In caso di mancato infortunio (near miss) nella mansione di ${mansione.nome}, il lavoratore deve:`, r:[
      {lettera:'A',testo:'Non segnalarlo perch\u00e9 non ha causato danni',corretta:false},
      {lettera:'B',testo:'Segnalarlo immediatamente al responsabile per prevenire futuri incidenti',corretta:true},
      {lettera:'C',testo:'Annotarlo solo se si verifica pi\u00f9 di una volta',corretta:false},
      {lettera:'D',testo:'Segnalarlo solo se ci sono testimoni',corretta:false},
    ]},
    {d:`Cosa si intende per stress lavoro-correlato nella mansione di ${mansione.nome}?`, r:[
      {lettera:'A',testo:'La stanchezza fisica dopo una giornata di lavoro intensa',corretta:false},
      {lettera:'B',testo:'Una condizione derivante da fattori di rischio psicosociali che possono nuocere alla salute',corretta:true},
      {lettera:'C',testo:'Un problema che riguarda solo i dirigenti',corretta:false},
      {lettera:'D',testo:'Un disturbo muscolare da sforzo eccessivo',corretta:false},
    ]},
  ];

  // ── quizExtra (calibrati sulla mansione) ───────────────────────────────
  const extra = (mansione.quizExtra || []).map(q => ({
    d: String(q.d).replace(/^\d+\.\s*/, ''),
    r: q.r,
  }));

  // ── ORDINAMENTO ────────────────────────────────────────────────────────
  // 1) tutti i DPI, poi tutte le misure, poi le trasversali: separa per
  //    costruzione la coppia dpi[i]/misura[i] dello stesso rischio.
  // 2) interleaving deterministico con i quizExtra (non più accodati in blocco)
  // 3) anti-adiacenza per rischio
  const auto = [...dpiQ, ...misQ, ...tras];
  let ordered = _spezzaAdiacenze(_interleave(auto, extra)).slice(0, 30);

  // Rotazione posizione corretta + numerazione progressiva finale.
  return ordered.map((q, idx) => ({
    d: `${idx + 1}. ${String(q.d).replace(/^\d+\.\s*/, '')}`,
    r: ruotaRisposte(q.r, idx),
  }));
}


async function genTestGenerale(cliente) {
  const domande = [
    {d:"1. Che cosa si intende per 'pericolo' in ambito lavorativo?", r:[
      {lettera:'A', testo:"Una proprietà o condizione che può causare un danno", corretta:true},
      {lettera:'B', testo:"Un evento senza conseguenze", corretta:false},
      {lettera:'C', testo:"Una persona incaricata della sicurezza", corretta:false},
      {lettera:'D', testo:"Una procedura amministrativa", corretta:false},
    ]},
    {d:"2. Come si definisce 'rischio'?", r:[
      {lettera:'A', testo:"La sola gravità del danno", corretta:false},
      {lettera:'B', testo:"Probabilità e gravità di un danno", corretta:true},
      {lettera:'C', testo:"Il costo della prevenzione", corretta:false},
      {lettera:'D', testo:"Un comportamento sicuro", corretta:false},
    ]},
    {d:"3. Cosa si intende per 'danno'?", r:[
      {lettera:'A', testo:"Una misura preventiva", corretta:false},
      {lettera:'B', testo:"Una segnalazione", corretta:false},
      {lettera:'C', testo:"Conseguenza negativa su persone o cose", corretta:true},
      {lettera:'D', testo:"Un obbligo formale", corretta:false},
    ]},
    {d:"4. Quale delle seguenti situazioni rappresenta un rischio elevato?", r:[
      {lettera:'A', testo:"Ambiente ordinato", corretta:false},
      {lettera:'B', testo:"Usare DPI adeguati", corretta:false},
      {lettera:'C', testo:"Lavorare in quota senza protezioni", corretta:true},
      {lettera:'D', testo:"Seguire la procedura", corretta:false},
    ]},
    {d:"5. Qual è l'obiettivo della valutazione dei rischi?", r:[
      {lettera:'A', testo:"Valutare la produttività", corretta:false},
      {lettera:'B', testo:"Individuare pericoli e definire misure di prevenzione", corretta:true},
      {lettera:'C', testo:"Pianificare le ferie", corretta:false},
      {lettera:'D', testo:"Organizzare i turni", corretta:false},
    ]},
    {d:"6. Quale intervento riduce maggiormente il rischio?", r:[
      {lettera:'A', testo:"Solo segnaletica", corretta:false},
      {lettera:'B', testo:"Uso sporadico dei DPI", corretta:false},
      {lettera:'C', testo:"Eliminazione del pericolo alla fonte", corretta:true},
      {lettera:'D', testo:"Controlli saltuari", corretta:false},
    ]},
    {d:"7. Che cosa significa 'prevenzione' sul luogo di lavoro?", r:[
      {lettera:'A', testo:"Limitare i danni dopo l'evento", corretta:false},
      {lettera:'B', testo:"Evitare che accadano eventi dannosi", corretta:true},
      {lettera:'C', testo:"Curare gli infortunati", corretta:false},
      {lettera:'D', testo:"Archiviare documenti", corretta:false},
    ]},
    {d:"8. Che cosa significa 'protezione'?", r:[
      {lettera:'A', testo:"Eliminare ogni rischio", corretta:false},
      {lettera:'B', testo:"Migliorare la produttività", corretta:false},
      {lettera:'C', testo:"Ridurre le conseguenze di un evento pericoloso", corretta:true},
      {lettera:'D', testo:"Trasferire il rischio ai fornitori", corretta:false},
    ]},
    {d:"9. La formazione dei lavoratori è una misura di:", r:[
      {lettera:'A', testo:"Protezione", corretta:false},
      {lettera:'B', testo:"Sanzione", corretta:false},
      {lettera:'C', testo:"Sorveglianza sanitaria", corretta:false},
      {lettera:'D', testo:"Prevenzione", corretta:true},
    ]},
    {d:"10. Un estintore è un esempio di misura di:", r:[
      {lettera:'A', testo:"Prevenzione", corretta:false},
      {lettera:'B', testo:"Partecipazione", corretta:false},
      {lettera:'C', testo:"Protezione", corretta:true},
      {lettera:'D', testo:"Organizzazione", corretta:false},
    ]},
    {d:"11. Nell'ordine delle priorità della prevenzione qual è la prima scelta?", r:[
      {lettera:'A', testo:"Fornire DPI", corretta:false},
      {lettera:'B', testo:"Eliminare il pericolo alla fonte", corretta:false},
      {lettera:'C', testo:"Mettere cartelli", corretta:false},
      {lettera:'D', testo:"Formare i lavoratori", corretta:true},
    ]},
    {d:"12. Che cos'è una procedura di lavoro sicuro?", r:[
      {lettera:'A', testo:"Una nota informale tra colleghi", corretta:false},
      {lettera:'B', testo:"Un documento contabile", corretta:false},
      {lettera:'C', testo:"Istruzione scritta che standardizza attività in sicurezza", corretta:true},
      {lettera:'D', testo:"Un suggerimento non vincolante", corretta:false},
    ]},
    {d:"13. Qual è il compito principale dell'RSPP?", r:[
      {lettera:'A', testo:"Redigere cartelle cliniche", corretta:false},
      {lettera:'B', testo:"Coordinare attività di prevenzione e protezione", corretta:true},
      {lettera:'C', testo:"Sostituire il datore di lavoro", corretta:false},
      {lettera:'D', testo:"Rappresentare i lavoratori", corretta:false},
    ]},
    {d:"14. Chi è l'RLS?", r:[
      {lettera:'A', testo:"Un revisore legale", corretta:false},
      {lettera:'B', testo:"Rappresentante dei Lavoratori per la Sicurezza", corretta:true},
      {lettera:'C', testo:"Responsabile delle risorse umane", corretta:false},
      {lettera:'D', testo:"Medico competente", corretta:false},
    ]},
    {d:"15. Chi nomina il medico competente in azienda?", r:[
      {lettera:'A', testo:"Il datore di lavoro", corretta:true},
      {lettera:'B', testo:"L'INAIL", corretta:false},
      {lettera:'C', testo:"L'RSPP", corretta:false},
      {lettera:'D', testo:"Il lavoratore", corretta:false},
    ]},
    {d:"16. Chi partecipa alla riunione periodica di prevenzione?", r:[
      {lettera:'A', testo:"Solo i lavoratori", corretta:false},
      {lettera:'B', testo:"Solo i dirigenti", corretta:false},
      {lettera:'C', testo:"Datore di lavoro, RSPP, medico competente, RLS", corretta:true},
      {lettera:'D', testo:"Solo il preposto", corretta:false},
    ]},
    {d:"17. Se un lavoratore individua un pericolo, cosa deve fare immediatamente?", r:[
      {lettera:'A', testo:"Risolverlo da solo", corretta:false},
      {lettera:'B', testo:"Aspettare la riunione periodica", corretta:false},
      {lettera:'C', testo:"Segnalare al preposto o datore di lavoro", corretta:true},
      {lettera:'D', testo:"Ignorarlo", corretta:false},
    ]},
    {d:"18. Qual è il ruolo del preposto?", r:[
      {lettera:'A', testo:"Emettere sanzioni disciplinari", corretta:false},
      {lettera:'B', testo:"Sovrintendere e vigilare sul rispetto delle misure di sicurezza", corretta:true},
      {lettera:'C', testo:"Redigere la valutazione dei rischi", corretta:false},
      {lettera:'D', testo:"Svolgere funzioni mediche", corretta:false},
    ]},
    {d:"19. Quale diritto fondamentale hanno i lavoratori in materia di sicurezza?", r:[
      {lettera:'A', testo:"Decidere le procedure", corretta:false},
      {lettera:'B', testo:"Rifiutare ogni mansione", corretta:false},
      {lettera:'C', testo:"Ricevere formazione, informazione e addestramento", corretta:true},
      {lettera:'D', testo:"Non usare DPI", corretta:false},
    ]},
    {d:"20. Qual è un obbligo del lavoratore?", r:[
      {lettera:'A', testo:"Nominare l'RSPP", corretta:false},
      {lettera:'B', testo:"Usare correttamente i DPI forniti", corretta:true},
      {lettera:'C', testo:"Gestire la sorveglianza sanitaria", corretta:false},
      {lettera:'D', testo:"Redigere la valutazione dei rischi", corretta:false},
    ]},
    {d:"21. Quale conseguenza rischia il datore di lavoro che non rispetta le norme?", r:[
      {lettera:'A', testo:"Nessuna conseguenza", corretta:false},
      {lettera:'B', testo:"Sanzioni amministrative e/o penali", corretta:true},
      {lettera:'C', testo:"Solo un richiamo verbale", corretta:false},
      {lettera:'D', testo:"Un bonus fiscale", corretta:false},
    ]},
    {d:"22. In quali casi il lavoratore può rifiutare una mansione?", r:[
      {lettera:'A', testo:"Se il compito è noioso", corretta:false},
      {lettera:'B', testo:"Se esiste un pericolo grave e immediato non controllato", corretta:true},
      {lettera:'C', testo:"Se manca un collega", corretta:false},
      {lettera:'D', testo:"Se non è il suo giorno", corretta:false},
    ]},
    {d:"23. Cosa deve garantire il datore di lavoro ai lavoratori?", r:[
      {lettera:'A', testo:"Solo produttività", corretta:false},
      {lettera:'B', testo:"Turni più lunghi", corretta:false},
      {lettera:'C', testo:"Tutela della salute e sicurezza sul lavoro", corretta:true},
      {lettera:'D', testo:"Premi di rendimento", corretta:false},
    ]},
    {d:"24. Le sanzioni per i lavoratori sono previste quando:", r:[
      {lettera:'A', testo:"Propongono miglioramenti", corretta:false},
      {lettera:'B', testo:"Violano consapevolmente le misure di sicurezza", corretta:true},
      {lettera:'C', testo:"Chiedono DPI", corretta:false},
      {lettera:'D', testo:"Raggiungono gli obiettivi", corretta:false},
    ]},
    {d:"25. Quali sono gli organi di vigilanza sulla sicurezza sul lavoro?", r:[
      {lettera:'A', testo:"Agenzie interinali", corretta:false},
      {lettera:'B', testo:"ASL/ATS, Ispettorato Nazionale del Lavoro, Vigili del Fuoco", corretta:true},
      {lettera:'C', testo:"RSPP e RLS", corretta:false},
      {lettera:'D', testo:"INPS e CAF", corretta:false},
    ]},
    {d:"26. Quale ente assicura contro gli infortuni sul lavoro?", r:[
      {lettera:'A', testo:"Prefettura", corretta:false},
      {lettera:'B', testo:"INPS", corretta:false},
      {lettera:'C', testo:"Ministero dell'Interno", corretta:false},
      {lettera:'D', testo:"INAIL", corretta:true},
    ]},
    {d:"27. Chi può svolgere ispezioni in azienda per la sicurezza?", r:[
      {lettera:'A', testo:"Il medico di base", corretta:false},
      {lettera:'B', testo:"I fornitori", corretta:false},
      {lettera:'C', testo:"Ispettorato Nazionale del Lavoro e ASL/ATS", corretta:true},
      {lettera:'D', testo:"Solo il datore di lavoro", corretta:false},
    ]},
    {d:"28. Un compito dell'ASL/ATS in materia di sicurezza è:", r:[
      {lettera:'A', testo:"Pagare stipendi", corretta:false},
      {lettera:'B', testo:"Vigilanza igienico-sanitaria e prevenzione", corretta:true},
      {lettera:'C', testo:"Vendere estintori", corretta:false},
      {lettera:'D', testo:"Gestire ferie", corretta:false},
    ]},
    {d:"29. A cosa serve la sorveglianza sanitaria?", r:[
      {lettera:'A', testo:"Aumentare i turni", corretta:false},
      {lettera:'B', testo:"Punire i lavoratori", corretta:false},
      {lettera:'C', testo:"Tutelare l'idoneità alla mansione e prevenire malattie lavoro-correlate", corretta:true},
      {lettera:'D', testo:"Stabilire le sanzioni", corretta:false},
    ]},
    {d:"30. Chi fornisce supporto tecnico e linee guida alle aziende in materia di prevenzione?", r:[
      {lettera:'A', testo:"Solo i sindacati", corretta:false},
      {lettera:'B', testo:"Camere di commercio", corretta:false},
      {lettera:'C', testo:"INAIL e organismi regionali di prevenzione", corretta:true},
      {lettera:'D', testo:"Agenzie di viaggio", corretta:false},
    ]},
  ];
  validaDomande(domande, 'Test Generale');
  // ── Costruisce e salva Test_Formazione_Generale (allievo + docente) ──
  const MARGIN = { top: 709, right: 1134, bottom: 1134, left: 1134 };
  const header = makeHeaderTest();
  const footer = makeFooterTest();

  function buildChildren(isDocente) {
    return [
      makeTopTable(),
      vuoto(20),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'TEST DI APPRENDIMENTO –',bold:true,font:FONT,size:28,color:C.BLU_DARK})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:isDocente?40:40},children:[new TextRun({text:'FORMAZIONE GENERALE',bold:true,font:FONT,size:28,color:C.BLU_DARK})]}),
      ...(isDocente?[new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'– VERSIONE DOCENTE – CON RISPOSTE EVIDENZIATE –',bold:true,font:FONT,size:20,color:C.ROSSO})]})]:[]),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:'Formazione obbligatoria – D.Lgs. 81/08 e Accordo Stato-Regioni 17/04/2025',italics:true,font:FONT,size:18,color:C.GRIGIO})]}),
      makeDiscente(),
      vuoto(20),
      ...(!isDocente?[new Paragraph({spacing:{after:200},children:[new TextRun({text:'ISTRUZIONI: Per ogni domanda, barrare la risposta che si ritiene corretta (A, B, C oppure D). È ammessa una sola risposta per domanda. Durata: 30 minuti. Punteggio minimo per il superamento: 21/30 (70%).',italics:true,font:FONT,size:18})]})]:[]),
      ...domande.flatMap(({d,r}) => [makeQuestion(d,r,isDocente), vuoto(10)]),
      ...firme(),
    ];
  }

  const doc = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(false)}]});
  await salvaDoc(doc, `${OUT}/03 - TEST FINALI DI APPRENDIMENTO/Generale/Test_Formazione_Generale.docx`);
  const docD = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(true)}]});
  await salvaDoc(docD, `${OUT}/03 - TEST FINALI DI APPRENDIMENTO/Generale/Test_Formazione_Generale DOCENTE.docx`);
}


async function genTestMansione(mansione) {
  const MARGIN = { top: 709, right: 1134, bottom: 1134, left: 1134 };
  const header = makeHeaderTest();
  const footer = makeFooterTest();
  const domande = validaDomande(domandeSpecifiche(mansione), `Test specifico – ${mansione.nome}`);

  function buildChildren(isDocente) {
    return [
      makeTopTable(),
      vuoto(20),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'TEST SPECIFICO –',bold:true,font:FONT,size:28,color:C.BLU_DARK})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:isDocente?40:40},children:[new TextRun({text:mansione.nome.toUpperCase(),bold:true,font:FONT,size:28,color:C.BLU_DARK})]}),
      ...(isDocente?[new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'– VERSIONE DOCENTE – CON RISPOSTE EVIDENZIATE –',bold:true,font:FONT,size:20,color:C.ROSSO})]})]:[]),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:'Formazione obbligatoria – D.Lgs. 81/08 e Accordo Stato-Regioni 17/04/2025',italics:true,font:FONT,size:18,color:C.GRIGIO})]}),
      makeDiscente(),
      vuoto(20),
      ...(!isDocente?[new Paragraph({spacing:{after:200},children:[new TextRun({text:'ISTRUZIONI: Per ogni domanda, barrare la risposta che si ritiene corretta (A, B, C oppure D). È ammessa una sola risposta per domanda. Durata: 30 minuti. Punteggio minimo per il superamento: 21/30 (70%).',italics:true,font:FONT,size:18})]})]:[]),
      ...domande.flatMap(({d,r}) => [makeQuestion(d,r,isDocente), vuoto(10)]),
      ...firme(),
    ];
  }

  const doc = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(false)}]});
  await salvaDoc(doc, `${OUT}/03 - TEST FINALI DI APPRENDIMENTO/Specifica/Test_${mansione.id}.docx`);
  const docD = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(true)}]});
  await salvaDoc(docD, `${OUT}/03 - TEST FINALI DI APPRENDIMENTO/Specifica/Test_${mansione.id} DOCENTE.docx`);
}

module.exports = { genTestGenerale, genTestMansione };
