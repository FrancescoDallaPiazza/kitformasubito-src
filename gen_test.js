'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  SimpleField,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, logoBytes,
  vuoto, cella, salvaDoc,
} = h;

const OUT = `/home/claude/kit/OUT/KIT FORMASUBITO - ${CLIENTE.ragioneSocialeBreve}`;

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
function domandeSpecifiche(mansione) {
  const domande = [];
  let n = 1;
  mansione.rischi.forEach(r => {
    domande.push({d:`${n++}. Quale DPI è specificamente richiesto per il rischio "${r.nome}"?`,r:[
      {lettera:'A',testo:r.dpi[0]||'Nessun DPI specifico',corretta:true},
      {lettera:'B',testo:'Otoprotettori',corretta:false},
      {lettera:'C',testo:'Maschera antigas integrale',corretta:false},
      {lettera:'D',testo:'Imbracatura anticaduta',corretta:false},
    ]});
    if (r.misure.length > 0) {
      domande.push({d:`${n++}. Qual è la principale misura di prevenzione per il rischio "${r.nome}"?`,r:[
        {lettera:'A',testo:r.misure[0],corretta:true},
        {lettera:'B',testo:'Ignorare il rischio se di breve durata',corretta:false},
        {lettera:'C',testo:'Continuare a lavorare senza interruzione',corretta:false},
        {lettera:'D',testo:'Attendere disposizioni del datore di lavoro prima di ogni azione',corretta:false},
      ]});
    }
  });
  domande.push({d:`${n++}. In caso di infortunio durante la mansione di ${mansione.nome}, il lavoratore deve:`,r:[{lettera:'A',testo:'Continuare a lavorare e segnalare a fine turno',corretta:false},{lettera:'B',testo:'Informare immediatamente il responsabile e ricevere le cure necessarie',corretta:true},{lettera:'C',testo:'Recarsi autonomamente in ospedale senza avvisare nessuno',corretta:false},{lettera:'D',testo:'Compilare il registro presenze e proseguire',corretta:false}]});
  domande.push({d:`${n++}. In caso di mancato infortunio (near miss) nella mansione di ${mansione.nome}, il lavoratore deve:`,r:[{lettera:'A',testo:'Non segnalarlo perché non ha causato danni',corretta:false},{lettera:'B',testo:'Segnalarlo immediatamente al responsabile per prevenire futuri incidenti',corretta:true},{lettera:'C',testo:'Annotarlo solo se si verifica più di una volta',corretta:false},{lettera:'D',testo:'Segnalarlo solo se ci sono testimoni',corretta:false}]});
  domande.push({d:`${n++}. Cosa si intende per stress lavoro-correlato nella mansione di ${mansione.nome}?`,r:[{lettera:'A',testo:'La stanchezza fisica dopo una giornata di lavoro intensa',corretta:false},{lettera:'B',testo:'Una condizione derivante da fattori di rischio psicosociali che possono nuocere alla salute',corretta:true},{lettera:'C',testo:'Un problema che riguarda solo i dirigenti',corretta:false},{lettera:'D',testo:'Un disturbo muscolare da sforzo eccessivo',corretta:false}]});
  // ── Quiz extra calibrati sulla mansione (da helpers.js → quizExtra) ──
  const extra = (mansione.quizExtra || []).map((q, i) => ({
    d: `${n + i} ${q.d.replace(/^\d+\.\s*/, '')}`,
    r: q.r,
  }));
  const all = [...domande, ...extra];
  // Rinumera in ordine progressivo
  return all.map((q, idx) => ({...q, d: q.d.replace(/^\d+\./, `${idx + 1}.`)}));
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
  return domande;
}


async function genTestMansione(mansione) {
  const MARGIN = { top: 709, right: 1134, bottom: 1134, left: 1134 };
  const header = makeHeaderTest();
  const footer = makeFooterTest();
  const domande = domandeSpecifiche(mansione);

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
  await salvaDoc(doc, `${OUT}/03 - TEST FINALI DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_${mansione.id}.docx`);
  const docD = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(true)}]});
  await salvaDoc(docD, `${OUT}/03 - TEST FINALI DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_${mansione.id} DOCENTE.docx`);
}

module.exports = { genTestGenerale, genTestMansione };
