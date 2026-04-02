'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  SimpleField,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, logoBytes,
  vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

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

// 30 domande generali
function domandeGenerali() {
  return [
    {d:'1. Cosa si intende per \'rischio\' nel D.Lgs. 81/08?',r:[{lettera:'A',testo:'La probabilità che si verifichi un evento dannoso in assenza di misure preventive',corretta:true},{lettera:'B',testo:'La certezza che un danno si verificherà',corretta:false},{lettera:'C',testo:'Un incidente già avvenuto in azienda',corretta:false},{lettera:'D',testo:'Una sanzione prevista dall\'autorità di controllo',corretta:false}]},
    {d:'2. Chi è il Datore di Lavoro ai sensi del D.Lgs. 81/08?',r:[{lettera:'A',testo:'Il responsabile del servizio di prevenzione',corretta:false},{lettera:'B',testo:'Il rappresentante dei lavoratori per la sicurezza',corretta:false},{lettera:'C',testo:'Il soggetto che ha la responsabilità dell\'impresa e il potere decisionale',corretta:true},{lettera:'D',testo:'Il medico competente',corretta:false}]},
    {d:'3. Qual è il principale obbligo del lavoratore in materia di sicurezza?',r:[{lettera:'A',testo:'Redigere il Documento di Valutazione dei Rischi',corretta:false},{lettera:'B',testo:'Prendersi cura della propria salute e sicurezza e di quella degli altri',corretta:true},{lettera:'C',testo:'Nominare il Responsabile del Servizio di Prevenzione e Protezione',corretta:false},{lettera:'D',testo:'Effettuare la sorveglianza sanitaria',corretta:false}]},
    {d:'4. Cosa si intende per DPI?',r:[{lettera:'A',testo:'Documento Progettuale Integrato',corretta:false},{lettera:'B',testo:'Dispositivo di Protezione Individuale',corretta:true},{lettera:'C',testo:'Direttiva sulla Protezione degli Impianti',corretta:false},{lettera:'D',testo:'Decreto di Prevenzione Industriale',corretta:false}]},
    {d:'5. Chi è l\'RSPP?',r:[{lettera:'A',testo:'Il rappresentante sindacale dei lavoratori',corretta:false},{lettera:'B',testo:'Il medico che effettua le visite di idoneità',corretta:false},{lettera:'C',testo:'Il Responsabile del Servizio di Prevenzione e Protezione',corretta:true},{lettera:'D',testo:'Il Responsabile della qualità aziendale',corretta:false}]},
    {d:'6. Cosa si intende per \'pericolo\' nel D.Lgs. 81/08?',r:[{lettera:'A',testo:'La probabilità che un danno si verifichi',corretta:false},{lettera:'B',testo:'La proprietà intrinseca di un agente o di una situazione di poter causare danni',corretta:true},{lettera:'C',testo:'Un infortunio già avvenuto',corretta:false},{lettera:'D',testo:'Un difetto di funzionamento di un macchinario',corretta:false}]},
    {d:'7. Cosa si intende per \'near miss\' (mancato infortunio)?',r:[{lettera:'A',testo:'Un infortunio che ha causato danni gravi',corretta:false},{lettera:'B',testo:'Una malattia professionale già dichiarata',corretta:false},{lettera:'C',testo:'Un evento che avrebbe potuto causare un infortunio ma che per puro caso non lo ha fatto',corretta:true},{lettera:'D',testo:'Un controllo ispettivo senza esito negativo',corretta:false}]},
    {d:'8. Qual è il significato del segnale di sicurezza circolare con bordo rosso?',r:[{lettera:'A',testo:'Indica una condizione di pericolo generico',corretta:false},{lettera:'B',testo:'Indica un obbligo da rispettare',corretta:false},{lettera:'C',testo:'Indica il divieto di compiere una determinata azione',corretta:true},{lettera:'D',testo:'Indica la posizione di un presidio di primo soccorso',corretta:false}]},
    {d:'9. Il Documento di Valutazione dei Rischi (DVR) deve essere redatto da:',r:[{lettera:'A',testo:'Il lavoratore più anziano in azienda',corretta:false},{lettera:'B',testo:'Il medico competente, autonomamente',corretta:false},{lettera:'C',testo:'L\'organo di vigilanza (ASL/INL)',corretta:false},{lettera:'D',testo:'Il Datore di Lavoro, con la collaborazione del RSPP e del Medico Competente',corretta:true}]},
    {d:'10. Cosa si intende per \'stress lavoro-correlato\'?',r:[{lettera:'A',testo:'Un tipo di infortunio sul lavoro dovuto a sforzi fisici eccessivi',corretta:false},{lettera:'B',testo:'Una condizione derivante da fattori di rischio psicosociali che possono nuocere alla salute',corretta:true},{lettera:'C',testo:'Un problema che riguarda esclusivamente i dirigenti aziendali',corretta:false},{lettera:'D',testo:'Una malattia infettiva contratta sul luogo di lavoro',corretta:false}]},
    {d:'11. Qual è la frequenza minima di aggiornamento della formazione per i lavoratori (ASR 2025)?',r:[{lettera:'A',testo:'Ogni anno',corretta:false},{lettera:'B',testo:'Ogni 3 anni',corretta:false},{lettera:'C',testo:'Ogni 5 anni',corretta:true},{lettera:'D',testo:'Solo in caso di cambio di mansione',corretta:false}]},
    {d:'12. Cosa deve fare il lavoratore se rileva un\'anomalia o un rischio non previsto?',r:[{lettera:'A',testo:'Continuare a lavorare e riferire al termine del turno',corretta:false},{lettera:'B',testo:'Risolvere autonomamente il problema senza informare nessuno',corretta:false},{lettera:'C',testo:'Segnalare immediatamente al preposto o al datore di lavoro',corretta:true},{lettera:'D',testo:'Attendere l\'ispezione periodica dell\'ASL',corretta:false}]},
    {d:'13. Quale segnale di sicurezza indica l\'obbligo di indossare i DPI?',r:[{lettera:'A',testo:'Segnale di forma triangolare con bordo giallo',corretta:false},{lettera:'B',testo:'Segnale di forma circolare con fondo azzurro',corretta:true},{lettera:'C',testo:'Segnale di forma rettangolare con fondo verde',corretta:false},{lettera:'D',testo:'Segnale di forma quadrata con bordo rosso',corretta:false}]},
    {d:'14. Il RLS (Rappresentante dei Lavoratori per la Sicurezza) viene:',r:[{lettera:'A',testo:'Nominato direttamente dal datore di lavoro',corretta:false},{lettera:'B',testo:'Eletto o designato dai lavoratori nell\'ambito delle rappresentanze sindacali',corretta:true},{lettera:'C',testo:'Nominato dall\'ASL competente per territorio',corretta:false},{lettera:'D',testo:'Scelto dal medico competente tra il personale sanitario',corretta:false}]},
    {d:'15. Cosa significa il triangolo giallo con punto esclamativo nella segnaletica?',r:[{lettera:'A',testo:'Obbligo di usare i DPI',corretta:false},{lettera:'B',testo:'Presenza di un estintore nelle vicinanze',corretta:false},{lettera:'C',testo:'Segnale di avvertimento di un pericolo generico',corretta:true},{lettera:'D',testo:'Uscita di emergenza nelle vicinanze',corretta:false}]},
    {d:'16. Chi effettua la sorveglianza sanitaria in azienda?',r:[{lettera:'A',testo:'L\'RSPP durante i sopralluoghi periodici',corretta:false},{lettera:'B',testo:'Il medico competente nominato dal datore di lavoro',corretta:true},{lettera:'C',testo:'Il medico di base del lavoratore',corretta:false},{lettera:'D',testo:'L\'ispettore dell\'ASL in occasione dei controlli',corretta:false}]},
    {d:'17. Cosa si intende per \'prevenzione\' ai sensi del D.Lgs. 81/08?',r:[{lettera:'A',testo:'Il complesso delle disposizioni e misure adottate per evitare o ridurre i rischi',corretta:true},{lettera:'B',testo:'Le cure mediche prestate dopo un infortunio',corretta:false},{lettera:'C',testo:'L\'insieme delle sanzioni per chi non rispetta le norme',corretta:false},{lettera:'D',testo:'La valutazione dei danni causati da un incidente',corretta:false}]},
    {d:'18. In caso di infortunio sul lavoro, il lavoratore deve:',r:[{lettera:'A',testo:'Recarsi autonomamente al pronto soccorso senza avvisare il datore',corretta:false},{lettera:'B',testo:'Attendere la fine del turno per riferire dell\'accaduto',corretta:false},{lettera:'C',testo:'Informare immediatamente il preposto o il datore di lavoro',corretta:true},{lettera:'D',testo:'Contattare direttamente l\'INAIL senza passare dall\'azienda',corretta:false}]},
    {d:'19. Qual è il colore dei segnali di salvataggio e di soccorso (uscite di emergenza)?',r:[{lettera:'A',testo:'Giallo su fondo nero',corretta:false},{lettera:'B',testo:'Verde su fondo bianco oppure bianco su fondo verde',corretta:true},{lettera:'C',testo:'Rosso su fondo bianco',corretta:false},{lettera:'D',testo:'Blu su fondo bianco',corretta:false}]},
    {d:'20. Cosa si intende per \'pericolo\' nel D.Lgs. 81/08?',r:[{lettera:'A',testo:'La probabilità che un danno si verifichi',corretta:false},{lettera:'B',testo:'La proprietà intrinseca di un agente di poter causare danni',corretta:true},{lettera:'C',testo:'Un infortunio già avvenuto',corretta:false},{lettera:'D',testo:'Un difetto di funzionamento di un macchinario',corretta:false}]},
    {d:'21. I lavoratori hanno l\'obbligo di partecipare ai programmi di formazione?',r:[{lettera:'A',testo:'No, è facoltativo',corretta:false},{lettera:'B',testo:'Sì, è un obbligo di legge sancito dal D.Lgs. 81/08',corretta:true},{lettera:'C',testo:'Solo i lavoratori con mansioni a rischio alto',corretta:false},{lettera:'D',testo:'Solo i nuovi assunti nei primi 6 mesi',corretta:false}]},
    {d:'22. In quale caso il datore di lavoro può essere esonerato dall\'obbligo di formare i lavoratori?',r:[{lettera:'A',testo:'Se l\'azienda ha meno di 5 dipendenti',corretta:false},{lettera:'B',testo:'Se i lavoratori hanno già un\'esperienza superiore a 10 anni',corretta:false},{lettera:'C',testo:'Se il medico competente certifica l\'idoneità dei lavoratori',corretta:false},{lettera:'D',testo:'Non esiste alcun caso di esonero: la formazione è sempre obbligatoria',corretta:true}]},
    {d:'23. Cosa indica il colore rosso nella segnaletica antincendio e di sicurezza?',r:[{lettera:'A',testo:'Percorsi sicuri e vie di esodo',corretta:false},{lettera:'B',testo:'Pericoli generici da tenere a mente',corretta:false},{lettera:'C',testo:'Attrezzature antincendio, divieti e pericolo imminente',corretta:true},{lettera:'D',testo:'Obbligo di utilizzare i DPI',corretta:false}]},
    {d:'24. Quale documento attesta che il lavoratore ha ricevuto la formazione obbligatoria?',r:[{lettera:'A',testo:'Il registro presenze del corso firmato dal lavoratore',corretta:true},{lettera:'B',testo:'Il solo certificato medico di idoneità',corretta:false},{lettera:'C',testo:'La lettera di assunzione',corretta:false},{lettera:'D',testo:'Il verbale dell\'ultima ispezione ASL',corretta:false}]},
    {d:'25. L\'art. 37 del D.Lgs. 81/08 stabilisce che la formazione dei lavoratori deve essere erogata:',r:[{lettera:'A',testo:'In orario extralavorativo a cura del lavoratore',corretta:false},{lettera:'B',testo:'Durante l\'orario di lavoro e con costi a carico del datore di lavoro',corretta:true},{lettera:'C',testo:'Annualmente, indipendentemente dai rischi presenti',corretta:false},{lettera:'D',testo:'Esclusivamente online tramite piattaforme accreditate',corretta:false}]},
    {d:'26. A chi spetta la responsabilità principale di garantire la sicurezza sul lavoro?',r:[{lettera:'A',testo:'All\'RSPP',corretta:false},{lettera:'B',testo:'Al medico competente',corretta:false},{lettera:'C',testo:'Al Datore di Lavoro',corretta:true},{lettera:'D',testo:'Ai lavoratori stessi',corretta:false}]},
    {d:'27. Cosa deve fare il lavoratore con i DPI ricevuti?',r:[{lettera:'A',testo:'Utilizzarli solo quando è presente il datore di lavoro',corretta:false},{lettera:'B',testo:'Conservarli in armadietto senza mai indossarli',corretta:false},{lettera:'C',testo:'Utilizzarli conformemente alle istruzioni, mantenerli in buono stato e segnalare eventuali difetti',corretta:true},{lettera:'D',testo:'Riporli a fine turno e non portarli a casa',corretta:false}]},
    {d:'28. Qual è l\'organo di vigilanza in materia di salute e sicurezza sul lavoro?',r:[{lettera:'A',testo:'L\'INPS',corretta:false},{lettera:'B',testo:'Il Ministero del Lavoro autonomamente',corretta:false},{lettera:'C',testo:'ASL/ATS e Ispettorato Nazionale del Lavoro (INL)',corretta:true},{lettera:'D',testo:'Solo le organizzazioni sindacali',corretta:false}]},
    {d:'29. Cosa si intende per \'mansione\' ai fini della sicurezza?',r:[{lettera:'A',testo:'La retribuzione percepita dal lavoratore',corretta:false},{lettera:'B',testo:'L\'insieme di compiti e attività assegnati a un lavoratore, che determinano la natura dei rischi cui è esposto',corretta:true},{lettera:'C',testo:'Il livello di inquadramento contrattuale',corretta:false},{lettera:'D',testo:'Il numero di ore lavorate nel mese',corretta:false}]},
    {d:'30. In presenza di un incendio in azienda, il lavoratore deve:',r:[{lettera:'A',testo:'Tentare di spegnere l\'incendio da solo prima di chiamare i soccorsi',corretta:false},{lettera:'B',testo:'Abbandonare immediatamente il luogo seguendo le vie di esodo, segnalare e attendere i soccorsi',corretta:true},{lettera:'C',testo:'Recuperare gli effetti personali prima di evacuare',corretta:false},{lettera:'D',testo:'Attendere le istruzioni telefoniche dei Vigili del Fuoco',corretta:false}]},
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

async function genTestGenerale() {
  const MARGIN = { top: 709, right: 1134, bottom: 1134, left: 1134 };
  const header = makeHeaderTest();
  const footer = makeFooterTest();
  const domande = domandeGenerali();

  function buildChildren(isDocente) {
    return [
      makeTopTable(),
      vuoto(20),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'TEST DI APPRENDIMENTO – FORMAZIONE GENERALE',bold:true,font:FONT,size:28,color:C.BLU_DARK})]}),
      ...(isDocente?[new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:40},children:[new TextRun({text:'– VERSIONE DOCENTE – CON RISPOSTE EVIDENZIATE –',bold:true,font:FONT,size:20,color:C.ROSSO})]})]:[]),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:'Formazione obbligatoria – D.Lgs. 81/08 e Accordo Stato-Regioni 17/04/2025',italic:true,font:FONT,size:18,color:C.GRIGIO})]}),
      makeDiscente(),
      vuoto(20),
      ...(!isDocente?[new Paragraph({spacing:{after:200},children:[new TextRun({text:'ISTRUZIONI: Per ogni domanda, barrare la risposta che si ritiene corretta (A, B, C oppure D). È ammessa una sola risposta per domanda. Durata: 30 minuti. Punteggio minimo per il superamento: 21/30 (70%).',italic:true,font:FONT,size:18})]})]:[]),
      ...domande.flatMap(({d,r}) => [makeQuestion(d,r,isDocente), vuoto(10)]),
      ...firme(),
    ];
  }

  const doc = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(false)}]});
  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_Generale.docx`);
  const docD = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(true)}]});
  await salvaDoc(docD, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_Generale_Docente.docx`);
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
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:'Formazione obbligatoria – D.Lgs. 81/08 e Accordo Stato-Regioni 17/04/2025',italic:true,font:FONT,size:18,color:C.GRIGIO})]}),
      makeDiscente(),
      vuoto(20),
      ...(!isDocente?[new Paragraph({spacing:{after:200},children:[new TextRun({text:'ISTRUZIONI: Per ogni domanda, barrare la risposta che si ritiene corretta (A, B, C oppure D). È ammessa una sola risposta per domanda. Durata: 30 minuti. Punteggio minimo per il superamento: 21/30 (70%).',italic:true,font:FONT,size:18})]})]:[]),
      ...domande.flatMap(({d,r}) => [makeQuestion(d,r,isDocente), vuoto(10)]),
      ...firme(),
    ];
  }

  const doc = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(false)}]});
  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_${mansione.id}.docx`);
  const docD = new Document({styles:docStyles,sections:[{properties:{page:{size:{width:11906,height:16838},margin:MARGIN}},headers:{default:header},footers:{default:footer},children:buildChildren(true)}]});
  await salvaDoc(docD, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_${mansione.id}_Docente.docx`);
}

module.exports = { genTestGenerale, genTestMansione };
