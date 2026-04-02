'use strict';
const h = require('./helpers');
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
  PageBreak,
  C, FONT, CLIENTE, MANSIONI, docStyles, A4_P, MARGIN_STD,
  makeHeader, makeFooter, titoloSezione, corpo, vuoto, cella, salvaDoc,
} = h;

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

// Domande Formazione Generale (30 domande fisse, distribuzione A/B/C/D bilanciata)
const DOMANDE_GENERALI = [
  { d:'Qual è il principale decreto legislativo italiano in materia di salute e sicurezza sul lavoro?', r:['D.Lgs. 81/2008','D.Lgs. 106/2009','D.Lgs. 151/2001','D.P.R. 303/1956'], c:0 },
  { d:'Chi ha l\'obbligo di valutare i rischi presenti in azienda?', r:['Il medico competente','Il Responsabile SPP','Il Datore di lavoro','Il Rappresentante dei lavoratori per la sicurezza'], c:2 },
  { d:'Cosa si intende per "rischio" in ambito sicurezza sul lavoro?', r:['La sola probabilità che si verifichi un evento dannoso','La gravità del danno che può derivare da un pericolo','La probabilità che si raggiunga il livello potenziale di danno nelle condizioni di impiego','Il numero di infortuni registrati in un anno'], c:2 },
  { d:'Quale colore identifica la segnaletica di divieto?', r:['Blu e bianco','Giallo e nero','Verde e bianco','Rosso e bianco'], c:3 },
  { d:'Cosa si intende per DPI?', r:['Documento di Prevenzione Interna','Dispositivo di Protezione Individuale','Direttiva di Protezione degli Impianti','Disposizione di Prevenzione e Intervento'], c:1 },
  { d:'Con quale periodicità deve essere effettuato l\'aggiornamento della formazione dei lavoratori (ASR 17/04/2025)?', r:['Ogni anno','Ogni tre anni','Ogni cinque anni','Ogni dieci anni'], c:2 },
  { d:'Quale figura aziendale coordina il servizio di prevenzione e protezione?', r:['Il medico competente','Il datore di lavoro','Il preposto','Il Responsabile del Servizio di Prevenzione e Protezione (RSPP)'], c:3 },
  { d:'Cosa si intende per "mancato infortunio" (near miss)?', r:['Un infortunio lieve senza giorni di assenza','Un evento che avrebbe potuto causare un infortunio ma non ha avuto conseguenze','Un infortunio occorso al di fuori dell\'orario di lavoro','Un infortunio non denunciato all\'INAIL'], c:1 },
  { d:'Quale segnaletica indica una via di fuga o di emergenza?', r:['Segnaletica di pericolo (giallo/nero)','Segnaletica di divieto (rosso/bianco)','Segnaletica di salvataggio (verde/bianco)','Segnaletica antincendio (rosso/bianco)'], c:2 },
  { d:'La sorveglianza sanitaria è affidata a:', r:['Il datore di lavoro','Il RSPP','Il medico competente','Il RLS'], c:2 },
  { d:'Cosa deve fare un lavoratore che rileva un pericolo imminente?', r:['Continuare il lavoro e segnalarlo alla fine del turno','Segnalarlo immediatamente al datore di lavoro o al preposto e, se necessario, allontanarsi','Rimuovere il pericolo autonomamente senza segnalazione','Attendere l\'ispezione periodica del RSPP'], c:1 },
  { d:'Quante ore di formazione generale prevede l\'ASR 17/04/2025?', r:['2 ore','4 ore','6 ore','8 ore'], c:1 },
  { d:'Il RLS (Rappresentante dei Lavoratori per la Sicurezza) è:', r:['Nominato direttamente dal datore di lavoro','Eletto o designato dai lavoratori','Nominato dal medico competente','Scelto dal servizio di prevenzione e protezione'], c:1 },
  { d:'Quale segnale di sicurezza indica l\'obbligo di indossare il casco?', r:['Segnaletica di divieto (rosso/bianco)','Segnaletica di avvertimento (giallo/nero)','Segnaletica di prescrizione (blu/bianco)','Segnaletica di salvataggio (verde/bianco)'], c:2 },
  { d:'Lo stress lavoro-correlato può causare:', r:['Solo disturbi fisici temporanei','Solo conflitti interpersonali irrilevanti','Effetti negativi sulla salute fisica e mentale del lavoratore','Esclusivamente cali di produttività senza rischi per la salute'], c:2 },
  { d:'Il Documento di Valutazione dei Rischi (DVR) deve essere:', r:['Redatto dal medico competente e conservato in archivio storico','Elaborato dal datore di lavoro e aggiornato periodicamente','Compilato annualmente dal RLS e approvato dal prefetto','Prodotto esclusivamente da un consulente esterno certificato'], c:1 },
  { d:'Quale delle seguenti è una misura di prevenzione primaria?', r:['Fornitura di DPI ai lavoratori','Eliminazione o riduzione del rischio alla fonte','Sorveglianza sanitaria periodica','Formazione sull\'uso dei DPI'], c:1 },
  { d:'In caso di emergenza e necessità di evacuazione, il lavoratore deve:', r:['Tornare alla propria postazione a prendere gli oggetti personali','Seguire le indicazioni delle vie di esodo e raggiungere il punto di raccolta','Attendere l\'arrivo dei soccorsi senza muoversi','Telefonare ai colleghi per coordinare l\'uscita autonomamente'], c:1 },
  { d:'Quale numero telefonico è dedicato al soccorso sanitario in Italia?', r:['112','113','115','118'], c:3 },
  { d:'La gerarchia delle misure di prevenzione prevede come priorità assoluta:', r:['La fornitura di DPI ai lavoratori','La formazione e l\'informazione','L\'eliminazione del rischio','La sorveglianza sanitaria'], c:2 },
  { d:'Quale tra i seguenti è un fattore di rischio per lo stress lavoro-correlato?', r:['Orario di lavoro regolare e prevedibile','Autonomia decisionale elevata','Eccesso di carico di lavoro e pressioni temporali costanti','Rapporti interpersonali positivi con i colleghi'], c:2 },
  { d:'Cosa significa la segnaletica triangolare con sfondo giallo?', r:['Obbligo','Divieto','Avvertimento di pericolo','Indicazione di salvataggio'], c:2 },
  { d:'Il preposto ha il compito di:', r:['Redigere il DVR insieme al RSPP','Effettuare la sorveglianza sanitaria dei lavoratori','Sovrintendere alle attività lavorative e garantire le direttive di sicurezza ricevute','Nominare i componenti del servizio di prevenzione e protezione'], c:2 },
  { d:'I DPI devono essere:', r:['Acquistati direttamente dal lavoratore a proprie spese','Forniti gratuitamente dal datore di lavoro e usati correttamente dal lavoratore','Utilizzati solo se il lavoratore li ritiene necessari','Sostituiti solo in caso di infortunio accertato'], c:1 },
  { d:'La segnalazione di un "near miss" da parte del lavoratore è:', r:['Facoltativa e senza particolari conseguenze procedurali','Obbligatoria solo per i preposti e non per i lavoratori','Un atto utile che contribuisce a prevenire futuri infortuni','Riservata esclusivamente alle figure di sistema (RSPP, MC, RLS)'], c:2 },
  { d:'Quante ore di formazione specifica sono previste per i lavoratori a rischio MEDIO (ASR 17/04/2025)?', r:['4 ore','6 ore','8 ore','12 ore'], c:2 },
  { d:'Il medico competente esegue la sorveglianza sanitaria per:', r:['Sostituire il RSPP nelle funzioni ispettive','Valutare l\'idoneità psico-fisica del lavoratore alla mansione specifica','Redigere il DVR in sostituzione del datore di lavoro','Controllare il rispetto delle misure di sicurezza in azienda'], c:1 },
  { d:'In quale caso un lavoratore può rifiutarsi di svolgere la propria mansione?', r:['Quando ritiene il lavoro troppo faticoso','In caso di pericolo grave e imminente per la propria salute o sicurezza','Quando non ha ricevuto la formazione generale','Solo con autorizzazione del RLS'], c:1 },
  { d:'Cosa si intende per "valutazione del rischio"?', r:['L\'elenco degli infortuni occorsi nell\'anno precedente','La stima quantitativa dei costi della sicurezza','La valutazione globale e documentata di tutti i rischi per la salute e sicurezza dei lavoratori','Il programma annuale di formazione e addestramento'], c:3 },
  { d:'Con quale periodicità il medico competente visita i luoghi di lavoro?', r:['Una volta al mese','Almeno una volta all\'anno','Ogni tre anni','Solo in caso di infortuni gravi'], c:1 },
];

// Domande specifiche per mansione
const DOMANDE_SPECIFICHE = {
  ImpiegataAmm: [
    { d:'Qual è la pausa visiva raccomandata durante l\'utilizzo prolungato del videoterminale?', r:['5 minuti ogni 30 minuti','15 minuti ogni 2 ore di lavoro continuativo','30 minuti ogni ora','1 ora ogni 4 ore'], c:1 },
    { d:'Per ridurre il rischio di posture incongrue alla postazione VDT è necessario:', r:['Tenere il monitor a meno di 30 cm dagli occhi','Posizionare lo schermo in modo da avere la luce diretta sul monitor','Usare una sedia ergonomica con schienale regolabile e appoggio lombare','Tenere le braccia sollevate durante la digitazione'], c:2 },
    { d:'In caso di pavimento bagnato in smart working, la misura più efficace è:', r:['Ignorare il rischio perché si è a casa propria','Segnalare l\'area con appositi segnali e asciugare immediatamente','Continuare a lavorare con particolare cautela','Interrompere la giornata lavorativa'], c:1 },
    { d:'Durante la sostituzione del toner della stampante è necessario:', r:['Non fare nulla di particolare, il rischio è trascurabile','Leggere attentamente le istruzioni del costruttore e indossare i DPI previsti','Smaltire immediatamente il vecchio toner nel cestino ordinario','Chiamare sempre un tecnico esterno anche per operazioni di routine'], c:1 },
    { d:'Lo stress lavoro-correlato nel lavoro d\'ufficio può essere ridotto mediante:', r:['Aumentare il carico di lavoro per abituarsi alla pressione','Organizzare i compiti con scadenze realistiche e pause regolari','Evitare qualsiasi comunicazione con colleghi e responsabili','Lavorare in isolamento per evitare conflitti interpersonali'], c:1 },
    { d:'Un\'illuminazione inadeguata alla postazione VDT può causare:', r:['Solo fastidio estetico senza rischi per la salute','Affaticamento visivo, cefalea e posture scorrette per compensare','Esclusivamente problemi alla vista a lungo termine irreversibili','Nessun problema poiché gli schermi moderni sono autoilluminanti'], c:1 },
    { d:'I cavi elettrici alla postazione di lavoro devono essere:', r:['Lasciati liberi sul pavimento per facilitare le connessioni','Protetti e ordinati per evitare rischi di inciampo e cortocircuiti','Collegati a prolunghe multiple per coprire più postazioni','Nascosti sotto i tappeti per non creare disordine visivo'], c:1 },
    { d:'In caso di incendio nell\'ambiente domestico di smart working, la prima azione è:', r:['Cercare di spegnere l\'incendio con qualsiasi mezzo disponibile','Allertare i soccorsi (115) e abbandonare l\'edificio seguendo le vie di fuga','Aprire tutte le finestre per ridurre il fumo','Raccogliere gli oggetti di valore prima di uscire'], c:1 },
    { d:'Gli arredi e scaffali della postazione home office devono essere:', r:['Posizionati in qualsiasi modo senza vincoli particolari','Ancorati alla parete e organizzati in modo da non causare ribaltamenti','Riempiti al massimo per ottimizzare lo spazio disponibile','Collocati solo sul lato sinistro della scrivania'], c:1 },
    { d:'Per accedere a ripiani alti in home office è corretto utilizzare:', r:['Una sedia girevole come sostituto improvvisato','Una scaletta idonea e stabile, evitando mezzi di fortuna','Una sedia fissa spingendola fino al ripiano','Il bordo della scrivania come supporto temporaneo'], c:1 },
    { d:'Quali sono le principali misure per prevenire lo stress da VDT?', r:['Lavorare molte ore consecutive senza interruzioni per finire prima','Pause regolari, illuminazione adeguata, postazione ergonomica e gestione dei carichi di lavoro','Aumentare la luminosità dello schermo al massimo','Utilizzare sempre monitor di grande dimensione indipendentemente dalla distanza'], c:1 },
    { d:'Il rischio di scivolamento in ambiente domestico è considerato:', r:['Assente in quanto l\'ambiente domestico è sicuro per definizione','Rilevante, soprattutto su superfici bagnate o in presenza di tappeti mal posizionati','Presente solo in caso di pavimenti in marmo o pietra','Trascurabile se il lavoratore indossa calzature chiuse'], c:1 },
    { d:'L\'impiego di videoterminali può causare disturbi muscolo-scheletrici a causa di:', r:['L\'intensità della luce emessa dal monitor','Posture fisse e mantenute per tempi prolungati senza pause attive','L\'emissione di radiazioni ionizzanti degli schermi LCD','La differenza di temperatura tra schermo e ambiente'], c:1 },
    { d:'La segnaletica di sicurezza di colore verde e bianco indica:', r:['Un divieto da rispettare tassativamente','Un pericolo imminente da evitare','Una via di fuga, uscita di emergenza o punto di soccorso','Un\'informazione normativa obbligatoria'], c:2 },
    { d:'In caso di emergenza in smart working, il lavoratore deve:', r:['Terminare il lavoro in corso prima di allertare i soccorsi','Contattare il datore di lavoro prima di qualsiasi altra azione','Abbandonare immediatamente e chiamare i soccorsi (118/112/115)','Attendere istruzioni via email dall\'azienda'], c:2 },
    { d:'Cosa si intende per "astenopia" nel lavoro con VDT?', r:['Una patologia oculare irreversibile causata dai monitor','Affaticamento visivo con bruciore, lacrimazione e visione annebbiata, che regredisce con il riposo','Una forma di stress cronico legata al lavoro notturno','Un disturbo muscolare delle spalle dovuto alla postura'], c:1 },
    { d:'La valutazione dei rischi per i lavoratori in smart working è:', r:['Non necessaria perché il lavoro è svolto in autonomia fuori sede','Obbligatoria anche per le postazioni domestiche ai sensi del D.Lgs. 81/2008','Facoltativa e affidata al lavoratore stesso','Applicabile solo se il lavoratore usa attrezzature aziendali'], c:1 },
    { d:'Per prevenire il sovraccarico muscolo-scheletrico durante il lavoro d\'ufficio è utile:', r:['Mantenere sempre la stessa posizione per favorire la concentrazione','Alternare la posizione seduta con brevi passeggiate e esercizi di stretching','Evitare qualsiasi movimento per ridurre la fatica','Usare esclusivamente mobili di design indipendentemente dall\'ergonomia'], c:1 },
    { d:'La distanza visiva ottimale dal monitor durante il lavoro al VDT è:', r:['Meno di 30 cm per vedere meglio i dettagli','Tra 50 e 70 cm, con il bordo superiore del monitor all\'altezza degli occhi','Oltre 1 metro per ridurre l\'esposizione alle radiazioni','Dipende esclusivamente dalle preferenze del lavoratore'], c:1 },
    { d:'Un lavoratore in smart working deve segnalare al datore di lavoro:', r:['Solo gli infortuni che richiedono l\'intervento del medico','Qualsiasi condizione di pericolo, rischio o near miss riscontrata nell\'ambiente domestico','Esclusivamente i guasti alle attrezzature aziendali','Nessuna situazione, avendo piena autonomia nell\'ambiente domestico'], c:1 },
    { d:'Il corretto smaltimento dei toner esauriti delle stampanti prevede:', r:['Il conferimento nel cestino ordinario dei rifiuti','Il rispetto delle indicazioni del produttore e la raccolta differenziata come rifiuto speciale','Il lavaggio con acqua e sapone e il riutilizzo','La conservazione in attesa di apposite istruzioni aziendali'], c:1 },
    { d:'Per ridurre i riflessi fastidiosi sullo schermo del computer è consigliabile:', r:['Aumentare la luminosità del monitor al massimo','Posizionare lo schermo perpendicolare alle finestre e usare tende o veneziane','Lavorare in ambienti completamente bui','Usare pellicole oscuranti su tutte le finestre della stanza'], c:1 },
    { d:'In smart working, il lavoratore ha l\'obbligo di:', r:['Gestire autonomamente tutti i rischi senza coinvolgere l\'azienda','Rispettare le istruzioni del datore di lavoro in materia di sicurezza anche nell\'ambiente domestico','Rinunciare a qualsiasi protezione essendo fuori dalla sede aziendale','Autodichiarare l\'idoneità della postazione senza controlli esterni'], c:1 },
    { d:'Il rischio incendio in home office è correlato principalmente a:', r:['La vicinanza agli uffici aziendali','Sovraccarico degli impianti elettrici, accumulo di materiale cartaceo e malfunzionamenti','L\'uso di monitor di grandi dimensioni ad alta potenza','La presenza di finestre non sigillate'], c:1 },
    { d:'Cosa deve fare il lavoratore se rileva un near miss durante lo smart working?', r:['Ignorarlo in quanto evento non verificatosi in sede aziendale','Annotarlo e segnalarlo al datore di lavoro per attivare le misure preventive','Gestirlo autonomamente senza coinvolgere l\'azienda','Aspettare il successivo incontro periodico di team'], c:1 },
    { d:'La pausa di 15 minuti ogni 2 ore al VDT è:', r:['Una raccomandazione facoltativa del medico competente','Un obbligo normativo previsto dal D.Lgs. 81/2008 per i lavoratori VDT','Una prassi aziendale non obbligatoria','Applicabile solo per turni lavorativi superiori alle 8 ore'], c:1 },
    { d:'Il rumore negli ambienti d\'ufficio:', r:['Non costituisce mai un rischio per la salute','Può aumentare lo stress e ridurre la concentrazione, pur raramente superando i limiti d\'azione','È considerato rischio elevato solo nelle fabbriche','Riguarda esclusivamente il personale che usa macchinari rumorosi'], c:1 },
    { d:'Il rispetto dell\'ergonomia alla postazione di lavoro ha lo scopo di:', r:['Migliorare solo l\'estetica e il comfort percepito','Prevenire disturbi muscolo-scheletrici, affaticamento visivo e posture scorrette','Aumentare la produttività senza considerare la salute del lavoratore','Ridurre i costi delle attrezzature aziendali'], c:1 },
    { d:'La segnaletica di sicurezza di colore blu e bianco indica:', r:['Un divieto da rispettare','Un pericolo imminente','Un obbligo (es. indossare i DPI)','Un\'uscita di emergenza'], c:2 },
    { d:'Il DVR deve essere aggiornato:', r:['Solo ogni 5 anni in modo sistematico','Ogni volta che intervengono modifiche significative nell\'organizzazione o nei processi','Solo se richiesto da un organo di vigilanza','Non è mai necessario aggiornarlo una volta redatto'], c:1 },
  ],
  Commerciale: [
    { d:'Durante la guida per ragioni di lavoro, il lavoratore deve assolutamente evitare:', r:['Rispettare i limiti di velocità','Verificare lo stato dei pneumatici prima della partenza','L\'uso del telefono cellulare non in viva voce','Fare pause di riposo durante i lunghi trasferimenti'], c:2 },
    { d:'Quale DPI è obbligatorio in caso di guasto o emergenza sulla strada?', r:['Il casco protettivo','Il gilet ad alta visibilità','Gli occhiali di sicurezza','I guanti antitaglio'], c:1 },
    { d:'In caso di forte stanchezza durante un lungo trasferimento in auto, il lavoratore deve:', r:['Aumentare la velocità per terminare prima il percorso','Aprire il finestrino e continuare la guida','Fermarsi in un\'area di sosta sicura e riposare prima di proseguire','Chiamare il cliente per avvisare del ritardo e continuare'], c:2 },
    { d:'Il rischio di stress lavoro-correlato per il commerciale è classificato come:', r:['Basso, perché il lavoro è prevalentemente all\'aperto','Alto, a causa delle pressioni sugli obiettivi, dei tempi di spostamento e delle relazioni con i clienti','Medio, ma solo per i commerciali over 50','Assente, in quanto il commerciale gode di autonomia lavorativa'], c:1 },
    { d:'In caso di comportamento aggressivo di un cliente durante una visita, la risposta corretta è:', r:['Rispondere in modo assertivo con atteggiamento minaccioso','Cedere immediatamente a qualsiasi richiesta del cliente','Mantenere la calma, de-escalare la situazione e, se necessario, abbandonare il luogo in sicurezza','Ignorare il comportamento e proseguire la presentazione commerciale'], c:2 },
    { d:'La revisione biennale dell\'automezzo utilizzato per lavoro è:', r:['Facoltativa se il veicolo è nuovo','Obbligatoria per legge e deve essere documentata','Solo consigliata dalle case produttrici','Richiesta solo per veicoli aziendali formalmente intestati all\'azienda'], c:1 },
    { d:'Prima di un lungo trasferimento in auto per ragioni lavorative, il commerciale deve verificare:', r:['Solo il livello del carburante','Freni, pneumatici, livelli, luci e documenti di guida','Esclusivamente la disponibilità del GPS','Solo la validità dell\'assicurazione'], c:1 },
    { d:'Il rischio di incidente stradale per il lavoratore commerciale è considerato:', r:['Trascurabile perché avviene fuori dall\'azienda','Un rischio professionale rilevante che deve essere valutato e gestito dal DVR','Un rischio privato non coperto dalla normativa sul lavoro','Significativo solo in caso di guida notturna'], c:1 },
    { d:'Quale delle seguenti è una misura di prevenzione efficace per il rischio stradale lavorativo?', r:['Guidare a velocità elevata per ridurre il tempo di esposizione al rischio','Effettuare manutenzione programmata del veicolo e rispettare le norme del codice della strada','Effettuare trasferimenti solo in condizioni meteorologiche ottimali','Non utilizzare mai il navigatore per evitare distrazioni'], c:1 },
    { d:'La comunicazione assertiva con i clienti serve principalmente a:', r:['Ottenere sempre il massimo vantaggio commerciale','Gestire i conflitti in modo costruttivo riducendo il rischio di aggressioni verbali o fisiche','Evitare qualsiasi forma di comunicazione diretta','Imporre i propri punti di vista al cliente'], c:1 },
    { d:'Cosa deve fare il commerciale che lavora da remoto (casa/auto) in caso di guasto al PC aziendale?', r:['Smettere immediatamente di lavorare e aspettare riparazioni','Segnalare il guasto all\'azienda e usare attrezzature alternative sicure fornite','Acquistare autonomamente un nuovo dispositivo a proprie spese','Continuare a lavorare con device personali senza informare l\'azienda'], c:1 },
    { d:'Lo stress da obiettivi commerciali può essere gestito mediante:', r:['Ignorare le pressioni e lavorare in modo instintivo','Pianificazione delle attività, comunicazione aperta con il responsabile e gestione del tempo','Aumentare il numero di visite giornaliere per raggiungere più velocemente i target','Evitare qualsiasi confronto con il responsabile commerciale'], c:1 },
    { d:'In quale condizione atmosferica il commerciale NON dovrebbe effettuare spostamenti in auto?', r:['Quando piove leggermente','Con nebbia fitta, ghiaccio o neve abbondante che compromettono la sicurezza stradale','Durante le ore diurne con sole','Con vento moderato'], c:1 },
    { d:'Quale segnaletica di sicurezza deve conoscere il commerciale che visita cantieri o siti produttivi?', r:['Solo la segnaletica stradale','La segnaletica di sicurezza prevista dal D.Lgs. 81/2008 (divieto, pericolo, obbligo, salvataggio)','La segnaletica commerciale dei clienti','Esclusivamente la segnaletica antincendio'], c:1 },
    { d:'Il divieto di assumere bevande alcoliche durante l\'orario di lavoro (inclusa la guida) è:', r:['Una raccomandazione facoltativa del medico competente','Un obbligo normativo specifico per chi guida per ragioni lavorative','Una politica aziendale non supportata da obblighi di legge','Applicabile solo per autisti professionisti con CQC'], c:1 },
    { d:'In caso di incidente stradale durante una trasferta lavorativa, il lavoratore deve:', r:['Risolvere autonomamente la situazione senza coinvolgere l\'azienda','Seguire le procedure di primo soccorso, allertare i soccorsi (118/112) e informare tempestivamente il datore di lavoro','Proseguire verso la destinazione se l\'incidente è lieve','Attendere il responsabile commerciale prima di allertare i soccorsi'], c:1 },
    { d:'Le posture incongrue durante la guida prolungata possono causare:', r:['Solo stanchezza temporanea senza conseguenze sulla salute','Disturbi muscolo-scheletrici al rachide cervicale e lombare, spalle e polsi','Esclusivamente problemi alla vista per la luce abbagliante','Nessun problema se il veicolo è dotato di sedile ergonomico'], c:1 },
    { d:'La pianificazione degli spostamenti lavorativi riduce il rischio stradale perché:', r:['Permette di guidare più velocemente conoscendo il percorso','Consente di evitare condizioni di stanchezza estrema, soste in luoghi non sicuri e percorsi ad alto rischio','Garantisce sempre l\'arrivo in anticipo all\'appuntamento','Elimina la necessità di verificare il meteo prima della partenza'], c:1 },
    { d:'Il DPI "gilet ad alta visibilità" serve a:', r:['Proteggere dalla pioggia durante le visite in cantiere','Rendere il lavoratore visibile agli altri utenti della strada in caso di emergenza o sosta','Identificare il lavoratore come rappresentante aziendale','Proteggere dagli agenti chimici presenti nei magazzini dei clienti'], c:1 },
    { d:'Cosa si intende per "rischio psicosociale" nel lavoro commerciale?', r:['Il rischio di perdita del portafoglio clienti','Il rischio derivante da fattori organizzativi e relazionali che possono influenzare la salute mentale del lavoratore','Il rischio legato all\'uso di social media durante l\'orario di lavoro','Il rischio di contrarre malattie durante le trasferte all\'estero'], c:1 },
    { d:'La pausa durante i trasferimenti lunghi in auto è raccomandata:', r:['Solo se il lavoratore si sente stanco in modo estremo','Ogni 2 ore di guida continuativa, come misura preventiva di sicurezza stradale','Solo per trasferimenti superiori alle 500 km','Esclusivamente nelle ore notturne'], c:1 },
    { d:'Il commerciale che visita cantieri o siti industriali dei clienti deve:', r:['Applicare le sole norme di sicurezza valide nella propria azienda','Rispettare le norme di sicurezza del cliente e indossare i DPI richiesti dal sito ospitante','Ignorare le disposizioni del sito cliente se non esplicitamente richiamate','Richiedere al cliente un\'esonero dalla normativa sulla sicurezza'], c:1 },
    { d:'La stanchezza da guida (fatigue) è pericolosa perché:', r:['Riduce solo la capacità di attenzione per brevi periodi recuperabili subito','Aumenta significativamente i tempi di reazione e la probabilità di incidenti stradali','Riguarda esclusivamente i conducenti di veicoli pesanti','Ha effetti solo dopo 10 ore continuative di guida'], c:1 },
    { d:'In caso di near miss durante una visita a un cliente (es. quasi scivolamento), il lavoratore deve:', r:['Non segnalarlo in quanto avvenuto nella struttura del cliente','Segnalarlo al proprio datore di lavoro come mancato infortunio','Farlo presente solo al cliente senza comunicarlo all\'azienda','Ignorarlo se non si sono verificate conseguenze fisiche'], c:1 },
    { d:'La gestione dello stress in ambito commerciale prevede come primo passo:', r:['Aumentare i carichi di lavoro per abituarsi alla pressione','Identificare e analizzare le fonti di stress per adottare misure organizzative e di supporto appropriate','Evitare qualsiasi comunicazione sullo stress con il proprio responsabile','Ricorrere immediatamente a supporto psicologico esterno'], c:1 },
    { d:'La segnalazione di un near miss da parte del commerciale contribuisce a:', r:['Creare documentazione inutile per l\'azienda','Prevenire futuri infortuni attraverso l\'analisi delle cause e l\'adozione di misure correttive','Dimostrare l\'inefficienza del sistema di prevenzione aziendale','Trasferire la responsabilità dell\'evento al datore di lavoro'], c:1 },
    { d:'Il rischio da posture incongrue nel lavoro del commerciale può essere ridotto mediante:', r:['Uso esclusivo di veicoli di alta gamma con sedili premium','Pause attive durante gli spostamenti, regolazione corretta del sedile e dell\'appoggiatesta','Limitazione degli spostamenti a non più di 1 km per visita','Utilizzo di cinture di sicurezza di tipo sportivo'], c:1 },
    { d:'La riunione periodica sulla sicurezza (art. 35 D.Lgs. 81/2008) è prevista per:', r:['Tutte le aziende con più di 15 lavoratori, con cadenza almeno annuale','Solo le aziende del settore industriale con oltre 50 dipendenti','Esclusivamente le aziende con cantieri mobili e temporanei','Qualsiasi azienda, indipendentemente dal numero di lavoratori, con cadenza mensile'], c:0 },
    { d:'Quale dei seguenti comportamenti riduce il rischio di aggressione durante una visita commerciale?', r:['Mostrare atteggiamento dominante e assertivo per imporre il rispetto','Recarsi sempre da soli nei contesti percepiti come potenzialmente ostili','Valutare preventivamente il contesto, evitare situazioni percepite come rischiose e comunicare la propria posizione all\'azienda','Continuare la visita anche in caso di comportamenti ostili del cliente'], c:2 },
    { d:'Il commerciale che utilizza il proprio veicolo personale per ragioni lavorative deve:', r:['Stipulare una polizza assicurativa specifica per uso promiscuo (privato + lavoro)','Usarlo senza modifiche all\'assicurazione poiché è un veicolo privato','Richiedere un rimborso chilometrico senza obblighi assicurativi aggiuntivi','Comunicarlo solo al responsabile commerciale, non all\'assicurazione'], c:0 },
  ],
  Manutentore: [
    { d:'Prima di salire su una scala portatile, il manutentore deve verificare:', r:['Solo che sia della misura giusta per raggiungere la quota','Pioli integri, piedini antiscivolo, apertura massima (per le doppie) e posizionamento su superficie stabile','Esclusivamente che la scala sia dello stesso colore dei DPI aziendali','Solo la portata massima in kg riportata sull\'etichetta'], c:1 },
    { d:'Durante il lavoro in quota su scala portatile, il manutentore non deve mai:', r:['Salire con gli strumenti in borsa o nella cintura porta-attrezzi','Portare con sé materiali pesanti da entrambe le mani perdendo i punti di appoggio','Verificare la stabilità della scala prima della salita','Indossare le calzature antinfortunistiche'], c:1 },
    { d:'Il DPI obbligatorio per lavori in quota con dislivello superiore a 2 metri è:', r:['Il casco di protezione del capo','L\'imbragatura anticaduta agganciata a un punto di ancoraggio certificato','I guanti antitaglio rinforzati','La mascherina FFP2'], c:1 },
    { d:'Prima di intervenire su un impianto elettrico, la prima operazione indispensabile è:', r:['Indossare guanti dielettrici e procedere con l\'intervento','Verificare l\'assenza di tensione con il multimetro dopo aver sezionato l\'impianto','Chiamare il cliente per conferma dei dati dell\'impianto','Attendere 30 minuti dalla disalimentazione prima di toccare i conduttori'], c:1 },
    { d:'Il rilevatore di perdite gas (analizzatore) deve essere utilizzato:', r:['Solo in caso di odore di gas percepito dal manutentore','Sistematicamente prima e durante interventi su impianti a gas come misura preventiva','Solo dal fornitore del gas e non dal manutentore','Esclusivamente su impianti GPL e non su quelli a metano'], c:1 },
    { d:'In caso di sospetta perdita di gas in un locale chiuso, il manutentore deve:', r:['Accendere l\'interruttore della luce per vedere meglio','Aerare immediatamente il locale, non creare scintille e allertare i soccorsi se necessario','Usare il telefono cellulare per segnalare l\'emergenza stando all\'interno del locale','Chiudere il locale per confinare la perdita ed evitare rischi agli altri'], c:1 },
    { d:'I guanti antitaglio sono obbligatori durante:', r:['Tutte le operazioni di manutenzione senza eccezioni','Le operazioni che comportano il contatto con parti taglienti, bordi acuminati o utensili da taglio','Solo durante l\'uso di saldatrici o flessibili','Esclusivamente quando si lavora su impianti in quota'], c:1 },
    { d:'Gli agenti chimici utilizzati nella manutenzione (es. sigillanti, refrigeranti) richiedono:', r:['Solo buona ventilazione del locale senza altri accorgimenti','Lettura della Scheda Dati di Sicurezza (SDS), uso dei DPI previsti e rispetto delle indicazioni di smaltimento','Esclusivamente l\'uso di guanti in lattice di qualsiasi tipo','Solo la verifica della data di scadenza del prodotto'], c:1 },
    { d:'Il rischio di "ambienti confinati" (es. locali tecnici, contatori gas) prevede come misura obbligatoria:', r:['Lavorare velocemente per ridurre il tempo di esposizione','La presenza di almeno un addetto esterno al locale durante tutta la durata del lavoro','Solo il monitoraggio della qualità dell\'aria con rilevatori portatili','L\'uso di autorespiratore in tutte le situazioni confinante'], c:1 },
    { d:'Prima di entrare in un locale tecnico sospetto di inquinamento atmosferico, il manutentore deve:', r:['Annusare l\'aria per valutare la presenza di gas','Verificare la qualità dell\'aria con idoneo strumento prima dell\'ingresso','Entrare rapidamente tenendo il respiro per i primi 30 secondi','Aprire solo una finestra e procedere immediatamente all\'interno'], c:1 },
    { d:'Il rischio elettrico negli impianti termosanitari è:', r:['Trascurabile perché gli impianti termici non sono collegati alla rete elettrica','Rilevante, in quanto caldaie, pompe di calore e impianti di climatizzazione sono apparecchiature elettriche','Presente solo negli impianti fotovoltaici e non in quelli termici','Assente se l\'impianto è certificato CE'], c:1 },
    { d:'Il metodo LOTO (Lockout/Tagout) serve per:', r:['Bloccare fisicamente i dispositivi di sezionamento impedendo la riaccensione accidentale durante la manutenzione','Etichettare le attrezzature per il riconoscimento durante gli inventari','Registrare le operazioni di manutenzione nel libro macchine','Verificare la taratura degli strumenti di misura prima dell\'uso'], c:0 },
    { d:'Le calzature antinfortunistiche (con puntale in acciaio) sono obbligatorie per il manutentore perché:', r:['Aumentano la stabilità durante la salita sulle scale','Proteggono dal rischio di caduta di oggetti pesanti sul piede e dalle scivolature','Sono richieste solo dai regolamenti condominiali','Sostituiscono completamente gli altri DPI negli interventi di routine'], c:1 },
    { d:'L\'esposizione al rumore durante l\'uso di utensili (trapano, flex) richiede:', r:['Nessun DPI specifico se la durata è inferiore a un\'ora','L\'uso di otoprotettori quando il livello supera 80 dB(A) di esposizione giornaliera','Solo la limitazione della durata dell\'esposizione senza protezioni individuali','La segnalazione del rischio ma non l\'uso sistematico di otoprotettori'], c:1 },
    { d:'Durante interventi su caldaie o bruciatori caldi, il manutentore deve:', r:['Procedere immediatamente per ridurre i tempi di intervento','Attendere il raffreddamento delle superfici calde e indossare guanti resistenti al calore','Usare solo guanti in lattice per il calore moderato','Lavorare rapidamente senza DPI specifici se l\'impianto è spento da meno di 30 minuti'], c:1 },
    { d:'Il refrigerante R32 (HFC) contenuto nelle pompe di calore deve essere gestito:', r:['Come un rifiuto ordinario, smaltito nel cestino','Secondo il Regolamento UE 517/2014 (F-Gas) da parte di personale certificato','Solo se le quantità superano i 3 kg','Liberandolo nell\'atmosfera durante le operazioni di manutenzione ordinaria'], c:1 },
    { d:'In caso di taglio accidentale durante un intervento di manutenzione, la prima azione è:', r:['Continuare il lavoro applicando un cerotto provvisorio','Fermare l\'emorragia con pressione diretta, disinfettare e, se necessario, recarsi al pronto soccorso','Aspettare che il cliente fornisca il kit di primo soccorso','Applicare ghiaccio direttamente sulla ferita senza medicazione'], c:1 },
    { d:'L\'occhiale di sicurezza è obbligatorio durante:', r:['Tutte le fasi del lavoro di manutenzione senza eccezioni','Operazioni che generano proiezione di schegge, liquidi o sostanze chimiche (foratura, taglio, saldatura)','Solo le operazioni di saldatura autogena','Esclusivamente quando si usano prodotti chimici spray'], c:1 },
    { d:'Il rischio di esposizione a vibrazioni per il manutentore è associato a:', r:['L\'uso di attrezzi manuali non vibranti (martello, cacciavite)','L\'uso prolungato di attrezzature vibranti (martelli pneumatici, trapani a percussione, flessibili)','Esclusivamente all\'uso di mezzi di trasporto pesanti','Solo alle lavorazioni in ambienti con temperatura elevata'], c:1 },
    { d:'Prima di lavorare su impianti a gas in un condominio o in un albergo, il manutentore deve:', r:['Iniziare subito il lavoro per ridurre i disagi ai residenti','Verificare la presenza di perdite con il rilevatore, aerare il locale e, se necessario, interrompere l\'erogazione del gas','Solo comunicare l\'intervento al portiere del condominio','Aspettare che il cliente firmi l\'autorizzazione prima di qualsiasi verifica'], c:1 },
    { d:'La Scheda Dati di Sicurezza (SDS) di un prodotto chimico deve essere:', r:['Letta solo dal datore di lavoro e non dal lavoratore','Disponibile e consultata prima dell\'uso del prodotto per conoscerne i rischi e le misure di prevenzione','Conservata esclusivamente in azienda e non portata in cantiere','Richiesta solo per prodotti con concentrazioni superiori al 10%'], c:1 },
    { d:'Il rischio di calore negli interventi su impianti termici è gestito mediante:', r:['L\'uso di indumenti ignifughi in tutte le operazioni di manutenzione','L\'attesa del raffreddamento, la segnalazione delle superfici calde e l\'uso dei guanti da calore','Solo la riduzione del tempo di esposizione senza DPI specifici','L\'applicazione di acqua fredda sulle superfici prima di toccarle'], c:1 },
    { d:'L\'elmetto di protezione è obbligatorio per il manutentore:', r:['Solo durante le attività svolte in quota su scale','Negli interventi all\'interno di cantieri edili e in ogni situazione con rischio di caduta di oggetti dall\'alto','Esclusivamente nelle demolizioni e nelle ristrutturazioni pesanti','Mai, in quanto la manutenzione termica non è considerata attività di cantiere'], c:1 },
    { d:'Durante l\'uso di sostanze chimiche (sigillanti, detergenti, flussanti), la mascherina FFP2 serve a:', r:['Proteggere esclusivamente dai virus durante il periodo post-pandemico','Filtrare polveri, aerosol e vapori nocivi che possono essere inalati durante l\'applicazione del prodotto','Sostituire la ventilazione del locale in tutti i casi','Essere indossata solo se il prodotto ha un odore sgradevole'], c:1 },
    { d:'Prima di intervenire in un locale tecnico angusto, è obbligatorio predisporre:', r:['Solo un\'estintore a portata di mano','Un piano di emergenza con addetto esterno di sorveglianza e dispositivi di comunicazione','Esclusivamente la documentazione dell\'impianto','Solo l\'illuminazione portatile adeguata'], c:1 },
    { d:'Il rischio di "caduta di oggetti dall\'alto" durante interventi su caldaie a parete richiede:', r:['Solo che il manutentore presti attenzione durante il lavoro','L\'uso di elmetto, calzature antinfortunistiche e organizzazione dell\'area di lavoro per evitare la presenza di persone sotto la zona di intervento','Esclusivamente la segnalazione dell\'area con nastro bianco e rosso','Solo la comunicazione verbale al cliente della presenza del rischio'], c:1 },
    { d:'Cosa significa che un\'attrezzatura è "Classe II" (doppio isolamento)?', r:['Ha due alimentatori distinti per sicurezza redundante','Non richiede collegamento a terra in quanto ha un doppio strato di isolamento che protegge dall\'elettrocuzione','Può essere usata in ambienti bagnati senza alcuna precauzione','Richiede una verifica elettrica ogni due anni invece che annuale'], c:1 },
    { d:'Il gilet ad alta visibilità è obbligatorio per il manutentore:', r:['Solo durante il lavoro notturno','In tutti i contesti stradali, nelle aree di parcheggio e nelle zone di cantiere dove circolano veicoli','Esclusivamente nelle visite ai cantieri edili','Mai, in quanto è un DPI specifico dei soli operai stradali'], c:1 },
    { d:'In quale situazione il manutentore può lavorare in quota senza imbragatura?', r:['Quando si lavora su una scala a pioli entro i 2 metri di altezza con assistente a terra','Mai al di sopra di 2 metri: l\'imbragatura è sempre obbligatoria con ancoraggio certificato','Quando il cliente garantisce verbalmente la sicurezza del luogo di lavoro','Quando si usa una scala doppia con gradini antiscivolo'], c:0 },
    { d:'Il piano di emergenza prima di un intervento in spazio confinato deve includere:', r:['Solo il numero di telefono del pronto soccorso più vicino','Procedure di evacuazione, comunicazione con l\'esterno, presenza dell\'addetto di sorveglianza e dispositivi di soccorso','Esclusivamente il permesso scritto del responsabile del sito cliente','Solo la verifica che l\'impianto sia spento e sezionato'], c:1 },
  ],
};

function genDocTest(mansione, domande, isDocente) {
  const titoloTipo = isDocente ? 'TEST DI APPRENDIMENTO – VERSIONE DOCENTE (con risposte)' : 'TEST DI APPRENDIMENTO';
  const header = makeHeader(CLIENTE.ragioneSocialeBreve, titoloTipo, CLIENTE.atecoCodice, CLIENTE.atecoDesc);
  const footer = makeFooter(`Test ${mansione.nome} – D.Lgs. 81/2008`);
  const lettere = ['A', 'B', 'C', 'D'];

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: titoloTipo.replace('– VERSIONE DOCENTE (con risposte)', ''), bold: true, font: FONT, size: 28, color: C.BLU_DARK })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [new TextRun({ text: `${mansione.nome} – Rischio ${mansione.livello}`, font: FONT, size: 22, color: C.BLU_HEADER })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'Soglia superamento: 18/30 (60%) – Una sola risposta corretta per domanda', font: FONT, size: 18, color: C.GRIGIO })],
    }),

    // Info candidato
    ...['Cognome e Nome: ___________________________________', `Mansione: ${mansione.nome}`, 'Data: ___/___/_____    Firma: ___________________________________']
      .map(t => new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: t, font: FONT, size: 20 })] })),
    vuoto(80),

    // Domande
    ...domande.map((dq, idx) => {
      const items = [];
      items.push(new Paragraph({
        keepLines: true, keepNext: true,
        spacing: { before: 0, after: 0 },
        children: [new TextRun({
          text: `${idx + 1}. ${dq.d}`,
          bold: true, font: FONT, size: 18,
          ...(isDocente ? {} : {}),
        })],
      }));
      dq.r.forEach((r, ri) => {
        const isCorretta = ri === dq.c;
        const prefix = isDocente && isCorretta ? '✓ ' : '   ';
        items.push(new Paragraph({
          keepLines: true, keepNext: ri < dq.r.length - 1,
          spacing: { before: 0, after: ri === dq.r.length - 1 ? 200 : 0 },
          indent: { left: 360 },
          children: [new TextRun({
            text: `${prefix}${lettere[ri]}) ${r}`,
            font: FONT, size: 18,
            bold: isDocente && isCorretta,
            color: isDocente && isCorretta ? C.VERDE : undefined,
          })],
        }));
      });
      return items;
    }).flat(),

    // Griglia risposte (solo versione studente)
    ...(!isDocente ? [
      new Paragraph({ children: [new PageBreak()] }),
      titoloSezione('GRIGLIA RISPOSTE'),
      corpo('Barrare la lettera corrispondente alla risposta scelta:', { before: 60 }),
      vuoto(40),
      new Table({
        width: { size: 9638, type: WidthType.DXA },
        columnWidths: Array(6).fill(1606),
        borders: {
          top:{style:BorderStyle.SINGLE,size:1,color:'999999'},bottom:{style:BorderStyle.SINGLE,size:1,color:'999999'},
          left:{style:BorderStyle.SINGLE,size:1,color:'999999'},right:{style:BorderStyle.SINGLE,size:1,color:'999999'},
          insideH:{style:BorderStyle.SINGLE,size:1,color:'999999'},insideV:{style:BorderStyle.SINGLE,size:1,color:'999999'},
        },
        rows: [
          new TableRow({ tableHeader: true, children: ['N°', 'A', 'B', 'C', 'D', 'RISPOSTA'].map((h2, i) =>
            cella(h2, { width: 1606, bold: true, fill: C.BLU_LIGHT, align: 'center' })
          )}),
          ...Array.from({length:30}, (_,i) => new TableRow({
            children: [
              cella(`${i+1}`, { width:1606, align:'center', fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
              cella('☐', { width:1606, align:'center', fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
              cella('☐', { width:1606, align:'center', fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
              cella('☐', { width:1606, align:'center', fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
              cella('☐', { width:1606, align:'center', fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
              cella('', { width:1606, fill: i%2===0?C.BIANCO:C.GRIGIO_ALT }),
            ],
          })),
        ],
      }),
    ] : []),
  ];

  return new Document({
    styles: docStyles,
    sections: [{
      properties: { page: { size: A4_P, margin: MARGIN_STD } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });
}

async function genTestGenerale() {
  const mFake = { nome: 'Formazione Generale', livello: '', id: 'Generale' };
  const doc = genDocTest(mFake, DOMANDE_GENERALI, false);
  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_Generale.docx`);
  const docD = genDocTest(mFake, DOMANDE_GENERALI, true);
  await salvaDoc(docD, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_Generale_Docente.docx`);
}

async function genTestMansione(mansione) {
  const domande = DOMANDE_SPECIFICHE[mansione.id];
  if (!domande) { console.log(`Nessuna domanda specifica per ${mansione.id}`); return; }
  const doc = genDocTest(mansione, domande, false);
  await salvaDoc(doc, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_${mansione.id}.docx`);
  const docD = genDocTest(mansione, domande, true);
  await salvaDoc(docD, `${OUT}/03 - TEST DI APPRENDIMENTO/00. GENERALE E SPECIFICA/Test_${mansione.id}_Docente.docx`);
}

module.exports = { genTestGenerale, genTestMansione };
