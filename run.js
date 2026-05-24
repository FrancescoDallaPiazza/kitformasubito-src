'use strict';
const { MANSIONI, CLIENTE, MODALITA } = require('./helpers');
const { genProgettoFormativo, genRegistroFormIniziale, genRegistroAggiornamento } = require('./docs1');
const { genColloquio, genGradimento, genAttestato, genAttestatiAggiornamento, genVerbaleVerifica, genVerificaEfficacia } = require('./docs2');
const { genTestGenerale, genTestMansione } = require('./gen_test');
const { genSchedaMansione, genSchedaAddestrativa } = require('./gen_scheda');
const { execSync } = require('child_process');

const NOME_BREVE = CLIENTE.ragioneSocialeBreve;
// Nome della cartella radice del kit: cambia in base a MODALITA.
const KIT_FOLDER_PREFIX = MODALITA === 'aggiornamento'
  ? 'KIT FORMASUBITO AGGIORNAMENTO'
  : 'KIT FORMASUBITO';
const OUT = `/home/claude/kit/OUT/${KIT_FOLDER_PREFIX} - ${NOME_BREVE}`;

// Helper: sceglie la mansione a rischio più alto (o la prima a parità di livello)
// per la Scheda Addestrativa, che viene generata per UNA sola mansione.
function mansionePiuRischiosa() {
  const score = { BASSO: 1, MEDIO: 2, ALTO: 3 };
  return MANSIONI.reduce((a, b) => (score[b.livello] || 0) >= (score[a.livello] || 0) ? b : a);
}

async function mainIniziale() {
  console.log('═══════════════════════════════════════════════════');
  console.log(`  KIT FORMASUBITO – FORMAZIONE INIZIALE`);
  console.log(`  ${CLIENTE.ragioneSociale}`);
  console.log('═══════════════════════════════════════════════════');

  // 00 – Progetto Formativo (versione iniziale, senza blocco 4.3)
  console.log('\n📋 [00] Progetto Formativo...');
  await genProgettoFormativo();

  // 01 – Schede Mansioni
  console.log('\n📊 [01] Schede Mansioni...');
  for (const m of MANSIONI) {
    await genSchedaMansione(m);
  }

  // 02 – Registro Presenze (solo iniziale)
  console.log('\n📝 [02] Registri Presenze...');
  for (const m of MANSIONI) {
    await genRegistroFormIniziale(m);
  }

  // 03 – Test di Apprendimento (Generale + Specifici per mansione)
  console.log('\n📚 [03] Test di Apprendimento...');
  await genTestGenerale();
  for (const m of MANSIONI) {
    await genTestMansione(m);
  }

  // 04 – Gradimento
  console.log('\n😊 [04] Scheda Gradimento...');
  await genGradimento();

  // 05 – Attestati (uno per mansione)
  console.log('\n🏆 [05] Attestati...');
  for (const m of MANSIONI) {
    await genAttestato(m);
  }

  // 06 – Verbale Verifica
  console.log('\n✅ [06] Verbale Verifica Finale...');
  await genVerbaleVerifica();

  // 07 – Verifica Efficacia
  console.log('\n🔍 [07] Verifica Efficacia...');
  await genVerificaEfficacia();

  // BONUS – Scheda Addestrativa (solo per la mansione a rischio più alto)
  console.log('\n🛠️  [BONUS] Scheda Addestrativa...');
  await genSchedaAddestrativa(mansionePiuRischiosa());
}

async function mainAggiornamento() {
  console.log('═══════════════════════════════════════════════════');
  console.log(`  KIT FORMASUBITO – AGGIORNAMENTO QUINQUENNALE`);
  console.log(`  ${CLIENTE.ragioneSociale}`);
  console.log('═══════════════════════════════════════════════════');

  // 00 – Progetto Formativo (versione aggiornamento, solo blocco 4.3)
  console.log('\n📋 [00] Progetto Formativo (aggiornamento)...');
  await genProgettoFormativo();

  // 02 – Registro Aggiornamento (unico)
  console.log('\n📝 [02] Registro Aggiornamento...');
  await genRegistroAggiornamento();

  // 03 – Colloqui di Apprendimento (uno per mansione)
  console.log('\n💬 [03] Colloqui di Apprendimento...');
  for (const m of MANSIONI) {
    await genColloquio(m);
  }

  // 04 – Gradimento
  console.log('\n😊 [04] Scheda Gradimento...');
  await genGradimento();

  // 05 – Attestato Aggiornamento (unico, generico per mansione da compilare)
  console.log('\n🏆 [05] Attestato Aggiornamento...');
  await genAttestatiAggiornamento();

  // 06 – Verbale Verifica
  console.log('\n✅ [06] Verbale Verifica Finale...');
  await genVerbaleVerifica();

  // 07 – Verifica Efficacia
  console.log('\n🔍 [07] Verifica Efficacia...');
  await genVerificaEfficacia();
}

async function main() {
  // ─── PRE-FLIGHT CHECK: 30 domande minime per mansione (SKILL §3.5) ────
  // Conta auto + extra per ciascuna mansione e fallisce se qualcuna è < 30.
  // Questo intercetta lo skip dello STEP 3.5b della SKILL (definizione quizExtra
  // in helpers.js): senza quizExtra, i test specifici hanno solo 2×rischi+3
  // domande, ben sotto il minimo normativo previsto dall'ASR 17/04/2025.
  if (MODALITA === 'iniziale' || MODALITA === undefined) {
    const errori = [];
    for (const m of MANSIONI) {
      const auto = 2 * m.rischi.length + 3;
      const extra = (m.quizExtra || []).length;
      const totale = auto + extra;
      if (totale < 30) {
        errori.push(`  ✗ ${m.id.padEnd(15)} ${m.rischi.length}r → ${auto}auto + ${extra}extra = ${totale} (servono ≥ 30)`);
      }
    }
    if (errori.length > 0) {
      console.error('\n⛔ PRE-FLIGHT CHECK FALLITO — Test specifici sotto le 30 domande minime (SKILL §3.5):');
      errori.forEach(e => console.error(e));
      console.error('\nAzione: torna allo STEP 3.5b della SKILL kitformasubito e popola il campo quizExtra');
      console.error('per le mansioni indicate. Vedi /mnt/skills/user/kitformasubito/SKILL.md righe 411-580.\n');
      process.exit(1);
    }
    console.log(`✓ Pre-flight check: ${MANSIONI.length} mansioni raggiungono ≥ 30 domande`);
  }

  // Pulisco OUT da eventuali residui di run precedenti per evitare di mischiare
  // documenti di iniziale e aggiornamento.
  try { execSync(`rm -rf "${OUT}"`); } catch (_) {}

  if (MODALITA === 'aggiornamento') {
    await mainAggiornamento();
  } else if (MODALITA === 'iniziale' || MODALITA === undefined) {
    await mainIniziale();
  } else {
    throw new Error(`MODALITA non riconosciuta: '${MODALITA}'. Valori ammessi: 'iniziale' | 'aggiornamento'.`);
  }

  // ZIP
  console.log('\n📦 Creazione ZIP...');
  const nomeZip = NOME_BREVE.replace(/[^a-zA-Z0-9]/g, '_');
  const zipPrefix = MODALITA === 'aggiornamento' ? 'KIT_FORMASUBITO_AGGIORNAMENTO' : 'KIT_FORMASUBITO';
  const zipPath = `/mnt/user-data/outputs/${zipPrefix}_${nomeZip}.zip`;
  // Rimuovo eventuale ZIP precedente: 'zip -r' fa append agli archivi esistenti
  // mantenendo i vecchi entry — comportamento indesiderato se la struttura cartelle
  // del kit è cambiata tra una run e l'altra.
  try { execSync(`rm -f "${zipPath}"`); } catch (_) {}
  execSync(`cd /home/claude/kit/OUT && zip -r "${zipPath}" "${KIT_FOLDER_PREFIX} - ${NOME_BREVE}" -q`);

  console.log('\n═══════════════════════════════════════════════════');
  console.log('  ✅ KIT FORMASUBITO COMPLETATO!');
  console.log('═══════════════════════════════════════════════════');

  // Riepilogo
  const count = execSync(`find "${OUT}" -name "*.docx" | wc -l`).toString().trim();
  const size = execSync(`du -sh "${OUT}" 2>/dev/null | cut -f1`).toString().trim();
  console.log(`\n📁 Documenti generati: ${count} file .docx`);
  console.log(`💾 Dimensione totale: ${size}`);
  console.log(`📂 Cartella: ${KIT_FOLDER_PREFIX} - ${NOME_BREVE}`);
}

main().catch(e => { console.error('ERRORE:', e); process.exit(1); });
