'use strict';
const { MANSIONI } = require('./helpers');
const { genProgettoFormativo, genRegistroFormIniziale, genRegistroAggiornamento } = require('./docs1');
const { genColloquio, genGradimento, genAttestato, genAttestatiAggiornamento, genVerbaleVerifica, genVerificaEfficacia } = require('./docs2');
const { genTestGenerale, genTestMansione } = require('./gen_test');
const { genSchedaMansione, genSchedaAddestrativa } = require('./gen_scheda');
const { execSync } = require('child_process');
const path = require('path');

const OUT = '/home/claude/kit/OUT/KIT FORMASUBITO - Calor Energy Verona';

async function main() {
  console.log('═══════════════════════════════════════════════════');
  console.log('  KIT FORMASUBITO – Calor Energy Verona Soc. Coop.');
  console.log('═══════════════════════════════════════════════════');

  // 00 – Progetto Formativo
  console.log('\n📋 [00] Progetto Formativo...');
  await genProgettoFormativo();

  // 01 – Schede Mansioni
  console.log('\n📊 [01] Schede Mansioni...');
  for (const m of MANSIONI) {
    await genSchedaMansione(m);
  }

  // 02 – Registro Presenze
  console.log('\n📝 [02] Registri Presenze...');
  for (const m of MANSIONI) {
    await genRegistroFormIniziale(m);
  }
  await genRegistroAggiornamento();

  // 03 – Test di Apprendimento
  console.log('\n📚 [03] Test di Apprendimento...');
  await genTestGenerale();
  for (const m of MANSIONI) {
    await genTestMansione(m);
  }
  for (const m of MANSIONI) {
    await genColloquio(m);
  }

  // 04 – Gradimento
  console.log('\n😊 [04] Scheda Gradimento...');
  await genGradimento();

  // 05 – Attestati
  console.log('\n🏆 [05] Attestati...');
  for (const m of MANSIONI) {
    await genAttestato(m);
  }
  await genAttestatiAggiornamento();

  // 06 – Verbale Verifica
  console.log('\n✅ [06] Verbale Verifica Finale...');
  await genVerbaleVerifica();

  // 07 – Verifica Efficacia
  console.log('\n🔍 [07] Verifica Efficacia...');
  await genVerificaEfficacia();

  // BONUS – Schede Addestrative
  console.log('\n🛠️  [BONUS] Schede Addestrative...');
  for (const m of MANSIONI) {
    await genSchedaAddestrativa(m);
  }

  // ZIP
  console.log('\n📦 Creazione ZIP...');
  execSync(`cd /home/claude/kit/OUT && zip -r "/mnt/user-data/outputs/KIT_FORMASUBITO_CalorEnergyVerona.zip" "KIT FORMASUBITO - Calor Energy Verona" -q`);

  console.log('\n═══════════════════════════════════════════════════');
  console.log('  ✅ KIT FORMASUBITO COMPLETATO!');
  console.log('═══════════════════════════════════════════════════');

  // Riepilogo
  const { execSync: ex } = require('child_process');
  const count = ex(`find "${OUT}" -name "*.docx" | wc -l`).toString().trim();
  const size = ex(`du -sh "${OUT}" 2>/dev/null | cut -f1`).toString().trim();
  console.log(`\n📁 Documenti generati: ${count} file .docx`);
  console.log(`💾 Dimensione totale: ${size}`);
}

main().catch(e => { console.error('ERRORE:', e); process.exit(1); });
