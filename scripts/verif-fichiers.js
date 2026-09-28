'use strict';
// Vérification des parseurs sur les fichiers réels de l'ERP (hors repo).
// Usage : node scripts/verif-fichiers.js <fichier-minimums.xlsx> <fichier-ventes.xlsx>

const fs = require('fs');
const { parserMinimums, parserVentes } = require('../lib/parseurs');

async function main() {
  const [fMin, fVentes] = process.argv.slice(2);
  if (!fMin || !fVentes) {
    console.error('Usage : node scripts/verif-fichiers.js <minimums.xlsx> <ventes.xlsx>');
    process.exit(1);
  }

  console.log(`— Minimums : ${fMin}`);
  const { articles, avertissements: a1 } = await parserMinimums(fs.readFileSync(fMin));
  console.log(`  ${articles.length} articles, dont ${articles.filter((x) => x.stockMin > 0).length} avec minimum > 0`);
  a1.forEach((m) => console.log(`  ⚠ ${m}`));
  const ex = articles.find((x) => x.stockMin > 0 && x.ean);
  console.log('  Exemple :', JSON.stringify(ex));

  console.log(`— Ventes : ${fVentes}`);
  const { lignes, avertissements: a2 } = await parserVentes(fs.readFileSync(fVentes));
  const totQ = lignes.reduce((s, l) => s + l.qteTotale, 0);
  const totF = lignes.reduce((s, l) => s + l.qteFami, 0);
  console.log(`  ${lignes.length} articles vendus — total ${totQ} pièces, dont Fami ${totF}`);
  a2.forEach((m) => console.log(`  ⚠ ${m}`));
  console.log('  Exemple :', JSON.stringify(lignes[0]));

  // Taux de raccordement ventes -> référentiel produits
  const parCode = new Set(articles.map((x) => x.codeArticle));
  const parEan = new Set(articles.filter((x) => x.ean).map((x) => x.ean));
  let viaCode = 0, viaEan = 0, orphelines = [];
  for (const l of lignes) {
    if (parCode.has(l.codeArticle)) viaCode++;
    else if (l.ean && parEan.has(l.ean)) viaEan++;
    else orphelines.push(l);
  }
  console.log(`— Raccordement : ${viaCode} par code article, ${viaEan} par EAN, ${orphelines.length} orphelines`);
  orphelines.slice(0, 5).forEach((l) => console.log(`  orpheline : ${l.codeArticle} / ${l.ean}`));
}

main().catch((e) => { console.error(e); process.exit(1); });
