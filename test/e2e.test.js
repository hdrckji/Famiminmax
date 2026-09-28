'use strict';
// E2E complet sur base pg-mem (pas besoin de Postgres local) :
// import minimums -> import ventes -> scan -> proposition -> export Excel.

const { test, before, after } = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');
const { newDb } = require('pg-mem');

process.env.ADMIN_CODE = 'test-code';

const db = require('../lib/db');
const { serveur } = require('../server');

let base; // http://127.0.0.1:PORT
const ADMIN = { 'X-Admin-Code': 'test-code' };

async function fabriquerMinimums() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Export');
  ws.addRow(['Article', 'Description', 'VK-12', 'Stock (de marchandise)', 'Minimale Stock', 'Maximale Stock',
    'Récolte', 'VPE', 'PA', 'Fournisseur', 'N°decommande.', 'EAN barcode']);
  ws.addRow(['57913', 'Goudspray 150ml', '248', '124', '120', '132', '0', '12', '1.16', 'Goodmark Europe NV', '022280', '5410764216374']);
  ws.addRow(['61702', 'Gazebo Lemax', '9', '5', '6', '12', '0', '1', '7.27', 'Lemax BV', 'L-99', '0728162041609']);
  ws.addRow(['99001', 'Article sans EAN', '0', '2', '0', '0', '0', '1', '2', 'Divers', '', '']);
  // EAN écrit en cellule numérique (cas réel possible)
  ws.addRow(['99002', 'EAN numérique', 3, 7, 4, 8, 0, 1, 1.5, 'Divers', '', 5412345678908]);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function fabriquerVentes(qteGoudspray) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Feuil1');
  ws.addRow(['Artikelnummer', 'EANBarcode', 'Omschrijving artikel', 'Aantal', 'Fami (#)']);
  ws.addRow(['57913', '5410764216374', 'Goudspray 150ml', String(qteGoudspray), qteGoudspray + ',00']);
  ws.addRow(['61702', '0728162041609', 'Gazebo Lemax', '3', '2,00']);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function api(chemin, options) {
  return fetch(base + chemin, options);
}

before(async () => {
  const mem = newDb();
  const { Pool } = mem.adapters.createPg();
  db.setPool(new Pool());
  await db.initSchema();
  await new Promise((ok) => serveur.listen(0, '127.0.0.1', ok));
  base = `http://127.0.0.1:${serveur.address().port}`;
});

after(() => serveur.close());

test('les routes admin exigent le code', async () => {
  const rep = await api('/api/etat');
  assert.strictEqual(rep.status, 401);
});

test('import du fichier minimums', async () => {
  const rep = await api('/api/import/minimums', { method: 'POST', headers: ADMIN, body: await fabriquerMinimums() });
  const data = await rep.json();
  assert.strictEqual(rep.status, 200, JSON.stringify(data));
  assert.strictEqual(data.nbArticles, 4);
  assert.strictEqual(data.nbAvecMin, 3);
  assert.match(data.avertissements[0] || '', /sans EAN/);
});

test('import des ventes de deux semaines', async () => {
  for (const [lundi, qte] of [['2026-09-14', 18], ['2026-09-21', 25]]) {
    const rep = await api(`/api/import/ventes?semaine=${lundi}&fichier=test.xlsx`,
      { method: 'POST', headers: ADMIN, body: await fabriquerVentes(qte) });
    const data = await rep.json();
    assert.strictEqual(rep.status, 200, JSON.stringify(data));
    assert.strictEqual(data.nbLignes, 2);
  }
  // Réimport de la même semaine : remplace, pas de doublon
  const rep = await api('/api/import/ventes?semaine=2026-09-21&fichier=test2.xlsx',
    { method: 'POST', headers: ADMIN, body: await fabriquerVentes(30) });
  assert.strictEqual(rep.status, 200);
});

test('scan par EAN : fiche produit + ventes', async () => {
  const rep = await api('/api/produit?code=5410764216374');
  const data = await rep.json();
  assert.strictEqual(rep.status, 200, JSON.stringify(data));
  assert.strictEqual(data.produit.codeArticle, '57913');
  assert.strictEqual(data.produit.stockMin, 120);
  assert.strictEqual(data.produit.stockMax, 132);
  assert.strictEqual(data.produit.vk12, 248);
  assert.strictEqual(data.ventes.length, 2);
  assert.deepStrictEqual(data.ventes.map((v) => v.semaine), ['2026-09-14', '2026-09-21']);
  assert.deepStrictEqual(data.ventes.map((v) => v.qteFami), [18, 30]);
  assert.strictEqual(data.proposition, null);
});

test('scan UPC-A 12 chiffres retrouve l’EAN-13 à zéro de tête', async () => {
  const rep = await api('/api/produit?code=728162041609');
  const data = await rep.json();
  assert.strictEqual(rep.status, 200, JSON.stringify(data));
  assert.strictEqual(data.produit.codeArticle, '61702');
});

test('recherche par code article et code inconnu', async () => {
  const ok = await api('/api/produit?code=99001');
  assert.strictEqual(ok.status, 200);
  const ko = await api('/api/produit?code=0000000000000');
  assert.strictEqual(ko.status, 404);
});

test('proposition : création puis écrasement par la dernière', async () => {
  let rep = await api('/api/proposition', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ codeArticle: '57913', minPropose: 96, commentaire: 'Trop haut', auteur: 'Jimmy' }),
  });
  assert.strictEqual(rep.status, 200, JSON.stringify(await rep.json()));

  rep = await api('/api/proposition', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ codeArticle: '57913', minPropose: 108, commentaire: 'Plutôt 108', auteur: 'Kim' }),
  });
  assert.strictEqual(rep.status, 200);

  const liste = await (await api('/api/propositions', { headers: ADMIN })).json();
  assert.strictEqual(liste.propositions.length, 1);
  assert.strictEqual(liste.propositions[0].minPropose, 108);
  assert.strictEqual(liste.propositions[0].auteur, 'Kim');
  assert.strictEqual(liste.propositions[0].minActuel, 120);
  assert.strictEqual(liste.propositions[0].exportee, false);

  const fiche = await (await api('/api/produit?code=5410764216374')).json();
  assert.strictEqual(fiche.proposition.minPropose, 108);
});

test('proposition invalide refusée', async () => {
  const rep = await api('/api/proposition', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ codeArticle: '57913', minPropose: -3 }),
  });
  assert.strictEqual(rep.status, 400);
});

test('export Excel des nouvelles propositions puis marquage', async () => {
  const rep = await api('/api/export?mode=nouvelles', { headers: ADMIN });
  assert.strictEqual(rep.status, 200);
  assert.match(rep.headers.get('content-disposition') || '', /adaptations-minimums-.*\.xlsx/);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await rep.arrayBuffer()));
  const ws = wb.getWorksheet('Adaptations');
  assert.strictEqual(ws.rowCount, 2); // entête + 1 proposition
  assert.strictEqual(ws.getRow(2).getCell(1).value, '57913');
  assert.strictEqual(ws.getRow(2).getCell(6).value, 108);
  assert.strictEqual(ws.getRow(2).getCell(7).value, -12); // écart 108 - 120

  // Plus rien de nouveau à exporter
  const vide = await api('/api/export?mode=nouvelles', { headers: ADMIN });
  assert.strictEqual(vide.status, 404);
  // Mais « tout exporter » les reprend
  const toutes = await api('/api/export?mode=toutes', { headers: ADMIN });
  assert.strictEqual(toutes.status, 200);
});

test('réimport des minimums : le référentiel est remplacé, la proposition reste', async () => {
  const rep = await api('/api/import/minimums', { method: 'POST', headers: ADMIN, body: await fabriquerMinimums() });
  assert.strictEqual(rep.status, 200);
  const etat = await (await api('/api/etat', { headers: ADMIN })).json();
  assert.strictEqual(etat.produits.nb, 4);
  assert.strictEqual(etat.propositions.total, 1);
  assert.strictEqual(etat.semaines.length, 2);
});

test('suppression d’une proposition', async () => {
  const rep = await api('/api/proposition-suppr?codeArticle=57913', { method: 'POST', headers: ADMIN });
  assert.strictEqual(rep.status, 200);
  const liste = await (await api('/api/propositions', { headers: ADMIN })).json();
  assert.strictEqual(liste.propositions.length, 0);
});
