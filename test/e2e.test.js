'use strict';
// E2E complet sur base pg-mem (pas besoin de Postgres local) :
// imports (minimums / ventes / collections) -> scan -> propositions min & collection
// -> circuit acheteur (valider / refuser / rouvrir) -> alertes -> exports Excel.

const { test, before, after } = require('node:test');
const assert = require('node:assert');
const ExcelJS = require('exceljs');
const { newDb } = require('pg-mem');

process.env.ADMIN_CODE = 'test-code';

const db = require('../lib/db');
const { serveur } = require('../server');

let base; // http://127.0.0.1:PORT
const ADMIN = { 'X-Admin-Code': 'test-code' };
const JSON_H = { 'Content-Type': 'application/json' };

async function fabriquerMinimums() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Export');
  ws.addRow(['Article', 'Description', 'VK-12', 'Stock (de marchandise)', 'Minimale Stock', 'Maximale Stock',
    'Récolte', 'VPE', 'PA', 'Fournisseur', 'N°decommande.', 'EAN barcode', 'Actief AK', 'FDepot', 'FDepot2', 'FDCM']);
  // Doublon de code article (cas réel dans Beta4) : la dernière ligne doit gagner
  ws.addRow(['57913', 'Goudspray 150ml (ancien)', '248', '124', '96', '132', '0', '12', '1.16', 'Goodmark Europe NV', '022280', '5410764216374', 'True', '0', '0', '0']);
  ws.addRow(['57913', 'Goudspray 150ml', '248', '124', '120', '132', '0', '12', '1.16', 'Goodmark Europe NV', '022280', '5410764216374', 'True', '687', '0', '3888']);
  ws.addRow(['61702', 'Gazebo Lemax', '9', '5', '6', '12', '0', '1', '7.27', 'Lemax BV', 'L-99', '0728162041609', 'False', '0', '0', '0']);
  ws.addRow(['99001', 'Article sans EAN', '0', '2', '0', '0', '0', '1', '2', 'Divers', '', '', 'True', '', '', '']);
  // EAN écrit en cellule numérique (cas réel possible)
  ws.addRow(['99002', 'EAN numérique', 3, 7, 4, 8, 0, 1, 1.5, 'Divers', '', 5412345678908, 'True', 5, 0, 12]);
  // Inactif mais avec du stock dépôt : la proposition doit rester possible
  ws.addRow(['99003', 'Inactif avec stock dépôt', '0', '0', '0', '0', '0', '1', '1', 'Divers', '', '5410999999990', 'False', '0', '0', '50']);
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

async function fabriquerCollections() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Feuil1');
  ws.addRow(['N° art', 'Marque', 'Collectie', 'Type', 'Description', 'Code EAN']);
  ws.addRow(['57913', '', 'Lemax26-03', '', 'Goudspray 150ml', '5410764216374']);
  ws.addRow(['99002', '', 'SugarCrush26-06', '', 'EAN numérique', '5412345678908']);
  ws.addRow(['61702', '', '', '', 'Gazebo Lemax', '0728162041609']); // sans collection
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function api(chemin, options) {
  return fetch(base + chemin, options);
}
async function fiche(code) {
  return (await api('/api/produit?code=' + code)).json();
}
async function decider(type, codeArticle, decision, motif) {
  return api('/api/proposition-decision', {
    method: 'POST', headers: { ...ADMIN, ...JSON_H },
    body: JSON.stringify({ type, codeArticle, decision, motif }),
  });
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
  assert.strictEqual((await api('/api/etat')).status, 401);
  assert.strictEqual((await api('/api/proposition-decision', { method: 'POST' })).status, 401);
});

test('import du fichier minimums', async () => {
  const rep = await api('/api/import/minimums', { method: 'POST', headers: ADMIN, body: await fabriquerMinimums() });
  const data = await rep.json();
  assert.strictEqual(rep.status, 200, JSON.stringify(data));
  assert.strictEqual(data.nbArticles, 5); // le doublon 57913 est dédupliqué
  assert.strictEqual(data.nbAvecMin, 3);
  assert.ok(data.avertissements.some((a) => /doublon/.test(a)), 'avertissement doublons attendu');
  assert.ok(data.avertissements.some((a) => /sans EAN/.test(a)), 'avertissement sans EAN attendu');
});

test('import des ventes de deux semaines', async () => {
  for (const [lundi, qte] of [['2026-09-14', 18], ['2026-09-21', 25]]) {
    const rep = await api(`/api/import/ventes?semaine=${lundi}&fichier=test.xlsx`,
      { method: 'POST', headers: ADMIN, body: await fabriquerVentes(qte) });
    assert.strictEqual(rep.status, 200);
  }
  // Réimport de la même semaine : remplace, pas de doublon
  const rep = await api('/api/import/ventes?semaine=2026-09-21&fichier=test2.xlsx',
    { method: 'POST', headers: ADMIN, body: await fabriquerVentes(30) });
  assert.strictEqual(rep.status, 200);
});

test('import du fichier collections', async () => {
  const rep = await api('/api/import/collections', { method: 'POST', headers: ADMIN, body: await fabriquerCollections() });
  const data = await rep.json();
  assert.strictEqual(rep.status, 200, JSON.stringify(data));
  assert.strictEqual(data.nbArticles, 2); // seuls les articles avec collection
  assert.strictEqual(data.nbCollections, 2);
  const { collections } = await (await api('/api/collections')).json();
  assert.deepStrictEqual(collections, ['Lemax26-03', 'SugarCrush26-06']);
});

test('scan par EAN : fiche produit + ventes + collection', async () => {
  const data = await fiche('5410764216374');
  assert.strictEqual(data.produit.codeArticle, '57913');
  assert.strictEqual(data.produit.stockMin, 120);
  assert.strictEqual(data.produit.vk12, 248);
  assert.strictEqual(data.produit.actif, true);
  assert.strictEqual(data.produit.stockDepot, 687);
  assert.strictEqual(data.produit.stockFdcm, 3888);
  assert.strictEqual(data.collection, 'Lemax26-03');
  assert.deepStrictEqual(data.ventes.map((v) => v.qteFami), [18, 30]);
  assert.strictEqual(data.proposition, null);
  assert.strictEqual(data.propositionCollection, null);
});

test('scan UPC-A 12 chiffres retrouve l’EAN-13 à zéro de tête', async () => {
  const data = await fiche('728162041609');
  assert.strictEqual(data.produit.codeArticle, '61702');
  assert.strictEqual(data.produit.actif, false);
  assert.strictEqual(data.collection, null); // pas de collection renseignée
});

test('les scans introuvables sont journalisés puis vidables', async () => {
  await api('/api/scans-inconnus-vider', { method: 'POST', headers: ADMIN });
  await api('/api/produit?code=5400924479374');
  await api('/api/produit?code=5400924479374');
  await api('/api/produit?code=1112223334445');
  let { scans } = await (await api('/api/scans-inconnus', { headers: ADMIN })).json();
  assert.strictEqual(scans.length, 2);
  assert.strictEqual(scans.find((s) => s.code === '5400924479374').nb, 2);
  await api('/api/scans-inconnus-vider', { method: 'POST', headers: ADMIN });
  ({ scans } = await (await api('/api/scans-inconnus', { headers: ADMIN })).json());
  assert.strictEqual(scans.length, 0);
});

test('proposition min : circuit complet avec alertes', async () => {
  // 1. Proposition douteuse : min 10 (ventes moy. 24/sem, 4575 en dépôt, actuel 120, VPE 12)
  let rep = await api('/api/proposition', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '57913', minPropose: 10, commentaire: 'Trop haut', auteur: 'Jimmy' }),
  });
  assert.strictEqual(rep.status, 200);
  let f = await fiche('57913');
  assert.strictEqual(f.proposition.statut, 'validee'); // validée d'office

  let { propositions } = await (await api('/api/propositions', { headers: ADMIN })).json();
  assert.strictEqual(propositions.length, 1);
  assert.deepStrictEqual(propositions[0].alertes.map((a) => a.code).sort(),
    ['depot', 'variation', 'ventes', 'vpe']);

  // 2. La dernière proposition écrase : min 108 (plus raisonnable, reste l'alerte dépôt)
  rep = await api('/api/proposition', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '57913', minPropose: 108, auteur: 'Kim' }),
  });
  assert.strictEqual(rep.status, 200);
  ({ propositions } = await (await api('/api/propositions', { headers: ADMIN })).json());
  assert.strictEqual(propositions.length, 1);
  assert.strictEqual(propositions[0].minPropose, 108);
  assert.strictEqual(propositions[0].auteur, 'Kim');
  assert.deepStrictEqual(propositions[0].alertes.map((a) => a.code), ['depot']);
});

test('proposition invalide refusée', async () => {
  const rep = await api('/api/proposition', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '57913', minPropose: -3 }),
  });
  assert.strictEqual(rep.status, 400);
});

test('refus avec motif, revalidation, puis nouvelle proposition', async () => {
  // Refus sans motif : rejeté ; décision inconnue : rejetée
  assert.strictEqual((await decider('min', '57913', 'refusee', '')).status, 400);
  assert.strictEqual((await decider('min', '57913', 'a_traiter')).status, 400);

  assert.strictEqual((await decider('min', '57913', 'refusee', 'Colisage imposé par le fournisseur')).status, 200);
  let f = await fiche('57913');
  assert.strictEqual(f.proposition.statut, 'refusee');
  assert.strictEqual(f.proposition.motifRefus, 'Colisage imposé par le fournisseur');

  // L'acheteur peut revenir sur son refus
  assert.strictEqual((await decider('min', '57913', 'validee')).status, 200);
  assert.strictEqual((await fiche('57913')).proposition.statut, 'validee');

  // Et le collègue peut reproposer : la dernière écrase, validée d'office
  const rep = await api('/api/proposition', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '57913', minPropose: 96, auteur: 'Kim' }),
  });
  assert.strictEqual(rep.status, 200);
  f = await fiche('57913');
  assert.strictEqual(f.proposition.statut, 'validee');
  assert.strictEqual(f.proposition.motifRefus, null);
});

test('export ERP : la proposition validée d’office passe traitée', async () => {
  const etat = await (await api('/api/etat', { headers: ADMIN })).json();
  assert.strictEqual(etat.propositions.aTraiter, 0);
  assert.strictEqual(etat.propositions.validees, 1);

  const rep = await api('/api/export?type=min&mode=validees', { headers: ADMIN });
  assert.strictEqual(rep.status, 200);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await rep.arrayBuffer()));
  const ws = wb.getWorksheet('Adaptations');
  assert.strictEqual(ws.rowCount, 2);
  assert.deepStrictEqual(ws.getRow(1).values.slice(1, 10),
    ['artikelnummer', 'minimale voorraad', 'Omschrijving', 'Magazijn', 'eanbarcode',
      'BestelNr.', 'afname', 'maximale voorraad', 'gereserveerde voorraad']);
  const ligne = ws.getRow(2);
  assert.strictEqual(ligne.getCell(1).value, '57913');
  assert.strictEqual(ligne.getCell(2).value, 96);
  assert.strictEqual(ligne.getCell(4).value, 'Fami');
  assert.strictEqual(ligne.getCell(6).value, '022280');
  assert.strictEqual(ligne.getCell(7).value, 12);

  const f = await fiche('57913');
  assert.strictEqual(f.proposition.statut, 'traitee');
  // Plus rien de validé à exporter, mais l'historique reste re-téléchargeable
  assert.strictEqual((await api('/api/export?type=min&mode=validees', { headers: ADMIN })).status, 404);
  assert.strictEqual((await api('/api/export?type=min&mode=toutes', { headers: ADMIN })).status, 200);
});

test('proposition de collection : liste fermée, circuit et export', async () => {
  // Collection inconnue : rejetée
  let rep = await api('/api/proposition-collection', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '57913', collectionProposee: 'NImporteQuoi26-01', auteur: 'Jimmy' }),
  });
  assert.strictEqual(rep.status, 400);

  rep = await api('/api/proposition-collection', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '57913', collectionProposee: 'SugarCrush26-06', auteur: 'Jimmy' }),
  });
  assert.strictEqual(rep.status, 200, JSON.stringify(await rep.json()));

  const f = await fiche('57913');
  assert.strictEqual(f.propositionCollection.statut, 'validee'); // validée d'office
  assert.strictEqual(f.propositionCollection.collectionProposee, 'SugarCrush26-06');
  assert.strictEqual(f.propositionCollection.collectionActuelle, 'Lemax26-03');

  const { propositions } = await (await api('/api/propositions', { headers: ADMIN })).json();
  const coll = propositions.find((p) => p.type === 'collection');
  assert.ok(coll, 'la proposition de collection doit être dans la liste admin');
  assert.strictEqual(coll.statut, 'validee');

  const exp = await api('/api/export?type=collection&mode=validees', { headers: ADMIN });
  assert.strictEqual(exp.status, 200);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await exp.arrayBuffer()));
  const ws = wb.getWorksheet('Collections');
  assert.strictEqual(ws.rowCount, 2);
  assert.strictEqual(ws.getRow(2).getCell(1).value, '57913');
  assert.strictEqual(ws.getRow(2).getCell(4).value, 'Lemax26-03');
  assert.strictEqual(ws.getRow(2).getCell(5).value, 'SugarCrush26-06');

  assert.strictEqual((await fiche('57913')).propositionCollection.statut, 'traitee');
});

test('réimport des minimums : propositions et collections survivent', async () => {
  const rep = await api('/api/import/minimums', { method: 'POST', headers: ADMIN, body: await fabriquerMinimums() });
  assert.strictEqual(rep.status, 200);
  const etat = await (await api('/api/etat', { headers: ADMIN })).json();
  assert.strictEqual(etat.produits.nb, 5);
  assert.strictEqual(etat.propositions.total, 2); // 1 min + 1 collection, traitées
  assert.strictEqual(etat.collections.nbArticles, 2);
  assert.strictEqual((await fiche('57913')).collection, 'Lemax26-03');
});

test('suppression des deux types de propositions', async () => {
  assert.strictEqual((await api('/api/proposition-suppr?type=min&codeArticle=57913', { method: 'POST', headers: ADMIN })).status, 200);
  assert.strictEqual((await api('/api/proposition-suppr?type=collection&codeArticle=57913', { method: 'POST', headers: ADMIN })).status, 200);
  const { propositions } = await (await api('/api/propositions', { headers: ADMIN })).json();
  assert.strictEqual(propositions.length, 0);
});

test('proposition bloquée pour un article inactif sans stock dépôt', async () => {
  // 61702 : inactif, dépôts 0/0/0 -> refusé
  const ko = await api('/api/proposition', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '61702', minPropose: 10 }),
  });
  assert.strictEqual(ko.status, 409);
  assert.match((await ko.json()).erreur, /inactif sans stock dépôt/);

  // 99003 : inactif mais 50 en FDCM -> autorisé
  const ok = await api('/api/proposition', {
    method: 'POST', headers: JSON_H,
    body: JSON.stringify({ codeArticle: '99003', minPropose: 4 }),
  });
  assert.strictEqual(ok.status, 200, JSON.stringify(await ok.json()));
  await api('/api/proposition-suppr?type=min&codeArticle=99003', { method: 'POST', headers: ADMIN });
});

test('décision sur une proposition inexistante : 404', async () => {
  assert.strictEqual((await decider('min', '00000', 'validee')).status, 404);
});
