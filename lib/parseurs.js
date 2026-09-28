'use strict';
// Parseurs des fichiers Excel de l'ERP.
// Fichier minimums : export type "Beta4" (feuille Export, ~27k articles).
// Fichier ventes   : export hebdomadaire agrégé (Artikelnummer / EANBarcode / Aantal / Fami (#)).

const ExcelJS = require('exceljs');

// Normalise un intitulé de colonne : minuscules, sans accents ni caractères spéciaux.
function normEntete(v) {
  return String(v == null ? '' : v)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9#()\- ]/gi, '')
    .trim()
    .toLowerCase();
}

// Valeur de cellule exceljs -> texte brut (gère formules, richText, nombres).
function celTexte(cell) {
  let v = cell && cell.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.result != null) v = v.result;
    else if (v.richText) v = v.richText.map((r) => r.text).join('');
    else if (v.text != null) v = v.text;
  }
  return String(v).trim();
}

// Nombre depuis une cellule qui peut contenir "32", "32,00", 32 ou "".
function celNombre(cell) {
  const t = celTexte(cell);
  if (t === '') return 0;
  const n = parseFloat(t.replace(/\s/g, '').replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}

// Repère l'index (1-based) de chaque colonne attendue d'après la ligne d'entêtes.
// `attendu` : { cle: (entete_normalisee) => bool }. Première colonne qui matche gagne.
function mapperColonnes(ligneEntetes, attendu) {
  const map = {};
  ligneEntetes.eachCell({ includeEmpty: false }, (cell, col) => {
    const h = normEntete(celTexte(cell));
    for (const [cle, test] of Object.entries(attendu)) {
      if (map[cle] == null && test(h)) map[cle] = col;
    }
  });
  return map;
}

/**
 * Parse le fichier des minimums rayon (type Beta4).
 * @param {Buffer} buffer contenu .xlsx
 * @returns {{articles: Array, avertissements: string[]}}
 */
async function parserMinimums(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.getWorksheet('Export') || wb.worksheets[0];
  if (!ws) throw new Error('Aucune feuille trouvée dans le fichier.');

  const cols = mapperColonnes(ws.getRow(1), {
    article: (h) => h === 'article' || h === 'artikelnummer',
    description: (h) => h === 'description' || h.startsWith('omschrijving'),
    vk12: (h) => h === 'vk-12' || h === 'vk12',
    stock: (h) => h.startsWith('stock'),
    stockMin: (h) => h.startsWith('minimale stock'),
    stockMax: (h) => h.startsWith('maximale stock'),
    vpe: (h) => h === 'vpe',
    prixAchat: (h) => h === 'pa',
    fournisseur: (h) => h === 'fournisseur' || h === 'leverancier',
    numCommande: (h) => h.includes('commande') || h.includes('bestelnummer'),
    ean: (h) => h.includes('ean'),
  });

  const manquantes = ['article', 'stockMin', 'ean'].filter((c) => cols[c] == null);
  if (manquantes.length) {
    throw new Error(`Colonnes introuvables dans le fichier : ${manquantes.join(', ')}. Est-ce bien l'export des minimums rayon ?`);
  }

  // Un même code article peut apparaître sur plusieurs lignes de l'export :
  // on garde la dernière (Postgres refuse deux fois la même clé dans un INSERT).
  const parCode = new Map();
  const avertissements = [];
  let doublons = 0;
  ws.eachRow((row, num) => {
    if (num === 1) return;
    const code = celTexte(row.getCell(cols.article));
    if (!code) return;
    if (parCode.has(code)) doublons++;
    parCode.set(code, {
      codeArticle: code,
      ean: cols.ean ? celTexte(row.getCell(cols.ean)) : '',
      description: cols.description ? celTexte(row.getCell(cols.description)) : '',
      stockMin: Math.round(celNombre(row.getCell(cols.stockMin))),
      stockMax: cols.stockMax ? Math.round(celNombre(row.getCell(cols.stockMax))) : 0,
      stock: cols.stock ? Math.round(celNombre(row.getCell(cols.stock))) : 0,
      vpe: cols.vpe ? Math.round(celNombre(row.getCell(cols.vpe))) : null,
      vk12: cols.vk12 ? Math.round(celNombre(row.getCell(cols.vk12))) : null,
      prixAchat: cols.prixAchat ? celNombre(row.getCell(cols.prixAchat)) : null,
      fournisseur: cols.fournisseur ? celTexte(row.getCell(cols.fournisseur)) : '',
      numCommande: cols.numCommande ? celTexte(row.getCell(cols.numCommande)) : '',
    });
  });

  const articles = [...parCode.values()];
  if (!articles.length) throw new Error('Aucun article trouvé dans le fichier.');
  if (doublons) avertissements.push(`${doublons} ligne(s) en doublon de code article (dernière ligne du fichier conservée).`);
  const sansEan = articles.filter((a) => !a.ean).length;
  if (sansEan) avertissements.push(`${sansEan} article(s) sans EAN (introuvables au scan, consultables par code article).`);
  return { articles, avertissements };
}

/**
 * Parse le fichier des ventes hebdomadaires agrégées.
 * @param {Buffer} buffer contenu .xlsx
 * @returns {{lignes: Array, avertissements: string[]}}
 */
async function parserVentes(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('Aucune feuille trouvée dans le fichier.');

  const cols = mapperColonnes(ws.getRow(1), {
    article: (h) => h === 'artikelnummer' || h === 'article',
    ean: (h) => h === 'eanbarcode' || h.includes('ean'),
    description: (h) => h.startsWith('omschrijving') || h === 'description',
    qteTotale: (h) => h === 'aantal',
    qteFami: (h) => h === 'fami (#)',
  });

  const manquantes = ['article', 'qteTotale'].filter((c) => cols[c] == null);
  if (manquantes.length) {
    throw new Error(`Colonnes introuvables dans le fichier : ${manquantes.join(', ')}. Est-ce bien l'export des ventes de la semaine ?`);
  }

  // Certains articles peuvent apparaître sur plusieurs lignes : on cumule.
  const parCode = new Map();
  ws.eachRow((row, num) => {
    if (num === 1) return;
    const code = celTexte(row.getCell(cols.article));
    if (!code) return;
    const l = parCode.get(code) || {
      codeArticle: code,
      ean: cols.ean ? celTexte(row.getCell(cols.ean)) : '',
      qteTotale: 0,
      qteFami: 0,
    };
    l.qteTotale += celNombre(row.getCell(cols.qteTotale));
    l.qteFami += cols.qteFami != null ? celNombre(row.getCell(cols.qteFami)) : 0;
    parCode.set(code, l);
  });

  const lignes = [...parCode.values()];
  if (!lignes.length) throw new Error('Aucune ligne de vente trouvée dans le fichier.');
  const avertissements = [];
  if (cols.qteFami == null) avertissements.push('Colonne "Fami (#)" absente : ventes Fami comptées à 0, seul le total est disponible.');
  return { lignes, avertissements };
}

// Candidats de recherche pour un code scanné : EAN tel quel, variantes avec/sans zéro
// de tête (un UPC-A scanné en 12 chiffres correspond à un EAN-13 commençant par 0).
function candidatsScan(brut) {
  const code = String(brut || '').trim().replace(/\s/g, '');
  if (!code) return [];
  const c = new Set([code]);
  if (/^\d+$/.test(code)) {
    if (code.length === 12) c.add('0' + code);
    if (code.length === 13 && code.startsWith('0')) c.add(code.slice(1));
    if (code.length === 14 && code.startsWith('0')) c.add(code.slice(1)); // GTIN-14 -> EAN-13
  }
  return [...c];
}

module.exports = { parserMinimums, parserVentes, candidatsScan, normEntete };
