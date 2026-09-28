'use strict';
// Famiminmax — serveur HTTP (Node pur, style Famipresta/cirque).
// Pages statiques dans /public, API JSON sous /api/*.

const http = require('http');
const fs = require('fs');
const path = require('path');
const db = require('./lib/db');
const { parserMinimums, parserVentes, candidatsScan } = require('./lib/parseurs');
const { genererExportPropositions } = require('./lib/export');

const PORT = process.env.PORT || 3000;
const ADMIN_CODE = process.env.ADMIN_CODE || '';
const MAX_UPLOAD = 30 * 1024 * 1024;

const MIMES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

function json(res, code, data) {
  const body = JSON.stringify(data);
  res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(body);
}

function lireCorps(req, limite = MAX_UPLOAD) {
  return new Promise((resolve, reject) => {
    const morceaux = [];
    let taille = 0;
    req.on('data', (c) => {
      taille += c.length;
      if (taille > limite) {
        reject(Object.assign(new Error('Fichier trop volumineux.'), { code: 413 }));
        req.destroy();
        return;
      }
      morceaux.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(morceaux)));
    req.on('error', reject);
  });
}

function estAdmin(req, url) {
  if (!ADMIN_CODE) return true; // pas de code configuré = accès libre (à éviter en prod)
  const code = req.headers['x-admin-code'] || url.searchParams.get('code') || '';
  return code === ADMIN_CODE;
}

function dateISO(d) {
  if (d == null) return null;
  if (d instanceof Date) return d.toISOString().slice(0, 10);
  return String(d).slice(0, 10);
}

// --- Handlers API -----------------------------------------------------------

async function apiProduit(req, res, url) {
  const brut = (url.searchParams.get('code') || '').trim();
  if (!brut) return json(res, 400, { erreur: 'Paramètre "code" manquant.' });

  const candidats = candidatsScan(brut);
  const placeholders = candidats.map((_, i) => `$${i + 1}`).join(',');
  const r = await db.requete(
    `SELECT * FROM produits WHERE ean IN (${placeholders}) OR code_article = $${candidats.length + 1} LIMIT 6`,
    [...candidats, brut]
  );
  if (!r.rows.length) return json(res, 404, { erreur: `Aucun produit trouvé pour « ${brut} ».` });
  if (r.rows.length > 1) {
    return json(res, 200, {
      choix: r.rows.map((p) => ({ codeArticle: p.code_article, ean: p.ean, description: p.description })),
    });
  }

  const p = r.rows[0];
  const ventes = await db.requete(
    `SELECT s.debut, s.nb_lignes, l.qte_totale, l.qte_fami
       FROM semaines_ventes s
       LEFT JOIN lignes_ventes l ON l.semaine = s.debut AND l.code_article = $1
      ORDER BY s.debut DESC
      LIMIT 5`,
    [p.code_article]
  );
  const prop = await db.requete('SELECT * FROM propositions WHERE code_article = $1', [p.code_article]);

  json(res, 200, {
    produit: {
      codeArticle: p.code_article,
      ean: p.ean,
      description: p.description,
      stockMin: p.stock_min,
      stockMax: p.stock_max,
      stock: p.stock,
      vpe: p.vpe,
      vk12: p.vk12,
      fournisseur: p.fournisseur,
      numCommande: p.num_commande,
      majLe: p.maj_le,
    },
    ventes: ventes.rows
      .map((v) => ({
        semaine: dateISO(v.debut),
        qteTotale: v.qte_totale == null ? 0 : Number(v.qte_totale),
        qteFami: v.qte_fami == null ? 0 : Number(v.qte_fami),
      }))
      .reverse(),
    proposition: prop.rows.length
      ? {
          minPropose: prop.rows[0].min_propose,
          minActuel: prop.rows[0].min_actuel,
          commentaire: prop.rows[0].commentaire,
          auteur: prop.rows[0].auteur,
          majLe: prop.rows[0].maj_le,
          exportee: prop.rows[0].exporte_le != null,
        }
      : null,
  });
}

async function apiProposer(req, res) {
  let corps;
  try {
    corps = JSON.parse((await lireCorps(req, 64 * 1024)).toString('utf8'));
  } catch {
    return json(res, 400, { erreur: 'Corps JSON invalide.' });
  }
  const codeArticle = String(corps.codeArticle || '').trim();
  const minPropose = Number(corps.minPropose);
  if (!codeArticle) return json(res, 400, { erreur: 'codeArticle manquant.' });
  if (!Number.isInteger(minPropose) || minPropose < 0 || minPropose > 100000) {
    return json(res, 400, { erreur: 'Le minimum proposé doit être un entier entre 0 et 100 000.' });
  }

  const r = await db.requete('SELECT * FROM produits WHERE code_article = $1', [codeArticle]);
  if (!r.rows.length) return json(res, 404, { erreur: 'Produit inconnu.' });
  const p = r.rows[0];

  const commentaire = String(corps.commentaire || '').trim().slice(0, 500) || null;
  const auteur = String(corps.auteur || '').trim().slice(0, 80) || null;

  // Une seule proposition par produit : la dernière écrase la précédente
  // et redevient « à exporter ».
  await db.requete(
    `INSERT INTO propositions (code_article, ean, description, min_actuel, min_propose, commentaire, auteur)
     VALUES ($1, $2, $3, $4, $5, $6, $7)
     ON CONFLICT (code_article) DO UPDATE SET
       ean = EXCLUDED.ean, description = EXCLUDED.description, min_actuel = EXCLUDED.min_actuel,
       min_propose = EXCLUDED.min_propose, commentaire = EXCLUDED.commentaire, auteur = EXCLUDED.auteur,
       maj_le = now(), exporte_le = NULL`,
    [codeArticle, p.ean, p.description, p.stock_min, minPropose, commentaire, auteur]
  );
  json(res, 200, { ok: true });
}

async function apiPropositions(req, res) {
  const r = await db.requete(
    `SELECT pr.*, p.fournisseur, p.stock_min AS min_courant
       FROM propositions pr LEFT JOIN produits p ON p.code_article = pr.code_article
      ORDER BY pr.maj_le DESC`
  );
  json(res, 200, {
    propositions: r.rows.map((x) => ({
      codeArticle: x.code_article,
      ean: x.ean,
      description: x.description,
      fournisseur: x.fournisseur,
      minActuel: x.min_actuel,
      minPropose: x.min_propose,
      commentaire: x.commentaire,
      auteur: x.auteur,
      majLe: x.maj_le,
      exportee: x.exporte_le != null,
    })),
  });
}

async function apiSupprimerProposition(req, res, url) {
  const code = (url.searchParams.get('codeArticle') || '').trim();
  if (!code) return json(res, 400, { erreur: 'codeArticle manquant.' });
  await db.requete('DELETE FROM propositions WHERE code_article = $1', [code]);
  json(res, 200, { ok: true });
}

async function apiImportMinimums(req, res) {
  const buffer = await lireCorps(req);
  if (!buffer.length) return json(res, 400, { erreur: 'Fichier vide.' });
  const { articles, avertissements } = await parserMinimums(buffer);
  await db.remplacerProduits(articles);
  json(res, 200, {
    ok: true,
    nbArticles: articles.length,
    nbAvecMin: articles.filter((a) => a.stockMin > 0).length,
    avertissements,
  });
}

async function apiImportVentes(req, res, url) {
  const semaine = url.searchParams.get('semaine') || '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(semaine)) {
    return json(res, 400, { erreur: 'Paramètre "semaine" attendu au format YYYY-MM-DD (lundi de la semaine).' });
  }
  const fichier = url.searchParams.get('fichier') || null;
  const buffer = await lireCorps(req);
  if (!buffer.length) return json(res, 400, { erreur: 'Fichier vide.' });
  const { lignes, avertissements } = await parserVentes(buffer);
  await db.remplacerVentesSemaine(semaine, fichier, lignes);
  json(res, 200, {
    ok: true,
    semaine,
    nbLignes: lignes.length,
    nbPieces: Math.round(lignes.reduce((s, l) => s + l.qteTotale, 0)),
    avertissements,
  });
}

async function apiEtat(req, res) {
  const prod = await db.requete('SELECT COUNT(*)::int AS nb, MAX(maj_le) AS maj FROM produits');
  const sem = await db.requete('SELECT debut, nb_lignes, importe_le FROM semaines_ventes ORDER BY debut DESC LIMIT 8');
  const prop = await db.requete(
    'SELECT COUNT(*)::int AS total, SUM(CASE WHEN exporte_le IS NULL THEN 1 ELSE 0 END)::int AS nouvelles FROM propositions'
  );
  json(res, 200, {
    produits: { nb: prod.rows[0].nb, majLe: prod.rows[0].maj },
    semaines: sem.rows.map((s) => ({ debut: dateISO(s.debut), nbLignes: s.nb_lignes, importeLe: s.importe_le })),
    propositions: prop.rows[0],
  });
}

async function apiExport(req, res, url) {
  const mode = url.searchParams.get('mode') === 'toutes' ? 'toutes' : 'nouvelles';
  const filtre = mode === 'nouvelles' ? 'WHERE pr.exporte_le IS NULL' : '';
  const r = await db.requete(
    `SELECT pr.*, p.fournisseur
       FROM propositions pr LEFT JOIN produits p ON p.code_article = pr.code_article
      ${filtre}
      ORDER BY pr.maj_le DESC`
  );
  if (!r.rows.length) return json(res, 404, { erreur: 'Aucune proposition à exporter.' });

  const buffer = await genererExportPropositions(r.rows);
  await db.requete('UPDATE propositions SET exporte_le = now() WHERE exporte_le IS NULL');
  const nom = `adaptations-minimums-${new Date().toISOString().slice(0, 10)}.xlsx`;
  res.writeHead(200, {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': `attachment; filename="${nom}"`,
    'Content-Length': buffer.length,
  });
  res.end(buffer);
}

// --- Routage ----------------------------------------------------------------

const PUBLIC_DIR = path.join(__dirname, 'public');

function servirStatique(res, fichier) {
  const cible = path.join(PUBLIC_DIR, fichier);
  if (!cible.startsWith(PUBLIC_DIR) || !fs.existsSync(cible) || !fs.statSync(cible).isFile()) {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    return res.end('Introuvable');
  }
  res.writeHead(200, { 'Content-Type': MIMES[path.extname(cible)] || 'application/octet-stream' });
  fs.createReadStream(cible).pipe(res);
}

const serveur = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const route = `${req.method} ${url.pathname}`;

  try {
    if (url.pathname.startsWith('/api/')) {
      const ADMIN_ROUTES = ['/api/propositions', '/api/proposition-suppr', '/api/import/minimums', '/api/import/ventes', '/api/etat', '/api/export'];
      if (ADMIN_ROUTES.includes(url.pathname) && !estAdmin(req, url)) {
        return json(res, 401, { erreur: 'Code d’accès admin invalide.' });
      }
      if (route === 'GET /api/sante') return json(res, 200, { ok: true });
      if (route === 'GET /api/produit') return await apiProduit(req, res, url);
      if (route === 'POST /api/proposition') return await apiProposer(req, res);
      if (route === 'GET /api/propositions') return await apiPropositions(req, res);
      if (route === 'POST /api/proposition-suppr') return await apiSupprimerProposition(req, res, url);
      if (route === 'POST /api/import/minimums') return await apiImportMinimums(req, res);
      if (route === 'POST /api/import/ventes') return await apiImportVentes(req, res, url);
      if (route === 'GET /api/etat') return await apiEtat(req, res);
      if (route === 'GET /api/export') return await apiExport(req, res, url);
      return json(res, 404, { erreur: 'Route inconnue.' });
    }

    if (req.method !== 'GET' && req.method !== 'HEAD') {
      return json(res, 405, { erreur: 'Méthode non autorisée.' });
    }
    if (url.pathname === '/') return servirStatique(res, 'index.html');
    if (url.pathname === '/admin') return servirStatique(res, 'admin.html');
    return servirStatique(res, url.pathname.slice(1));
  } catch (e) {
    const code = e.code === 413 ? 413 : 500;
    console.error(`[erreur] ${route} :`, e.message);
    json(res, code, { erreur: e.message || 'Erreur interne.' });
  }
});

async function demarrer() {
  await db.initSchema();
  serveur.listen(PORT, () => {
    console.log(`Famiminmax en écoute sur le port ${PORT}`);
    if (!ADMIN_CODE) console.warn('⚠ ADMIN_CODE non défini : les pages admin sont accessibles sans code.');
  });
}

if (require.main === module) {
  demarrer().catch((e) => {
    console.error('Démarrage impossible :', e);
    process.exit(1);
  });
}

module.exports = { serveur, demarrer };
