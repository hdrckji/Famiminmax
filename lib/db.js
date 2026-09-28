'use strict';
// Accès Postgres. Le pool peut être injecté (tests pg-mem) via setPool().

const SCHEMA = `
CREATE TABLE IF NOT EXISTS produits (
  code_article TEXT PRIMARY KEY,
  ean TEXT,
  description TEXT NOT NULL DEFAULT '',
  stock_min INTEGER NOT NULL DEFAULT 0,
  stock_max INTEGER NOT NULL DEFAULT 0,
  stock INTEGER NOT NULL DEFAULT 0,
  vpe INTEGER,
  vk12 INTEGER,
  prix_achat NUMERIC,
  fournisseur TEXT,
  num_commande TEXT,
  maj_le TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_produits_ean ON produits(ean);

CREATE TABLE IF NOT EXISTS semaines_ventes (
  debut DATE PRIMARY KEY,
  fichier TEXT,
  nb_lignes INTEGER NOT NULL DEFAULT 0,
  importe_le TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS lignes_ventes (
  semaine DATE NOT NULL,
  code_article TEXT NOT NULL,
  ean TEXT,
  qte_totale NUMERIC NOT NULL DEFAULT 0,
  qte_fami NUMERIC NOT NULL DEFAULT 0,
  PRIMARY KEY (semaine, code_article)
);
CREATE INDEX IF NOT EXISTS idx_lignes_code ON lignes_ventes(code_article);

CREATE TABLE IF NOT EXISTS propositions (
  code_article TEXT PRIMARY KEY,
  ean TEXT,
  description TEXT,
  min_actuel INTEGER,
  min_propose INTEGER NOT NULL,
  commentaire TEXT,
  auteur TEXT,
  cree_le TIMESTAMPTZ NOT NULL DEFAULT now(),
  maj_le TIMESTAMPTZ NOT NULL DEFAULT now(),
  exporte_le TIMESTAMPTZ
);
`;

let pool = null;

function setPool(p) { pool = p; }

function getPool() {
  if (!pool) {
    const { Pool } = require('pg');
    if (!process.env.DATABASE_URL) {
      throw new Error('DATABASE_URL manquant : renseigne la variable d’environnement (Postgres Railway).');
    }
    pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      max: 5,
      ssl: /localhost|127\.0\.0\.1/.test(process.env.DATABASE_URL) ? false : { rejectUnauthorized: false },
    });
  }
  return pool;
}

async function initSchema() {
  const client = await getPool().connect();
  try {
    for (const stmt of SCHEMA.split(';').map((s) => s.trim()).filter(Boolean)) {
      await client.query(stmt);
    }
  } finally {
    client.release();
  }
}

const requete = (text, params) => getPool().query(text, params);

// Remplace tout le référentiel produits (import du fichier minimums).
async function remplacerProduits(articles) {
  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM produits');
    const TAILLE = 500;
    for (let i = 0; i < articles.length; i += TAILLE) {
      const lot = articles.slice(i, i + TAILLE);
      const valeurs = [];
      const params = [];
      lot.forEach((a, j) => {
        const b = j * 11;
        valeurs.push(`($${b + 1},$${b + 2},$${b + 3},$${b + 4},$${b + 5},$${b + 6},$${b + 7},$${b + 8},$${b + 9},$${b + 10},$${b + 11})`);
        params.push(a.codeArticle, a.ean || null, a.description, a.stockMin, a.stockMax, a.stock,
          a.vpe, a.vk12, a.prixAchat, a.fournisseur || null, a.numCommande || null);
      });
      await client.query(
        `INSERT INTO produits (code_article, ean, description, stock_min, stock_max, stock, vpe, vk12, prix_achat, fournisseur, num_commande)
         VALUES ${valeurs.join(',')}
         ON CONFLICT (code_article) DO UPDATE SET
           ean = EXCLUDED.ean, description = EXCLUDED.description, stock_min = EXCLUDED.stock_min,
           stock_max = EXCLUDED.stock_max, stock = EXCLUDED.stock, vpe = EXCLUDED.vpe, vk12 = EXCLUDED.vk12,
           prix_achat = EXCLUDED.prix_achat, fournisseur = EXCLUDED.fournisseur, num_commande = EXCLUDED.num_commande`,
        params
      );
    }
    await client.query('COMMIT');
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
}

// Remplace les ventes d'une semaine donnée (debut = lundi, 'YYYY-MM-DD').
async function remplacerVentesSemaine(debut, fichier, lignes) {
  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM lignes_ventes WHERE semaine = $1', [debut]);
    await client.query('DELETE FROM semaines_ventes WHERE debut = $1', [debut]);
    await client.query(
      'INSERT INTO semaines_ventes (debut, fichier, nb_lignes) VALUES ($1, $2, $3)',
      [debut, fichier || null, lignes.length]
    );
    const TAILLE = 500;
    for (let i = 0; i < lignes.length; i += TAILLE) {
      const lot = lignes.slice(i, i + TAILLE);
      const valeurs = [];
      const params = [];
      lot.forEach((l, j) => {
        const b = j * 5;
        valeurs.push(`($${b + 1},$${b + 2},$${b + 3},$${b + 4},$${b + 5})`);
        params.push(debut, l.codeArticle, l.ean || null, l.qteTotale, l.qteFami);
      });
      await client.query(
        `INSERT INTO lignes_ventes (semaine, code_article, ean, qte_totale, qte_fami) VALUES ${valeurs.join(',')}`,
        params
      );
    }
    await client.query('COMMIT');
  } catch (e) {
    await client.query('ROLLBACK');
    throw e;
  } finally {
    client.release();
  }
}

module.exports = { setPool, getPool, initSchema, requete, remplacerProduits, remplacerVentesSemaine };
