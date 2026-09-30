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
  actif BOOLEAN NOT NULL DEFAULT TRUE,
  stock_depot INTEGER,
  stock_depot2 INTEGER,
  stock_fdcm INTEGER,
  maj_le TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_produits_ean ON produits(ean);

CREATE TABLE IF NOT EXISTS scans_inconnus (
  code TEXT PRIMARY KEY,
  nb INTEGER NOT NULL DEFAULT 1,
  premier_le TIMESTAMPTZ NOT NULL DEFAULT now(),
  dernier_le TIMESTAMPTZ NOT NULL DEFAULT now()
);

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
  statut TEXT NOT NULL DEFAULT 'a_traiter',
  motif_refus TEXT,
  decide_le TIMESTAMPTZ,
  cree_le TIMESTAMPTZ NOT NULL DEFAULT now(),
  maj_le TIMESTAMPTZ NOT NULL DEFAULT now(),
  exporte_le TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS collections_articles (
  code_article TEXT PRIMARY KEY,
  ean TEXT,
  collection TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS collections (
  nom TEXT PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS collegues (
  nom TEXT PRIMARY KEY,
  ajoute_le TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS propositions_collection (
  code_article TEXT PRIMARY KEY,
  ean TEXT,
  description TEXT,
  collection_actuelle TEXT,
  collection_proposee TEXT NOT NULL,
  auteur TEXT,
  statut TEXT NOT NULL DEFAULT 'a_traiter',
  motif_refus TEXT,
  decide_le TIMESTAMPTZ,
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

// Colonnes ajoutées après la première mise en production : appliquées une à une,
// l'échec « colonne existe déjà » est ignoré.
const MIGRATIONS = [
  'ALTER TABLE produits ADD COLUMN actif BOOLEAN NOT NULL DEFAULT TRUE',
  'ALTER TABLE produits ADD COLUMN stock_depot INTEGER',
  'ALTER TABLE produits ADD COLUMN stock_depot2 INTEGER',
  'ALTER TABLE produits ADD COLUMN stock_fdcm INTEGER',
  "ALTER TABLE propositions ADD COLUMN statut TEXT NOT NULL DEFAULT 'a_traiter'",
  'ALTER TABLE propositions ADD COLUMN motif_refus TEXT',
  'ALTER TABLE propositions ADD COLUMN decide_le TIMESTAMPTZ',
  // Reprise de l'existant : ce qui avait déjà été exporté passe « traitée » (idempotent)
  "UPDATE propositions SET statut = 'traitee' WHERE exporte_le IS NOT NULL AND statut = 'a_traiter'",
  // Validation automatique : plus d'état « à traiter », tout est validé sauf refus explicite
  "UPDATE propositions SET statut = 'validee' WHERE statut = 'a_traiter'",
  "UPDATE propositions_collection SET statut = 'validee' WHERE statut = 'a_traiter'",
];

async function initSchema() {
  const client = await getPool().connect();
  try {
    for (const stmt of SCHEMA.split(';').map((s) => s.trim()).filter(Boolean)) {
      await client.query(stmt);
    }
    for (const stmt of MIGRATIONS) {
      try {
        await client.query(stmt);
      } catch (e) {
        if (!/existe déjà|already exists|duplicate column/i.test(e.message)) throw e;
      }
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
        const b = j * 15;
        valeurs.push(`($${b + 1},$${b + 2},$${b + 3},$${b + 4},$${b + 5},$${b + 6},$${b + 7},$${b + 8},$${b + 9},$${b + 10},$${b + 11},$${b + 12},$${b + 13},$${b + 14},$${b + 15})`);
        params.push(a.codeArticle, a.ean || null, a.description, a.stockMin, a.stockMax, a.stock,
          a.vpe, a.vk12, a.prixAchat, a.fournisseur || null, a.numCommande || null,
          a.actif !== false, a.stockDepot, a.stockDepot2, a.stockFdcm);
      });
      await client.query(
        `INSERT INTO produits (code_article, ean, description, stock_min, stock_max, stock, vpe, vk12, prix_achat, fournisseur, num_commande, actif, stock_depot, stock_depot2, stock_fdcm)
         VALUES ${valeurs.join(',')}
         ON CONFLICT (code_article) DO UPDATE SET
           ean = EXCLUDED.ean, description = EXCLUDED.description, stock_min = EXCLUDED.stock_min,
           stock_max = EXCLUDED.stock_max, stock = EXCLUDED.stock, vpe = EXCLUDED.vpe, vk12 = EXCLUDED.vk12,
           prix_achat = EXCLUDED.prix_achat, fournisseur = EXCLUDED.fournisseur, num_commande = EXCLUDED.num_commande,
           actif = EXCLUDED.actif, stock_depot = EXCLUDED.stock_depot, stock_depot2 = EXCLUDED.stock_depot2, stock_fdcm = EXCLUDED.stock_fdcm`,
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

// Remplace tout le référentiel collections (import du fichier collections).
async function remplacerCollections(lignes, collections) {
  const client = await getPool().connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM collections_articles');
    await client.query('DELETE FROM collections');
    const TAILLE = 500;
    for (let i = 0; i < lignes.length; i += TAILLE) {
      const lot = lignes.slice(i, i + TAILLE);
      const valeurs = [];
      const params = [];
      lot.forEach((l, j) => {
        const b = j * 3;
        valeurs.push(`($${b + 1},$${b + 2},$${b + 3})`);
        params.push(l.codeArticle, l.ean || null, l.collection);
      });
      await client.query(
        `INSERT INTO collections_articles (code_article, ean, collection) VALUES ${valeurs.join(',')}`,
        params
      );
    }
    for (let i = 0; i < collections.length; i += TAILLE) {
      const lot = collections.slice(i, i + TAILLE);
      await client.query(
        `INSERT INTO collections (nom) VALUES ${lot.map((_, j) => `($${j + 1})`).join(',')}`,
        lot
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

module.exports = { setPool, getPool, initSchema, requete, remplacerProduits, remplacerVentesSemaine, remplacerCollections };
