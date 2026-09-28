# Famiminmax

Consultation et adaptation des **minimums rayon** en magasin, via scan sur Zebra TC26.

L'ERP ne permet pas de voir le minimum rayon paramétré par produit. Famiminmax comble ce trou :
les collègues scannent un produit en magasin, voient le minimum paramétré, le stock, et les ventes
des 5 dernières semaines, puis proposent une adaptation. Les propositions sont récupérées en Excel
depuis la page admin pour être réinjectées dans l'ERP.

## Écrans

- **`/` — page scan (TC26)** : champ de scan toujours actif (le TC26 scanne en mode clavier dans
  Chrome, aucune app à installer). Affiche la fiche produit : minimum rayon, stock, max, VPE,
  ventes 12 mois (VK-12) et graphique des ventes Fami des 5 dernières semaines. Formulaire de
  proposition avec pas d'incrément = VPE. Une seule proposition par produit : la dernière écrase
  la précédente.
- **`/admin` — administration (PC)** : protégée par code d'accès. Import du fichier minimums,
  import des ventes hebdomadaires, tableau des propositions, export Excel (les propositions
  exportées sont marquées pour ne pas être retraitées).

## Fichiers attendus

| Import | Source ERP | Colonnes clés |
|---|---|---|
| Minimums rayon | Export type « Beta4 », feuille `Export` | `Article`, `EAN barcode`, `Minimale Stock`, `Maximale Stock`, `Stock`, `VPE`, `VK-12`, `Fournisseur` |
| Ventes d'une semaine | Export hebdo agrégé | `Artikelnummer`, `EANBarcode`, `Aantal`, `Fami (#)` |

Les colonnes sont repérées par leur intitulé (l'ordre n'a pas d'importance). Un article présent
sur plusieurs lignes de ventes est cumulé. Les EAN scannés en UPC-A (12 chiffres) retrouvent
automatiquement l'EAN-13 à zéro de tête.

## Routine hebdomadaire

1. Exporter depuis l'ERP le fichier minimums et le fichier des ventes de la semaine écoulée.
2. Sur `/admin` : importer les deux (choisir le **lundi** de la semaine pour les ventes).
3. Récupérer les propositions : « Exporter les nouvelles » → fichier
   `adaptations-minimums-AAAA-MM-JJ.xlsx` → adapter les minimums dans l'ERP.

## Déploiement Railway

1. Nouveau projet Railway → **Deploy from GitHub repo** → `hdrckji/Famiminmax`.
2. Ajouter une base **PostgreSQL** au projet.
3. Variables du service :
   - `DATABASE_URL` = `${{Postgres.DATABASE_URL}}`
   - `ADMIN_CODE` = code d'accès de la page admin (obligatoire en production)
4. Settings → Networking → **Generate Domain**.

Le schéma de base se crée tout seul au démarrage, aucune migration à lancer.

## Développement

```bash
npm install
npm test          # E2E complet sur pg-mem (aucun Postgres local requis)
npm run verif -- minimums.xlsx ventes.xlsx   # vérifie les parseurs sur des fichiers réels
```

`npm start` exige `DATABASE_URL` (et éventuellement `ADMIN_CODE`, sinon admin en accès libre).
