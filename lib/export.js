'use strict';
// Génération du fichier Excel des propositions d'adaptation.
// Les 9 premières colonnes suivent le format d'import de l'ERP (intitulés exacts) ;
// les colonnes suivantes sont informatives (« vue d'ensemble »).

const ExcelJS = require('exceljs');

const VERT_FONCE = 'FF245D39';
const VERT_CLAIR = 'FFE6F0D7';
const MAGAZIJN = 'Fami';

async function genererExportPropositions(propositions) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Famiminmax';
  const ws = wb.addWorksheet('Adaptations');

  ws.columns = [
    // --- Format ERP ---
    { header: 'artikelnummer', key: 'artikelnummer', width: 14 },
    { header: 'minimale voorraad', key: 'minimale_voorraad', width: 16 },
    { header: 'Omschrijving', key: 'omschrijving', width: 36 },
    { header: 'Magazijn', key: 'magazijn', width: 10 },
    { header: 'eanbarcode', key: 'eanbarcode', width: 16 },
    { header: 'BestelNr.', key: 'bestelnr', width: 14 },
    { header: 'afname', key: 'afname', width: 8 },
    { header: 'maximale voorraad', key: 'maximale_voorraad', width: 16 },
    { header: 'gereserveerde voorraad', key: 'gereserveerde_voorraad', width: 20 },
    // --- Vue d'ensemble ---
    { header: 'Min actuel', key: 'min_actuel', width: 11 },
    { header: 'Écart', key: 'ecart', width: 8 },
    { header: 'Commentaire', key: 'commentaire', width: 40 },
    { header: 'Auteur', key: 'auteur', width: 16 },
    { header: 'Proposé le', key: 'maj_le', width: 18 },
    { header: 'Fournisseur', key: 'fournisseur', width: 26 },
    { header: 'Famille', key: 'famille', width: 12 },
  ];

  const entete = ws.getRow(1);
  entete.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  entete.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERT_FONCE } };

  propositions.forEach((p, i) => {
    const row = ws.addRow({
      artikelnummer: p.code_article,
      minimale_voorraad: p.min_propose,
      omschrijving: p.description || '',
      magazijn: MAGAZIJN,
      eanbarcode: p.ean || '',
      bestelnr: p.num_commande || '',
      afname: p.vpe == null ? '' : p.vpe,
      maximale_voorraad: p.stock_max == null ? '' : p.stock_max,
      gereserveerde_voorraad: 0,
      min_actuel: p.min_actuel,
      ecart: p.min_propose - (p.min_actuel || 0),
      commentaire: p.commentaire || '',
      auteur: p.auteur || '',
      maj_le: p.maj_le ? new Date(p.maj_le) : null,
      famille: p.famille || '',
    });
    if (i % 2 === 1) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERT_CLAIR } };
    }
    row.getCell('artikelnummer').numFmt = '@';
    row.getCell('eanbarcode').numFmt = '@';
    row.getCell('maj_le').numFmt = 'dd/mm/yyyy hh:mm';
  });

  ws.autoFilter = { from: 'A1', to: 'P1' };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// Export des propositions de collection — format provisoire, à caler sur
// le format d'import ERP quand il sera défini.
async function genererExportCollections(propositions) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Famiminmax';
  const ws = wb.addWorksheet('Collections');

  ws.columns = [
    { header: 'N° art', key: 'code_article', width: 14 },
    { header: 'Code EAN', key: 'ean', width: 16 },
    { header: 'Description', key: 'description', width: 36 },
    { header: 'Collectie actuelle', key: 'collection_actuelle', width: 22 },
    { header: 'Collectie proposée', key: 'collection_proposee', width: 22 },
    { header: 'Auteur', key: 'auteur', width: 16 },
    { header: 'Proposé le', key: 'maj_le', width: 18 },
    { header: 'Famille', key: 'famille', width: 12 },
  ];

  const entete = ws.getRow(1);
  entete.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  entete.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERT_FONCE } };

  propositions.forEach((p, i) => {
    const row = ws.addRow({
      code_article: p.code_article,
      ean: p.ean || '',
      description: p.description || '',
      collection_actuelle: p.collection_actuelle || '',
      collection_proposee: p.collection_proposee,
      auteur: p.auteur || '',
      maj_le: p.maj_le ? new Date(p.maj_le) : null,
      famille: p.famille || '',
    });
    if (i % 2 === 1) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERT_CLAIR } };
    }
    row.getCell('code_article').numFmt = '@';
    row.getCell('ean').numFmt = '@';
    row.getCell('maj_le').numFmt = 'dd/mm/yyyy hh:mm';
  });

  ws.autoFilter = { from: 'A1', to: 'H1' };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  return Buffer.from(await wb.xlsx.writeBuffer());
}

module.exports = { genererExportPropositions, genererExportCollections };
