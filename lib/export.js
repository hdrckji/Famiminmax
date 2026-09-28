'use strict';
// Génération du fichier Excel des propositions d'adaptation.

const ExcelJS = require('exceljs');

const VERT_FONCE = 'FF245D39';
const VERT_CLAIR = 'FFE6F0D7';

async function genererExportPropositions(propositions) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Famiminmax';
  const ws = wb.addWorksheet('Adaptations');

  ws.columns = [
    { header: 'Code article', key: 'code_article', width: 14 },
    { header: 'EAN', key: 'ean', width: 16 },
    { header: 'Description', key: 'description', width: 36 },
    { header: 'Fournisseur', key: 'fournisseur', width: 26 },
    { header: 'Min actuel', key: 'min_actuel', width: 11 },
    { header: 'Min proposé', key: 'min_propose', width: 12 },
    { header: 'Écart', key: 'ecart', width: 8 },
    { header: 'Commentaire', key: 'commentaire', width: 40 },
    { header: 'Auteur', key: 'auteur', width: 16 },
    { header: 'Proposé le', key: 'maj_le', width: 18 },
  ];

  const entete = ws.getRow(1);
  entete.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  entete.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERT_FONCE } };

  propositions.forEach((p, i) => {
    const row = ws.addRow({
      code_article: p.code_article,
      ean: p.ean || '',
      description: p.description || '',
      fournisseur: p.fournisseur || '',
      min_actuel: p.min_actuel,
      min_propose: p.min_propose,
      ecart: p.min_propose - (p.min_actuel || 0),
      commentaire: p.commentaire || '',
      auteur: p.auteur || '',
      maj_le: p.maj_le ? new Date(p.maj_le) : null,
    });
    if (i % 2 === 1) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: VERT_CLAIR } };
    }
    row.getCell('ean').numFmt = '@';
    row.getCell('maj_le').numFmt = 'dd/mm/yyyy hh:mm';
  });

  ws.autoFilter = { from: 'A1', to: 'J1' };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  return Buffer.from(await wb.xlsx.writeBuffer());
}

module.exports = { genererExportPropositions };
