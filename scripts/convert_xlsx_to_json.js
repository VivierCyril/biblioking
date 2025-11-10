const path = require('path');
const fs = require('fs').promises;
const fsSync = require('fs');
const xlsx = require('xlsx');

const xlsxPath = process.argv[2];
if (!xlsxPath) {
  console.error('Usage: node convert_xlsx_to_json.js "<chemin/vers/Liste SK 2025.xlsx>"');
  process.exit(1);
}

function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (v === null || v === undefined) return false;
  const s = String(v).trim().toLowerCase();
  return ['oui','o','yes','y','true','vrai','1','x'].includes(s);
}

function pick(row, keys) {
  for (const k of keys) {
    if (k in row) return row[k];
  }
  return undefined;
}

(async () => {
  try {
    const workbook = xlsx.readFile(xlsxPath);
    const sheetName = workbook.SheetNames[0];
    const rows = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });

    const result = {};

    for (const row of rows) {
      const type = (pick(row, ['type','Type','TYPE','categorie','Categorie','catégorie']) || '').toString().trim();
      if (!type) continue;

      const titre_fr = (pick(row, ['titre_fr','Titre_fr','Titre Français','Titre Français','titre','Titre']) || '').toString().trim();
      const titre_original = (pick(row, ['titre_original','Titre_original','titre_original','Titre Original','original']) || '').toString().trim();
      const anneeRaw = pick(row, ['annee','Année','Annee','year','Year']) || '';
      const annee = anneeRaw === '' ? null : (Number(anneeRaw) || anneeRaw);
      const possedeRaw = pick(row, ['possede','Possede','possédé','Possédé','possesse','possession']) || '';
      const possede = toBool(possedeRaw);

      const item = {
        titre_fr: titre_fr,
        titre_original: titre_original,
        annee: annee,
        possede: possede
      };

      if (!result[type]) result[type] = [];
      result[type].push(item);
    }

    const dataDir = path.join(__dirname, '..', 'data');
    const outPath = path.join(dataDir, 'stephen_king_oeuvres.json');

    // backup existing file if present
    if (fsSync.existsSync(outPath)) {
      const backupPath = outPath + '.bak.' + Date.now();
      await fs.copyFile(outPath, backupPath);
      console.log('Sauvegarde créée:', backupPath);
    }

    await fs.mkdir(dataDir, { recursive: true });
    await fs.writeFile(outPath, JSON.stringify(result, null, 2), 'utf8');
    console.log('Fichier écrit :', outPath);
  } catch (err) {
    console.error('Erreur :', err);
    process.exit(1);
  }
})();