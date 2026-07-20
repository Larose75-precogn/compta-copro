/**
 * Compta Copro — Communicator Google Sheets
 *
 * Installation :
 * 1. Dans le sheet, Extensions → Apps Script → coller ce code
 * 2. Configurer ORG_ID et API_BASE ci-dessous
 * 3. Ajouter le déclencheur : Déclencheurs → + Ajouter → onEdit → Source = Feuille de calcul → Type = Au moment de la modification
 * 4. Autoriser l'accès (UrlFetchApp demande une permission au premier run)
 *
 * Fonctionnement :
 * - Tout onglet qui commence par "1 - " est un compte comptable (section 1)
 * - Structure attendue : ligne 1 = en-tête compte [code, nom, ...], ligne 2 = colonnes [Date, Libellé, Débit, Crédit, Solde, ...]
 * - Quand tu complètes une ligne (Date + Libellé + Débit OU Crédit), l'écriture part dans ledger-cli
 * - La colonne STATUT (dernière colonne + 1) affiche ✅ ou ❌ erreur
 */

const API_BASE = 'http://vps-03db771f.vps.ovh.net:8080';
const ORG_ID   = 'copro_1crE1G2RerFeXQfHNh0yERfvfAjVKGUz53LE9szCqMMs';

// Colonne fixe pour le statut (H = 8). Si ton sheet a plus de colonnes, ajuste.
const COL_DATE    = 1;  // A
const COL_LIBELLE = 2;  // B
const COL_DEBIT   = 3;  // C
const COL_CREDIT  = 4;  // D
const COL_STATUT  = 7;  // G (après Solde col E et une col vide F pour 451xxx)

function onEdit(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Seulement les onglets section 1 (comptes comptables)
  if (!sheetName.startsWith('1 - ')) return;
  if (sheetName === '1 - Balance') return;  // onglet synthèse, pas un compte

  const row = e.range.getRow();
  if (row <= 2) return;  // ligne 1 = en-tête compte, ligne 2 = colonnes

  // Lire la ligne
  const dateVal  = sheet.getRange(row, COL_DATE).getValue();
  const libelle  = String(sheet.getRange(row, COL_LIBELLE).getValue() || '').trim();
  const debit    = parseFloat(sheet.getRange(row, COL_DEBIT).getValue()  || 0);
  const credit   = parseFloat(sheet.getRange(row, COL_CREDIT).getValue() || 0);

  // Ligne incomplète : on attend
  if (!dateVal || !libelle) return;
  if (!debit && !credit) return;

  // Déjà envoyé
  const statut = sheet.getRange(row, COL_STATUT).getValue();
  if (statut === '✅') return;

  // Compte depuis ligne 1 col A (ex: "451001.0" → "451001")
  const compteRaw = String(sheet.getRange(1, 1).getValue() || '');
  const compte = compteRaw.replace(/\.0$/, '').trim();
  if (!compte) return;

  // Format date
  let dateStr;
  if (dateVal instanceof Date) {
    dateStr = Utilities.formatDate(dateVal, 'Europe/Paris', 'yyyy/MM/dd');
  } else {
    dateStr = String(dateVal).replace(/-/g, '/');
  }

  // Appel API
  const payload = {
    orgId: ORG_ID,
    compte: compte,
    libelle: libelle,
    montant_debit: debit > 0 ? debit : 0,
    montant_credit: credit > 0 ? credit : 0,
    date: dateStr
  };

  try {
    const resp = UrlFetchApp.fetch(API_BASE + '/api/ledger/sheet-entry', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    if (result.success) {
      sheet.getRange(row, COL_STATUT).setValue('✅');
      sheet.getRange(row, COL_STATUT).setNote(
        'Contrepartie : ' + result.contrepartie.compte + ' — ' + result.contrepartie.nom
      );
    } else {
      sheet.getRange(row, COL_STATUT).setValue('❌ ' + (result.error || 'erreur'));
    }
  } catch (err) {
    sheet.getRange(row, COL_STATUT).setValue('❌ ' + err.toString());
  }
}


/**
 * À lancer UNE SEULE FOIS après installation sur un sheet qui contient déjà
 * des données historiques déjà importées dans ledger-cli (via sheettojournal).
 * Marque toutes les lignes existantes comme ✅ SANS les renvoyer à l'API.
 * Après ça, seulement les nouvelles lignes saisies déclenchent un appel API.
 */
function markAllExistingAsDone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;
  for (const sheet of ss.getSheets()) {
    const name = sheet.getName();
    if (!name.startsWith('1 - ') || name === '1 - Balance') continue;
    const lastRow = sheet.getLastRow();
    for (let row = 3; row <= lastRow; row++) {
      const dateVal = sheet.getRange(row, COL_DATE).getValue();
      const libelle = String(sheet.getRange(row, COL_LIBELLE).getValue() || '').trim();
      const debit   = parseFloat(sheet.getRange(row, COL_DEBIT).getValue()  || 0);
      const credit  = parseFloat(sheet.getRange(row, COL_CREDIT).getValue() || 0);
      if (!dateVal || !libelle || (!debit && !credit)) continue;
      sheet.getRange(row, COL_STATUT).setValue('✅ historique');
      count++;
    }
  }
  SpreadsheetApp.getUi().alert(`${count} lignes historiques marquées ✅ (non renvoyées à l'API).`);
}


/**
 * Fonction utilitaire : renvoyer en masse toutes les lignes non encore envoyées
 * (utile après une panne ou pour rattraper des lignes manquées).
 * NE PAS lancer sur les données historiques → utiliser markAllExistingAsDone() d'abord.
 * Lancer depuis Apps Script → Exécuter → syncAll
 */
function syncAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let total = 0, ok = 0, err = 0;

  for (const sheet of sheets) {
    const name = sheet.getName();
    if (!name.startsWith('1 - ') || name === '1 - Balance') continue;

    const compteRaw = String(sheet.getRange(1, 1).getValue() || '');
    const compte = compteRaw.replace(/\.0$/, '').trim();
    if (!compte) continue;

    const lastRow = sheet.getLastRow();
    for (let row = 3; row <= lastRow; row++) {
      const statut = sheet.getRange(row, COL_STATUT).getValue();
      if (statut === '✅') continue;  // déjà envoyé

      const dateVal = sheet.getRange(row, COL_DATE).getValue();
      const libelle = String(sheet.getRange(row, COL_LIBELLE).getValue() || '').trim();
      const debit   = parseFloat(sheet.getRange(row, COL_DEBIT).getValue()  || 0);
      const credit  = parseFloat(sheet.getRange(row, COL_CREDIT).getValue() || 0);

      if (!dateVal || !libelle || (!debit && !credit)) continue;

      let dateStr;
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, 'Europe/Paris', 'yyyy/MM/dd');
      } else {
        dateStr = String(dateVal).replace(/-/g, '/');
      }

      total++;
      try {
        const resp = UrlFetchApp.fetch(API_BASE + '/api/ledger/sheet-entry', {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            orgId: ORG_ID, compte, libelle,
            montant_debit: debit > 0 ? debit : 0,
            montant_credit: credit > 0 ? credit : 0,
            date: dateStr
          }),
          muteHttpExceptions: true
        });
        const result = JSON.parse(resp.getContentText());
        if (result.success) {
          sheet.getRange(row, COL_STATUT).setValue('✅');
          ok++;
        } else {
          sheet.getRange(row, COL_STATUT).setValue('❌ ' + (result.error || '?'));
          err++;
        }
      } catch (e) {
        sheet.getRange(row, COL_STATUT).setValue('❌ ' + e.toString());
        err++;
      }

      Utilities.sleep(200);  // éviter le rate-limiting API
    }
  }

  SpreadsheetApp.getUi().alert(`syncAll terminé : ${total} lignes traitées, ${ok} OK, ${err} erreurs`);
}
