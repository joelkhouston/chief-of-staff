/**
 * Phase 1: Local Pseudonymization Layer (Colombia Ley 1581 Compliance)
 * Strips student and colleague PII locally before any payload touches the Gemini API.
 */

function assignMissingUUIDs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pseudonym_Map');
  if (!sheet) throw new Error('Missing "Pseudonym_Map" tab.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 4); // Real_Name, Email, UUID_Alias, Role
  const values = range.getValues();
  let updated = false;

  for (let i = 0; i < values.length; i++) {
    const [realName, email, uuid, role] = values[i];
    if ((realName || email) && !uuid) {
      const prefix = (role || 'USR').toString().trim().toUpperCase().slice(0, 3);
      values[i][2] = `${prefix}_${Utilities.getUuid().slice(0, 8)}`;
      updated = true;
    }
  }

  if (updated) {
    range.setValues(values);
  }
}

function scrubPII(rawText) {
  if (!rawText) return '';
  assignMissingUUIDs();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pseudonym_Map');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return rawText;

  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  // Sort longest names first so full names are scrubbed before single first names
  rows.sort((a, b) => (b[0] || '').length - (a[0] || '').length);

  let cleanText = rawText;
  rows.forEach(([realName, email, uuid]) => {
    if (!uuid) return;
    if (email) {
      const emailRegex = new RegExp(escapeRegex_(email.toString().trim()), 'gi');
      cleanText = cleanText.replace(emailRegex, uuid);
    }
    if (realName) {
      const nameRegex = new RegExp('\\b' + escapeRegex_(realName.toString().trim()) + '\\b', 'gi');
      cleanText = cleanText.replace(nameRegex, uuid);
    }
  });

  return cleanText;
}

function restorePII(aiText) {
  if (!aiText) return '';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pseudonym_Map');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return aiText;

  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  let restoredText = aiText;

  rows.forEach(([realName, , uuid]) => {
    if (realName && uuid) {
      const uuidRegex = new RegExp(escapeRegex_(uuid.toString().trim()), 'g');
      restoredText = restoredText.replace(uuidRegex, realName.toString().trim());
    }
  });

  return restoredText;
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}