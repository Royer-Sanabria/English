/**
 * Google Apps Script — backend para el Atajo "Inglés" de iOS.
 * Se despliega como Web App (Extensiones > Apps Script en la hoja "English Words").
 *
 * Uso desde el Atajo: GET a la URL /exec con ?text=<palabra o frase>
 */

const HEADERS = ['Fecha', 'Palabra'];

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = HEADERS.every((h, i) => firstRow[i] === h);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders_(sheet);

  const text = e.parameter.text;

  if (text && text.trim()) {
    sheet.appendRow([new Date(), text.trim()]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, saved: text.trim() }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'No text provided' }))
    .setMimeType(ContentService.MimeType.JSON);
}
