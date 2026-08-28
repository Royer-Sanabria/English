/**
 * Google Apps Script — backend para el Atajo "Inglés" de iOS y la página web.
 * Proyecto standalone en script.google.com, no vinculado a la hoja.
 *
 * Uso:
 *  - Agregar palabra:  GET /exec?text=<palabra o frase>
 *  - Listar palabras:  GET /exec?action=list
 *  - Borrar palabra:   GET /exec?action=delete&row=<numero de fila>
 */

const SHEET_ID = '1-S-yUPdSiHzM40Ff7WcXlygHVeZ2sl5IZEmNEpUS26c';
const HEADERS = ['Fecha', 'Palabra'];

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = HEADERS.every((h, i) => firstRow[i] === h);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function listWords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse_({ ok: true, words: [] });

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const words = data.map((row, i) => ({
    row: i + 2,
    date: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
    text: row[1]
  })).reverse();

  return jsonResponse_({ ok: true, words });
}

function deleteWord_(sheet, rowParam) {
  const row = parseInt(rowParam, 10);
  if (!row || row < 2) return jsonResponse_({ ok: false, error: 'Invalid row' });
  sheet.deleteRow(row);
  return jsonResponse_({ ok: true });
}

function doGet(e) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  ensureHeaders_(sheet);

  const action = e.parameter.action;

  if (action === 'list') {
    return listWords_(sheet);
  }

  if (action === 'delete') {
    return deleteWord_(sheet, e.parameter.row);
  }

  const text = e.parameter.text;
  if (text && text.trim()) {
    sheet.appendRow([new Date(), text.trim()]);
    return jsonResponse_({ ok: true, saved: text.trim() });
  }

  return jsonResponse_({ ok: false, error: 'No text provided' });
}
