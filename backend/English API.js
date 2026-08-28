/**
 * Google Apps Script — backend para el Atajo "Inglés" de iOS y la página web.
 * Proyecto standalone en script.google.com, no vinculado a la hoja.
 *
 * La generación de definición/ejemplos (Gemini) la hace la página web en el navegador;
 * este script solo guarda el resultado.
 *
 * Uso:
 *  - Agregar palabra:  GET /exec?text=<palabra o frase>   (si ya existe, suma 1 al conteo)
 *  - Listar palabras:  GET /exec?action=list
 *  - Borrar palabra:   GET /exec?action=delete&row=<numero de fila>
 *  - Guardar definición/ejemplos: GET /exec?action=update&row=<n>&definition=<...>&examples=<...>
 */

const SHEET_ID = '1-S-yUPdSiHzM40Ff7WcXlygHVeZ2sl5IZEmNEpUS26c';
const HEADERS = ['Primera vez', 'Palabra', 'Veces', 'Última vez', 'Definición', 'Ejemplos'];

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

function findRowByText_(sheet, text) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const normalized = text.trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === normalized) {
      return i + 2;
    }
  }
  return -1;
}

function addWord_(sheet, text) {
  const trimmed = text.trim();
  const existingRow = findRowByText_(sheet, trimmed);
  const now = new Date();

  if (existingRow > 0) {
    const countCell = sheet.getRange(existingRow, 3);
    const currentCount = Number(countCell.getValue()) || 1;
    const newCount = currentCount + 1;
    countCell.setValue(newCount);
    sheet.getRange(existingRow, 4).setValue(now);
    return jsonResponse_({ ok: true, saved: trimmed, count: newCount });
  }

  sheet.appendRow([now, trimmed, 1, now, '', '']);
  return jsonResponse_({ ok: true, saved: trimmed, count: 1 });
}

function listWords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse_({ ok: true, words: [] });

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const words = data.map((row, i) => ({
    row: i + 2,
    firstDate: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
    text: row[1],
    count: row[2] || 1,
    lastDate: row[3] instanceof Date ? row[3].toISOString() : String(row[3] || row[0]),
    definition: row[4] || '',
    examples: row[5] ? String(row[5]).split(' | ') : []
  })).sort((a, b) => new Date(b.lastDate) - new Date(a.lastDate));

  return jsonResponse_({ ok: true, words });
}

function deleteWord_(sheet, rowParam) {
  const row = parseInt(rowParam, 10);
  if (!row || row < 2) return jsonResponse_({ ok: false, error: 'Invalid row' });
  sheet.deleteRow(row);
  return jsonResponse_({ ok: true });
}

function updateWordInfo_(sheet, rowParam, definition, examplesParam) {
  const row = parseInt(rowParam, 10);
  if (!row || row < 2) return jsonResponse_({ ok: false, error: 'Invalid row' });

  sheet.getRange(row, 5).setValue(definition || '');
  sheet.getRange(row, 6).setValue(examplesParam || '');
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

  if (action === 'update') {
    return updateWordInfo_(sheet, e.parameter.row, e.parameter.definition, e.parameter.examples);
  }

  const text = e.parameter.text;
  if (text && text.trim()) {
    return addWord_(sheet, text);
  }

  return jsonResponse_({ ok: false, error: 'No text provided' });
}
