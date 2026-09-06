/**
 * Google Apps Script — backend para el Atajo "Inglés" de iOS y la página web.
 * Proyecto standalone en script.google.com, no vinculado a la hoja.
 *
 * La generación de definición/ejemplos/traducción/tiempos verbales (Gemini) corre
 * automáticamente cada 5 minutos vía un trigger de tiempo (processPendingWords),
 * y también puede hacerla la página web como respaldo.
 *
 * Uso:
 *  - Agregar palabra:  GET /exec?text=<palabra o frase>   (si ya existe, suma 1 al conteo)
 *  - Listar palabras:  GET /exec?action=list
 *  - Borrar palabra:   GET /exec?action=delete&row=<numero de fila>
 *  - Guardar info generada:
 *      GET /exec?action=update&row=<n>&definition=<...>&examples=<...>&definitionEs=<...>
 *          &examplesEs=<...>&isVerb=<true|false>&baseForm=<...>&tenses=<...>
 *  - Registrar repaso: GET /exec?action=review&row=<n>&result=know|dontknow
 */

const SHEET_ID = '1-S-yUPdSiHzM40Ff7WcXlygHVeZ2sl5IZEmNEpUS26c';
const HEADERS = [
  'Primera vez', 'Palabra', 'Veces', 'Última vez', 'Definición', 'Ejemplos',
  'Nivel', 'Última revisión', 'Traducción', 'Ejemplos traducidos',
  'Es verbo', 'Forma base', 'Tiempos verbales'
];
const MAX_LEVEL = 5;
const TENSE_KEYS = ['presente', 'pasado', 'participio', 'gerundio'];

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

function tensesToString_(tenses) {
  if (!tenses) return '';
  return TENSE_KEYS.map(k => k + ':' + (tenses[k] || '')).join(';');
}

function tensesFromString_(str) {
  if (!str) return null;
  const result = {};
  String(str).split(';').forEach(part => {
    const idx = part.indexOf(':');
    if (idx === -1) return;
    result[part.slice(0, idx)] = part.slice(idx + 1);
  });
  return result;
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

  sheet.appendRow([now, trimmed, 1, now, '', '', 0, '', '', '', '', '', '']);
  return jsonResponse_({ ok: true, saved: trimmed, count: 1 });
}

function listWords_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse_({ ok: true, words: [] });

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const words = data.map((row, i) => ({
    row: i + 2,
    firstDate: row[0] instanceof Date ? row[0].toISOString() : String(row[0]),
    text: row[1],
    count: row[2] || 1,
    lastDate: row[3] instanceof Date ? row[3].toISOString() : String(row[3] || row[0]),
    definition: row[4] || '',
    examples: row[5] ? String(row[5]).split(' | ') : [],
    level: row[6] || 0,
    lastReview: row[7] instanceof Date ? row[7].toISOString() : (row[7] || ''),
    definitionEs: row[8] || '',
    examplesEs: row[9] ? String(row[9]).split(' | ') : [],
    isVerb: row[10] === true || row[10] === 'true',
    baseForm: row[11] || '',
    tenses: tensesFromString_(row[12])
  })).sort((a, b) => new Date(b.lastDate) - new Date(a.lastDate));

  return jsonResponse_({ ok: true, words });
}

function deleteWord_(sheet, rowParam) {
  const row = parseInt(rowParam, 10);
  if (!row || row < 2) return jsonResponse_({ ok: false, error: 'Invalid row' });
  sheet.deleteRow(row);
  return jsonResponse_({ ok: true });
}

function updateWordInfo_(sheet, rowParam, definition, examplesParam, definitionEs, examplesEsParam, isVerbParam, baseForm, tensesParam) {
  const row = parseInt(rowParam, 10);
  if (!row || row < 2) return jsonResponse_({ ok: false, error: 'Invalid row' });

  sheet.getRange(row, 5).setValue(definition || '');
  sheet.getRange(row, 6).setValue(examplesParam || '');
  sheet.getRange(row, 9).setValue(definitionEs || '');
  sheet.getRange(row, 10).setValue(examplesEsParam || '');
  sheet.getRange(row, 11).setValue(isVerbParam === 'true');
  sheet.getRange(row, 12).setValue(baseForm || '');
  sheet.getRange(row, 13).setValue(tensesParam || '');
  return jsonResponse_({ ok: true });
}

function generateWordInfoOnce_(word) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('No GEMINI_API_KEY set');

  const prompt = 'Analyze this English word or phrase: "' + word + '". ' +
    'Determine if it is a verb (in any tense or form). ' +
    'If it is a verb: set isVerb to true, give its base/infinitive form (no "to"), ' +
    'and write the definition and 2 example sentences using the PRESENT TENSE of that verb, ' +
    'even if the original word was in another tense. Also give its conjugation: ' +
    'presente (e.g. "run / runs"), pasado, participio (past participle), gerundio (-ing form). ' +
    'If it is NOT a verb: set isVerb to false, baseForm equal to the original word, ' +
    'tenses as null, and give a normal definition and 2 example sentences for it as-is. ' +
    'Also give the Spanish translation of the definition and of each example sentence. ' +
    'Respond ONLY with valid JSON, no markdown formatting, in exactly this shape: ' +
    '{"isVerb":true,"baseForm":"...","definition":"...","examples":["...","..."],' +
    '"definitionEs":"...","examplesEs":["...","..."],' +
    '"tenses":{"presente":"...","pasado":"...","participio":"...","gerundio":"..."}}';

  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + apiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true
    }
  );
  const data = JSON.parse(res.getContentText());
  if (!data.candidates || !data.candidates[0]) {
    throw new Error((data.error && data.error.message) || 'Sin respuesta de Gemini');
  }
  const raw = data.candidates[0].content.parts[0].text;
  const cleaned = raw.replace(/```json|```/g, '').trim();
  const parsed = JSON.parse(cleaned);
  return {
    definition: parsed.definition || '',
    examples: parsed.examples || [],
    definitionEs: parsed.definitionEs || '',
    examplesEs: parsed.examplesEs || [],
    isVerb: !!parsed.isVerb,
    baseForm: parsed.baseForm || word,
    tenses: parsed.isVerb ? parsed.tenses : null
  };
}

function generateWordInfo_(word, attempts) {
  attempts = attempts || 3;
  let lastError;
  for (let i = 0; i < attempts; i++) {
    try {
      return generateWordInfoOnce_(word);
    } catch (err) {
      lastError = err;
      if (i < attempts - 1) Utilities.sleep(1500 * (i + 1));
    }
  }
  throw lastError;
}

/**
 * Corre en un trigger de tiempo (no como Web App), así que sí tiene permiso
 * para hacer llamadas externas (UrlFetchApp) sin la restricción de acceso anónimo.
 * Procesa hasta MAX_PER_RUN palabras por ejecución para no exceder el tiempo límite.
 */
function processPendingWords() {
  const MAX_PER_RUN = 5;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  let processed = 0;

  for (let i = 0; i < data.length && processed < MAX_PER_RUN; i++) {
    const row = i + 2;
    const word = data[i][1];
    const definition = data[i][4];
    const definitionEs = data[i][8];

    if (!word || (definition && definitionEs)) continue;

    try {
      const info = generateWordInfo_(word);
      sheet.getRange(row, 5).setValue(info.definition);
      sheet.getRange(row, 6).setValue(info.examples.join(' | '));
      sheet.getRange(row, 9).setValue(info.definitionEs);
      sheet.getRange(row, 10).setValue(info.examplesEs.join(' | '));
      sheet.getRange(row, 11).setValue(info.isVerb);
      sheet.getRange(row, 12).setValue(info.baseForm);
      sheet.getRange(row, 13).setValue(tensesToString_(info.tenses));
    } catch (err) {
      // Se reintenta en la próxima ejecución del trigger.
    }

    processed++;
  }
}

/**
 * Ejecutar UNA vez manualmente desde el editor para instalar el trigger.
 * Vuelve a ejecutarse sola cada 5 minutos después de esto.
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processPendingWords') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('processPendingWords')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function reviewWord_(sheet, rowParam, result) {
  const row = parseInt(rowParam, 10);
  if (!row || row < 2) return jsonResponse_({ ok: false, error: 'Invalid row' });

  const levelCell = sheet.getRange(row, 7);
  const currentLevel = Number(levelCell.getValue()) || 0;
  const newLevel = result === 'know' ? Math.min(currentLevel + 1, MAX_LEVEL) : 0;

  levelCell.setValue(newLevel);
  sheet.getRange(row, 8).setValue(new Date());
  return jsonResponse_({ ok: true, level: newLevel });
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
    return updateWordInfo_(
      sheet,
      e.parameter.row,
      e.parameter.definition,
      e.parameter.examples,
      e.parameter.definitionEs,
      e.parameter.examplesEs,
      e.parameter.isVerb,
      e.parameter.baseForm,
      e.parameter.tenses
    );
  }

  if (action === 'review') {
    return reviewWord_(sheet, e.parameter.row, e.parameter.result);
  }

  const text = e.parameter.text;
  if (text && text.trim()) {
    return addWord_(sheet, text);
  }

  return jsonResponse_({ ok: false, error: 'No text provided' });
}
