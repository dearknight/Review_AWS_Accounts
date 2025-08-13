/************************************************************
 * AWS Org → Gmail → Google Sheets (Unificado)
 * - Lee correos con Subject: "Reporte Cuentas - <Cliente>"
 * - Cuerpo CSV: "AccountId,AccountName\n111111111111,Management\n..."
 * - Crea/actualiza pestaña <Cliente> con [AccountName, AccountId]
 * - Etiqueta hilos procesados para no re-leerlos
 * - Trigger diario que solo ejecuta en el 2º lunes de cada mes
 ************************************************************/

/** ===================== CONFIG ===================== **/
const SPREADSHEET_ID = 'PON_AQUI_TU_SPREADSHEET_ID'; // <-- CAMBIA ESTO
const SUBJECT_PREFIX  = 'Reporte Cuentas - ';
const PROCESSED_LABEL = 'processed/aws-org-reporter';

/** Query amplia para depuración; puedes endurecerla luego.
 *  Ejemplos para endurecer:
 *    - añadir from:no-reply@sns.amazonaws.com
 *    - añadir newer_than:30d (ventana)
 *    - añadir in:inbox
 */
function buildQuery() {
  return `subject:"${SUBJECT_PREFIX}" -label:${PROCESSED_LABEL}`;
}

/** ===================== MENÚ ===================== **/
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('AWS Reporter')
      .addItem('Procesar correos ahora', 'processAccountReports')
      .addItem('Debug: procesar el último', 'processLatestForDebug')
      .addSeparator()
      .addItem('Crear trigger (2º lunes/mes 09:00)', 'createDailyTriggerForSecondWeek')
      .addToUi();
  } catch (e) {
    console.log('onOpen error:', e);
  }
}

/** ===================== VALIDACIÓN SHEET ===================== **/
function getSpreadsheet() {
  if (!SPREADSHEET_ID || /PON_AQUI_TU_SPREADSHEET_ID/i.test(SPREADSHEET_ID)) {
    throw new Error('SPREADSHEET_ID no está configurado. Reemplaza PON_AQUI_TU_SPREADSHEET_ID por el ID real del Google Sheet.');
  }
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error('No se pudo abrir el Spreadsheet con ese ID. Verifica permisos y el ID de la URL. Detalle: ' + e);
  }
}

/** ===================== FLUJO NORMAL ===================== **/
function processAccountReports() {
  const query = buildQuery();
  console.log('Query Gmail:', query);

  const threads = GmailApp.search(query, 0, 50);
  console.log('Hilos encontrados:', threads.length);
  if (!threads.length) return;

  const processedLabel = ensureLabel(PROCESSED_LABEL);
  const ss = getSpreadsheet();

  threads.forEach((thread, tIdx) => {
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];
    const subject = (msg.getSubject() || '').trim();
    console.log(`[Thread ${tIdx}] Subject: "${subject}"`);

    if (!subject.startsWith(SUBJECT_PREFIX)) {
      console.log('Subject no coincide con el prefijo, se omite.');
      return;
    }

    const clientName = subject.substring(SUBJECT_PREFIX.length).trim();
    if (!clientName) {
      console.log('No se pudo extraer clientName del subject, se omite.');
      return;
    }

    const text = safeGetBodyText(msg);
    if (!text) {
      console.log('Cuerpo vacío/no legible, se omite.');
      return;
    }

    const rows = parseCsvIdName(text); // [[AccountId, AccountName], ...]
    console.log(`Filas parseadas para "${clientName}":`, rows.length);

    if (!rows.length) {
      console.log('No se parseó ninguna fila. Primeros 200 chars cuerpo:', text.substring(0, 200));
      return;
    }

    const affected = syncSheetForClient(ss, clientName, rows);
    console.log(`Sync "${clientName}" → add:${affected.added} upd:${affected.updated} del:${affected.deleted} total:${affected.total}`);

    // Marca hilo como procesado (opcionalmente archivar)
    thread.addLabel(processedLabel);
    // thread.moveToArchive();
  });
}

/** ===================== DEBUG RÁPIDO ===================== **/
function processLatestForDebug() {
  const q = buildQuery();
  console.log('[DEBUG] Query:', q);

  const threads = GmailApp.search(q, 0, 1);
  console.log('[DEBUG] Hilos encontrados:', threads.length);
  if (!threads.length) { console.log('[DEBUG] Nada que procesar.'); return; }

  const thread = threads[0];
  const messages = thread.getMessages();
  const msg = messages[messages.length - 1];

  const subject = (msg.getSubject() || '').trim();
  console.log('[DEBUG] Subject:', subject);

  if (!subject.startsWith(SUBJECT_PREFIX)) { console.log('[DEBUG] Prefijo no coincide.'); return; }

  const clientName = subject.substring(SUBJECT_PREFIX.length).trim();
  if (!clientName) { console.log('[DEBUG] clientName vacío.'); return; }
  console.log('[DEBUG] ClientName:', clientName);

  const text = safeGetBodyText(msg) || '';
  console.log('[DEBUG] Body (primeros 300 chars):', text.substring(0, 300).replace(/\n/g, '\\n'));

  const rows = parseCsvIdName(text);
  console.log('[DEBUG] Filas parseadas:', rows.length, rows);
  if (!rows.length) return;

  const ss = getSpreadsheet();
  const affected = syncSheetForClient(ss, clientName, rows);
  console.log(`[DEBUG] Sync "${clientName}" → add:${affected.added} upd:${affected.updated} del:${affected.deleted} total:${affected.total}`);

  const processedLabel = ensureLabel(PROCESSED_LABEL);
  thread.addLabel(processedLabel);
  console.log('[DEBUG] Marcado con label:', PROCESSED_LABEL);
}

/** ===================== UTILIDADES GMAIL ===================== **/
function ensureLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

/** Intenta texto plano; si no, convierte HTML a texto básico. */
function safeGetBodyText(msg) {
  try {
    const plain = msg.getPlainBody();
    if (plain && plain.trim()) return plain;
  } catch (e) {
    console.log('getPlainBody error:', e);
  }
  try {
    const html = msg.getBody() || '';
    const stripped = html
      .replace(/<style[\\s\\S]*?<\\/style>/gi, '')
      .replace(/<script[\\s\\S]*?<\\/script>/gi, '')
      .replace(/<br\\s*\\/?>(?=.)/gi, '\\n')
      .replace(/<\\/p>/gi, '\\n')
      .replace(/<\\/div>/gi, '\\n')
      .replace(/<[^>]+>/g, '')
      .replace(/\\&nbsp;/g, ' ')
      .replace(/\\&amp;/g, '&')
      .replace(/\\&lt;/g, '<')
      .replace(/\\&gt;/g, '>');
    return stripped;
  } catch (e) {
    console.log('getBody/stripped error:', e);
    return '';
  }
}

/** ===================== PARSER CSV ===================== **/
/**
 * Espera cabecera "AccountId,AccountName" (o tolera sin cabecera).
 * Devuelve [[AccountId, AccountName], ...]
 */
function parseCsvIdName(text) {
  if (!text) return [];
  const lines = text.split(/\\r?\\n/).map(l => l.trim()).filter(Boolean);

  // Intenta ubicar el encabezado exacto
  let startIndex = lines.findIndex(l => /^AccountId\\s*,\\s*AccountName$/i.test(l));
  const data = [];

  const parseLine = (line) => {
    const parts = line.split(',');
    if (parts.length < 2) return null;
    const id   = parts[0].trim();
    const name = parts.slice(1).join(',').trim(); // permite nombre con comas si llegaran
    if (!/^\\d{12}$/.test(id)) return null;       // valida 12 dígitos
    return [id, name];
  };

  if (startIndex >= 0) {
    for (let i = startIndex + 1; i < lines.length; i++) {
      const row = parseLine(lines[i]);
      if (row) data.push(row);
    }
  } else {
    // Sin encabezado: intenta todo
    lines.forEach(line => {
      const row = parseLine(line);
      if (row) data.push(row);
    });
  }

  return data;
}

/** ===================== SYNC A SHEET ===================== **/
/**
 * Sincroniza la pestaña del cliente:
 * - Crea si no existe, con encabezado [AccountName, AccountId]
 * - Agrega cuentas nuevas
 * - Actualiza nombres cambiados
 * - Elimina cuentas que ya no vienen
 * - Ordena por AccountName asc
 * Retorna {added, updated, deleted, total}
 */
function syncSheetForClient(ss, sheetName, incomingIdNamePairs) {
  console.log(`Sincronizando hoja "${sheetName}" con ${incomingIdNamePairs.length} registros entrantes…`);

  // incoming: [[AccountId, AccountName], ...] → Map id->name
  const incomingMap = new Map();
  incomingIdNamePairs.forEach(([id, name]) => incomingMap.set(id, name));

  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, 2).setValues([['AccountName', 'AccountId']]);
    sh.setFrozenRows(1);
  }

  // Lee existentes (sin encabezado)
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  let existing = []; // [ [name, id], ... ]
  if (lastRow >= 2 && lastCol >= 2) {
    existing = sh.getRange(2, 1, lastRow - 1, 2).getValues()
      .filter(r => r[0] && r[1]);
  }

  // Índice actual por id
  const currentById = new Map();
  existing.forEach(([name, id], i) => currentById.set(id, { name, row: i + 2 }));

  // 1) Borrar ausentes
  const toRemoveRows = [];
  currentById.forEach((val, id) => { if (!incomingMap.has(id)) toRemoveRows.push(val.row); });
  toRemoveRows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));

  // 2) Agregar nuevos
  const toAdd = [];
  incomingMap.forEach((name, id) => { if (!currentById.has(id)) toAdd.push([name, id]); });
  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
  }

  // 3) Actualizar nombres cambiados
  const lr = sh.getLastRow();
  let current = [];
  if (lr >= 2) current = sh.getRange(2, 1, lr - 1, 2).getValues();
  const newIndexById = new Map();
  current.forEach(([name, id], i) => newIndexById.set(id, { name, row: i + 2 }));

  let updated = 0;
  incomingMap.forEach((name, id) => {
    const found = newIndexById.get(id);
    if (found && found.name !== name) {
      sh.getRange(found.row, 1).setValue(name);
      updated++;
    }
  });

  // 4) Ordenar por AccountName asc
  const rowsCount = sh.getLastRow();
  if (rowsCount > 2) sh.getRange(2, 1, rowsCount - 1, 2).sort({ column: 1, ascending: true });

  // Formato
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 2);
  const range = sh.getRange(1, 1, sh.getLastRow(), 2);
  range.setBorder(true, true, true, true, true, true);

  const total = sh.getLastRow() - 1;
  return { added: toAdd.length, updated, deleted: toRemoveRows.length, total };
}

/** ===================== TRIGGER: 2ª semana del mes ===================== **/
/**
 * Crea un trigger diario a la hora indicada (UTC de tu cuenta de Google)
 * que llama a scheduledRunner(). Este runner solo ejecuta el flujo
 * si HOY es lunes de la segunda semana del mes (días 8–14).
 */
function createDailyTriggerForSecondWeek() {
  // Limpieza opcional de triggers previos del mismo handler
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'scheduledRunner') {
      ScriptApp.deleteTrigger(tr);
    }
  });

  // Ajusta la hora (24h). Ej: 9 = 09:00
  ScriptApp.newTrigger('scheduledRunner')
    .timeBased()
    .everyDays(1)  // ejecutar todos los días
    .atHour(9)     // a las 09:00
    .create();

  console.log('Trigger creado: scheduledRunner → diario 09:00. Ejecutará solo en 2º lunes del mes.');
}

/**
 * Runner llamado por el trigger diario.
 * Solo ejecuta processAccountReports() si es lunes de la segunda semana del mes (8–14).
 */
function scheduledRunner() {
  const now = new Date();                   // hora local de tu cuenta
  const dow = now.getDay();                 // 0=Dom, 1=Lun, ... 6=Sáb
  const dom = now.getDate();                // día del mes (1..31)
  const isSecondWeek = dom >= 8 && dom <= 14;
  const isMonday = dow === 1;

  if (!(isSecondWeek && isMonday)) {
    console.log('No es el lunes de la segunda semana. Saliendo.');
    return;
  }

  processAccountReports();
}
