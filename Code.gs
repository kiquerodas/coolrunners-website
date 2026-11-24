/**
 * CoolRunners Guatemala — Registro Histórico de Maratones
 * Google Apps Script Web App. Deploy as: Execute as me; access Anyone.
 * The linked sheet is private unless you explicitly share it.
 */
const CONFIG = {
  SHEET_NAME: 'Maratones',
  MAX_PER_SUBMISSION: 100,
  HEADERS: [
    'Fecha de registro', 'ID de envío', 'Nombre', 'Apellido', 'Correo',
    'Maratón', 'Ciudad', 'País', 'Año', 'Fecha de carrera',
    'Tiempo oficial (HH:MM:SS)', 'Tiempo en segundos',
    'Enlace de resultados', 'Permiso para publicar', 'Aceptación de privacidad'
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CoolRunners | Registro Histórico de Maratones')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once in the Apps Script editor to provision the response spreadsheet. */
function configurarRegistro() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let spreadsheet;
  if (id) {
    spreadsheet = SpreadsheetApp.openById(id);
  } else {
    spreadsheet = SpreadsheetApp.create('CoolRunners — Registro Histórico de Maratones (respuestas)');
    props.setProperty('SHEET_ID', spreadsheet.getId());
  }
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length)
      .setBackground('#174A86').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('J:J').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('L:L').setNumberFormat('0');
    sheet.setColumnWidths(1, CONFIG.HEADERS.length, 160);
    sheet.setColumnWidth(6, 230);
    sheet.setColumnWidth(13, 260);
  }
  Logger.log('HOJA DE RESPUESTAS: ' + spreadsheet.getUrl());
  Logger.log('ID DE HOJA: ' + spreadsheet.getId());
  return spreadsheet.getUrl();
}

function registrarMaratones(payload) {
  // HTML Service exposes this server-side function to the web app.
  if (!payload || typeof payload !== 'object') throw new Error('Datos inválidos.');
  const p = payload.participante || {};
  const nombre = saneText(p.nombre, 80);
  const apellido = saneText(p.apellido, 80);
  const correo = saneText(p.correo, 160).toLowerCase();
  if (!nombre || !apellido || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) {
    throw new Error('Ingrese nombre, apellido y un correo válido.');
  }
  if (p.aceptaPrivacidad !== true) {
    throw new Error('Necesitamos su autorización para guardar estos datos.');
  }
  const permiso = p.publicar === true ? 'Sí' : 'No';
  if (!Array.isArray(payload.maratones) || !payload.maratones.length ||
      payload.maratones.length > CONFIG.MAX_PER_SUBMISSION) {
    throw new Error('Ingrese de 1 a ' + CONFIG.MAX_PER_SUBMISSION + ' maratones.');
  }
  const currentYear = new Date().getFullYear();
  const events = payload.maratones.map((m, i) => {
    const carrera = saneText(m.carrera, 180);
    const ciudad = saneText(m.ciudad, 120);
    const pais = saneText(m.pais, 120);
    const year = Number(m.anio);
    const tiempo = saneText(m.tiempo, 12);
    const dateText = saneText(m.fecha, 10);
    const link = saneText(m.enlace, 500);
    if (!carrera || !ciudad || !pais || !Number.isInteger(year) || year < 1896 || year > currentYear) {
      throw new Error('Revise nombre, ciudad, país y año en el maratón #' + (i + 1) + '.');
    }
    if (!/^\d{1,2}:\d{2}:\d{2}$/.test(tiempo)) {
      throw new Error('Use HH:MM:SS para el tiempo del maratón #' + (i + 1) + '.');
    }
    const [h, min, sec] = tiempo.split(':').map(Number);
    if (h > 23 || min > 59 || sec > 59 || (h === 0 && min === 0 && sec === 0)) {
      throw new Error('Tiempo inválido en el maratón #' + (i + 1) + '.');
    }
    const normalizedTime = String(h).padStart(2, '0') + ':' + String(min).padStart(2, '0') + ':' + String(sec).padStart(2, '0');
    let fecha = '';
    if (dateText) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) throw new Error('Fecha inválida en el maratón #' + (i + 1) + '.');
      const [y, mo, d] = dateText.split('-').map(Number);
      const test = new Date(y, mo - 1, d);
      if (test.getFullYear() !== y || test.getMonth() !== mo - 1 || test.getDate() !== d || y !== year || test > new Date()) {
        throw new Error('La fecha no coincide con el año en el maratón #' + (i + 1) + '.');
      }
      fecha = test;
    }
    if (link && !/^https?:\/\/\S+$/i.test(link)) {
      throw new Error('El enlace de resultados debe iniciar con https:// en el maratón #' + (i + 1) + '.');
    }
    return { carrera, ciudad, pais, year, fecha, tiempo: normalizedTime, segundos: h*3600+min*60+sec, link };
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) throw new Error('El administrador debe ejecutar configurarRegistro() primero.');
    const sheet = SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error('No se encontró la hoja de registro.');
    const timestamp = new Date();
    const batch = Utilities.getUuid();
    const rows = events.map(m => [timestamp, batch, nombre, apellido, correo,
      m.carrera, m.ciudad, m.pais, m.year, m.fecha, m.tiempo, m.segundos,
      m.link, permiso, 'Sí']);
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, CONFIG.HEADERS.length).setValues(rows);
    return { ok: true, total: rows.length, mensaje: '¡Sus maratones quedaron registrados!' };
  } finally {
    lock.releaseLock();
  }
}

function saneText(value, maxLen) {
  const raw = String(value == null ? '' : value).trim().replace(/[\u0000-\u001f]/g, ' ');
  if (raw.length > maxLen) throw new Error('Uno de los campos supera el largo permitido.');
  // Prevent spreadsheet formula injection from user-supplied text.
  return /^[=+\-@]/.test(raw) ? "'" + raw : raw;
}


/** Receives submissions from a separately hosted CoolRunners website. */
function doPost(e) {
  let response;
  try {
    if (!e || !e.parameter || !e.parameter.payload) throw new Error('No se recibieron datos.');
    const data = JSON.parse(e.parameter.payload);
    response = registrarMaratones(data);
  } catch (err) {
    response = {ok:false, mensaje:String(err && err.message || err)};
  }
  // Posted into a hidden iframe. Notify the originating website after Sheets saves.
  // Data is not returned to any outside origin beyond the success/error status.
  const json = JSON.stringify(response).replace(/</g, '\\u003c');
  const html = '<!doctype html><html><body><script>window.parent.postMessage(' + json + ', "*");</' + 'script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('CoolRunners - Confirmación');
}
