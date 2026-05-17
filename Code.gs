// ============================================================
// GESTIÓN DE TARJETAS - Club Deportivo Mitre
// Google Apps Script - Backend v2
// ============================================================
// Todas las operaciones usan GET para compatibilidad CORS
// desde GitHub Pages.
// Desplegar como: Ejecutar como "Yo" | Acceso "Cualquiera"
// ============================================================

const SHEET_REG = 'Registros';
const SHEET_CFG = 'Config';
const SHEET_MOV = 'Movimientos';

// --- Respuesta JSON estándar ---
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Handler principal (todo vía GET) ---
function doGet(e) {
  try {
    const action  = e.parameter.action;
    const rawData = e.parameter.data;
    const data    = rawData ? JSON.parse(rawData) : null;

    switch (action) {
      case 'getConfig':          return jsonResponse(getConfig());
      case 'saveConfig':         return jsonResponse(saveConfig(data));
      case 'getRegistrations':   return jsonResponse(getRegistrations());
      case 'addRegistration':    return jsonResponse(addRegistration(data));
      case 'updateRegistration': return jsonResponse(updateRegistration(data.id, data.fields));
      case 'deleteRegistration': return jsonResponse(deleteRegistration(data.id));
      case 'getMovimientos':     return jsonResponse(getMovimientos());
      case 'addMovimiento':      return jsonResponse(addMovimiento(data));
      default:
        return jsonResponse({ success: false, error: 'Acción no reconocida: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ============================================================
// CONFIG
// ============================================================

function getConfig() {
  const sheet   = getOrCreateSheet(SHEET_CFG);
  const numRows = sheet.getLastRow();

  if (numRows === 0) {
    return { success: true, config: { eventName: '', eventDate: '', priceTypes: [] } };
  }

  const data   = sheet.getRange(1, 1, numRows, 2).getValues();
  const config = {};
  data.forEach(function(row) {
    if (row[0]) config[row[0]] = row[1];
  });

  if (config.priceTypes) {
    try { config.priceTypes = JSON.parse(config.priceTypes); }
    catch (e) { config.priceTypes = []; }
  } else {
    config.priceTypes = [];
  }

  return { success: true, config: config };
}

function saveConfig(config) {
  const sheet = getOrCreateSheet(SHEET_CFG);
  sheet.clearContents();

  const rows = [
    ['eventName',  config.eventName  || ''],
    ['eventDate',  config.eventDate  || ''],
    ['priceTypes', JSON.stringify(config.priceTypes || [])]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  return { success: true };
}

// ============================================================
// REGISTROS
// ============================================================

const HEADERS = [
  'id', 'timestamp', 'nombre', 'genero', 'cantidad',
  'tipoPrecio', 'precioUnitario', 'totalDebe',
  'montoPagado', 'estado', 'metodoPago', 'notas'
];

function ensureRegHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#1a5276').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } else {
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (existing.join(',') !== HEADERS.join(',')) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold').setBackground('#1a5276').setFontColor('#ffffff');
    }
  }
}

function getRegistrations() {
  const sheet   = getOrCreateSheet(SHEET_REG);
  ensureRegHeaders(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, registrations: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  const registrations = data
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) {
      const obj = {};
      HEADERS.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });

  return { success: true, registrations: registrations };
}

function addRegistration(reg) {
  const sheet = getOrCreateSheet(SHEET_REG);
  ensureRegHeaders(sheet);

  const id        = generateId();
  const timestamp = new Date().toISOString();

  sheet.appendRow([
    id,
    timestamp,
    reg.nombre          || '',
    reg.genero          || '',
    Number(reg.cantidad)        || 1,
    reg.tipoPrecio      || '',
    Number(reg.precioUnitario)  || 0,
    Number(reg.totalDebe)       || 0,
    Number(reg.montoPagado)     || 0,
    reg.estado          || 'pendiente',
    reg.metodoPago      || '',
    reg.notas           || ''
  ]);

  return { success: true, id: id };
}

function updateRegistration(id, fields) {
  const sheet   = getOrCreateSheet(SHEET_REG);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No hay registros' };

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowNum = i + 2;
      Object.keys(fields).forEach(function(key) {
        const colIndex = HEADERS.indexOf(key);
        if (colIndex >= 0) {
          sheet.getRange(rowNum, colIndex + 1).setValue(fields[key]);
        }
      });
      return { success: true };
    }
  }

  return { success: false, error: 'Registro no encontrado: ' + id };
}

function deleteRegistration(id) {
  const sheet   = getOrCreateSheet(SHEET_REG);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No hay registros' };

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }

  return { success: false, error: 'Registro no encontrado: ' + id };
}

// ============================================================
// MOVIMIENTOS
// ============================================================
// Un movimiento = un cobro puntual registrado en el evento.
// tipo: 'completo' | 'tarjetas' | 'seña'
// cantidadTarjetas: cuántas tarjetas completas se pagaron
//                   (0 si fue una seña libre)
// ============================================================

const MOV_HEADERS = [
  'id', 'timestamp', 'registrationId', 'nombre',
  'tipoPrecio', 'cantidadTarjetas', 'montoCobrado',
  'metodoPago', 'notas', 'tipo'
];

function ensureMovHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MOV_HEADERS);
    sheet.getRange(1, 1, 1, MOV_HEADERS.length)
      .setFontWeight('bold').setBackground('#1e8449').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

function getMovimientos() {
  const sheet   = getOrCreateSheet(SHEET_MOV);
  ensureMovHeaders(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, movimientos: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, MOV_HEADERS.length).getValues();

  const movimientos = data
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) {
      const obj = {};
      MOV_HEADERS.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    })
    .reverse(); // más nuevo primero

  return { success: true, movimientos: movimientos };
}

function addMovimiento(mov) {
  const sheet = getOrCreateSheet(SHEET_MOV);
  ensureMovHeaders(sheet);

  const id        = generateId();
  const timestamp = new Date().toISOString();

  sheet.appendRow([
    id,
    timestamp,
    mov.registrationId    || '',
    mov.nombre            || '',
    mov.tipoPrecio        || '',
    Number(mov.cantidadTarjetas) || 0,
    Number(mov.montoCobrado)     || 0,
    mov.metodoPago        || '',
    mov.notas             || '',
    mov.tipo              || 'pago'   // 'completo' | 'tarjetas' | 'seña'
  ]);

  return { success: true, id: id };
}

// ============================================================
// UTILIDADES
// ============================================================

function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function generateId() {
  const ts   = new Date().getTime().toString(36);
  const rand = Math.random().toString(36).substring(2, 7);
  return ts + rand;
}

// ============================================================
// TEST - ejecutar manualmente para verificar
// ============================================================
function testSetup() {
  const config = {
    eventName:  'Peña Folklórica - Prueba',
    eventDate:  '2025-06-01',
    priceTypes: [
      { name: 'General',  amount: 5000 },
      { name: 'Especial', amount: 4000 },
      { name: 'Menor',    amount: 2500 }
    ]
  };
  Logger.log('saveConfig: ' + JSON.stringify(saveConfig(config)));
  Logger.log('getConfig:  ' + JSON.stringify(getConfig()));

  const reg = {
    nombre: 'Carlos González', cantidad: 10,
    tipoPrecio: 'General', precioUnitario: 5000,
    totalDebe: 50000, montoPagado: 5000,
    estado: 'seña', metodoPago: 'efectivo', notas: ''
  };
  const addResult = addRegistration(reg);
  Logger.log('addReg: ' + JSON.stringify(addResult));

  const mov = {
    registrationId: addResult.id, nombre: 'Carlos González',
    tipoPrecio: 'General', cantidadTarjetas: 1,
    montoCobrado: 5000, metodoPago: 'efectivo',
    notas: '', tipo: 'tarjetas'
  };
  Logger.log('addMov: ' + JSON.stringify(addMovimiento(mov)));
  Logger.log('getMov: ' + JSON.stringify(getMovimientos()));
}
