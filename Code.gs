// ============================================================
// GESTIÓN DE TARJETAS - Club Deportivo Mitre
// Google Apps Script - Backend
// ============================================================
// Todas las operaciones usan GET para compatibilidad con CORS
// desde GitHub Pages.
// Desplegar como: Ejecutar como "Yo" | Acceso "Cualquiera"
// ============================================================

const SHEET_REG = 'Registros';
const SHEET_CFG = 'Config';

// --- Respuesta JSON estándar ---
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- Handler principal (todo vía GET) ---
function doGet(e) {
  try {
    const action = e.parameter.action;
    const rawData = e.parameter.data;
    const data = rawData ? JSON.parse(rawData) : null;

    switch (action) {
      case 'getConfig':
        return jsonResponse(getConfig());
      case 'saveConfig':
        return jsonResponse(saveConfig(data));
      case 'getRegistrations':
        return jsonResponse(getRegistrations());
      case 'addRegistration':
        return jsonResponse(addRegistration(data));
      case 'updateRegistration':
        return jsonResponse(updateRegistration(data.id, data.fields));
      case 'deleteRegistration':
        return jsonResponse(deleteRegistration(data.id));
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
  const sheet = getOrCreateSheet(SHEET_CFG);
  const numRows = sheet.getLastRow();

  if (numRows === 0) {
    // Hoja vacía: retornar config por defecto
    return { success: true, config: { eventName: '', eventDate: '', priceTypes: [] } };
  }

  const data = sheet.getRange(1, 1, numRows, 2).getValues();
  const config = {};
  data.forEach(function(row) {
    if (row[0]) config[row[0]] = row[1];
  });

  // Parsear tipos de precio desde JSON
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
    ['eventName', config.eventName || ''],
    ['eventDate', config.eventDate || ''],
    ['priceTypes', JSON.stringify(config.priceTypes || [])]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  return { success: true };
}

// ============================================================
// REGISTROS
// ============================================================

// Columnas del sheet de registros
const HEADERS = [
  'id', 'timestamp', 'nombre', 'cantidad',
  'tipoPrecio', 'precioUnitario', 'totalDebe',
  'montoPagado', 'estado', 'metodoPago', 'notas'
];

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    // Formatear encabezados
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#34495e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

function getRegistrations() {
  const sheet = getOrCreateSheet(SHEET_REG);
  ensureHeaders(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, registrations: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  const registrations = data
    .filter(function(row) { return row[0] !== ''; }) // filtrar filas vacías
    .map(function(row) {
      const obj = {};
      HEADERS.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });

  return { success: true, registrations: registrations };
}

function addRegistration(reg) {
  const sheet = getOrCreateSheet(SHEET_REG);
  ensureHeaders(sheet);

  const id = generateId();
  const timestamp = new Date().toISOString();

  const row = [
    id,
    timestamp,
    reg.nombre || '',
    Number(reg.cantidad) || 1,
    reg.tipoPrecio || '',
    Number(reg.precioUnitario) || 0,
    Number(reg.totalDebe) || 0,
    Number(reg.montoPagado) || 0,
    reg.estado || 'pendiente',   // 'pagado' | 'seña' | 'pendiente'
    reg.metodoPago || '',
    reg.notas || ''
  ];

  sheet.appendRow(row);
  return { success: true, id: id };
}

function updateRegistration(id, fields) {
  const sheet = getOrCreateSheet(SHEET_REG);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: false, error: 'No hay registros' };

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowNum = i + 2; // +1 por header, +1 por base-1
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
  const sheet = getOrCreateSheet(SHEET_REG);
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
// UTILIDADES
// ============================================================

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function generateId() {
  const ts = new Date().getTime().toString(36);
  const rand = Math.random().toString(36).substring(2, 7);
  return ts + rand;
}

// ============================================================
// TEST - ejecutar manualmente para verificar
// ============================================================
function testSetup() {
  const config = {
    eventName: 'Peña Folklórica - Prueba',
    eventDate: '2025-06-01',
    priceTypes: [
      { name: 'General', amount: 5000 },
      { name: 'Especial', amount: 4000 },
      { name: 'Menor', amount: 2500 }
    ]
  };
  Logger.log(saveConfig(config));
  Logger.log(getConfig());

  const reg = {
    nombre: 'Juan Pérez',
    cantidad: 2,
    tipoPrecio: 'General',
    precioUnitario: 5000,
    totalDebe: 10000,
    montoPagado: 10000,
    estado: 'pagado',
    metodoPago: 'efectivo',
    notas: ''
  };
  Logger.log(addRegistration(reg));
  Logger.log(getRegistrations());
}
