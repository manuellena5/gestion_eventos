// ============================================================
// GESTIÓN DE TARJETAS - Club Deportivo Mitre
// Google Apps Script - Backend
// v1 (tarjetas) arriba · v2 (eventos, pedidos, contactos) al final del archivo
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
// Las acciones v2 (sección "V2" al final) responden {ok, ...}.
// Las acciones viejas siguen respondiendo {success, ...} hasta la fase 8.
function doGet(e) {
  const action  = e.parameter.action;
  if (V2_ACTIONS[action]) {
    return jsonResponse(runV2(action, e.parameter.data));
  }
  try {
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
    if (!row[0]) return;
    var val = row[1];
    if (val instanceof Date) {
      var y = val.getFullYear();
      var m = String(val.getMonth() + 1).padStart(2, '0');
      var d = String(val.getDate()).padStart(2, '0');
      val = y + '-' + m + '-' + d;
    }
    config[row[0]] = val;
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

  // Asignar ID automáticamente a filas que no lo tienen (registros manuales)
  const idColIndex = HEADERS.indexOf('id');
  const tsColIndex = HEADERS.indexOf('timestamp');
  for (var i = 0; i < data.length; i++) {
    if (data[i][idColIndex] === '' || data[i][idColIndex] === null || data[i][idColIndex] === undefined) {
      // Verificar que la fila tenga al menos el nombre para no asignar ID a filas vacías
      var nombreColIndex = HEADERS.indexOf('nombre');
      if (data[i][nombreColIndex] === '' || data[i][nombreColIndex] === null) continue;

      var newId = generateId();
      data[i][idColIndex] = newId;
      sheet.getRange(i + 2, idColIndex + 1).setValue(newId);

      // Si tampoco tiene timestamp, asignarlo
      if (data[i][tsColIndex] === '' || data[i][tsColIndex] === null) {
        var newTs = new Date().toISOString();
        data[i][tsColIndex] = newTs;
        sheet.getRange(i + 2, tsColIndex + 1).setValue(newTs);
      }
    }
  }

  const numericFields = ['cantidad', 'precioUnitario', 'totalDebe', 'montoPagado'];

  const registrations = data
    .filter(function(row) { return row[idColIndex] !== '' && row[idColIndex] !== null; })
    .map(function(row) {
      const obj = {};
      HEADERS.forEach(function(h, i) {
        var val = row[i];
        // Normalizar campos numéricos
        if (numericFields.indexOf(h) >= 0) {
          val = Number(val) || 0;
        }
        obj[h] = val;
      });
      // Valor por defecto para estado
      if (!obj['estado']) obj['estado'] = 'pendiente';
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


// ############################################################
// ############################################################
//  V2 — Eventos, productos, pedidos, pagos y contactos
// ############################################################
// ############################################################
//
// Todas las acciones reciben `data` (JSON) por GET y responden
// { ok: true, ... }  o  { ok: false, error: 'mensaje para mostrar' }.
// Toda escritura corre dentro de LockService.getScriptLock().
// ============================================================

// --- Esquema de hojas -----------------------------------------
// Tipos: text (formato texto plano), num, numOpt (vacío = sin valor),
//        bool, boolT (vacío = TRUE), date (aaaa-mm-dd), time (HH:mm), ts (fecha y hora)
const V2_SCHEMA = {
  Usuarios: {
    headers: ['id', 'nombre', 'pin', 'activo'],
    types:   { activo: 'boolT' }
  },
  Eventos: {
    headers: ['id', 'nombre', 'fecha', 'hora', 'lugar', 'estado', 'campos', 'tiposComensal', 'plantillaWhatsApp', 'creado'],
    types:   { fecha: 'date', hora: 'time', creado: 'ts' }
  },
  Productos: {
    headers: ['id', 'eventoId', 'nombre', 'precio', 'cupo', 'orden', 'activo', 'cuentaComensales'],
    types:   { precio: 'num', cupo: 'numOpt', orden: 'num', activo: 'boolT', cuentaComensales: 'bool' }
  },
  Pedidos: {
    headers: ['id', 'fecha', 'eventoId', 'evento', 'contactoId', 'nombre', 'apellido', 'telefono',
              'total', 'pagado', 'saldo', 'estado', 'entregado', 'retiro', 'nota', 'anoto',
              'anulado', 'anuladoPor', 'motivoAnulacion'],
    types:   { fecha: 'ts', total: 'num', pagado: 'num', saldo: 'num', entregado: 'bool', anulado: 'bool' }
  },
  Items: {
    headers: ['id', 'pedidoId', 'eventoId', 'productoId', 'producto', 'tipoComensal', 'cantidad', 'precioUnitario', 'subtotal'],
    types:   { cantidad: 'num', precioUnitario: 'num', subtotal: 'num' }
  },
  Pagos: {
    headers: ['id', 'fecha', 'pedidoId', 'eventoId', 'nombre', 'monto', 'medio', 'tipo', 'cobro', 'nota'],
    types:   { fecha: 'ts', monto: 'num' }
  },
  Contactos: {
    headers: ['id', 'nombre', 'apellido', 'telefono', 'telefonoNorm', 'primerEvento', 'ultimoEvento',
              'cantEventos', 'eventos', 'productos', 'creado', 'actualizado'],
    types:   { cantEventos: 'num', creado: 'ts', actualizado: 'ts' }
  }
};

const V2_CAMPOS_DEFAULT = { apellido: true, sena: true, nota: true, retiro: false, entregado: true, comensales: false };
const V2_TIPOS_DEFAULT  = ['Hombre', 'Mujer', 'Niño'];
const V2_MEDIOS         = ['efectivo', 'transferencia'];
const V2_SEP            = ' · ';

// --- Acciones ------------------------------------------------------
// write: true → corre bajo lock y valida el usuario
const V2_ACTIONS = {
  getInit:       { fn: v2GetInit },
  login:         { fn: v2Login },
  getEvento:     { fn: v2GetEvento },
  buscarContacto:{ fn: v2BuscarContacto },
  getContactos:  { fn: v2GetContactos },
  saveEvento:    { fn: v2SaveEvento,    write: true },
  copiarEvento:  { fn: v2CopiarEvento,  write: true },
  savePedido:    { fn: v2SavePedido,    write: true },
  updatePedido:  { fn: v2UpdatePedido,  write: true },
  addPago:       { fn: v2AddPago,       write: true },
  setEntregado:  { fn: v2SetEntregado,  write: true },
  anularPedido:  { fn: v2AnularPedido,  write: true }
};

// Error con mensaje para mostrar a la persona
function V2Error(msg, extra) {
  const e = new Error(msg);
  e.v2user = true;
  e.extra  = extra || null;
  return e;
}
function fail(msg, extra) { throw V2Error(msg, extra); }

function runV2(action, rawData) {
  const def = V2_ACTIONS[action];
  let lock = null;
  try {
    let data = {};
    if (rawData) {
      try { data = JSON.parse(rawData) || {}; }
      catch (e) { fail('No se pudo leer el pedido. Probá de nuevo.'); }
    }
    _v2Reset();
    if (def.write) {
      lock = LockService.getScriptLock();
      if (!lock.tryLock(25000)) { lock = null; fail('Hay mucha gente cargando a la vez. Probá de nuevo en un momento.'); }
      _v2HoldsLock = true;
      data._usuario = v2Auth(data);
    }
    const res = def.fn(data) || {};
    return Object.assign({ ok: true }, res);
  } catch (err) {
    if (err && err.v2user) return Object.assign({ ok: false, error: err.message }, err.extra || {});
    console.error(action + ': ' + (err && err.stack || err));
    return { ok: false, error: 'Algo salió mal. Probá de nuevo en un rato.' };
  } finally {
    if (lock) { try { SpreadsheetApp.flush(); } catch (e) {} lock.releaseLock(); }
    _v2HoldsLock = false;
  }
}

// ============================================================
// Acceso a hojas (lectura de una sola vez con getValues)
// ============================================================

let _v2Tables = {};
let _v2Ensured = {};   // hojas ya revisadas en esta ejecución
let _v2Sheets = {};
let _v2Ss = null;
let _v2Tz = null;
let _v2HoldsLock = false;
function _v2Reset() { _v2Tables = {}; }
function v2Tz() { if (!_v2Tz) _v2Tz = Session.getScriptTimeZone(); return _v2Tz; }

function v2Type(name, h) { return (V2_SCHEMA[name].types || {})[h] || 'text'; }

// Reintenta ante errores temporales de Google ("no ha podido acceder al documento", etc.)
// Solo se usa para lecturas y para crear hojas, nunca para agregar filas (no duplicar).
function v2Retry(fn) {
  for (let i = 0; ; i++) {
    try { return fn(); }
    catch (e) {
      const msg = String(e && e.message || e);
      const temporal = /no ha podido acceder|failed while accessing|Service (error|Spreadsheets)|servicio|timed out|tiempo de espera|try again|vuelve a intentarlo|Internal error|error interno/i.test(msg);
      if (!temporal || i >= 3) throw e;
      Utilities.sleep(400 * (i + 1) + Math.floor(Math.random() * 300));
      _v2Ss = null; _v2Sheets = {};
    }
  }
}

function v2Ss() {
  if (!_v2Ss) _v2Ss = SpreadsheetApp.getActiveSpreadsheet();
  return _v2Ss;
}

// Devuelve la hoja; si no existe la crea con encabezados en negrita, fila 1 fija y formato texto.
function v2EnsureSheet(name) {
  if (_v2Sheets[name]) return _v2Sheets[name];
  const ss = v2Ss();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    try { sh = ss.insertSheet(name); }
    catch (e) { sh = ss.getSheetByName(name); if (!sh) throw e; } // otro pedido la creó al mismo tiempo
  }
  _v2Sheets[name] = sh;
  return sh;
}

function v2SetupHeaders(name, sh, values) {
  const headers = V2_SCHEMA[name].headers;
  const vacia = values.length <= 1 && values[0].every(function (v) { return v === '' || v === null; });
  if (vacia) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    const maxRows = sh.getMaxRows();
    headers.forEach(function (h, i) {
      const t = v2Type(name, h);
      if (t === 'text' || t === 'date' || t === 'time') sh.getRange(1, i + 1, maxRows, 1).setNumberFormat('@');
    });
    return [headers.slice()];
  }
  // Agregar al final los encabezados que falten (no se toca nada más)
  const current = values[0].map(function (v) { return String(v).trim(); });
  const missing = headers.filter(function (h) { return current.indexOf(h) < 0; });
  if (missing.length) {
    let lastCol = current.length;
    while (lastCol > 0 && !current[lastCol - 1]) lastCol--;
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    return v2Retry(function () { return sh.getDataRange().getValues(); });
  }
  return values;
}

function v2FromCell(v, type) {
  const tz = v2Tz();
  switch (type) {
    case 'num':    return Number(v) || 0;
    case 'numOpt': return (v === '' || v === null || v === undefined) ? null : (Number(v) || 0);
    case 'bool':   return v === true || /^(true|verdadero|si|sí|1|x)$/i.test(String(v).trim());
    case 'boolT':  return (v === '' || v === null || v === undefined) ? true
                          : (v === true || /^(true|verdadero|si|sí|1|x)$/i.test(String(v).trim()));
    case 'date':   return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '').trim();
    case 'time':   return v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm') : String(v || '').trim();
    case 'ts':     return v instanceof Date ? v.toISOString() : String(v || '');
    default:       return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : (v === null || v === undefined ? '' : String(v));
  }
}

function v2ToCell(v, type) {
  switch (type) {
    case 'num':    return Number(v) || 0;
    case 'numOpt': return (v === '' || v === null || v === undefined) ? '' : (Number(v) || 0);
    case 'bool':
    case 'boolT':  return !!v;
    case 'ts':     return v ? (v instanceof Date ? v : new Date(v)) : '';
    default:       return (v === null || v === undefined) ? '' : (typeof v === 'object' ? JSON.stringify(v) : String(v));
  }
}

// Devuelve { name, sheet, idx, rows, width } — rows: objetos con _row (n° de fila) y _raw
function v2Table(name) {
  if (_v2Tables[name]) return _v2Tables[name];
  let sh = null;
  let values = v2Retry(function () {
    sh = v2EnsureSheet(name);
    return sh.getDataRange().getValues();
  });
  if (!_v2Ensured[name]) {
    values = v2SetupHeaders(name, sh, values);
    _v2Ensured[name] = true;
  }
  const head = values[0].map(function (v) { return String(v).trim(); });
  const idx = {};
  head.forEach(function (h, i) { if (h && idx[h] === undefined) idx[h] = i; });
  const headers = V2_SCHEMA[name].headers;
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const raw = values[r];
    if (raw.every(function (v) { return v === '' || v === null; })) continue;
    const o = { _row: r + 1, _raw: raw };
    headers.forEach(function (h) {
      o[h] = v2FromCell(idx[h] === undefined ? '' : raw[idx[h]], v2Type(name, h));
    });
    rows.push(o);
  }
  _v2Tables[name] = { name: name, sheet: sh, idx: idx, rows: rows, width: head.length };
  return _v2Tables[name];
}

function v2RowArray(t, o, base) {
  const arr = base ? base.slice() : [];
  while (arr.length < t.width) arr.push('');
  V2_SCHEMA[t.name].headers.forEach(function (h) {
    if (o[h] !== undefined) arr[t.idx[h]] = v2ToCell(o[h], v2Type(t.name, h));
  });
  return arr;
}

function v2TextCols(t) {
  return V2_SCHEMA[t.name].headers
    .filter(function (h) { const ty = v2Type(t.name, h); return ty === 'text' || ty === 'date' || ty === 'time'; })
    .map(function (h) { return t.idx[h] + 1; });
}

function v2Append(name, objs) {
  if (!objs.length) return;
  const t = v2Table(name);
  const sh = t.sheet;
  const start = sh.getLastRow() + 1;
  const need = start + objs.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  v2TextCols(t).forEach(function (c) { sh.getRange(start, c, objs.length, 1).setNumberFormat('@'); });
  const rows = objs.map(function (o) { return v2RowArray(t, o); });
  sh.getRange(start, 1, rows.length, t.width).setValues(rows);
  delete _v2Tables[name];
}

function v2Update(name, o) {
  const t = v2Table(name);
  const arr = v2RowArray(t, o, o._raw);
  t.sheet.getRange(o._row, 1, 1, t.width).setValues([arr]);
  o._raw = arr;
}

function v2DeleteRows(name, rowNums) {
  if (!rowNums.length) return;
  const sh = v2Table(name).sheet;
  rowNums.slice().sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  delete _v2Tables[name];
}

function v2Clean(o) {
  const c = {};
  Object.keys(o).forEach(function (k) { if (k.charAt(0) !== '_') c[k] = o[k]; });
  return c;
}

function v2Id(prefix) {
  return prefix + new Date().getTime().toString(36) + Math.random().toString(36).substring(2, 7);
}

// ============================================================
// Utilidades de texto, teléfono y montos
// ============================================================

function v2Norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim();
}

function v2NombreKey(nombre, apellido) { return v2Norm((nombre || '') + ' ' + (apellido || '')); }

// Normaliza un teléfono argentino a 10 dígitos (característica + número).
// Si no se puede, devuelve los dígitos tal cual.
function v2NormTel(tel) {
  let d = String(tel || '').replace(/\D/g, '');
  if (!d) return '';
  if (d.indexOf('00') === 0) d = d.substring(2);
  if (d.indexOf('54') === 0 && d.length >= 12) d = d.substring(2);
  if (d.charAt(0) === '9' && (d.length === 11 || d.length === 13)) d = d.substring(1);
  if (d.charAt(0) === '0') d = d.substring(1);
  if (d.length === 12) {
    // característica (2 a 4 dígitos) + 15 + número
    for (let a = 2; a <= 4; a++) {
      if (d.substring(a, a + 2) === '15') { d = d.substring(0, a) + d.substring(a + 2); break; }
    }
  }
  return d;
}

function v2TelValido(norm) { return /^\d{10}$/.test(norm || ''); }

function v2Money(n) {
  n = Math.round(Number(n) || 0);
  const s = String(Math.abs(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (n < 0 ? '-$' : '$') + s;
}

function v2Plural(n, sing, plur) { return n + ' ' + (n === 1 ? sing : plur); }

function v2Estado(total, pagado) {
  const saldo = total - pagado;
  if (saldo <= 0) return 'pagado';
  if (pagado > 0) return 'senado';
  return 'debe';
}

function v2ParseJSON(s, def) {
  if (s && typeof s === 'object') return s;
  try { const v = JSON.parse(s); return v === null || v === undefined ? def : v; } catch (e) { return def; }
}

function v2Int(v) { const n = Math.round(Number(v)); return isFinite(n) ? n : 0; }

// ============================================================
// Usuarios
// ============================================================

function v2Usuarios() {
  const t = v2Table('Usuarios');
  // Filas cargadas a mano sin id: se les asigna uno
  const sinId = t.rows.filter(function (u) { return !u.id && String(u.nombre).trim(); });
  if (sinId.length) {
    const asignar = function () {
      _v2Tables = {};
      v2Table('Usuarios').rows.forEach(function (u) {
        if (!u.id && String(u.nombre).trim()) { u.id = v2Id('U'); v2Update('Usuarios', u); }
      });
    };
    if (_v2HoldsLock) asignar();
    else {
      const lock = LockService.getScriptLock();
      if (lock.tryLock(10000)) { try { asignar(); SpreadsheetApp.flush(); } finally { lock.releaseLock(); } }
    }
    delete _v2Tables.Usuarios;
  }
  return v2Table('Usuarios').rows.filter(function (u) { return u.id && u.activo && String(u.nombre).trim(); });
}

function v2Auth(data) {
  const u = v2Usuarios().filter(function (x) { return x.id === data.usuarioId; })[0];
  if (!u) fail('Elegí de nuevo quién sos.', { reLogin: true });
  if (String(u.pin).trim() && String(data.pin || '').trim() !== String(u.pin).trim()) fail('PIN incorrecto.', { reLogin: true });
  return u;
}

function v2GetInit() {
  const usuarios = v2Usuarios()
    .map(function (u) { return { id: u.id, nombre: String(u.nombre).trim(), tienePin: !!String(u.pin).trim() }; })
    .sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });
  const eventos = v2Table('Eventos').rows.map(v2EventoPublico);
  return { usuarios: usuarios, eventos: v2OrdenarEventos(eventos), eventoActualId: v2EventoActualId(eventos) };
}

function v2Login(data) {
  const u = v2Usuarios().filter(function (x) { return x.id === data.usuarioId; })[0];
  if (!u) fail('Elegí de nuevo quién sos.');
  if (String(u.pin).trim() && String(data.pin || '').trim() !== String(u.pin).trim()) fail('PIN incorrecto.');
  return { usuario: { id: u.id, nombre: String(u.nombre).trim(), tienePin: !!String(u.pin).trim() } };
}

// ============================================================
// Eventos y productos
// ============================================================

function v2EventoPublico(e) {
  const campos = Object.assign({}, V2_CAMPOS_DEFAULT, v2ParseJSON(e.campos, {}));
  let tipos = v2ParseJSON(e.tiposComensal, null);
  if (!Array.isArray(tipos) || !tipos.length) tipos = V2_TIPOS_DEFAULT.slice();
  return {
    id: e.id, nombre: e.nombre, fecha: e.fecha, hora: e.hora, lugar: e.lugar,
    estado: e.estado === 'cerrado' ? 'cerrado' : 'abierto',
    campos: campos, tiposComensal: tipos, plantillaWhatsApp: e.plantillaWhatsApp || '', creado: e.creado
  };
}

function v2OrdenarEventos(evs) {
  return evs.slice().sort(function (a, b) {
    return String(b.fecha || '').localeCompare(String(a.fecha || '')) || String(b.creado).localeCompare(String(a.creado));
  });
}

function v2EventoActualId(evs) {
  const ab = v2OrdenarEventos(evs.filter(function (e) { return e.estado === 'abierto'; }));
  return ab.length ? ab[0].id : (evs.length ? v2OrdenarEventos(evs)[0].id : null);
}

function v2FindEvento(eventoId) {
  const e = v2Table('Eventos').rows.filter(function (x) { return x.id === eventoId; })[0];
  if (!e) fail('No encontramos ese evento.');
  return e;
}

// Vendidos por producto (sin contar anulados ni, opcionalmente, un pedido)
function v2Vendidos(eventoId, exceptoPedidoId) {
  const anulados = {};
  v2Table('Pedidos').rows.forEach(function (p) { if (p.eventoId === eventoId && p.anulado) anulados[p.id] = true; });
  const v = {};
  v2Table('Items').rows.forEach(function (it) {
    if (it.eventoId !== eventoId || anulados[it.pedidoId] || it.pedidoId === exceptoPedidoId) return;
    v[it.productoId] = (v[it.productoId] || 0) + it.cantidad;
  });
  return v;
}

function v2ProductosEvento(eventoId) {
  const vend = v2Vendidos(eventoId);
  return v2Table('Productos').rows
    .filter(function (p) { return p.eventoId === eventoId; })
    .sort(function (a, b) { return a.orden - b.orden; })
    .map(function (p) {
      const o = v2Clean(p);
      o.vendidos = vend[p.id] || 0;
      o.quedan = p.cupo === null ? null : Math.max(0, p.cupo - o.vendidos);
      return o;
    });
}

function v2EventoData(eventoId) {
  const ev = v2FindEvento(eventoId);
  const de = function (name) {
    return v2Table(name).rows.filter(function (r) { return r.eventoId === eventoId; }).map(v2Clean);
  };
  return {
    evento:    v2EventoPublico(ev),
    productos: v2ProductosEvento(eventoId),
    pedidos:   de('Pedidos'),
    items:     de('Items'),
    pagos:     de('Pagos')
  };
}

function v2GetEvento(data) {
  if (!data.eventoId) fail('Elegí un evento.');
  return v2EventoData(data.eventoId);
}

function v2ValidarEvento(ev) {
  const nombre = String(ev.nombre || '').trim();
  if (!nombre) fail('Poné el nombre del evento.');
  const fecha = String(ev.fecha || '').trim();
  if (fecha && !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) fail('Revisá la fecha del evento.');
  const hora = String(ev.hora || '').trim();
  if (hora && !/^\d{1,2}:\d{2}$/.test(hora)) fail('Revisá la hora del evento.');
  const campos = {};
  Object.keys(V2_CAMPOS_DEFAULT).forEach(function (k) {
    campos[k] = ev.campos && ev.campos[k] !== undefined ? !!ev.campos[k] : V2_CAMPOS_DEFAULT[k];
  });
  let tipos = Array.isArray(ev.tiposComensal) ? ev.tiposComensal : V2_TIPOS_DEFAULT;
  const vistos = {};
  tipos = tipos.map(function (x) { return String(x || '').trim(); })
    .filter(function (x) { const k = v2Norm(x); if (!x || vistos[k]) return false; vistos[k] = true; return true; });
  if (campos.comensales && !tipos.length) fail('Agregá al menos un tipo de comensal.');
  if (!tipos.length) tipos = V2_TIPOS_DEFAULT.slice();
  return {
    nombre: nombre, fecha: fecha, hora: hora, lugar: String(ev.lugar || '').trim(),
    estado: ev.estado === 'cerrado' ? 'cerrado' : 'abierto',
    campos: JSON.stringify(campos), tiposComensal: JSON.stringify(tipos),
    plantillaWhatsApp: ev.plantillaWhatsApp !== undefined ? String(ev.plantillaWhatsApp || '') : undefined
  };
}

// data: { evento: {id?, nombre, fecha, hora, lugar, estado, campos, tiposComensal, plantillaWhatsApp}, productos?: [...] }
// Los productos que no vienen en la lista no se tocan (para ocultar uno: activo=false).
function v2SaveEvento(data) {
  let ev = data.evento || {};
  let eventoId = ev.id;
  let row = null;
  if (eventoId) {
    row = v2FindEvento(eventoId);
    // Lo que no viene se mantiene (permite, por ejemplo, solo cerrar el evento)
    const base = v2EventoPublico(row);
    ev = Object.assign({}, base, ev, { campos: Object.assign({}, base.campos, ev.campos || {}) });
  }
  const val = v2ValidarEvento(ev);

  // Validar productos antes de escribir nada
  const prods = Array.isArray(data.productos) ? data.productos : null;
  const existentes = {};
  if (eventoId) v2Table('Productos').rows.forEach(function (p) { if (p.eventoId === eventoId) existentes[p.id] = p; });
  const vend = eventoId ? v2Vendidos(eventoId) : {};
  const nuevos = [];
  const cambios = [];
  if (prods) {
    prods.forEach(function (p, i) {
      const nombre = String(p.nombre || '').trim();
      if (!nombre) fail('Falta el nombre de un producto.');
      const precio = Number(p.precio);
      if (!isFinite(precio) || precio < 0) fail('Revisá el precio de ' + nombre + '.');
      let cupo = (p.cupo === '' || p.cupo === null || p.cupo === undefined) ? null : Number(p.cupo);
      if (cupo !== null && (!isFinite(cupo) || cupo < 0 || Math.round(cupo) !== cupo)) fail('Revisá el cupo de ' + nombre + '.');
      const o = {
        nombre: nombre, precio: precio, cupo: cupo,
        orden: p.orden !== undefined && p.orden !== null && p.orden !== '' ? Number(p.orden) : i + 1,
        activo: p.activo === undefined ? true : !!p.activo,
        cuentaComensales: !!p.cuentaComensales
      };
      if (p.id && existentes[p.id]) {
        const vendidos = vend[p.id] || 0;
        if (cupo !== null && cupo < vendidos) {
          fail('Ya se ' + (vendidos === 1 ? 'vendió 1' : 'vendieron ' + vendidos) + ' de ' + nombre + '. El cupo no puede ser menor.');
        }
        cambios.push(Object.assign(existentes[p.id], o));
      } else if (p.id) {
        fail('No encontramos el producto ' + nombre + '. Recargá la pantalla.');
      } else {
        nuevos.push(o);
      }
    });
  }

  // Escribir evento
  if (row) {
    Object.keys(val).forEach(function (k) { if (val[k] !== undefined) row[k] = val[k]; });
    v2Update('Eventos', row);
    // Mantener legible el nombre del evento en sus pedidos
    v2Table('Pedidos').rows.forEach(function (p) {
      if (p.eventoId === eventoId && p.evento !== row.nombre) { p.evento = row.nombre; v2Update('Pedidos', p); }
    });
  } else {
    eventoId = v2Id('E');
    v2Append('Eventos', [Object.assign({ id: eventoId, creado: new Date() }, val, {
      plantillaWhatsApp: val.plantillaWhatsApp || ''
    })]);
  }

  cambios.forEach(function (p) { v2Update('Productos', p); });
  if (cambios.length) {
    // Mantener legible el nombre del producto en los items
    const porId = {};
    cambios.forEach(function (p) { porId[p.id] = p.nombre; });
    v2Table('Items').rows.forEach(function (it) {
      if (porId[it.productoId] !== undefined && it.producto !== porId[it.productoId]) {
        it.producto = porId[it.productoId]; v2Update('Items', it);
      }
    });
  }
  v2Append('Productos', nuevos.map(function (p) { return Object.assign({ id: v2Id('P'), eventoId: eventoId }, p); }));
  return Object.assign({ eventoId: eventoId }, v2EventoData(eventoId));
}

// data: { eventoId, nuevoNombre, nuevaFecha }
function v2CopiarEvento(data) {
  const orig = v2FindEvento(data.eventoId);
  const pub = v2EventoPublico(orig);
  const prods = v2Table('Productos').rows.filter(function (p) { return p.eventoId === orig.id; })
    .map(function (p) { return { nombre: p.nombre, precio: p.precio, cupo: p.cupo, orden: p.orden, activo: p.activo, cuentaComensales: p.cuentaComensales }; });
  return v2SaveEvento({
    _usuario: data._usuario,
    evento: {
      nombre: data.nuevoNombre, fecha: data.nuevaFecha || '', hora: pub.hora, lugar: pub.lugar, estado: 'abierto',
      campos: pub.campos, tiposComensal: pub.tiposComensal, plantillaWhatsApp: pub.plantillaWhatsApp
    },
    productos: prods
  });
}

// ============================================================
// Contactos
// ============================================================

function v2ContactoPublico(c) {
  return {
    id: c.id, nombre: c.nombre, apellido: c.apellido, telefono: c.telefono,
    whatsapp: v2TelValido(c.telefonoNorm) ? c.telefonoNorm : '',
    cantEventos: c.cantEventos, eventos: c.eventos, productos: c.productos
  };
}

// data: { texto }
function v2BuscarContacto(data) {
  const texto = String(data.texto || '').trim();
  if (texto.length < 2) return { contactos: [] };
  const digitos = texto.replace(/\D/g, '');
  const rows = v2Table('Contactos').rows;
  let res;
  if (digitos.length >= 6 && digitos.length >= texto.replace(/\s/g, '').length - 2) {
    const n = v2NormTel(digitos);
    res = rows.filter(function (c) {
      return c.telefonoNorm && (c.telefonoNorm === n || c.telefonoNorm.indexOf(n) >= 0 || c.telefonoNorm.indexOf(digitos) >= 0);
    });
  } else {
    const q = v2Norm(texto);
    const palabras = q.split(' ');
    res = rows.filter(function (c) {
      const k = v2NombreKey(c.nombre, c.apellido);
      return palabras.every(function (w) { return k.indexOf(w) >= 0; });
    });
    res.sort(function (a, b) {
      const ka = v2NombreKey(a.nombre, a.apellido), kb = v2NombreKey(b.nombre, b.apellido);
      return (ka.indexOf(q) === 0 ? 0 : 1) - (kb.indexOf(q) === 0 ? 0 : 1) || b.cantEventos - a.cantEventos;
    });
  }
  return { contactos: res.slice(0, 5).map(v2ContactoPublico) };
}

// Busca o crea el contacto para un pedido. Devuelve la fila (puede ser nueva, aún sin guardar si _nuevo).
// datos: { nombre, apellido, telefono, contactoId? }
function v2ResolverContacto(datos) {
  const t = v2Table('Contactos');
  const nombre = String(datos.nombre || '').trim();
  const apellido = String(datos.apellido || '').trim();
  const telefono = String(datos.telefono || '').trim();
  const norm = v2NormTel(telefono);
  const key = v2NombreKey(nombre, apellido);
  let c = null;

  const porTel = norm ? t.rows.filter(function (x) { return x.telefonoNorm === norm; })[0] : null;
  if (datos.contactoId) c = t.rows.filter(function (x) { return x.id === datos.contactoId; })[0] || null;
  if (c && porTel && porTel.id !== c.id) c = porTel; // el teléfono es de otro contacto
  if (!c && porTel) c = porTel;
  if (!c) {
    const mismos = t.rows.filter(function (x) { return v2NombreKey(x.nombre, x.apellido) === key; });
    if (norm) {
      // Completar un contacto que existía sin teléfono
      c = mismos.filter(function (x) { return !x.telefonoNorm; })[0] || null;
    } else {
      c = mismos.filter(function (x) { return !x.telefonoNorm; })[0] || mismos[0] || null;
    }
  }
  if (!c) {
    c = { _nuevo: true, id: v2Id('C'), creado: new Date(), cantEventos: 0, eventos: '', productos: '', primerEvento: '', ultimoEvento: '' };
  }
  c.nombre = nombre || c.nombre;
  if (apellido || c._nuevo) c.apellido = apellido;
  if (norm) { c.telefono = telefono; c.telefonoNorm = norm; }
  else if (c._nuevo) { c.telefono = ''; c.telefonoNorm = ''; }
  c.actualizado = new Date();
  return c;
}

function v2GuardarContacto(c) {
  if (c._nuevo) { delete c._nuevo; v2Append('Contactos', [c]); }
  else v2Update('Contactos', c);
}

// Recalcula eventos, productos y cantEventos de los contactos a partir de los pedidos (sin anulados)
function v2RecalcularContactos(ids) {
  const want = {};
  ids.forEach(function (id) { if (id) want[id] = true; });
  if (!Object.keys(want).length) return;
  const eventos = {};
  v2Table('Eventos').rows.forEach(function (e) { eventos[e.id] = e; });
  const pedidosPorContacto = {};
  const pedidoContacto = {};
  v2Table('Pedidos').rows.forEach(function (p) {
    if (!want[p.contactoId] || p.anulado) return;
    (pedidosPorContacto[p.contactoId] = pedidosPorContacto[p.contactoId] || []).push(p);
    pedidoContacto[p.id] = p.contactoId;
  });
  const prodsPorContacto = {};
  v2Table('Items').rows.forEach(function (it) {
    const cid = pedidoContacto[it.pedidoId];
    if (!cid) return;
    const l = prodsPorContacto[cid] = prodsPorContacto[cid] || [];
    if (l.indexOf(it.producto) < 0) l.push(it.producto);
  });
  v2Table('Contactos').rows.forEach(function (c) {
    if (!want[c.id]) return;
    const ps = pedidosPorContacto[c.id] || [];
    const evIds = [];
    ps.forEach(function (p) { if (evIds.indexOf(p.eventoId) < 0) evIds.push(p.eventoId); });
    const evs = evIds.map(function (id) { return eventos[id] || { id: id, nombre: '', fecha: '' }; })
      .sort(function (a, b) { return String(a.fecha).localeCompare(String(b.fecha)); });
    const nuevo = {
      cantEventos: evs.length,
      eventos: evs.map(function (e) { return e.nombre; }).join(V2_SEP),
      productos: (prodsPorContacto[c.id] || []).join(V2_SEP),
      primerEvento: evs.length ? evs[0].nombre : '',
      ultimoEvento: evs.length ? evs[evs.length - 1].nombre : ''
    };
    const cambio = Object.keys(nuevo).some(function (k) { return c[k] !== nuevo[k]; });
    if (cambio) { Object.assign(c, nuevo, { actualizado: new Date() }); v2Update('Contactos', c); }
  });
}

// data: { eventoId?, productoId?, producto? (nombre), texto? }
function v2GetContactos(data) {
  let rows = v2Table('Contactos').rows;
  const eventoId = data.eventoId || '';
  const productoId = data.productoId || '';
  const productoNom = v2Norm(data.producto || '');
  if (eventoId || productoId || productoNom) {
    const pedidosOk = {};
    v2Table('Pedidos').rows.forEach(function (p) {
      if (p.anulado) return;
      if (eventoId && p.eventoId !== eventoId) return;
      pedidosOk[p.id] = p.contactoId;
    });
    const ids = {};
    if (productoId || productoNom) {
      v2Table('Items').rows.forEach(function (it) {
        if (!pedidosOk[it.pedidoId]) return;
        if (productoId && it.productoId !== productoId) return;
        if (!productoId && productoNom && v2Norm(it.producto) !== productoNom) return;
        ids[pedidosOk[it.pedidoId]] = true;
      });
    } else {
      Object.keys(pedidosOk).forEach(function (k) { ids[pedidosOk[k]] = true; });
    }
    rows = rows.filter(function (c) { return ids[c.id]; });
  }
  const texto = String(data.texto || '').trim();
  if (texto) {
    const q = v2Norm(texto);
    const dig = texto.replace(/\D/g, '');
    rows = rows.filter(function (c) {
      return v2NombreKey(c.nombre, c.apellido).indexOf(q) >= 0 || (dig.length >= 3 && String(c.telefonoNorm).indexOf(dig) >= 0);
    });
  }
  const lista = rows.map(v2ContactoPublico).sort(function (a, b) {
    return v2NombreKey(a.nombre, a.apellido).localeCompare(v2NombreKey(b.nombre, b.apellido), 'es');
  });
  return { contactos: lista, total: lista.length };
}

// ============================================================
// Pedidos
// ============================================================

// Arma las líneas validando productos, tipos de comensal y cupo.
// lineas: [{productoId, tipoComensal, cantidad}]
// previas: items actuales del pedido (al editar) para mantener el precio congelado
function v2ArmarItems(ev, lineas, pedidoId, previas) {
  const pub = v2EventoPublico(ev);
  const prods = {};
  v2Table('Productos').rows.forEach(function (p) { if (p.eventoId === ev.id) prods[p.id] = p; });
  const previo = {}, previoCant = {};
  (previas || []).forEach(function (it) {
    previo[it.productoId + '|' + it.tipoComensal] = it;
    previoCant[it.productoId] = (previoCant[it.productoId] || 0) + it.cantidad;
  });

  // Unificar líneas repetidas
  const merged = {}, orden = [];
  (lineas || []).forEach(function (l) {
    const cant = v2Int(l.cantidad);
    if (cant < 0) fail('Revisá las cantidades.');
    if (!cant) return;
    const p = prods[l.productoId];
    if (!p) fail('Uno de los productos ya no está. Recargá la pantalla.');
    let tipo = '';
    if (p.cuentaComensales && pub.campos.comensales) {
      tipo = String(l.tipoComensal || '').trim();
      const ok = pub.tiposComensal.filter(function (x) { return v2Norm(x) === v2Norm(tipo); })[0];
      if (!ok && !previo[p.id + '|' + tipo]) fail('Elegí el tipo de comensal de ' + p.nombre + '.');
      tipo = ok || tipo;
    }
    const k = p.id + '|' + tipo;
    if (!merged[k]) { merged[k] = { p: p, tipo: tipo, cantidad: 0 }; orden.push(k); }
    merged[k].cantidad += cant;
  });
  if (!orden.length) fail('Agregá al menos un producto.');

  // Productos ocultos y cupo
  const vend = v2Vendidos(ev.id, pedidoId);
  const porProducto = {};
  orden.forEach(function (k) { const m = merged[k]; porProducto[m.p.id] = (porProducto[m.p.id] || 0) + m.cantidad; });
  Object.keys(porProducto).forEach(function (pid) {
    const p = prods[pid];
    const pide = porProducto[pid];
    if (!p.activo && pide > (previoCant[pid] || 0)) fail(p.nombre + ' ya no se ofrece.');
    if (p.cupo !== null) {
      const disp = Math.max(0, p.cupo - (vend[pid] || 0));
      if (pide > disp) {
        if (disp === 0) fail('Se completó el cupo de ' + p.nombre + '.');
        fail('Solo ' + (disp === 1 ? 'queda 1' : 'quedan ' + disp) + ' de ' + p.nombre + '.');
      }
    }
  });

  // Precios desde la hoja (o congelados si la línea ya existía)
  return orden.map(function (k) {
    const m = merged[k];
    const prev = previo[k];
    const precio = prev ? prev.precioUnitario : m.p.precio;
    return {
      productoId: m.p.id, producto: m.p.nombre, tipoComensal: m.tipo,
      cantidad: m.cantidad, precioUnitario: precio, subtotal: precio * m.cantidad
    };
  });
}

function v2FindPedido(pedidoId) {
  const p = v2Table('Pedidos').rows.filter(function (x) { return x.id === pedidoId; })[0];
  if (!p) fail('No encontramos ese pedido. Recargá la pantalla.');
  return p;
}

function v2PagadoDe(pedidoId) {
  return v2Table('Pagos').rows.reduce(function (s, x) {
    return x.pedidoId === pedidoId ? s + (x.tipo === 'devolucion' ? -Math.abs(x.monto) : x.monto) : s;
  }, 0);
}

function v2Recalcular(p) {
  p.total = v2Table('Items').rows.reduce(function (s, it) { return it.pedidoId === p.id ? s + it.subtotal : s; }, 0);
  p.pagado = v2PagadoDe(p.id);
  p.saldo = p.total - p.pagado;
  p.estado = v2Estado(p.total, p.pagado);
}

function v2PedidoRespuesta(pedidoId) {
  const p = v2FindPedido(pedidoId);
  return {
    pedido: v2Clean(p),
    items: v2Table('Items').rows.filter(function (it) { return it.pedidoId === pedidoId; }).map(v2Clean),
    pagos: v2Table('Pagos').rows.filter(function (x) { return x.pedidoId === pedidoId; }).map(v2Clean),
    productos: v2ProductosEvento(p.eventoId)
  };
}

function v2NombreCompleto(p) { return [p.nombre, p.apellido].filter(function (x) { return String(x || '').trim(); }).join(' '); }

// data: { idemKey, pedido: {eventoId, nombre, apellido, telefono, nota, retiro, contactoId?},
//         items: [{productoId, tipoComensal, cantidad}], pagoInicial?: {monto, medio} }
function v2SavePedido(data) {
  const u = data._usuario;
  const cache = CacheService.getScriptCache();
  const idemKey = data.idemKey ? 'idem_' + String(data.idemKey).substring(0, 100) : null;
  if (idemKey) {
    const ya = cache.get(idemKey);
    if (ya) {
      const p = v2Table('Pedidos').rows.filter(function (x) { return x.id === ya; })[0];
      if (p) return Object.assign({ repetido: true }, v2PedidoRespuesta(ya));
    }
  }

  const ped = data.pedido || {};
  const ev = v2FindEvento(ped.eventoId);
  const pub = v2EventoPublico(ev);
  if (pub.estado === 'cerrado') fail('Este evento está cerrado. No se pueden cargar pedidos.');
  const nombre = String(ped.nombre || '').trim();
  if (!nombre) fail('Poné el nombre.');
  const apellido = pub.campos.apellido ? String(ped.apellido || '').trim() : '';

  const items = v2ArmarItems(ev, data.items, null, null);
  const total = items.reduce(function (s, it) { return s + it.subtotal; }, 0);

  // Pago inicial
  let pago = null;
  const pi = data.pagoInicial;
  if (pi && Number(pi.monto) > 0) {
    const monto = Math.round(Number(pi.monto));
    if (monto > total) fail('La seña no puede ser mayor que el total (' + v2Money(total) + ').');
    const medio = V2_MEDIOS.indexOf(pi.medio) >= 0 ? pi.medio : 'efectivo';
    pago = { monto: monto, medio: medio, tipo: monto >= total ? 'pago' : 'sena' };
  }

  const contacto = v2ResolverContacto({ nombre: nombre, apellido: apellido, telefono: ped.telefono, contactoId: ped.contactoId });
  v2GuardarContacto(contacto);

  const pedidoId = v2Id('O');
  const pagado = pago ? pago.monto : 0;
  const ahora = new Date();
  v2Append('Pedidos', [{
    id: pedidoId, fecha: ahora, eventoId: ev.id, evento: ev.nombre, contactoId: contacto.id,
    nombre: nombre, apellido: apellido, telefono: String(ped.telefono || '').trim(),
    total: total, pagado: pagado, saldo: total - pagado, estado: v2Estado(total, pagado),
    entregado: false, retiro: pub.campos.retiro ? String(ped.retiro || '').trim() : '',
    nota: pub.campos.nota ? String(ped.nota || '').trim() : '',
    anoto: u.nombre, anulado: false, anuladoPor: '', motivoAnulacion: ''
  }]);
  v2Append('Items', items.map(function (it) {
    return Object.assign({ id: v2Id('I'), pedidoId: pedidoId, eventoId: ev.id }, it);
  }));
  if (pago) {
    v2Append('Pagos', [{
      id: v2Id('G'), fecha: ahora, pedidoId: pedidoId, eventoId: ev.id, nombre: v2NombreCompleto({ nombre: nombre, apellido: apellido }),
      monto: pago.monto, medio: pago.medio, tipo: pago.tipo, cobro: u.nombre, nota: ''
    }]);
  }
  v2RecalcularContactos([contacto.id]);
  if (idemKey) cache.put(idemKey, pedidoId, 21600);
  return v2PedidoRespuesta(pedidoId);
}

// data: { pedidoId, cambios: {nombre, apellido, telefono, nota, retiro, entregado}, items?: [...] }
function v2UpdatePedido(data) {
  const p = v2FindPedido(data.pedidoId);
  if (p.anulado) fail('Este pedido está anulado.');
  const ev = v2FindEvento(p.eventoId);
  const pub = v2EventoPublico(ev);
  const c = data.cambios || {};
  const contactoAnterior = p.contactoId;

  if (c.nombre !== undefined) {
    if (!String(c.nombre).trim()) fail('Poné el nombre.');
    p.nombre = String(c.nombre).trim();
  }
  ['apellido', 'telefono', 'nota', 'retiro'].forEach(function (k) {
    if (c[k] !== undefined) p[k] = String(c[k] || '').trim();
  });
  if (c.entregado !== undefined) p.entregado = !!c.entregado;

  if (Array.isArray(data.items)) {
    const previas = v2Table('Items').rows.filter(function (it) { return it.pedidoId === p.id; });
    const nuevos = v2ArmarItems(ev, data.items, p.id, previas);
    v2DeleteRows('Items', previas.map(function (it) { return it._row; }));
    v2Append('Items', nuevos.map(function (it) {
      return Object.assign({ id: v2Id('I'), pedidoId: p.id, eventoId: ev.id }, it);
    }));
  }

  if (c.nombre !== undefined || c.apellido !== undefined || c.telefono !== undefined) {
    const contacto = v2ResolverContacto({ nombre: p.nombre, apellido: p.apellido, telefono: p.telefono, contactoId: p.contactoId });
    v2GuardarContacto(contacto);
    p.contactoId = contacto.id;
  }

  v2Recalcular(p);
  v2Update('Pedidos', p);
  // Nombre legible en los pagos
  const nom = v2NombreCompleto(p);
  v2Table('Pagos').rows.forEach(function (x) {
    if (x.pedidoId === p.id && x.nombre !== nom) { x.nombre = nom; v2Update('Pagos', x); }
  });
  v2RecalcularContactos([contactoAnterior, p.contactoId]);
  return v2PedidoRespuesta(p.id);
}

// data: { pedidoId, monto, medio, nota, confirmar? }
function v2AddPago(data) {
  const u = data._usuario;
  const p = v2FindPedido(data.pedidoId);
  if (p.anulado) fail('Este pedido está anulado.');
  v2Recalcular(p);
  const monto = Math.round(Number(data.monto) || 0);
  if (monto <= 0) fail('Poné cuánto cobrás.');
  if (!data.confirmar) {
    if (p.saldo <= 0) fail('Este pedido ya está pagado.', { confirmar: true });
    if (monto > p.saldo) {
      fail('Le falta pagar ' + v2Money(p.saldo) + '. ¿Cobrás ' + v2Money(monto) + ' igual?', { confirmar: true });
    }
  }
  const medio = V2_MEDIOS.indexOf(data.medio) >= 0 ? data.medio : 'efectivo';
  v2Append('Pagos', [{
    id: v2Id('G'), fecha: new Date(), pedidoId: p.id, eventoId: p.eventoId, nombre: v2NombreCompleto(p),
    monto: monto, medio: medio, tipo: 'pago', cobro: u.nombre, nota: String(data.nota || '').trim()
  }]);
  const p2 = v2FindPedido(data.pedidoId);
  v2Recalcular(p2);
  v2Update('Pedidos', p2);
  return v2PedidoRespuesta(p2.id);
}

// data: { pedidoId, entregado }
function v2SetEntregado(data) {
  const p = v2FindPedido(data.pedidoId);
  if (p.anulado) fail('Este pedido está anulado.');
  p.entregado = !!data.entregado;
  v2Update('Pedidos', p);
  return v2PedidoRespuesta(p.id);
}

// data: { pedidoId, motivo }
function v2AnularPedido(data) {
  const p = v2FindPedido(data.pedidoId);
  if (p.anulado) fail('Este pedido ya estaba anulado.');
  const motivo = String(data.motivo || '').trim();
  if (!motivo) fail('Contá brevemente por qué se anula.');
  p.anulado = true;
  p.anuladoPor = data._usuario.nombre;
  p.motivoAnulacion = motivo;
  v2Update('Pedidos', p);
  v2RecalcularContactos([p.contactoId]);
  return v2PedidoRespuesta(p.id);
}

// ============================================================
// MIGRACIÓN DESDE V1 — ejecutar a mano desde el editor: migrarDesdeV1
// ============================================================
// - "Registros" + "Movimientos" + "Config"  → evento abierto (el actual).
// - Cada par archivado "Registros X" + "Movimientos X" → evento cerrado "X".
// - No borra ni modifica ninguna hoja vieja. Es idempotente: si ya se migró, no duplica.

const V2_MIGRAR_ARCHIVADAS = true;

function migrarDesdeV1() {
  _v2Reset();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  _v2HoldsLock = true;
  const log = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fuentes = [];
    if (ss.getSheetByName(SHEET_REG)) {
      fuentes.push({ key: 'principal', reg: SHEET_REG, mov: SHEET_MOV, principal: true, eventoId: 'E_V1' });
    }
    if (V2_MIGRAR_ARCHIVADAS) {
      ss.getSheets().forEach(function (sh) {
        const m = sh.getName().match(/^Registros\s+(.+)$/);
        if (!m) return;
        const suf = m[1].trim();
        fuentes.push({
          key: 'archivo:' + suf, reg: sh.getName(), mov: 'Movimientos ' + suf, nombre: suf,
          eventoId: 'E_V1_' + v2Norm(suf).replace(/[^a-z0-9]+/g, '_')
        });
      });
    }
    if (!fuentes.length) { log.push('No hay hojas de la versión anterior para migrar.'); return log; }

    const props = PropertiesService.getScriptProperties();
    const hechos = v2ParseJSON(props.getProperty('V2_MIGRADO'), {});
    const contactosTocados = {};
    fuentes.forEach(function (f) {
      const yaEvento = v2Table('Eventos').rows.some(function (e) { return e.id === f.eventoId; });
      if (hechos[f.key] || yaEvento) {
        log.push('— "' + f.reg + '": ya estaba migrado, no se duplica nada.');
        hechos[f.key] = f.eventoId;
        return;
      }
      const r = v2MigrarFuente(f);
      r.contactos.forEach(function (id) { contactosTocados[id] = true; });
      hechos[f.key] = f.eventoId;
      props.setProperty('V2_MIGRADO', JSON.stringify(hechos));
      r.log.forEach(function (l) { log.push(l); });
    });
    props.setProperty('V2_MIGRADO', JSON.stringify(hechos));
    _v2Tables = {};
    v2RecalcularContactos(Object.keys(contactosTocados));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
    _v2HoldsLock = false;
    log.forEach(function (l) { Logger.log(l); });
  }
  return log;
}

function v2LeerHojaVieja(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const head = values[0].map(function (v) { return String(v).trim(); });
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const o = { _fila: r + 1 };
    head.forEach(function (h, i) { if (h) o[h] = values[r][i]; });
    out.push(o);
  }
  return out;
}

function v2LeerConfigVieja() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CFG);
  const cfg = { eventName: '', eventDate: '', priceTypes: [] };
  if (!sh || sh.getLastRow() === 0) return cfg;
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (row) {
    if (!row[0]) return;
    let v = row[1];
    if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    cfg[String(row[0]).trim()] = v;
  });
  cfg.priceTypes = v2ParseJSON(cfg.priceTypes, []);
  if (!Array.isArray(cfg.priceTypes)) cfg.priceTypes = [];
  return cfg;
}

function v2AFecha(v) {
  if (v instanceof Date) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function v2MigrarFuente(f) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const log = [];
  const eventoId = f.eventoId;

  // Si quedó una migración a medias (sin la fila del evento), se limpia lo que haya quedado
  ['Productos', 'Pedidos', 'Items', 'Pagos'].forEach(function (name) {
    const rows = v2Table(name).rows.filter(function (r) { return r.eventoId === eventoId; }).map(function (r) { return r._row; });
    if (rows.length) { v2DeleteRows(name, rows); log.push('  (se limpiaron ' + rows.length + ' filas de una migración anterior incompleta en ' + name + ')'); }
  });

  const regs = v2LeerHojaVieja(ss.getSheetByName(f.reg)).filter(function (r) { return String(r.nombre || '').trim(); });
  const movs = v2LeerHojaVieja(ss.getSheetByName(f.mov)).filter(function (m) { return m.registrationId !== '' && m.registrationId !== undefined; });

  // --- Evento
  let nombreEv, fechaEv;
  let priceTypes = [];
  if (f.principal) {
    const cfg = v2LeerConfigVieja();
    nombreEv = String(cfg.eventName || '').trim() || 'Evento anterior';
    fechaEv = String(cfg.eventDate || '').substring(0, 10);
    priceTypes = cfg.priceTypes;
  } else {
    nombreEv = f.nombre;
    let max = null;
    regs.concat(movs).forEach(function (x) { const d = v2AFecha(x.timestamp); if (d && (!max || d > max)) max = d; });
    fechaEv = max ? Utilities.formatDate(max, tz, 'yyyy-MM-dd') : '';
  }

  // --- Productos (uno por tipo de precio)
  const productos = [];
  const prodPorNombre = {};
  const addProd = function (name, amount, activo) {
    const k = v2Norm(name);
    if (prodPorNombre[k]) return prodPorNombre[k];
    const p = {
      id: eventoId + '_P' + (productos.length + 1), eventoId: eventoId, nombre: String(name).trim() || 'Tarjeta',
      precio: Number(amount) || 0, cupo: null, orden: productos.length + 1, activo: activo, cuentaComensales: true
    };
    productos.push(p); prodPorNombre[k] = p;
    return p;
  };
  priceTypes.forEach(function (t) { if (t && String(t.name || '').trim()) addProd(t.name, t.amount, true); });
  regs.forEach(function (r) {
    const name = String(r.tipoPrecio || '').trim() || 'Tarjeta';
    if (!prodPorNombre[v2Norm(name)]) {
      addProd(name, r.precioUnitario, !f.principal);
      if (f.principal) log.push('  Aviso: el tipo "' + name + '" no estaba en la configuración; se creó oculto.');
    }
  });

  // --- Tipos de comensal
  const tipos = V2_TIPOS_DEFAULT.slice();
  regs.forEach(function (r) {
    const g = String(r.genero || '').trim();
    if (g && !tipos.some(function (t) { return v2Norm(t) === v2Norm(g); })) tipos.push(g);
  });

  // --- Movimientos por registro
  const movsPorReg = {};
  const regIds = {};
  regs.forEach(function (r) { if (r.id) regIds[String(r.id)] = true; });
  const huerfanos = [];
  movs.forEach(function (m) {
    const k = String(m.registrationId);
    if (!regIds[k]) { huerfanos.push(m); return; }
    (movsPorReg[k] = movsPorReg[k] || []).push(m);
  });

  // --- Contactos existentes (por nombre normalizado)
  const tc = v2Table('Contactos');
  const contactoPorKey = {};
  tc.rows.forEach(function (c) {
    const k = v2NombreKey(c.nombre, c.apellido);
    if (!contactoPorKey[k] || (!c.telefonoNorm && contactoPorKey[k].telefonoNorm)) contactoPorKey[k] = c;
  });
  const contactosNuevos = [];
  const contactosIds = {};

  const pedidos = [], items = [], pagos = [];
  const viejo = { pedidos: 0, total: 0, cobrado: 0, porTipo: {}, porProducto: {}, movs: 0 };
  let ajustes = 0, sinMovs = 0;

  regs.forEach(function (r) {
    const regId = r.id ? String(r.id) : 'fila' + r._fila;
    const nombre = String(r.nombre).trim();
    const cantidad = Math.max(1, v2Int(r.cantidad) || 1);
    let precio = Number(r.precioUnitario) || 0;
    let total = (r.totalDebe === '' || r.totalDebe === undefined) ? precio * cantidad : (Number(r.totalDebe) || 0);
    if (Math.abs(precio * cantidad - total) > 0.5) {
      log.push('  Aviso: ' + nombre + ' tenía total ' + v2Money(total) + ' con ' + cantidad + ' × ' + v2Money(precio) + '; se respetó el total.');
      precio = total / cantidad;
    }
    const montoPagado = Number(r.montoPagado) || 0;
    const genero = String(r.genero || '').trim();
    const tipo = genero ? tipos.filter(function (t) { return v2Norm(t) === v2Norm(genero); })[0] : '';
    const prod = prodPorNombre[v2Norm(String(r.tipoPrecio || '').trim() || 'Tarjeta')];

    viejo.pedidos++; viejo.total += total; viejo.cobrado += montoPagado;
    viejo.porTipo[genero || '(sin tipo)'] = (viejo.porTipo[genero || '(sin tipo)'] || 0) + cantidad;
    viejo.porProducto[prod.nombre] = (viejo.porProducto[prod.nombre] || 0) + cantidad;

    // Contacto
    const key = v2NombreKey(nombre, '');
    let c = contactoPorKey[key];
    if (!c) {
      c = { id: v2Id('C'), nombre: nombre, apellido: '', telefono: '', telefonoNorm: '', primerEvento: '', ultimoEvento: '',
            cantEventos: 0, eventos: '', productos: '', creado: new Date(), actualizado: new Date() };
      contactoPorKey[key] = c; contactosNuevos.push(c);
    }
    contactosIds[c.id] = true;

    const pedidoId = 'O_' + eventoId + '_' + regId;
    const fecha = v2AFecha(r.timestamp) || new Date();

    // Pagos
    let suma = 0;
    (movsPorReg[regId] || []).forEach(function (m, i) {
      const monto = Number(m.montoCobrado) || 0;
      viejo.movs += monto;
      suma += monto;
      const medio = String(m.metodoPago || '').trim().toLowerCase();
      pagos.push({
        id: pedidoId + '_G' + (i + 1), fecha: v2AFecha(m.timestamp) || fecha, pedidoId: pedidoId, eventoId: eventoId,
        nombre: nombre, monto: monto, medio: V2_MEDIOS.indexOf(medio) >= 0 ? medio : '',
        tipo: /se[ñn]a/i.test(String(m.tipo)) ? 'sena' : 'pago', cobro: 'Migrado', nota: String(m.notas || '').trim()
      });
    });
    const dif = montoPagado - suma;
    if (Math.abs(dif) > 0.5) {
      const medio = String(r.metodoPago || '').trim().toLowerCase();
      const sinMov = !(movsPorReg[regId] || []).length;
      if (sinMov) sinMovs++; else ajustes++;
      pagos.push({
        id: pedidoId + '_GM', fecha: fecha, pedidoId: pedidoId, eventoId: eventoId, nombre: nombre,
        monto: Math.abs(dif), medio: V2_MEDIOS.indexOf(medio) >= 0 ? medio : '',
        tipo: dif > 0 ? (montoPagado < total ? 'sena' : 'pago') : 'devolucion',
        cobro: 'Migrado', nota: sinMov ? 'Migrado' : 'Migrado (ajuste para que coincida con lo cobrado)'
      });
    }

    pedidos.push({
      id: pedidoId, fecha: fecha, eventoId: eventoId, evento: nombreEv, contactoId: c.id,
      nombre: nombre, apellido: '', telefono: '', total: total, pagado: montoPagado, saldo: total - montoPagado,
      estado: v2Estado(total, montoPagado), entregado: false, retiro: '', nota: String(r.notas || '').trim(),
      anoto: 'Migrado', anulado: false, anuladoPor: '', motivoAnulacion: ''
    });
    items.push({
      id: pedidoId + '_I1', pedidoId: pedidoId, eventoId: eventoId, productoId: prod.id, producto: prod.nombre,
      tipoComensal: tipo || '', cantidad: cantidad, precioUnitario: precio, subtotal: total
    });
  });

  // --- Escribir (la fila del evento va al final: marca que la migración terminó)
  v2Append('Productos', productos);
  v2Append('Pedidos', pedidos);
  v2Append('Items', items);
  v2Append('Pagos', pagos);
  v2Append('Contactos', contactosNuevos);
  v2Append('Eventos', [{
    id: eventoId, nombre: nombreEv, fecha: fechaEv, hora: '', lugar: '', estado: f.principal ? 'abierto' : 'cerrado',
    campos: JSON.stringify({ apellido: false, sena: true, nota: true, retiro: false, entregado: true, comensales: true }),
    tiposComensal: JSON.stringify(tipos), plantillaWhatsApp: '', creado: new Date()
  }]);

  // --- Verificación
  const nuevo = { pedidos: pedidos.length, total: 0, cobrado: 0, porTipo: {}, porProducto: {} };
  pedidos.forEach(function (p) { nuevo.total += p.total; });
  pagos.forEach(function (g) { nuevo.cobrado += g.tipo === 'devolucion' ? -g.monto : g.monto; });
  items.forEach(function (it) {
    const t = it.tipoComensal || '(sin tipo)';
    nuevo.porTipo[t] = (nuevo.porTipo[t] || 0) + it.cantidad;
    nuevo.porProducto[it.producto] = (nuevo.porProducto[it.producto] || 0) + it.cantidad;
  });
  const fmtMap = function (m) { return Object.keys(m).sort().map(function (k) { return k + ' ' + m[k]; }).join(', '); };
  const coincide = viejo.pedidos === nuevo.pedidos && Math.abs(viejo.total - nuevo.total) < 1 &&
    Math.abs(viejo.cobrado - nuevo.cobrado) < 1 && fmtMap(viejo.porTipo) === fmtMap(nuevo.porTipo) &&
    fmtMap(viejo.porProducto) === fmtMap(nuevo.porProducto);

  log.unshift('✔ "' + f.reg + '" → evento "' + nombreEv + '"' + (fechaEv ? ' (' + fechaEv + ')' : '') + (f.principal ? ' — abierto' : ' — cerrado'));
  log.push('  Creados: ' + v2Plural(pedidos.length, 'pedido', 'pedidos') + ', ' + v2Plural(items.length, 'ítem', 'ítems') + ', ' +
    v2Plural(pagos.length, 'pago', 'pagos') + ', ' + v2Plural(contactosNuevos.length, 'contacto nuevo', 'contactos nuevos') + ', ' +
    v2Plural(productos.length, 'producto', 'productos') + '.');
  if (sinMovs) log.push('  ' + v2Plural(sinMovs, 'pago "Migrado" por registros con monto pagado y sin movimientos', 'pagos "Migrado" por registros con monto pagado y sin movimientos') + '.');
  if (ajustes) log.push('  ' + v2Plural(ajustes, 'ajuste', 'ajustes') + ' porque los movimientos no sumaban lo mismo que el monto pagado.');
  if (huerfanos.length) {
    log.push('  Aviso: ' + v2Plural(huerfanos.length, 'movimiento sin registro', 'movimientos sin registro') + ' (no se migran): ' +
      huerfanos.map(function (m) { return (m.nombre || '?') + ' ' + v2Money(m.montoCobrado); }).join('; '));
  }
  log.push('  Antes:  ' + v2Plural(viejo.pedidos, 'registro', 'registros') + ' · total ' + v2Money(viejo.total) + ' · cobrado ' + v2Money(viejo.cobrado) +
    ' (movimientos ' + v2Money(viejo.movs) + ') · por tipo: ' + fmtMap(viejo.porTipo) + ' · por tarjeta: ' + fmtMap(viejo.porProducto));
  log.push('  Ahora:  ' + v2Plural(nuevo.pedidos, 'pedido', 'pedidos') + ' · total ' + v2Money(nuevo.total) + ' · cobrado ' + v2Money(nuevo.cobrado) +
    ' · por tipo: ' + fmtMap(nuevo.porTipo) + ' · por producto: ' + fmtMap(nuevo.porProducto));
  log.push('  ¿Coinciden los totales? ' + (coincide ? 'SÍ' : 'NO — revisar'));
  return { eventoId: eventoId, log: log, contactos: Object.keys(contactosIds), coincide: coincide };
}

// ============================================================
// PRUEBAS — ejecutar a mano: testV2  (después: borrarDatosDePrueba)
// ============================================================
// Crea dos usuarios de prueba, un evento de comida (con cupo, sin comensales)
// y uno de cena (con comensales), y prueba todas las acciones por el mismo
// camino que usa la app (runV2). Al final cierra los eventos de prueba.

function testV2() {
  const res = [];
  const check = function (nombre, cond, detalle) {
    res.push((cond ? 'OK    ' : 'FALLA ') + nombre + (cond || !detalle ? '' : ' → ' + JSON.stringify(detalle)));
    return cond;
  };
  const call = function (action, data) { return runV2(action, JSON.stringify(data || {})); };
  const props = PropertiesService.getScriptProperties();
  const prueba = v2ParseJSON(props.getProperty('V2_PRUEBA'), { eventos: [], usuarios: [], contactos: [] });
  const sufijo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm:ss');

  // --- Usuarios de prueba
  _v2Reset();
  const uSin = v2Id('U'), uPin = v2Id('U');
  v2Append('Usuarios', [
    { id: uSin, nombre: 'Prueba ' + sufijo, pin: '', activo: true },
    { id: uPin, nombre: 'Prueba PIN ' + sufijo, pin: '4321', activo: true }
  ]);
  prueba.usuarios.push(uSin, uPin);
  props.setProperty('V2_PRUEBA', JSON.stringify(prueba));
  const yo = { usuarioId: uSin };

  // --- Teléfonos
  [['+54 9 11 1234-5678', '1112345678'], ['011 15 1234-5678', '1112345678'], ['+54 9 11 15 1234-5678', '1112345678'],
   ['2477 15 412345', '2477412345'], ['(0341) 155-123456', '3415123456'], ['3415123456', '3415123456'],
   ['5493415123456', '3415123456'], ['15 412345', '15412345']].forEach(function (c) {
    check('teléfono ' + c[0] + ' → ' + c[1], v2NormTel(c[0]) === c[1], v2NormTel(c[0]));
  });
  check('teléfono incompleto no ofrece WhatsApp', !v2TelValido(v2NormTel('15 412345')));

  // --- getInit / login
  const init = call('getInit');
  check('getInit responde', init.ok, init);
  check('getInit nunca devuelve el PIN', JSON.stringify(init).indexOf('4321') < 0 && JSON.stringify(init).indexOf('"pin"') < 0);
  check('getInit marca quién tiene PIN', init.usuarios.some(function (u) { return u.id === uPin && u.tienePin; }));
  check('login sin PIN entra', call('login', { usuarioId: uSin }).ok);
  check('login con PIN incorrecto no entra', call('login', { usuarioId: uPin, pin: '1111' }).error === 'PIN incorrecto.');
  check('login con PIN correcto entra', call('login', { usuarioId: uPin, pin: '4321' }).ok);
  check('escritura con PIN incorrecto rechazada', !call('saveEvento', { usuarioId: uPin, pin: '0000', evento: { nombre: 'X' } }).ok);
  check('escritura sin usuario rechazada', !call('saveEvento', { evento: { nombre: 'X' } }).ok);

  // --- Evento de comida (con cupo, sin comensales)
  let r = call('saveEvento', Object.assign({}, yo, {
    evento: { nombre: 'PRUEBA Comida ' + sufijo, fecha: '2026-11-21', hora: '21:00', lugar: 'Club Deportivo Mitre',
              campos: { apellido: true, sena: true, nota: true, retiro: true, entregado: true, comensales: false } },
    productos: [
      { nombre: 'Pollo con papas', precio: 10000, cupo: 3, cuentaComensales: false },
      { nombre: 'Empanadas (docena)', precio: 12000, cupo: '', cuentaComensales: false }
    ]
  }));
  check('crear evento de comida', r.ok && r.productos.length === 2, r);
  const evC = r.eventoId; prueba.eventos.push(evC); props.setProperty('V2_PRUEBA', JSON.stringify(prueba));
  const pollo = r.productos.filter(function (p) { return p.nombre === 'Pollo con papas'; })[0];
  const empa = r.productos.filter(function (p) { return p.nombre !== 'Pollo con papas'; })[0];
  check('el cupo vacío queda sin límite', empa.cupo === null && empa.quedan === null, empa);
  check('hora y fecha se guardan como texto', r.evento.fecha === '2026-11-21' && r.evento.hora === '21:00', r.evento);

  // --- Pedido con 2 productos y seña (precio "trucho" desde la app)
  const tel = '+54 9 11 ' + String(Math.floor(1000 + Math.random() * 8999)) + '-' + String(Math.floor(1000 + Math.random() * 8999));
  const idem = v2Id('K');
  r = call('savePedido', Object.assign({}, yo, {
    idemKey: idem,
    pedido: { eventoId: evC, nombre: 'Laura', apellido: 'Giménez', telefono: tel, nota: 'sin sal', retiro: '20:30' },
    items: [{ productoId: pollo.id, cantidad: 2, precioUnitario: 1 }, { productoId: empa.id, cantidad: 1, precio: 1 }],
    pagoInicial: { monto: 20000, medio: 'transferencia' }
  }));
  check('pedido con 2 productos y seña', r.ok, r);
  const ped1 = r.pedido;
  check('total calculado con precios de la hoja (no los de la app)', ped1.total === 32000, ped1);
  check('estado señado, saldo correcto', ped1.estado === 'senado' && ped1.pagado === 20000 && ped1.saldo === 12000, ped1);
  check('2 líneas en Items', r.items.length === 2);
  check('pago inicial tipo seña', r.pagos.length === 1 && r.pagos[0].tipo === 'sena' && r.pagos[0].medio === 'transferencia', r.pagos);
  check('queda anotado quién lo cargó', ped1.anoto === 'Prueba ' + sufijo);
  check('el cupo de Pollo baja a 1', r.productos.filter(function (p) { return p.id === pollo.id; })[0].quedan === 1);
  prueba.contactos.push(ped1.contactoId); props.setProperty('V2_PRUEBA', JSON.stringify(prueba));

  // --- Doble toque
  r = call('savePedido', Object.assign({}, yo, {
    idemKey: idem, pedido: { eventoId: evC, nombre: 'Laura', apellido: 'Giménez', telefono: tel },
    items: [{ productoId: pollo.id, cantidad: 2 }, { productoId: empa.id, cantidad: 1 }]
  }));
  check('doble toque no duplica', r.ok && r.repetido && r.pedido.id === ped1.id, r);

  // --- Contacto
  _v2Reset();
  const c1 = v2Table('Contactos').rows.filter(function (c) { return c.id === ped1.contactoId; })[0];
  check('contacto con teléfono normalizado a 10 dígitos', c1 && v2TelValido(c1.telefonoNorm), c1 && c1.telefonoNorm);
  r = call('buscarContacto', { texto: tel.replace(/\D/g, '').slice(-10) });
  check('buscar por teléfono encuentra a Laura', r.ok && r.contactos.length >= 1 && r.contactos[0].nombre === 'Laura', r);
  r = call('buscarContacto', { texto: 'laura gimenez' });
  check('buscar por nombre sin tilde encuentra a Laura', r.ok && r.contactos.some(function (c) { return c.id === ped1.contactoId; }), r);

  // --- Cupo
  r = call('savePedido', Object.assign({}, yo, { pedido: { eventoId: evC, nombre: 'Otro' }, items: [{ productoId: pollo.id, cantidad: 2 }] }));
  check('pedir más que el cupo da mensaje claro', !r.ok && r.error === 'Solo queda 1 de Pollo con papas.', r);
  r = call('savePedido', Object.assign({}, yo, { pedido: { eventoId: evC, nombre: 'Pedro', apellido: 'Sin Teléfono' }, items: [{ productoId: pollo.id, cantidad: 1 }] }));
  check('sin teléfono se guarda igual', r.ok && r.pedido.estado === 'debe', r);
  const ped2 = r.pedido;
  prueba.contactos.push(ped2.contactoId); props.setProperty('V2_PRUEBA', JSON.stringify(prueba));
  r = call('savePedido', Object.assign({}, yo, { pedido: { eventoId: evC, nombre: 'Otro' }, items: [{ productoId: pollo.id, cantidad: 1 }] }));
  check('producto sin cupo no se puede pedir', !r.ok && r.error === 'Se completó el cupo de Pollo con papas.', r);

  // --- Bajar cupo por debajo de lo vendido
  r = call('saveEvento', Object.assign({}, yo, { evento: { id: evC }, productos: [Object.assign({}, pollo, { cupo: 2 })] }));
  check('no deja bajar el cupo por debajo de lo vendido', !r.ok && r.error === 'Ya se vendieron 3 de Pollo con papas. El cupo no puede ser menor.', r);

  // --- Cobrar
  r = call('addPago', Object.assign({}, yo, { pedidoId: ped1.id, monto: 15000, medio: 'efectivo' }));
  check('cobrar más que el saldo pide confirmación', !r.ok && r.confirmar === true && r.error.indexOf('$12.000') >= 0, r);
  r = call('addPago', Object.assign({}, yo, { pedidoId: ped1.id, monto: 12000, medio: 'efectivo' }));
  check('cobrar el saldo pasa a Pagado', r.ok && r.pedido.estado === 'pagado' && r.pedido.saldo === 0, r);
  check('queda anotado quién cobró', r.pagos[1].cobro === 'Prueba ' + sufijo);

  // --- Entregado
  r = call('setEntregado', Object.assign({}, yo, { pedidoId: ped1.id, entregado: true }));
  check('marcar entregado', r.ok && r.pedido.entregado === true, r);

  // --- Precio congelado al editar
  r = call('saveEvento', Object.assign({}, yo, { evento: { id: evC }, productos: [Object.assign({}, empa, { precio: 15000, cupo: '' })] }));
  check('cambiar precio de Empanadas', r.ok, r);
  r = call('updatePedido', Object.assign({}, yo, {
    pedidoId: ped2.id, cambios: { telefono: '0341 15 555-1234' },
    items: [{ productoId: pollo.id, cantidad: 1 }, { productoId: empa.id, cantidad: 2 }]
  }));
  check('editar líneas recalcula el total (precio nuevo para línea nueva)', r.ok && r.pedido.total === 10000 + 30000, r.pedido || r);
  r = call('updatePedido', Object.assign({}, yo, { pedidoId: ped1.id, cambios: {}, items: [{ productoId: pollo.id, cantidad: 2 }, { productoId: empa.id, cantidad: 2 }] }));
  check('el precio de una línea ya cargada queda congelado', r.ok && r.pedido.total === 20000 + 24000, r.pedido || r);
  check('al subir el total vuelve a deber', r.ok && r.pedido.estado === 'senado' && r.pedido.saldo === 12000, r.pedido);

  // --- Contacto sin teléfono que después lo carga: no se duplica
  _v2Reset();
  const c2 = v2Table('Contactos').rows.filter(function (c) { return c.id === ped2.contactoId; })[0];
  check('al cargar el teléfono se completa el mismo contacto', c2 && c2.telefonoNorm === '3415551234', c2 && c2.telefonoNorm);

  // --- Anular
  r = call('anularPedido', Object.assign({}, yo, { pedidoId: ped2.id }));
  check('anular sin motivo pide el motivo', !r.ok, r);
  r = call('anularPedido', Object.assign({}, yo, { pedidoId: ped2.id, motivo: 'Prueba' }));
  check('anular libera el cupo', r.ok && r.productos.filter(function (p) { return p.id === pollo.id; })[0].quedan === 1, r);
  r = call('addPago', Object.assign({}, yo, { pedidoId: ped2.id, monto: 100 }));
  check('no se cobra un pedido anulado', !r.ok, r);

  // --- Evento de cena con comensales
  r = call('saveEvento', Object.assign({}, yo, {
    evento: { nombre: 'PRUEBA Cena ' + sufijo, fecha: '2026-12-12', campos: { comensales: true, retiro: false }, tiposComensal: ['Hombre', 'Mujer', 'Niño'] },
    productos: [{ nombre: 'Tarjeta', precio: 25000, cupo: 100, cuentaComensales: true }, { nombre: 'Vino', precio: 8000 }]
  }));
  check('crear evento de cena con comensales', r.ok && r.evento.campos.comensales, r);
  const evD = r.eventoId; prueba.eventos.push(evD); props.setProperty('V2_PRUEBA', JSON.stringify(prueba));
  const tarjeta = r.productos[0], vino = r.productos[1];
  r = call('savePedido', Object.assign({}, yo, {
    pedido: { eventoId: evD, nombre: 'Laura', apellido: 'Giménez', telefono: tel },
    items: [{ productoId: tarjeta.id, tipoComensal: 'Hombre', cantidad: 2 }, { productoId: tarjeta.id, tipoComensal: 'Mujer', cantidad: 1 },
            { productoId: tarjeta.id, tipoComensal: 'Niño', cantidad: 0 }, { productoId: vino.id, cantidad: 1 }],
    pagoInicial: { monto: 83000, medio: 'efectivo' }
  }));
  check('pedido con comensales: una línea por tipo', r.ok && r.items.length === 3 && r.pedido.estado === 'pagado', r);
  check('mismo teléfono → mismo contacto', r.ok && r.pedido.contactoId === ped1.contactoId, r.pedido);
  r = call('savePedido', Object.assign({}, yo, { pedido: { eventoId: evD, nombre: 'X' }, items: [{ productoId: tarjeta.id, cantidad: 1 }] }));
  check('tarjeta sin tipo de comensal pide el tipo', !r.ok && r.error.indexOf('tipo de comensal') >= 0, r);

  // --- Contactos
  _v2Reset();
  const c3 = v2Table('Contactos').rows.filter(function (c) { return c.id === ped1.contactoId; })[0];
  check('contacto que vino a 2 eventos: cantEventos 2', c3 && c3.cantEventos === 2, c3 && c3.cantEventos);
  r = call('getContactos', { eventoId: evC, productoId: pollo.id });
  check('filtro evento + producto', r.ok && r.total === 1 && r.contactos[0].id === ped1.contactoId, r);
  r = call('getContactos', { producto: 'tarjeta', eventoId: evD });
  check('filtro por nombre de producto', r.ok && r.total === 1, r);

  // --- Copiar y cerrar evento
  r = call('copiarEvento', Object.assign({}, yo, { eventoId: evC, nuevoNombre: 'PRUEBA Copia ' + sufijo, nuevaFecha: '2027-01-10' }));
  check('copiar evento copia productos sin pedidos', r.ok && r.productos.length === 2 && r.pedidos.length === 0 &&
    r.productos.every(function (p) { return p.vendidos === 0; }), r);
  const evCopia = r.eventoId; prueba.eventos.push(evCopia); props.setProperty('V2_PRUEBA', JSON.stringify(prueba));
  [evC, evD, evCopia].forEach(function (id) { call('saveEvento', Object.assign({}, yo, { evento: { id: id, estado: 'cerrado' } })); });
  r = call('savePedido', Object.assign({}, yo, { pedido: { eventoId: evC, nombre: 'Tarde' }, items: [{ productoId: empa.id, cantidad: 1 }] }));
  check('evento cerrado no acepta pedidos', !r.ok && r.error.indexOf('cerrado') >= 0, r);
  r = call('getEvento', { eventoId: evC });
  check('getEvento trae todo junto', r.ok && r.pedidos.length === 2 && r.items.length === 4 && r.pagos.length === 2, r.ok ? [r.pedidos.length, r.items.length, r.pagos.length] : r);

  // --- Resumen
  const fallas = res.filter(function (l) { return l.indexOf('FALLA') === 0; }).length;
  res.forEach(function (l) { Logger.log(l); });
  Logger.log(fallas ? '✘ ' + v2Plural(fallas, 'prueba falló', 'pruebas fallaron') : '✔ Todas las pruebas pasaron (' + res.length + ').');
  Logger.log('Para borrar lo que creó esta prueba, ejecutá borrarDatosDePrueba.');
  return { ok: !fallas, resultados: res };
}

// Deja un producto con 1 solo lugar y muestra un código para pegar en la consola del
// navegador que manda dos pedidos a la vez: solo uno tiene que pasar.
function prepararPruebaCupo() {
  _v2Reset();
  const props = PropertiesService.getScriptProperties();
  const prueba = v2ParseJSON(props.getProperty('V2_PRUEBA'), { eventos: [], usuarios: [], contactos: [] });
  const uid = v2Id('U');
  v2Append('Usuarios', [{ id: uid, nombre: 'Prueba cupo', pin: '', activo: true }]);
  prueba.usuarios.push(uid);
  const r = runV2('saveEvento', JSON.stringify({ usuarioId: uid, evento: { nombre: 'PRUEBA Cupo', fecha: '2026-12-31' },
    productos: [{ nombre: 'Último lugar', precio: 1000, cupo: 1 }] }));
  prueba.eventos.push(r.eventoId);
  props.setProperty('V2_PRUEBA', JSON.stringify(prueba));
  const url = ScriptApp.getService().getUrl();
  const mk = function (n) {
    return JSON.stringify({ usuarioId: uid, idemKey: 'cupo' + n + Date.now(), pedido: { eventoId: r.eventoId, nombre: 'Persona ' + n },
      items: [{ productoId: r.productos[0].id, cantidad: 1 }] });
  };
  const code = "const u='" + url + "';Promise.all([" + [1, 2].map(function (n) {
    return "fetch(u+'?action=savePedido&data='+encodeURIComponent('" + mk(n).replace(/'/g, "\\'") + "')).then(r=>r.json())";
  }).join(',') + "]).then(r=>console.log(r.map(x=>x.ok?'PASÓ':'NO PASÓ: '+x.error)))";
  Logger.log('Abrí la app en el navegador, abrí la consola (F12) y pegá esto:');
  Logger.log(code);
  Logger.log('Tiene que decir una vez PASÓ y otra NO PASÓ: Se completó el cupo de Último lugar.');
}

// Borra todo lo que crearon testV2 y prepararPruebaCupo (eventos, productos, pedidos,
// items, pagos, contactos y usuarios de prueba). No toca nada más.
function borrarDatosDePrueba() {
  _v2Reset();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  _v2HoldsLock = true;
  try {
    const props = PropertiesService.getScriptProperties();
    const prueba = v2ParseJSON(props.getProperty('V2_PRUEBA'), { eventos: [], usuarios: [], contactos: [] });
    const ev = {}; prueba.eventos.forEach(function (id) { ev[id] = true; });
    const us = {}; prueba.usuarios.forEach(function (id) { us[id] = true; });
    // Contactos de los pedidos de prueba, salvo que también tengan pedidos reales
    const ct = {};
    prueba.contactos.forEach(function (id) { ct[id] = true; });
    v2Table('Pedidos').rows.forEach(function (p) { if (ev[p.eventoId]) ct[p.contactoId] = true; });
    v2Table('Pedidos').rows.forEach(function (p) { if (!ev[p.eventoId]) delete ct[p.contactoId]; });
    let total = 0;
    const borrar = function (name, cond) {
      const rows = v2Table(name).rows.filter(cond).map(function (r) { return r._row; });
      v2DeleteRows(name, rows); total += rows.length;
      Logger.log(name + ': ' + rows.length);
    };
    borrar('Items', function (r) { return ev[r.eventoId]; });
    borrar('Pagos', function (r) { return ev[r.eventoId]; });
    borrar('Pedidos', function (r) { return ev[r.eventoId]; });
    borrar('Productos', function (r) { return ev[r.eventoId]; });
    borrar('Eventos', function (r) { return ev[r.id]; });
    borrar('Contactos', function (r) { return ct[r.id]; });
    borrar('Usuarios', function (r) { return us[r.id]; });
    props.deleteProperty('V2_PRUEBA');
    Logger.log('Listo: se borraron ' + total + ' filas de prueba.');
  } finally {
    lock.releaseLock();
    _v2HoldsLock = false;
  }
}
