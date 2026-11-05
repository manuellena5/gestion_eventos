// Simulador mínimo de Apps Script para probar Code.gs con Node.
const fs = require('fs'), vm = require('vm');
process.env.TZ = 'America/Argentina/Buenos_Aires';

function autoConvert(v) {
  // Emula la conversión automática de Sheets cuando la celda NO tiene formato texto
  if (typeof v !== 'string') return v;
  if (/^-?\d+(\.\d+)?$/.test(v.trim()) && v.trim() !== '') return Number(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) { const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d); }
  if (/^\d{1,2}:\d{2}$/.test(v)) { const [h, mi] = v.split(':').map(Number); return new Date(1899, 11, 30, h, mi); }
  if (/^(true|false)$/i.test(v)) return v.toLowerCase() === 'true';
  return v;
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; this.fmt = {}; this.maxRows = 1000; this.frozen = 0; this.ops = 0; }
  getName() { return this.name; }
  getLastRow() { for (let r = this.data.length; r > 0; r--) if ((this.data[r - 1] || []).some(v => v !== '' && v !== null && v !== undefined)) return r; return 0; }
  getLastColumn() { let m = 0; this.data.forEach(row => { for (let c = row.length; c > 0; c--) if (row[c - 1] !== '' && row[c - 1] !== undefined) { m = Math.max(m, c); break; } }); return m; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(after, n) { this.maxRows += n; }
  deleteRow(r) { this.ops++; this.data.splice(r - 1, 1); const nf = {}; Object.keys(this.fmt).forEach(k => { const [rr, cc] = k.split(',').map(Number); if (rr < r) nf[k] = this.fmt[k]; else if (rr > r) nf[(rr - 1) + ',' + cc] = this.fmt[k]; }); this.fmt = nf; this.maxRows--; }
  setFrozenRows(n) { this.frozen = n; }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  get(r, c) { const row = this.data[r - 1]; const v = row ? row[c - 1] : undefined; return v === undefined || v === null ? '' : v; }
  set(r, c, v) { if (r > this.maxRows) throw new Error('Fuera de rango: fila ' + r); while (this.data.length < r) this.data.push([]); const row = this.data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = this.fmt[r + ',' + c] === '@' ? (v === null || v === undefined ? '' : (v instanceof Date || typeof v === 'boolean' ? v : String(v))) : autoConvert(v); }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const st = Range.stats; if (st && st.failEvery && ++st.reads % st.failEvery === 0) { st.failed++; throw new Error('Exception: El servicio Hojas de cálculo no ha podido acceder al documento con ID X.'); } this.sh.ops++; const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = this.sh.get(this.r + i, this.c + j); row.push(v instanceof Date ? new Date(v) : v); } out.push(row); } return out; }
  setValues(vals) { this.sh.ops++; if (vals.length !== this.nr || vals.some(r => r.length !== this.nc)) throw new Error('Tamaño incorrecto en setValues: ' + vals.length + 'x' + (vals[0] || []).length + ' vs ' + this.nr + 'x' + this.nc); vals.forEach((row, i) => row.forEach((v, j) => this.sh.set(this.r + i, this.c + j, v))); return this; }
  setValue(v) { this.sh.set(this.r, this.c, v); return this; }
  setNumberFormat(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.fmt[(this.r + i) + ',' + (this.c + j)] = f; return this; }
  setFontWeight() { return this; } setBackground() { return this; } setFontColor() { return this; }
}
class SS {
  constructor() { this.sheets = []; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
  getSheets() { return this.sheets.slice(); }
}

function pad(n, l = 2) { return String(n).padStart(l, '0'); }
function formatDate(d, tz, f) {
  return f.replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate()))
    .replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes())).replace('ss', pad(d.getSeconds()));
}

function makeEnv() {
  const ss = new SS();
  const props = {}; const cache = {};
  const lockState = { held: false, log: [] };
  const logs = [];
  const stats = { tz: 0, failEvery: 0, reads: 0, failed: 0 };
  Range.stats = stats;
  const ctx = {
    console: { log: (...a) => logs.push(a.join(' ')), error: (...a) => logs.push('ERROR ' + a.join(' ')) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
    LockService: { getScriptLock: () => ({
      tryLock: () => { if (lockState.held) return false; lockState.held = true; lockState.log.push('lock'); return true; },
      waitLock: () => { if (lockState.held) throw new Error('lock ocupado'); lockState.held = true; },
      releaseLock: () => { lockState.held = false; lockState.log.push('unlock'); } }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] === undefined ? null : props[k], setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
    Utilities: { formatDate, sleep: () => {} },
    Session: { getScriptTimeZone: () => { stats.tz++; return 'America/Argentina/Buenos_Aires'; } },
    Logger: { log: (s) => logs.push(String(s)) },
    ContentService: { createTextOutput: s => ({ content: s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
    Date, Math, JSON, Object, Array, String, Number, isFinite, isNaN, RegExp, Error,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(require('path').join(__dirname, '..', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
  return { ss, ctx, props, cache, logs, lockState, stats, run: code => vm.runInContext(code, ctx) };
}
module.exports = { makeEnv, Sheet };
