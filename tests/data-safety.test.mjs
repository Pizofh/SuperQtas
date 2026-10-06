import fs from 'node:fs';
import vm from 'node:vm';
import crypto from 'node:crypto';
import assert from 'node:assert/strict';
import test from 'node:test';

function iterator(items) { let i = 0; return { hasNext: () => i < items.length, next: () => items[i++] }; }
class Sheet {
  constructor(name, grid, hidden = []) {
    this.name = name;
    this.records = grid.map((values, i) => ({ values: structuredClone(values), hidden: hidden.includes(i + 1), format: 'unchanged' }));
    this.id = name;
    this.deleted = [];
    this.filtered = true;
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getLastRow() { return this.records.length; }
  getLastColumn() { return this.records[0].values.length; }
  getFilter() { return this.filtered ? { description: 'Medio_Pago = NequiSteve' } : null; }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  getRange(r, c, n, m) {
    return {
      getValues: () => this.records.slice(r - 1, r - 1 + n).map(row => structuredClone(row.values.slice(c - 1, c - 1 + m))),
      getFormulas: () => Array.from({ length: n }, () => Array(m).fill('')),
      getNumRows: () => n, getNumColumns: () => m,
      clearContent: () => { throw new Error('Prohibido vaciar historicos'); },
      setValues: () => { throw new Error('Prohibido volver a escribir historicos'); }
    };
  }
  deleteRows(r, n) {
    assert.ok(r >= 2, 'Nunca borrar encabezados');
    this.deleted.push([r, n]);
    this.records.splice(r - 1, n);
  }
  clone() { const s = new Sheet(this.name, []); s.records = structuredClone(this.records); return s; }
}
function environment() {
  const clock = { now: '2026-10-06T08:30:00Z' };
  function ClockDate(...args) { return args.length ? new Date(...args) : new Date(clock.now); }
  ClockDate.prototype = Date.prototype;
  ClockDate.now = () => new Date(clock.now).getTime();
  const store = new Map(), properties = new Map();
  const sheets = [new Sheet('Ventas', [['Venta_ID', 'Fecha_Venta'], [1, new Date('2026-01-01T05:00:00Z')]])];
  const source = { getId: () => 'source', getName: () => 'SUPERQTAS', getSpreadsheetTimeZone: () => 'America/Bogota', getSheets: () => sheets };
  let active = source, sequence = 0;
  const files = [], triggers = [];
  const faults = { copy: false, corrupt: false, trigger: false };
  const folder = { getId: () => 'folder', getName: () => 'QTAS_Backups', getFiles: () => iterator(files.filter(f => !f.trashed)), getFoldersByName: () => iterator([]), createFolder: () => folder };
  function file(id, name, spreadsheet, description = '') {
    const f = { id, name, spreadsheet, description, trashed: false, created: new Date(clock.now),
      getId: () => id, getName: () => f.name, getDescription: () => f.description,
      setName: name => { f.name = name; }, setDescription: text => { f.description = text; },
      getDateCreated: () => f.created, setTrashed: value => { assert.notEqual(id, 'source'); f.trashed = value; },
      getParents: () => iterator([folder]),
      makeCopy: name => {
        if (faults.copy) throw new Error('makeCopy fallo');
        const copiedSheets = spreadsheet.getSheets().map(s => s.clone());
        if (faults.corrupt) copiedSheets[0].records[1].values[0] = 'CORRUPTO';
        const copied = { getSpreadsheetTimeZone: spreadsheet.getSpreadsheetTimeZone, getSheets: () => copiedSheets };
        const newFile = file(`copy-${++sequence}`, name, copied);
        files.push(newFile); return newFile;
      }
    };
    store.set(id, f); return f;
  }
  file('source', 'SUPERQTAS', source);
  const props = {
    getProperty: k => properties.get(k) ?? null,
    setProperty: (k, v) => { properties.set(k, v); return props; },
    setProperties: data => { Object.entries(data).forEach(([k, v]) => properties.set(k, v)); return props; },
    deleteProperty: k => { properties.delete(k); return props; }
  };
  function formatDate(date, timeZone, pattern) {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(date).map(x => [x.type, x.value]));
    const day = `${parts.year}-${parts.month}-${parts.day}`;
    return pattern === 'yyyy-MM-dd' ? day : `${day}_${parts.hour}-${parts.minute}-${parts.second}`;
  }
  const context = vm.createContext({ Date: ClockDate, console, Set,
    SpreadsheetApp: { getActive: () => active, openById: id => { if (!store.has(id)) throw new Error('Missing'); return store.get(id).spreadsheet; } },
    DriveApp: { getFileById: id => store.get(id), getFolderById: id => { assert.equal(id, 'folder'); return folder; }, getRootFolder: () => folder },
    PropertiesService: { getScriptProperties: () => props },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: { formatDate, getUuid: () => `uuid${++sequence}`, DigestAlgorithm: { SHA_256: 'sha256' },
      computeDigest: (algorithm, data) => [...crypto.createHash(algorithm).update(data).digest()],
      base64Encode: bytes => Buffer.from(bytes).toString('base64') },
    ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger: t => { triggers.splice(triggers.indexOf(t), 1); },
      newTrigger: handler => {
        const config = { handler };
        const builder = { timeBased: () => builder, everyDays: days => { config.days = days; return builder; },
          atHour: hour => { config.hour = hour; return builder; }, inTimezone: zone => { config.zone = zone; return builder; },
          create: () => { if (faults.trigger) throw new Error('trigger fallo'); const t = { ...config, getHandlerFunction: () => handler, getUniqueId: () => `trigger-${++sequence}` }; triggers.push(t); return t; } };
        return builder;
      } }
  });
  for (const name of ['QTAS_Utils.gs', 'QTAS_Backups.gs', 'QTAS_Ventas.gs', 'QTAS_Inventario.gs']) {
    vm.runInContext(fs.readFileSync(new URL(`../${name}`, import.meta.url), 'utf8'), context, { filename: name });
  }
  context.getHeaders_ = sheet => sheet.records[0].values.slice();
  return { context, clock, files, properties, source, sheets, triggers, faults, file,
    setActive: value => { active = value; } };
}

test('6,7,8,9: conserva 7,8,9 y retira el 6 solo despues de verificar', () => {
  const e = environment(), ids = [];
  for (const day of ['06', '07', '08', '09']) {
    e.clock.now = `2026-10-${day}T08:30:00Z`;
    ids.push(e.context.ejecutarBackupDiarioQTAS().backupFileId);
    assert.equal(e.files.filter(f => !f.trashed).length, Math.min(ids.length, 3));
  }
  assert.equal(e.files.find(f => f.id === ids[0]).trashed, true);
  assert.deepEqual(e.files.filter(f => !f.trashed).map(f => JSON.parse(f.description).day), ['2026-10-07', '2026-10-08', '2026-10-09']);
});
test('reintentar el mismo dia conserva la foto original, incluso con nuevas ventas', () => {
  const e = environment(), first = e.context.ejecutarBackupDiarioQTAS();
  e.sheets[0].records.push({ values: [2, new Date('2026-10-06T12:00:00Z')] });
  const retry = e.context.ejecutarBackupDiarioQTAS();
  assert.equal(first.backupFileId, retry.backupFileId); assert.equal(retry.created, false);
  assert.equal(e.files.length, 1); assert.equal(e.files[0].spreadsheet.getSheets()[0].records.length, 2);
});
test('un fallo al copiar conserva todos los backups anteriores', () => {
  const e = environment();
  for (const day of ['06', '07', '08']) { e.clock.now = `2026-10-${day}T08:30:00Z`; e.context.ejecutarBackupDiarioQTAS(); }
  e.clock.now = '2026-10-09T08:30:00Z'; e.faults.copy = true;
  assert.throws(() => e.context.ejecutarBackupDiarioQTAS(), /makeCopy/);
  assert.ok(e.files.every(f => !f.trashed)); assert.match(e.properties.get('QTAS_BACKUP_LAST_ERROR'), /makeCopy/);
});
test('una copia incompleta no causa rotacion ni se marca como verificada', () => {
  const e = environment(); e.context.ejecutarBackupDiarioQTAS(); e.clock.now = '2026-10-07T08:30:00Z'; e.faults.corrupt = true;
  assert.throws(() => e.context.ejecutarBackupDiarioQTAS(), /no coincide/);
  assert.ok(e.files.every(f => !f.trashed)); assert.ok(e.files[1].name.endsWith('__PENDIENTE')); assert.equal(e.files[1].description, '');
});
test('rotacion ignora manuales, antiguos, evidencia y copias de otros libros', () => {
  const e = environment(); const manual = e.context.crearBackupManualQTAS();
  const other = e.file('other', 'QTAS_BACKUP_DIARIO__other__2026-10-01__OTRO', e.source);
  const legacy = e.file('legacy', 'QTAS_BACKUP_ACTIVO__SUPERQTAS', e.source);
  const evidence = e.file('evidence', 'SUPERQTAS_EVIDENCIA_2026-09-29', e.source);
  e.files.push(other, legacy, evidence);
  for (const day of ['06', '07', '08', '09']) { e.clock.now = `2026-10-${day}T08:30:00Z`; e.context.ejecutarBackupDiarioQTAS(); }
  assert.ok([other, legacy, evidence, e.files.find(f => f.id === manual.backupFileId)].every(f => !f.trashed));
});
test('el dia del backup se calcula en Bogota, no en UTC', () => {
  const e = environment(); e.clock.now = '2026-10-07T02:00:00Z';
  assert.equal(e.context.ejecutarBackupDiarioQTAS().day, '2026-10-06');
});
test('el activador funciona sin un libro activo y se instala a las 0:00', () => {
  const e = environment(); const first = e.context.instalarBackupDiarioQTAS({ hour: 0 });
  assert.equal(first.hour, 0); assert.equal(e.triggers[0].zone, 'America/Bogota');
  e.setActive(null); assert.equal(e.context.ejecutarBackupDiarioQTAS().spreadsheetId, 'source');
  e.context.instalarBackupDiarioQTAS({ hour: 3 }); assert.equal(e.triggers.length, 1);
});
test('fallar al crear un activador no elimina el anterior', () => {
  const e = environment(); e.context.instalarBackupDiarioQTAS(); e.faults.trigger = true;
  assert.throws(() => e.context.instalarBackupDiarioQTAS(), /trigger fallo/); assert.equal(e.triggers.length, 1);
});
test('un libro activo distinto del configurado se rechaza sin copiar ni rotar', () => {
  const e = environment(); e.context.instalarBackupDiarioQTAS(); e.setActive({ getId: () => 'other' });
  assert.throws(() => e.context.ejecutarBackupDiarioQTAS(), /origen configurado/); assert.equal(e.files.length, 0);
});
test('un backup previo alterado impide retirar el mas antiguo', () => {
  const e = environment();
  for (const day of ['06', '07', '08']) { e.clock.now = `2026-10-${day}T08:30:00Z`; e.context.ejecutarBackupDiarioQTAS(); }
  e.files[2].spreadsheet.getSheets()[0].records[1].values[0] = 'ALTERADO'; e.clock.now = '2026-10-09T08:30:00Z';
  assert.throws(() => e.context.ejecutarBackupDiarioQTAS(), /no coincide/); assert.ok(e.files.every(f => !f.trashed));
});

test('borrado selectivo con filtro: fechas, formatos y formulas de supervivientes intactos', () => {
  const e = environment(); const sheet = new Sheet('Distribucion_Ingresos', [
    ['Venta_ID', 'Fecha_Base', 'Fecha_Venta', 'Fecha_Pago', 'Medio_Pago'],
    [1, new Date('2026-01-01T05:00:00Z'), new Date('2026-01-01T05:00:00Z'), '', 'Daviplata'],
    [9, new Date('2026-10-04T20:21:00Z'), new Date('2026-10-04T20:21:00Z'), '', 'NequiSteve'],
    [2, new Date('2026-02-02T05:00:00Z'), new Date('2026-02-02T05:00:00Z'), '', 'NequiMajo'],
    [9, new Date('2026-10-04T20:21:00Z'), new Date('2026-10-04T20:21:00Z'), new Date('2026-10-04T20:21:00Z'), 'Daviplata'],
    [9, new Date('2026-10-04T20:21:00Z'), new Date('2026-10-04T20:21:00Z'), '', 'Bancolombia']
  ], [2, 4, 5, 6]);
  sheet.records[3].formula = '=1+2';
  const retained = structuredClone([sheet.records[0], sheet.records[1], sheet.records[3]]);
  const rows = e.context.leerObjetosConMeta_(sheet).filter(r => r.Venta_ID === 9);
  const plan = e.context.planificarEliminacionFilasQTAS_(sheet, sheet.records[0].values, rows);
  assert.equal(e.context.ejecutarEliminacionFilasQTAS_(plan), 3);
  assert.deepEqual(sheet.records, retained); assert.deepEqual(sheet.deleted, [[5, 2], [3, 1]]);
});
test('rechaza encabezados, filas duplicadas y filas desactualizadas antes de borrar', () => {
  const e = environment(); const sheet = new Sheet('Ventas', [['Venta_ID'], [1], [2]]); const rows = e.context.leerObjetosConMeta_(sheet);
  assert.throws(() => e.context.planificarEliminacionFilasQTAS_(sheet, ['Venta_ID'], [{ ...rows[0], __rowNumber: 1 }]), /invalida/);
  assert.throws(() => e.context.planificarEliminacionFilasQTAS_(sheet, ['Venta_ID'], [rows[0], rows[0]]), /duplicada/);
  sheet.records[1].values[0] = 100;
  assert.throws(() => e.context.planificarEliminacionFilasQTAS_(sheet, ['Venta_ID'], rows), /cambio/); assert.deepEqual(sheet.deleted, []);
});
test('valida otra vez las filas inmediatamente antes de borrarlas', () => {
  const e = environment(); const sheet = new Sheet('Ventas', [['Venta_ID'], [1]]);
  const plan = e.context.planificarEliminacionFilasQTAS_(sheet, ['Venta_ID'], e.context.leerObjetosConMeta_(sheet));
  sheet.records[1].values[0] = 3;
  assert.throws(() => e.context.ejecutarEliminacionFilasQTAS_(plan), /cambiaron/); assert.deepEqual(sheet.deleted, []);
});
test('eliminarVentaReciente recorre todas sus tablas sin vaciar historicos', () => {
  const e = environment(), c = e.context, tables = new Map();
  const names = { ventas: 'Ventas', detalle: 'Venta_Detalle', pagos: 'Pagos', distribucionIngresos: 'Distribucion_Ingresos', ventasEnvio: 'Ventas_Envio', ventaDetalleCostosCalculado: 'Venta_Detalle_Costos_Calc', inventarioMovimientos: 'Inventario_Movimientos' };
  for (const name of Object.values(names)) tables.set(name, new Sheet(name, [
    ['Venta_ID', 'Detalle_ID', 'Nombre', 'Estado_Registro', 'Fecha_Venta', 'Fecha_Pago', 'Medio_Pago'],
    [1, 'old1', 'Historica Majo', 'Activo', new Date('2025-01-01T05:00:00Z'), new Date('2025-01-02T05:00:00Z'), 'NequiMajo'],
    [9, 'new1', 'Venta reciente', 'Activo', new Date('2026-10-04T01:02:03.456Z'), '', 'NequiSteve'],
    [2, 'old2', 'Historica Steve', 'Activo', new Date('2025-02-01T05:00:00Z'), '', 'NequiSteve']
  ], [2]));
  const survivors = new Map([...tables].map(([name, s]) => [name, structuredClone([s.records[0], s.records[1], s.records[3]])]));
  const ss = { getSheetByName: name => tables.get(name) };
  c.SpreadsheetApp.getActive = () => ss;
  c.QTAS = { sheets: names, schemas: {}, status: { registro: { anulado: 'Anulado' } } };
  c.validarModeloSoloLecturaQTAS_ = () => {};
  c.construirEstadoVentasQTAS_ = () => ({});
  c.ventasRecientesDesdeEstadoQTAS_ = () => [{ ventaId: 9 }];
  c.obtenerHojaVentasEnvioQTAS_ = () => tables.get(names.ventasEnvio);
  c.resolverHojaCanonicaOperativaQTAS_ = (_, name) => ({ ok: true, sheet: tables.get(name), headers: tables.get(name).records[0].values });
  c.resincronizarIdNumericoPersistenteQTAS_ = () => 2;
  c.reconstruirSnapshotInventarioQTAS_ = () => {};
  c.listarSnapshotInventarioQTAS_ = () => [];
  c.limpiarCachesEjecucionQTAS_ = () => {};
  c.dashboardVentasConsistenteQTAS_ = () => ({});
  const result = c.eliminarVentaRecienteQTAS({ ventaId: 9 });
  assert.equal(result.ok, true); assert.equal(result.removed.distribucionIngresos, 1);
  assert.equal(result.inventario.movimientos, 1);
  for (const [name, sheet] of tables) assert.deepEqual(sheet.records, survivors.get(name), name);
});


test('una zona horaria vacia usa Bogota al copiar e instalar el activador', () => {
  const e = environment(); e.source.getSpreadsheetTimeZone = () => null;
  assert.equal(e.context.ejecutarBackupDiarioQTAS().day, '2026-10-06');
  assert.equal(e.context.instalarBackupDiarioQTAS().timeZone, 'America/Bogota');
  assert.equal(e.triggers[0].zone, 'America/Bogota');
});
