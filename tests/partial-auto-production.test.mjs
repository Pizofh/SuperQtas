import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';

const code = fs.readFileSync(new URL('../QTAS_Inventario.gs', import.meta.url), 'utf8');

function fixture() {
  const sheets = {
    Producciones: [{
      Produccion_ID: 'PRD-000002', Producto_Estandar: '100mg', Unidad: 'und',
      Cantidad_Producida: 1, Estado_Registro: 'Activo',
      Comentario_Produccion: 'Produccion automatica por venta V2 | DET-000002-01 | Desde receta base'
    }],
    Produccion_Detalle: [],
    Inventario_Movimientos: []
  };
  const ss = { getSheetByName: name => ({ name, rows: sheets[name] }) };
  let reconstructionCount = 0;
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActive: () => ss },
    QTAS: { sheets: {
      producciones: 'Producciones', produccionDetalle: 'Produccion_Detalle',
      inventarioMovimientos: 'Inventario_Movimientos'
    }},
    texto_: v => String(v ?? ''),
    numero_: v => Number(v) || 0,
    redondear_: v => Math.round((Number(v) || 0) * 10000) / 10000,
    resolverFechaOperacion_: v => new Date(v),
    combinarFechaYHora_: (v) => v,
    normalizarUnidadCanonicaQTAS_: v => v,
    esRegistroAnulado_: v => v === 'Anulado'
  });
  vm.runInContext(code, sandbox);
  const detalleEsperado = [
    { Produccion_ID: 'PRD-000002', Produccion_Detalle_ID: 'D-1',
      Operacion: 'Entrada', Tipo_Item: 'Producto', Item: '100mg',
      Unidad: 'und', Cantidad: 1 },
    { Produccion_ID: 'PRD-000002', Produccion_Detalle_ID: 'D-2',
      Operacion: 'Salida', Tipo_Item: 'Insumo', Item: 'Capsulas',
      Unidad: 'und', Cantidad: 1 }
  ];
  const asMovimientos = rows => rows.map(row => ({
    Produccion_ID: 'PRD-000002', Detalle_ID: row.Produccion_Detalle_ID,
    Operacion: row.Operacion, Tipo_Item: row.Tipo_Item, Item: row.Item,
    Unidad: row.Unidad, Cantidad: row.Cantidad,
    Cantidad_Signada: row.Operacion === 'Entrada' ? row.Cantidad : -row.Cantidad
  }));
  Object.assign(sandbox, {
    construirIndiceControlesInventarioQTAS_: () => ({}),
    listarControlesInventarioQTAS_: () => [],
    obtenerControlInventarioEfectivoQTAS_: () => ({ modoStock: 'Fabricado' }),
    leerComponentesProductoActivosQTAS_: () => [],
    esMismaClaveProductoQTAS_: (a, unitA, b, unitB) => a === b && unitA === unitB,
    getHeaders_: () => [],
    siguienteIdConPrefijoPersistenteQTAS_: () => { throw new Error('Nunca reservar otro ID.'); },
    construirDetalleProduccionMaterializadoQTAS_: () => detalleEsperado,
    construirDisponibilidadCalcasInventarioQTAS_: () => ({}),
    leerObjetos_: sheet => [...sheet.rows],
    filaDesdeHeaders_: (_headers, row) => row,
    escribirFilas_: (sheet, rows) => { sheet.rows.push(...rows); },
    construirMovimientosInventarioDesdeProduccionDetalleRowsQTAS_: asMovimientos,
    sincronizarInventarioDesdeProduccionDetalleQTAS_: ({ detalleRows }) => {
      sheets.Inventario_Movimientos.push(...asMovimientos(detalleRows));
      return { ok: true, movimientos: detalleRows.length };
    },
    reconstruirSnapshotInventarioQTAS_: () => { reconstructionCount += 1; return []; }
  });
  const call = () => sandbox.registrarProduccionInternaQTAS_({
    ss, produccionIdExistente: 'PRD-000002',
    producto: '100mg', unidad: 'und', cantidad: 1,
    componentes: [], componentesProduccion: [ { producto:'100mg', unidadVenta:'und' } ],
    fechaProduccion: '2026-06-20', comentarioProduccion: 'auto test'
  });
  return { call, sheets, expected: detalleEsperado, getRebuildCount: () => reconstructionCount };
}

test('se recupera encabezado incompleto sin crear otra produccion', () => {
  const h = fixture();
  const result = h.call();
  assert.equal(result.ok, true);
  assert.equal(result.recuperada, true);
  assert.equal(result.produccionId, 'PRD-000002');
  assert.equal(h.sheets.Producciones.length, 1);
  assert.equal(h.sheets.Produccion_Detalle.length, 2);
  assert.equal(h.sheets.Inventario_Movimientos.length, 2);
  assert.equal(h.getRebuildCount(), 1);
});

test('reintentar la misma recuperacion no duplica detalles ni movimientos', () => {
  const h = fixture();
  h.call();
  const result = h.call();
  assert.equal(result.ok, true);
  assert.equal(result.inventario.movimientos, 0);
  assert.equal(h.sheets.Produccion_Detalle.length, 2);
  assert.equal(h.sheets.Inventario_Movimientos.length, 2);
  assert.equal(h.sheets.Producciones.length, 1);
});

test('completa solo el movimiento faltante de una produccion parcialmente persistida', () => {
  const h = fixture();
  h.call();
  h.sheets.Inventario_Movimientos.pop();
  h.call();
  assert.equal(h.sheets.Inventario_Movimientos.length, 2);
  assert.equal(h.sheets.Produccion_Detalle.length, 2);
});

test('no sobrescribe una produccion cuyo movimiento es inconsistente', () => {
  const h = fixture();
  h.call();
  h.sheets.Inventario_Movimientos[0].Cantidad = 999;
  assert.throws(() => h.call(), /Cantidad inconsistente/);
  assert.equal(h.sheets.Inventario_Movimientos.length, 2);
});

test('no acepta encabezados de produccion con cantidad distinta', () => {
  const h = fixture();
  h.sheets.Producciones[0].Cantidad_Producida = 2;
  assert.throws(() => h.call(), /no coincide/);
  assert.equal(h.sheets.Produccion_Detalle.length, 0);
});
