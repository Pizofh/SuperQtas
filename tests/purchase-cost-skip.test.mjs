import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';

const code = fs.readFileSync(new URL('../QTAS_Compras.gs', import.meta.url), 'utf8');

function harness(detalle) {
  let rebuilds = 0;
  let inventoryCalls = 0;
  const sheet = { name: 'Compra_Detalle' };
  const ss = { getSheetByName: name => name === 'Compra_Detalle' ? sheet : null };
  const ctx = vm.createContext({
    Date, Logger: { log() {} },
    QTAS: { sheets: { compraDetalle: 'Compra_Detalle' } },
    SpreadsheetApp: { getActive: () => ss },
    numero_: value => Number(value) || 0,
    normalizarClaveTexto_: value => String(value ?? '').trim().toLowerCase(),
    esRegistroAnulado_: () => false,
    leerObjetos_: () => detalle,
    fechaInput_: date => date.toISOString().slice(0, 10),
    sincronizarInventarioDesdeCompraQTAS_: () => {
      inventoryCalls += 1;
      return { ok: true, movimientos: 0 };
    },
    reconstruirCostoProductoCalculadoInternoQTAS_: () => {
      rebuilds += 1;
      return { ok: true, rows: 1, skipped: false };
    }
  });
  vm.runInContext(code, ctx, { filename: 'QTAS_Compras.gs' });
  return { ctx, ss, getRebuilds: () => rebuilds, getInventoryCalls: () => inventoryCalls };
}

test('una compra exclusivamente de gastos no reconstruye costos calculados', () => {
  const h = harness([{
    Compra_ID: 138,
    Compra_Detalle_ID: 'CD-138-01',
    Tipo_Item: 'Gasto',
    Impacta_Costo: false,
    Costo_Unitario: 50000
  }]);
  const result = h.ctx.procesarPostCompraDetalleQTAS_(h.ss, [138], new Date());
  assert.equal(h.getRebuilds(), 0);
  assert.equal(h.getInventoryCalls(), 1);
  assert.equal(result.costoProductoCalculado.ok, true);
  assert.equal(result.costoProductoCalculado.skipped, true);
  assert.equal(result.costoProductoCalculado.rows, 0);
});

test('una compra con productos costeables conserva el recalculo de costos', () => {
  const h = harness([{
    Compra_ID: 138,
    Compra_Detalle_ID: 'CD-138-01',
    Tipo_Item: 'Producto',
    Impacta_Costo: true,
    Costo_Unitario: 5000
  }]);
  const result = h.ctx.procesarPostCompraDetalleQTAS_(h.ss, [138], new Date());
  assert.equal(h.getRebuilds(), 1);
  assert.equal(h.getInventoryCalls(), 1);
  assert.equal(result.costoProductoCalculado.skipped, false);
});

test('una compra mixta con gasto y producto conserva el recalculo', () => {
  const h = harness([
    { Compra_ID: 138, Tipo_Item: 'Gasto', Impacta_Costo: false, Costo_Unitario: 15000 },
    { Compra_ID: 138, Tipo_Item: 'Insumo', Impacta_Costo: true, Costo_Unitario: 1200 }
  ]);
  h.ctx.procesarPostCompraDetalleQTAS_(h.ss, [138], new Date());
  assert.equal(h.getRebuilds(), 1);
});
