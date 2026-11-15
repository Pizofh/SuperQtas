import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';

function context() {
  const sandbox = vm.createContext({ console });
  for (const name of ['QTAS_Utils.gs', 'QTAS_Inventario.gs']) {
    vm.runInContext(fs.readFileSync(new URL('../' + name, import.meta.url), 'utf8'), sandbox, { filename: name });
  }
  vm.runInContext("function normalizarTipoCompraItemQTAS_(v) { return String(v || 'Insumo'); }", sandbox);
  sandbox.QTAS = { status: { registro: { anulado: 'Anulado' } } };
  return sandbox;
}

test('capsulas con y sin tilde comparten identidad de inventario', () => {
  const c = context();
  assert.equal(
    c.claveControlInventarioQTAS_('Insumo', 'Cápsulas', 'und'),
    c.claveControlInventarioQTAS_('Insumo', 'Capsulas', 'und')
  );
  const rows = [
    { Tipo_Item: 'Insumo', Item: 'Cápsulas', Unidad: 'und', Cantidad_Signada: 3000 },
    { Tipo_Item: 'Insumo', Item: 'Capsulas', Unidad: 'und', Cantidad_Signada: -519 }
  ];
  const controls = c.construirIndiceControlesInventarioQTAS_([
    { tipoItem: 'Insumo', item: 'Cápsulas', unidad: 'und', modoStock: 'Directo', activo: true }
  ]);
  const stock = c.construirSnapshotInventarioQTAS_(rows, controls);
  assert.equal(stock.length, 1);
  assert.equal(stock[0].Stock_Actual, 2481);
  assert.equal(stock[0].Item, 'Capsulas');
});

test('el consumo y la entrada de compra usan la misma identidad', () => {
  const c = context();
  const a = c.crearMovimientoInventarioQTAS_({
    operacion: 'Entrada', tipoItem: 'Insumo', item: 'Cápsulas', unidad: 'und', cantidad: 3000
  });
  const b = c.crearMovimientoInventarioQTAS_({
    operacion: 'Salida', tipoItem: 'Insumo', item: 'Capsulas', unidad: 'und', cantidad: 519
  });
  assert.equal(a.Item, b.Item);
});
