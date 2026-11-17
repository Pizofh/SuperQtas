import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';

const app = fs.readFileSync(new URL('../App.html', import.meta.url), 'utf8');

function view(name) {
  const start = app.indexOf('<section class="app-view hidden" data-view="' + name + '">');
  assert.notEqual(start, -1, 'Falta la vista ' + name);
  const end = app.indexOf('<section class="app-view hidden" data-view="', start + 1);
  return app.slice(start, end < 0 ? app.length : end);
}

test('Ventas solo tiene formulario de registro y pagos iniciales', () => {
  const ventas = view('ventas');
  assert.match(ventas, /id="guardarVentaBtn"/);
  assert.match(ventas, /id="productos"/);
  assert.match(ventas, /id="pagos"/);
  for (const id of ['ventasRecientes', 'cuadresPendientes', 'deudores', 'enviosPendientes', 'ventaPendiente', 'cuadreVentaPanel']) {
    assert.doesNotMatch(ventas, new RegExp('id="' + id + '"'), id + ' no pertenece al registro');
  }
});

test('Historial concentra gestion postventa sin duplicar identificadores', () => {
  const historial = view('historial');
  for (const id of ['ventasRecientes', 'cuadresPendientes', 'deudores', 'enviosPendientes', 'ventaPendiente', 'cuadreVentaPanel']) {
    assert.match(historial, new RegExp('id="' + id + '"'), id + ' debe estar en Historial');
    assert.equal(app.split('id="' + id + '"').length - 1, 1, id + ' tiene ID duplicado');
  }
  assert.match(app, /data-view="historial" onclick="mostrarVistaApp\('historial'\)"/);
  assert.match(app, /if \(nextView === 'historial'\) initHistorialVentas\(\)/);
  assert.match(app, /if \(appState\.currentView === 'historial'\)/);
});

test('Editar conserva el formulario compartido y regresa al historial', () => {
  assert.match(app, /function iniciarEdicionVenta\(ventaId\)/);
  assert.match(app, /function cancelarEdicionVenta\(options\)/);
  assert.match(app, /cancelarEdicionVenta\(\{ sinNavegar: true \}\)/);
  assert.match(app, /mostrarVistaApp\('historial'\)/);
  assert.match(app, /mostrarVistaApp\('ventas'\)/);
  assert.match(app, /function iniciarEdicionCompra\(compraId\)/);
});

test('Los mensajes para el operador no exponen detalles de implementacion', () => {
  assert.doesNotMatch(app, /Se conservan IDs y fechas de pagos/);
  assert.doesNotMatch(app, /Respaldo:.*resp\.backupId/);
  assert.doesNotMatch(app, /Conserva ID; se recalculan/);
  assert.match(app, /'Editando venta V' \+ ventaId/);
  assert.match(app, /'Editando compra C' \+ compraId/);
});
