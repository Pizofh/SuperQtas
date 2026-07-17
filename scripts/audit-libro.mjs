import fs from 'node:fs';
import path from 'node:path';

const libroDir = path.resolve(process.cwd(), process.argv[2] || 'libro');
const protectedProducts = new Set(['Feria', 'Vino', 'Kit', '500mg']);

function readTsv(fileName) {
  const filePath = path.join(libroDir, fileName);
  const lines = fs.readFileSync(filePath, 'utf8')
    .replace(/^\uFEFF/, '')
    .trimEnd()
    .split(/\r?\n/);
  const headers = lines.shift().split('\t');
  return lines.filter(Boolean).map(line => {
    const values = line.split('\t');
    return Object.fromEntries(headers.map((header, index) => [header, values[index] || '']));
  });
}

function number(value) {
  let text = String(value || '').trim().replace(/[$\s]/g, '');
  if (!text) return 0;

  if (text.includes(',')) {
    text = text.replace(/\./g, '').replace(',', '.');
  } else if (/^\d{1,3}(\.\d{3})+$/.test(text)) {
    text = text.replace(/\./g, '');
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function round(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function isActive(row) {
  return String(row.Estado_Registro || '').toLowerCase() !== 'anulado' &&
    String(row.Activo || 'TRUE').toLowerCase() !== 'false';
}

function dateKey(value) {
  return String(value || '').slice(0, 10);
}

function sum(rows, field) {
  return round((rows || []).reduce((total, row) => total + number(row[field]), 0));
}

function groupBy(rows, selector) {
  return rows.reduce((groups, row) => {
    const key = typeof selector === 'function' ? selector(row) : row[selector];
    const normalized = String(key || '');
    if (!groups[normalized]) groups[normalized] = [];
    groups[normalized].push(row);
    return groups;
  }, {});
}

function countBy(rows, field) {
  const result = {};
  rows.forEach(row => {
    const key = String(row[field] || '').trim() || '(vacio)';
    result[key] = (result[key] || 0) + 1;
  });
  return Object.fromEntries(Object.entries(result).sort((a, b) => b[1] - a[1]));
}

function canonicalMedium(value) {
  return String(value || '').trim().toLowerCase() === 'efectivo'
    ? 'Efectivo'
    : String(value || '').trim();
}

function canonicalOrigin(value) {
  const key = String(value || '').trim().toLowerCase();
  if (key === 'sm' || key === 'ms') return 'SM';
  if (key === 'mush') return 'Mush';
  if (key === 'caja') return 'Caja';
  return String(value || '').trim();
}

function sample(rows, mapper, limit = 12) {
  return rows.slice(0, limit).map(mapper);
}

function main() {
  const ventas = readTsv('05_Ventas.tsv');
  const compras = readTsv('06_Compras.tsv');
  const compraDetalle = readTsv('07_Compra_Detalle.tsv');
  const costos = readTsv('08_Costos_Referencia.tsv');
  const detalle = readTsv('09_Venta_Detalle.tsv');
  const pagos = readTsv('10_Pagos.tsv');
  const distribucion = readTsv('13_Distribucion_Ingresos.tsv');
  const precios = readTsv('03_Precios_Referencia.tsv');
  const inventarioMovimientos = readTsv('20_Inventario_Movimientos.tsv');
  const inventarioSnapshot = readTsv('21_Inventario_Snapshot.tsv');
  const analitica = readTsv('22_Venta_Detalle_Costos_Calc.tsv');
  const compraOrigenes = readTsv('24_Compra_Origenes_Fondos.tsv');

  const ventasActivas = ventas.filter(isActive);
  const comprasActivas = compras.filter(isActive);
  const detalleActivo = detalle.filter(isActive);
  const pagosActivos = pagos.filter(isActive);
  const distribucionActiva = distribucion.filter(isActive);
  const compraDetalleActivo = compraDetalle.filter(isActive);
  const detallePorVenta = groupBy(detalleActivo, 'Venta_ID');
  const pagosPorVenta = groupBy(pagosActivos, 'Venta_ID');
  const detallePorCompra = groupBy(compraDetalleActivo, 'Compra_ID');
  const fondosPorDetalle = groupBy(compraOrigenes, 'Compra_Detalle_ID');
  const movimientosPorItem = groupBy(
    inventarioMovimientos,
    row => [row.Tipo_Item, row.Item, row.Unidad].join('|')
  );
  const ventaPorId = Object.fromEntries(ventas.map(row => [String(row.Venta_ID), row]));
  const pagoPorId = Object.fromEntries(pagos.map(row => [String(row.Pago_ID), row]));

  const ventasDetalleInconsistentes = ventasActivas.filter(row =>
    Math.abs(number(row.Total_Venta) - sum(detallePorVenta[row.Venta_ID], 'Subtotal_Neto')) > 0.01
  );
  const ventasPagosInconsistentes = ventasActivas.filter(row =>
    Math.abs(number(row.Total_Pagado) - sum(pagosPorVenta[row.Venta_ID], 'Monto_Pago')) > 0.01
  );
  const ventasExtremas = ventasActivas.filter(row =>
    number(row.Total_Venta) > 3000000
  );
  const saldosInconsistentes = ventasActivas.filter(row =>
    Math.abs(number(row.Saldo) - Math.max(number(row.Total_Venta) - number(row.Total_Pagado), 0)) > 0.01
  );
  const comprasDetalleInconsistentes = comprasActivas.filter(row =>
    Math.abs(number(row.Total_Compra) - sum(detallePorCompra[row.Compra_ID], 'Costo_Total_Linea')) > 0.01
  );
  const fondosConDiferencia = compraDetalleActivo.map(row => {
    const fondos = fondosPorDetalle[row.Compra_Detalle_ID] || [];
    return {
      row,
      diferencia: Math.abs(number(row.Costo_Total_Linea) - sum(fondos, 'Monto_Asignado')),
      tieneFondos: fondos.length > 0
    };
  });
  const fondosRedondeo = fondosConDiferencia.filter(item =>
    item.tieneFondos && item.diferencia > 0.001 && item.diferencia <= 0.05
  );
  const fondosInconsistentes = fondosConDiferencia.filter(item =>
    item.tieneFondos && item.diferencia > 0.05
  ).map(item => item.row);
  const lineasSinFondos = compraDetalleActivo.filter(row =>
    !(fondosPorDetalle[row.Compra_Detalle_ID] || []).length
  );

  const distribucionConSplitInconsistente = distribucionActiva.filter(row =>
    Math.abs(number(row.Monto_Base) - (
      number(row.Steve_Valor) + number(row.Majo_Valor) + number(row.Mush_Valor)
    )) > 0.01
  );
  const distribucionPagoInconsistente = distribucionActiva.filter(row => {
    if (row.Fuente_Tipo !== 'Pago') return false;
    const pago = pagoPorId[row.Pago_ID];
    return !pago || Math.abs(number(row.Monto_Base) - number(pago.Monto_Pago)) > 0.01;
  });
  const distribucionVentaInconsistente = distribucionActiva.filter(row => {
    if (row.Fuente_Tipo !== 'Venta') return false;
    const venta = ventaPorId[row.Venta_ID];
    return !venta || Math.abs(number(row.Monto_Base) - number(venta.Total_Venta)) > 0.01;
  });

  const preciosActivos = precios.filter(row => String(row.Activo).toLowerCase() !== 'false').map(row => ({
    producto: row.Producto_Estandar,
    unidad: row.Unidad,
    precio: number(row.Precio),
    desde: dateKey(row.Fecha_Desde),
    hasta: dateKey(row.Fecha_Hasta)
  }));
  const precioVigente = (producto, unidad, fecha) => preciosActivos
    .filter(row => row.producto === producto && row.unidad === unidad &&
      row.desde <= fecha && (!row.hasta || fecha <= row.hasta))
    .sort((a, b) => b.desde.localeCompare(a.desde))[0] || null;
  const precioListaInconsistente = detalleActivo.map(row => ({
    row,
    precio: precioVigente(row.Producto_Estandar, row.Unidad, dateKey(row.Fecha_Venta))
  })).filter(item => item.precio &&
    Math.abs(number(item.row.Precio_Lista) - item.precio.precio) > 0.01
  );
  const ventasSobrePrecioVigente = detalleActivo.map(row => ({
    row,
    precio: precioVigente(row.Producto_Estandar, row.Unidad, dateKey(row.Fecha_Venta))
  })).filter(item => item.precio && item.precio.precio > 0 &&
    !protectedProducts.has(item.row.Producto_Estandar) &&
    number(item.row.Precio_Vendido_Unitario) / item.precio.precio >= 3
  );

  const costosPorClave = groupBy(
    costos.filter(row => String(row.Activo).toLowerCase() !== 'false'),
    row => [row.Tipo_Item, row.Item, row.Unidad].join('|')
  );
  const costosTraslapados = [];
  Object.entries(costosPorClave).forEach(([key, rows]) => {
    const sorted = rows.slice().sort((a, b) => dateKey(a.Fecha_Desde).localeCompare(dateKey(b.Fecha_Desde)));
    for (let index = 0; index < sorted.length - 1; index += 1) {
      const current = sorted[index];
      const next = sorted[index + 1];
      if (!current.Fecha_Hasta || dateKey(current.Fecha_Hasta) >= dateKey(next.Fecha_Desde)) {
        costosTraslapados.push({
          key,
          actual: current.Costo_ID,
          hastaActual: current.Fecha_Hasta,
          siguiente: next.Costo_ID,
          desdeSiguiente: next.Fecha_Desde
        });
      }
    }
  });

  const snapshotInconsistente = inventarioSnapshot.filter(row => {
    const key = [row.Tipo_Item, row.Item, row.Unidad].join('|');
    return Math.abs(number(row.Stock_Actual) - sum(movimientosPorItem[key], 'Cantidad_Signada')) > 0.01;
  });

  const margenBajo = analitica.filter(row =>
    row.Estado_Costo === 'Completo' &&
    !protectedProducts.has(row.Producto_Estandar) &&
    number(row.Subtotal_Neto) > 0 &&
    number(row.Margen_Porcentaje_Estimado) < 15
  ).sort((a, b) => number(a.Margen_Porcentaje_Estimado) - number(b.Margen_Porcentaje_Estimado));
  const costosExtremos = analitica.filter(row =>
    !protectedProducts.has(row.Producto_Estandar) &&
    number(row.Costo_Unitario_Usado) > 0 && (
      number(row.Costo_Unitario_Usado) >= 1000000 ||
      number(row.Costo_Total_Estimado) > number(row.Subtotal_Neto) * 5
    )
  );
  const lmPow = analitica.filter(row => row.Producto_Estandar === 'LmPow')
    .sort((a, b) => number(a.Margen_Porcentaje_Estimado) - number(b.Margen_Porcentaje_Estimado));
  const analiticaCompleta = analitica.filter(row => row.Estado_Costo === 'Completo');
  const subtotalConCosto = sum(analiticaCompleta, 'Subtotal_Neto');
  const costoConCosto = sum(analiticaCompleta, 'Costo_Total_Estimado');
  const ventaSinCostoCompleto = analitica.filter(row => row.Estado_Costo !== 'Completo');
  const gastosRegistrados = compraDetalleActivo.filter(row => row.Tipo_Item === 'Gasto');

  const venta2342 = {
    venta: ventas.find(row => row.Venta_ID === '2342') || null,
    detalles: detalle.filter(row => row.Venta_ID === '2342'),
    pagos: pagos.filter(row => row.Venta_ID === '2342'),
    distribucion: distribucion.filter(row => row.Venta_ID === '2342'),
    analitica: analitica.filter(row => row.Venta_ID === '2342')
  };

  const report = {
    rows: {
      ventas: ventas.length,
      ventaDetalle: detalle.length,
      pagos: pagos.length,
      compras: compras.length,
      compraDetalle: compraDetalle.length,
      distribucionIngresos: distribucion.length,
      compraOrigenesFondos: compraOrigenes.length,
      inventarioMovimientos: inventarioMovimientos.length,
      inventarioSnapshot: inventarioSnapshot.length,
      ventaDetalleCostos: analitica.length
    },
    conciliacion: {
      ventasVsDetalle: {
        count: ventasDetalleInconsistentes.length,
        samples: sample(ventasDetalleInconsistentes, row => ({
          ventaId: row.Venta_ID,
          venta: row.Total_Venta,
          detalle: sum(detallePorVenta[row.Venta_ID], 'Subtotal_Neto')
        }))
      },
      ventasVsPagos: {
        count: ventasPagosInconsistentes.length,
        samples: sample(ventasPagosInconsistentes, row => ({
          ventaId: row.Venta_ID,
          registrado: row.Total_Pagado,
          pagos: sum(pagosPorVenta[row.Venta_ID], 'Monto_Pago')
        }))
      },
      saldos: {
        count: saldosInconsistentes.length,
        samples: sample(saldosInconsistentes, row => ({ ventaId: row.Venta_ID, saldo: row.Saldo }))
      },
      ventasSobreTresMillones: {
        count: ventasExtremas.length,
        samples: sample(ventasExtremas, row => ({
          ventaId: row.Venta_ID,
          fecha: dateKey(row.Fecha_Venta),
          cliente: row.Nombre,
          total: row.Total_Venta
        }))
      },
      comprasVsDetalle: {
        count: comprasDetalleInconsistentes.length,
        samples: sample(comprasDetalleInconsistentes, row => ({
          compraId: row.Compra_ID,
          compra: row.Total_Compra,
          detalle: sum(detallePorCompra[row.Compra_ID], 'Costo_Total_Linea')
        }))
      },
      fondosCompras: {
        inconsistentes: fondosInconsistentes.length,
        redondeos: fondosRedondeo.length,
        sinAsignacion: lineasSinFondos.length,
        samples: sample(fondosInconsistentes, row => ({
          compraDetalleId: row.Compra_Detalle_ID,
          costo: row.Costo_Total_Linea,
          asignado: sum(fondosPorDetalle[row.Compra_Detalle_ID], 'Monto_Asignado')
        }))
      },
      distribucion: {
        splitInconsistente: distribucionConSplitInconsistente.length,
        pagosInconsistentes: distribucionPagoInconsistente.length,
        ventasInconsistentes: distribucionVentaInconsistente.length
      },
      inventario: {
        snapshotVsMovimientos: snapshotInconsistente.length,
        samples: sample(snapshotInconsistente, row => ({
          item: row.Item,
          unidad: row.Unidad,
          snapshot: row.Stock_Actual,
          movimientos: sum(movimientosPorItem[[row.Tipo_Item, row.Item, row.Unidad].join('|')], 'Cantidad_Signada')
        }))
      }
    },
    normalizacion: {
      mediosPagoCompras: countBy(comprasActivas, 'Medio_Pago'),
      mediosPagoVentas: countBy(pagosActivos, 'Medio_Pago'),
      mediosPagoVentasCanonicos: countBy(
        pagosActivos.map(row => Object.assign({}, row, { Medio_Pago: canonicalMedium(row.Medio_Pago) })),
        'Medio_Pago'
      ),
      origenesFondos: countBy(compraOrigenes, 'Origen_Fondos'),
      origenesFondosCanonicos: countBy(
        compraOrigenes.map(row => Object.assign({}, row, { Origen_Fondos: canonicalOrigin(row.Origen_Fondos) })),
        'Origen_Fondos'
      ),
      aportantesFondos: countBy(compraOrigenes, 'Aportante')
    },
    precios: {
      listaNoCoincideConVigente: precioListaInconsistente.length,
      listaSamples: sample(precioListaInconsistente, item => ({
        detalleId: item.row.Detalle_ID,
        ventaId: item.row.Venta_ID,
        producto: item.row.Producto_Estandar,
        fecha: dateKey(item.row.Fecha_Venta),
        listaRegistrada: item.row.Precio_Lista,
        vigente: item.precio.precio,
        vendido: item.row.Precio_Vendido_Unitario
      })),
      ventasSobreTresVecesVigente: ventasSobrePrecioVigente.length,
      samples: sample(ventasSobrePrecioVigente, item => ({
        detalleId: item.row.Detalle_ID,
        ventaId: item.row.Venta_ID,
        producto: item.row.Producto_Estandar,
        fecha: dateKey(item.row.Fecha_Venta),
        vendido: item.row.Precio_Vendido_Unitario,
        vigente: item.precio.precio,
        subtotal: item.row.Subtotal_Neto
      }))
    },
    costos: {
      estadoCosto: countBy(analitica, 'Estado_Costo'),
      traslapados: costosTraslapados.length,
      margenBajo: {
        count: margenBajo.length,
        negativos: margenBajo.filter(row => number(row.Margen_Bruto_Estimado) < 0).length,
        samples: sample(margenBajo, row => ({
          detalleId: row.Detalle_ID,
          ventaId: row.Venta_ID,
          producto: row.Producto_Estandar,
          margenPct: row.Margen_Porcentaje_Estimado,
          subtotal: row.Subtotal_Neto,
          costo: row.Costo_Total_Estimado,
          nota: row.Nota
        }))
      },
      extremos: {
        count: costosExtremos.length,
        samples: sample(costosExtremos, row => ({
          detalleId: row.Detalle_ID,
          ventaId: row.Venta_ID,
          producto: row.Producto_Estandar,
          unitario: row.Costo_Unitario_Usado,
          costo: row.Costo_Total_Estimado,
          subtotal: row.Subtotal_Neto
        }))
      },
      lmPow: sample(lmPow, row => ({
        detalleId: row.Detalle_ID,
        ventaId: row.Venta_ID,
        margenPct: row.Margen_Porcentaje_Estimado,
        subtotal: row.Subtotal_Neto,
        costo: row.Costo_Total_Estimado,
        estado: row.Estado_Costo,
        nota: row.Nota
      }))
    },
    rentabilidad: {
      ingresosNetosConCosto: subtotalConCosto,
      costoVentasEstimado: costoConCosto,
      margenBrutoEstimado: round(subtotalConCosto - costoConCosto),
      margenBrutoPct: subtotalConCosto
        ? round(((subtotalConCosto - costoConCosto) / subtotalConCosto) * 100)
        : 0,
      lineasExcluidasPorCosto: ventaSinCostoCompleto.length,
      ingresosExcluidosPorCosto: sum(ventaSinCostoCompleto, 'Subtotal_Neto'),
      gastosRegistradosEnCompras: sum(gastosRegistrados, 'Costo_Total_Linea'),
      nota: 'Los pagos son caja recibida; no deben sumarse al ingreso ni al margen bruto.'
    },
    venta2342
  };

  console.log(JSON.stringify(report, null, 2));
}

main();
