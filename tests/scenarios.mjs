function ventaPayloadBase(overrides = {}) {
  return {
    fechaVenta: '2026-06-20',
    cliente: {
      nombre: 'Cliente Test'
    },
    comentarioVenta: 'Escenario automatizado',
    lineas: [],
    pagos: [],
    devolverDashboard: false,
    validarModelo: false,
    ...overrides
  };
}

function ventaLinea(producto, cantidad, unidad, precioVendidoUnitario, descuentoLinea = 0, comentarioLinea = '') {
  return {
    producto,
    cantidad,
    unidad,
    precioVendidoUnitario,
    descuentoLinea,
    comentarioLinea
  };
}

function pagoLinea(medio, monto, comentarioPago = '') {
  return {
    medio,
    monto,
    comentarioPago
  };
}

function compraPayloadBase(overrides = {}) {
  return {
    fechaCompra: '2026-06-20',
    proveedor: 'Proveedor Test',
    medioPago: 'Efectivo',
    comentarioCompra: 'Escenario automatizado',
    lineas: [],
    ...overrides
  };
}

function compraLinea(tipoItem, item, cantidad, unidad, costoTotalLinea, impactaCosto, comentarioLinea = '') {
  return {
    tipoItem,
    item,
    cantidad,
    unidad,
    costoTotalLinea,
    impactaCosto,
    comentarioLinea
  };
}

function snapshotLigero(ctx, overrides = {}) {
  return ctx.snapshot({
    sheetNames: [],
    includeDashboard: false,
    includeCompras: false,
    includeConfig: false,
    ...overrides
  });
}

function batchStep(functionName, ...parameters) {
  return {
    functionName,
    parameters
  };
}

export const SCENARIOS = [
  {
    id: 'producto_y_precio_nuevo',
    title: 'Producto nuevo y cambio de precio desde configuracion',
    tags: ['admin', 'catalogo', 'precios'],
    run: async ctx => {
      await ctx.reset();

      const producto = 'ProdTestAuto';
      const unidad = 'und';
      const results = await ctx.batch([
        batchStep('guardarProductoConfiguracionQTAS', {
          producto,
          unidad,
          nota: 'Creado por test',
          activo: true
        }),
        batchStep('guardarCambioPrecioFrontendQTAS', {
          producto,
          unidad,
          precio: 12345,
          fechaDesde: '2026-06-20',
          nota: 'Precio de prueba'
        }),
        batchStep('getCatalogoQTAS', '2026-06-21')
      ]);
      const catalogo = results[2];
      const encontrado = (catalogo.productos || []).find(item => item.producto === producto);
      ctx.assert(encontrado, 'El producto nuevo debe aparecer en el catalogo de ventas.');
      ctx.equal(ctx.num(encontrado.precio), 12345, 'El precio configurado debe verse en el catalogo.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Productos', 'Precios_Referencia']
      });
      const productos = ctx.sheetRows(state, 'Productos');
      const precios = ctx.sheetRows(state, 'Precios_Referencia');
      ctx.assert(productos.some(row => row.Producto_Estandar === producto), 'La hoja Productos debe contener el producto nuevo.');
      ctx.assert(
        precios.some(row => row.Producto_Estandar === producto && ctx.num(row.Precio) === 12345),
        'La hoja de precios debe contener el nuevo precio.'
      );
    }
  },
  {
    id: 'venta_con_deuda',
    title: 'Venta sin pago inicial',
    tags: ['ventas', 'deuda', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Deuda Test' },
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000)
        ]
      }));

      ctx.equal(ctx.num(resp.totalVenta), 20000, 'El total de la venta debe ser 20000.');
      ctx.equal(ctx.num(resp.saldo), 20000, 'El saldo debe quedar pendiente.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas', 'Venta_Detalle', 'Pagos', 'Distribucion_Ingresos'],
        includeDashboard: true
      });
      const venta = ctx.findRow(state, 'Ventas', row => ctx.num(row.Venta_ID) === ctx.num(resp.ventaId), 'No se encontro la venta guardada.');
      ctx.equal(ctx.num(venta.Total_Pagado), 0, 'La venta pendiente no debe registrar pago inicial.');
      ctx.equal(String(venta.Estado_Pago), 'Pendiente', 'La venta debe quedar pendiente.');
      ctx.equal(ctx.sheetRows(state, 'Venta_Detalle').length, 1, 'Debe existir una sola linea de detalle.');
      ctx.equal(ctx.sheetRows(state, 'Pagos').length, 0, 'No debe existir fila en Pagos.');
      ctx.equal(ctx.sheetRows(state, 'Distribucion_Ingresos').length, 1, 'Solo debe existir la distribucion de la venta.');
      ctx.equal((state.dashboard.ventasPendientes || []).length, 1, 'La venta debe aparecer como pendiente en dashboard.');
      ctx.equal((state.dashboard.deudores || []).length, 1, 'Debe existir un deudor.');
    }
  },
  {
    id: 'venta_sin_deuda',
    title: 'Venta completamente pagada desde el inicio',
    tags: ['ventas', 'sin-deuda', 'pagos'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Pagado Test' },
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000)
        ],
        pagos: [
          pagoLinea('efectivo', 20000, 'Pago total')
        ]
      }));

      ctx.equal(ctx.num(resp.saldo), 0, 'La venta pagada no debe dejar saldo.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas', 'Pagos', 'Distribucion_Ingresos'],
        includeDashboard: true
      });
      const venta = ctx.findRow(state, 'Ventas', row => ctx.num(row.Venta_ID) === ctx.num(resp.ventaId), 'No se encontro la venta pagada.');
      ctx.equal(String(venta.Estado_Pago), 'Pagado', 'La venta debe quedar pagada.');
      ctx.equal(ctx.sheetRows(state, 'Pagos').length, 1, 'Debe existir un pago inicial.');
      ctx.equal(ctx.sheetRows(state, 'Distribucion_Ingresos').length, 2, 'Deben existir filas de distribucion para venta y pago.');
      ctx.equal((state.dashboard.ventasPendientes || []).length, 0, 'No debe quedar pendiente en dashboard.');
    }
  },
  {
    id: 'venta_pagada_con_envio_pendiente',
    title: 'Venta pagada que sigue pendiente de envio',
    tags: ['ventas', 'envios', 'sin-deuda', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const venta = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Envio Test' },
        pendienteEnvio: true,
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000)
        ],
        pagos: [
          pagoLinea('efectivo', 20000, 'Pago total')
        ]
      }));

      ctx.equal(String(venta.estadoEnvio), 'Pendiente', 'La venta debe quedar marcada como pendiente de envio.');

      let state = await snapshotLigero(ctx, {
        sheetNames: ['Pagos'],
        includeDashboard: true
      });
      ctx.equal((state.dashboard.ventasPendientes || []).length, 0, 'La venta pagada no debe quedar como deuda.');
      ctx.equal((state.dashboard.enviosPendientes || []).length, 1, 'La venta debe aparecer en el panel de envios pendientes.');
      ctx.equal(
        String(ctx.sheetRows(state, 'Pagos')[0].Medio_Pago),
        'Efectivo',
        'El pago nuevo debe guardar Efectivo de forma canonica.'
      );

      await ctx.call('actualizarEstadoEnvioVentaQTAS', {
        ventaId: venta.ventaId,
        estadoEnvio: 'Enviado',
        devolverDashboard: false,
        validarModelo: false
      });

      state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas_Envio'],
        includeDashboard: true
      });
      const envio = ctx.findRow(state, 'Ventas_Envio', row => ctx.num(row.Venta_ID) === ctx.num(venta.ventaId), 'No se encontro el registro de envio.');
      ctx.equal(String(envio.Estado_Envio), 'Enviado', 'El envio debe quedar marcado como enviado.');
      ctx.equal((state.dashboard.enviosPendientes || []).length, 0, 'La venta ya no debe quedar pendiente de envio.');
    }
  },
  {
    id: 'venta_varios_productos_varios_medios',
    title: 'Venta parcial con varios productos y varios medios de pago',
    tags: ['ventas', 'pagos', 'multi-linea'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Multi Test' },
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000),
          ventaLinea('100mg', 2, 'und', 2000),
          ventaLinea('AcMed', 1, 'g', 14000, 3000)
        ],
        pagos: [
          pagoLinea('Efectivo', 10000, 'Abono 1'),
          pagoLinea('NequiSteve', 5000, 'Abono 2')
        ]
      }));

      ctx.equal(ctx.num(resp.totalVenta), 35000, 'El total esperado de la venta multi-linea debe ser 35000.');
      ctx.equal(ctx.num(resp.totalPagado), 15000, 'El total pagado inicial debe ser 15000.');
      ctx.equal(ctx.num(resp.saldo), 20000, 'El saldo pendiente debe ser 20000.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas', 'Venta_Detalle', 'Pagos', 'Distribucion_Ingresos']
      });
      const venta = ctx.findRow(state, 'Ventas', row => ctx.num(row.Venta_ID) === ctx.num(resp.ventaId), 'No se encontro la venta multi-linea.');
      const pagos = ctx.sheetRows(state, 'Pagos');
      ctx.equal(String(venta.Estado_Pago), 'Parcial', 'La venta debe quedar parcial.');
      ctx.equal(ctx.sheetRows(state, 'Venta_Detalle').length, 3, 'La venta debe escribir tres lineas de detalle.');
      ctx.equal(pagos.length, 2, 'Deben existir dos pagos.');
      ctx.equal(ctx.sheetRows(state, 'Distribucion_Ingresos').length, 3, 'Deben existir tres filas de distribucion.');
      ctx.assert(pagos.some(row => row.Medio_Pago === 'Efectivo'), 'Debe persistir el pago en efectivo.');
      ctx.assert(pagos.some(row => row.Medio_Pago === 'NequiSteve'), 'Debe persistir el pago en NequiSteve.');
    }
  },
  {
    id: 'venta_actualiza_analitica_incremental',
    title: 'Venta actualiza Venta_Detalle_Costos_Calc sin reconstruccion manual',
    tags: ['ventas', 'costos', 'analitica-incremental', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const producto = 'ProdTestAuto';
      const results = await ctx.batch([
        batchStep('guardarProductoConfiguracionQTAS', {
          producto,
          unidad: 'und',
          nota: 'Creado por test',
          activo: true
        }),
        batchStep('registrarCompraQTAS', compraPayloadBase({
          proveedor: 'Proveedor Costo Test',
          lineas: [
            compraLinea('Producto', producto, 10, 'und', 100000, true, 'Costo incremental base')
          ]
        })),
        batchStep('registrarVentaQTAS', ventaPayloadBase({
          cliente: { nombre: 'Cliente Test' },
          lineas: [
            ventaLinea(producto, 1, 'und', 15000)
          ]
        }))
      ]);
      const venta = results[2];

      ctx.assert(venta.analiticaCostos && venta.analiticaCostos.ok === true, 'La venta debe devolver estado ok de analitica incremental.');
      ctx.equal(ctx.num(venta.analiticaCostos.rows), 1, 'La analitica incremental debe procesar una sola linea.');
      ctx.equal(ctx.num(venta.analiticaCostos.inserted), 1, 'La analitica incremental debe insertar una fila nueva.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Venta_Detalle_Costos_Calc']
      });
      const costo = ctx.findRow(
        state,
        'Venta_Detalle_Costos_Calc',
        row => ctx.num(row.Venta_ID) === ctx.num(venta.ventaId) && row.Producto_Estandar === producto,
        'No se encontro la fila incremental de costo para la venta.'
      );

      ctx.equal(ctx.num(costo.Costo_Unitario_Usado), 10000, 'La venta debe usar el costo directo vigente del producto.');
      ctx.equal(ctx.num(costo.Costo_Total_Estimado), 10000, 'El costo total estimado debe quedar sincronizado.');
      ctx.equal(String(costo.Metodo_Costo), 'Costo directo', 'La fila incremental debe marcar costo directo.');
      ctx.equal(String(costo.Estado_Costo), 'Directo', 'La fila incremental debe quedar en estado Directo.');
    }
  },
  {
    id: 'pago_pendiente_completa_deuda',
    title: 'Pago pendiente que completa una venta parcial',
    tags: ['ventas', 'pagos', 'deuda'],
    run: async ctx => {
      await ctx.reset();

      const venta = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Pago Pendiente' },
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000)
        ],
        pagos: [
          pagoLinea('Efectivo', 5000, 'Abono inicial')
        ]
      }));

      const pago = await ctx.call('registrarPagoPendienteQTAS', {
        ventaId: venta.ventaId,
        fechaPago: '2026-06-21',
        medio: 'Bancolombia',
        monto: 15000,
        comentarioPago: 'Pago final',
        devolverDashboard: false,
        validarModelo: false
      });

      ctx.equal(ctx.num(pago.saldoRestante), 0, 'El pago pendiente debe completar la deuda.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas', 'Pagos', 'Distribucion_Ingresos'],
        includeDashboard: true
      });
      const ventaActualizada = ctx.findRow(state, 'Ventas', row => ctx.num(row.Venta_ID) === ctx.num(venta.ventaId), 'No se encontro la venta actualizada.');
      ctx.equal(String(ventaActualizada.Estado_Pago), 'Pagado', 'La venta debe terminar pagada.');
      ctx.equal(ctx.num(ventaActualizada.Saldo), 0, 'El saldo final debe ser cero.');
      ctx.equal(ctx.sheetRows(state, 'Pagos').length, 2, 'Deben existir dos pagos acumulados.');
      ctx.equal(ctx.sheetRows(state, 'Distribucion_Ingresos').length, 3, 'Deben existir tres filas de distribucion tras sincronizar.');
      ctx.equal((state.dashboard.ventasPendientes || []).length, 0, 'La venta ya no debe aparecer como pendiente.');
    }
  },
  {
    id: 'compra_producto_impacta_costo',
    title: 'Compra de producto que actualiza costos',
    tags: ['compras', 'costos'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Costo Test',
        lineas: [
          compraLinea('Producto', 'AcSup', 10, 'g', 100000, true, 'Costo de prueba')
        ]
      }));

      ctx.equal(ctx.num(resp.totalCompra), 100000, 'El total de la compra debe coincidir.');
      ctx.equal(ctx.num(resp.costosActualizados), 1, 'La compra debe actualizar un costo.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Costos_Referencia'],
        includeCompras: true
      });
      const costos = ctx.sheetRows(state, 'Costos_Referencia');
      ctx.equal(costos.length, 1, 'Debe existir una fila de costo historico.');
      ctx.equal(ctx.num(costos[0].Costo_Unitario), 10000, 'El costo unitario debe ser 10000.');
      ctx.equal(String(costos[0].Tipo_Item), 'Producto', 'El costo debe quedar como producto.');
      ctx.assert(
        (state.costosVigentes || []).some(row => row.item === 'AcSup' && ctx.num(row.costoUnitario) === 10000),
        'El costo vigente debe reflejar la compra.'
      );
    }
  },
  {
    id: 'compra_refresca_costo_producto_calc_incremental',
    title: 'Compra refresca Costo_Producto_Calc sin reconstruccion manual',
    tags: ['compras', 'costos', 'analitica-incremental', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Costo Test',
        lineas: [
          compraLinea('Producto', 'AcSup', 10, 'g', 100000, true, 'Refresh incremental de costo producto')
        ]
      }));

      ctx.assert(
        resp.costoProductoCalculado && resp.costoProductoCalculado.ok === true,
        'La compra debe devolver estado ok del refresh incremental de Costo_Producto_Calc.'
      );
      ctx.assert(
        ctx.num(resp.costoProductoCalculado.rows) > 0,
        'El refresh incremental debe recalcular al menos una fila de costo producto.'
      );
      const state = await snapshotLigero(ctx, {
        sheetNames: ['Costo_Producto_Calc', 'Costos_Referencia']
      });
      const costo = ctx.findRow(
        state,
        'Costo_Producto_Calc',
        row => row.Producto_Estandar === 'AcSup' && row.Unidad_Venta === 'g',
        'No se encontro la fila incremental en Costo_Producto_Calc.'
      );

      ctx.equal(ctx.num(costo.Costo_Unitario_Total), 10000, 'El snapshot incremental de costo producto debe reflejar el costo directo vigente.');
      ctx.equal(String(costo.Metodo_Costo), 'Costo directo', 'El producto debe quedar costeado por costo directo.');
      ctx.equal(String(costo.Estado_Costo), 'Directo', 'El estado del costo producto debe quedar Directo.');
      ctx.equal(
        ctx.sheetRows(state, 'Costos_Referencia').length,
        1,
        'La compra debe persistir una unica referencia de costo vigente.'
      );
    }
  },
  {
    id: 'compra_producto_no_canonico_rechazada',
    title: 'Compra de producto con item no canonico rechazada al impactar costo',
    tags: ['compras', 'costos', 'validacion'],
    run: async ctx => {
      await ctx.reset();

      await ctx.expectError(
        'registrarCompraQTAS',
        [compraPayloadBase({
          proveedor: 'Proveedor Producto Invalido Test',
          lineas: [
            compraLinea('Producto', 'Ac Sup typo', 2, 'g', 18000, true, 'Debe fallar por producto no canonico')
          ]
        })],
        'producto canonico'
      );
    }
  },
  {
    id: 'compra_gasto_no_impacta_costo',
    title: 'Compra de gasto que no afecta costos de referencia',
    tags: ['compras', 'gastos'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Gasto Test',
        medioPago: 'Daviplata',
        lineas: [
          compraLinea('Gasto', 'Arriendo', 1, 'und', 50000, false, 'No impacta costo')
        ]
      }));

      ctx.equal(ctx.num(resp.costosActualizados), 0, 'Un gasto puro no debe actualizar costos.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Compra_Detalle', 'Costos_Referencia']
      });
      ctx.equal(ctx.sheetRows(state, 'Costos_Referencia').length, 0, 'No deben crearse costos historicos.');
      const detalle = ctx.findRow(state, 'Compra_Detalle', row => row.Item === 'Arriendo', 'No se encontro el detalle del gasto.');
      ctx.assert(
        detalle.Impacta_Costo === false || String(detalle.Impacta_Costo).toLowerCase() === 'false',
        'La linea de gasto no debe impactar costo.'
      );
    }
  },
  {
    id: 'compra_mixta_varias_lineas',
    title: 'Compra mixta con producto, insumo y gasto',
    tags: ['compras', 'gastos', 'multi-linea'],
    run: async ctx => {
      await ctx.reset();

      const resp = await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Mixto Test',
        medioPago: 'NequiMajo',
        lineas: [
          compraLinea('Producto', 'AcSup', 4, 'g', 24000, true, 'Lote pequeno'),
          compraLinea('Insumo', 'Caja', 10, 'und', 5000, true, 'Empaque'),
          compraLinea('Gasto', 'Domicilio', 1, 'und', 7000, false, 'Envio')
        ]
      }));

      ctx.equal(ctx.num(resp.totalCompra), 36000, 'La compra mixta debe sumar 36000.');
      ctx.equal(ctx.num(resp.costosActualizados), 2, 'Solo dos lineas deben actualizar costos.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Compra_Detalle', 'Costos_Referencia'],
        includeCompras: true
      });
      ctx.equal(ctx.sheetRows(state, 'Compra_Detalle').length, 3, 'Deben escribirse tres lineas de detalle.');
      ctx.equal(ctx.sheetRows(state, 'Costos_Referencia').length, 2, 'Deben existir dos costos historicos.');
      ctx.assert(
        (state.comprasRecientes || []).some(row => ctx.num(row.compraId) === ctx.num(resp.compraId) && row.medioPago === 'NequiMajo'),
        'La compra reciente debe conservar el medio de pago.'
      );
    }
  },
  {
    id: 'compra_origen_fondos_por_fecha',
    title: 'Compra asigna aportantes segun origen de fondos y fecha',
    tags: ['compras', 'fondos', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const results = await ctx.batch([
        batchStep('guardarReglaOrigenFondosFrontendQTAS', {
          origenFondos: 'MS',
          fechaDesde: '2026-06-01',
          steve: 50,
          majo: 50,
          mush: 0,
          nota: 'Base SM'
        }),
        batchStep('guardarReglaOrigenFondosFrontendQTAS', {
          origenFondos: 'sm',
          fechaDesde: '2026-07-01',
          steve: 40,
          majo: 60,
          mush: 0,
          nota: 'Ajuste SM'
        }),
        batchStep('registrarCompraQTAS', compraPayloadBase({
          fechaCompra: '2026-07-02',
          proveedor: 'Proveedor Fondo Test',
          origenFondos: 'sm',
          lineas: [
            compraLinea('Insumo', 'Caja', 2, 'und', 10000, true, 'Compra con reparto')
          ]
        }))
      ]);
      const resp = results[2];

      ctx.equal(String(resp.origenFondos), 'SM', 'La compra debe conservar el origen de fondos usado.');
      ctx.equal(ctx.num(resp.origenesFondosAsignados), 2, 'La compra debe generar dos asignaciones de fondos.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Origenes_Fondos_Reglas', 'Compra_Origenes_Fondos'],
        includeCompras: true
      });
      const reglas = ctx.sheetRows(state, 'Origenes_Fondos_Reglas');
      const asignaciones = ctx.sheetRows(state, 'Compra_Origenes_Fondos')
        .filter(row => ctx.num(row.Compra_ID) === ctx.num(resp.compraId));

      ctx.equal(reglas.length, 4, 'Deben existir cuatro filas historicas para las dos versiones de la regla SM.');
      ctx.equal(asignaciones.length, 2, 'La compra debe repartir una linea entre dos aportantes.');

      const filaSteve = asignaciones.find(row => row.Aportante === 'Steve');
      const filaMajo = asignaciones.find(row => row.Aportante === 'Majo');

      ctx.assert(filaSteve, 'Debe existir la asignacion para Steve.');
      ctx.assert(filaMajo, 'Debe existir la asignacion para Majo.');
      ctx.equal(ctx.num(filaSteve.Porcentaje), 40, 'Steve debe tomar 40% en la fecha nueva.');
      ctx.equal(ctx.num(filaMajo.Porcentaje), 60, 'Majo debe tomar 60% en la fecha nueva.');
      ctx.equal(ctx.num(filaSteve.Monto_Asignado), 4000, 'A Steve deben asignarse 4000.');
      ctx.equal(ctx.num(filaMajo.Monto_Asignado), 6000, 'A Majo deben asignarse 6000.');
      ctx.assert(
        asignaciones.every(row => row.Origen_Fondos === 'SM'),
        'Las asignaciones nuevas deben guardar SM de forma canonica.'
      );
      ctx.assert(
        (state.comprasRecientes || []).some(row => ctx.num(row.compraId) === ctx.num(resp.compraId) && row.origenFondos === 'SM'),
        'Las compras recientes deben exponer el origen de fondos.'
      );
    }
  },
  {
    id: 'compra_estandariza_item_catalogo',
    title: 'Compra reutiliza items conocidos y los deja sugeridos en catalogo',
    tags: ['compras', 'catalogo'],
    run: async ctx => {
      await ctx.reset();

      const results = await ctx.batch([
        batchStep('registrarCompraQTAS', compraPayloadBase({
          proveedor: 'Proveedor Catalogo Test',
          lineas: [
            compraLinea('Insumo', 'BolsaCatalogo', 10, 'und', 5000, true, 'Alta catalogo')
          ]
        })),
        batchStep('getCatalogoComprasQTAS')
      ]);
      const catalogo = results[1];
      const sugerencia = (catalogo.itemsSugeridos || []).find(row =>
        row.tipoItem === 'Insumo' && row.item === 'BolsaCatalogo'
      );

      ctx.assert(sugerencia, 'El item comprado debe quedar sugerido en el catalogo de compras.');
      ctx.equal(ctx.num(sugerencia.costoUnitario), 500, 'El catalogo debe recordar el ultimo costo unitario conocido.');
      ctx.equal(String(sugerencia.unidad), 'und', 'El catalogo debe conservar la unidad del item.');

      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Catalogo Test 2',
        lineas: [
          compraLinea('Insumo', 'bolsacatalogo', 6, 'und', 3600, true, 'Reuso catalogo')
        ]
      }));

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Compra_Detalle']
      });
      const detallesCanonicos = ctx.sheetRows(state, 'Compra_Detalle')
        .filter(row => row.Item === 'BolsaCatalogo');

      ctx.equal(detallesCanonicos.length, 2, 'Las compras repetidas deben quedar canonizadas con el mismo nombre de item.');
    }
  },
  {
    id: 'costo_producto_por_componentes',
    title: 'Costo de producto compuesto y margen por venta',
    tags: ['costos', 'receta', 'ventas'],
    run: async ctx => {
      await ctx.reset();

      const producto = 'ProdCostoCompuestoReceta';
      await ctx.batch([
        batchStep('guardarProductoConfiguracionQTAS', {
          producto,
          unidad: 'und',
          nota: 'Creado por test',
          activo: true
        }),
        batchStep('guardarCambioPrecioFrontendQTAS', {
          producto,
          unidad: 'und',
          precio: 2000,
          fechaDesde: '2026-06-20',
          nota: 'Precio producto compuesto'
        }),
        batchStep('registrarCompraQTAS', compraPayloadBase({
          proveedor: 'Proveedor Receta Test',
          lineas: [
            compraLinea('Insumo', 'CajaTest', 10, 'und', 5000, true, 'Caja para test')
          ]
        })),
        batchStep('registrarCompraQTAS', compraPayloadBase({
          proveedor: 'Proveedor Receta Test',
          lineas: [
            compraLinea('Insumo', 'EtiquetaTest', 20, 'und', 2000, true, 'Etiqueta para test')
          ]
        })),
        batchStep('guardarComponenteProductoQTAS', {
          producto,
          unidadVenta: 'und',
          orden: 1,
          tipoComponente: 'Insumo',
          itemComponente: 'CajaTest',
          cantidadComponente: 1,
          unidadComponente: 'und',
          mermaPct: 0,
          nota: 'Componente de prueba',
          activo: true
        }),
        batchStep('guardarComponenteProductoQTAS', {
          producto,
          unidadVenta: 'und',
          orden: 2,
          tipoComponente: 'Insumo',
          itemComponente: 'EtiquetaTest',
          cantidadComponente: 2,
          unidadComponente: 'und',
          mermaPct: 0,
          nota: 'Componente de prueba',
          activo: true
        }),
        batchStep('registrarVentaQTAS', ventaPayloadBase({
          cliente: { nombre: 'Cliente Test' },
          lineas: [
            ventaLinea(producto, 1, 'und', 2000)
          ]
        }))
      ]);

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Costo_Producto_Calc', 'Venta_Detalle_Costos_Calc']
      });
      const costosProducto = ctx.findRow(
        state,
        'Costo_Producto_Calc',
        row => row.Producto_Estandar === producto,
        'No se encontro el costo calculado del producto compuesto.'
      );
      const ventaCosto = ctx.findRow(
        state,
        'Venta_Detalle_Costos_Calc',
        row => row.Producto_Estandar === producto,
        'No se encontro la analitica de costo por detalle de venta.'
      );

      ctx.equal(ctx.num(costosProducto.Costo_Unitario_Total), 700, 'El costo total del producto compuesto debe ser 700.');
      ctx.equal(String(costosProducto.Metodo_Costo), 'Receta', 'El producto debe calcularse por receta.');
      ctx.equal(String(costosProducto.Estado_Costo), 'Completo', 'La receta debe quedar completa.');
      ctx.equal(ctx.num(ventaCosto.Costo_Total_Estimado), 700, 'La venta debe estimar costo total 700.');
      ctx.equal(ctx.num(ventaCosto.Margen_Bruto_Estimado), 1300, 'El margen bruto estimado debe ser 1300.');
    }
  },
  {
    id: 'regla_distribucion_nueva_aplicada',
    title: 'Nueva regla de distribucion aplicada en ventas posteriores',
    tags: ['admin', 'distribucion', 'ventas', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const results = await ctx.batch([
        batchStep('guardarReglaDistribucionFrontendQTAS', {
          fechaDesde: '2026-07-01',
          steve: 50,
          majo: 30,
          mush: 20,
          nota: 'Regla automatizada'
        }),
        batchStep('registrarVentaQTAS', ventaPayloadBase({
          fechaVenta: '2026-07-02',
          cliente: { nombre: 'Cliente Regla Test' },
          lineas: [
            ventaLinea('AcSup', 1, 'g', 20000)
          ]
        }))
      ]);
      const configuracion = results[0];
      const venta = results[1];

      ctx.assert(
        (configuracion.reglasDistribucion || []).some(row => row.fechaDesde === '2026-07-01' && ctx.num(row.steve) === 50),
        'La nueva regla debe quedar registrada.'
      );

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Distribucion_Reglas', 'Ventas', 'Distribucion_Ingresos']
      });
      const reglas = ctx.sheetRows(state, 'Distribucion_Reglas');
      const ventaRow = ctx.findRow(state, 'Ventas', row => ctx.num(row.Venta_ID) === ctx.num(venta.ventaId), 'No se encontro la venta bajo nueva regla.');
      const dist = ctx.findRow(state, 'Distribucion_Ingresos', row => String(row.Fuente_Tipo) === 'Venta', 'No se encontro la distribucion de la venta.');

      ctx.equal(reglas.length, 3, 'Deben existir tres reglas historicas tras agregar una nueva.');
      ctx.equal(String(ventaRow.Regla_Distribucion_Venta_ID), 'DIST-0003', 'La venta debe usar la nueva regla.');
      ctx.equal(ctx.num(ventaRow.Steve_Pct_Venta), 50, 'Steve debe quedar con 50%.');
      ctx.equal(ctx.num(ventaRow.Majo_Pct_Venta), 30, 'Majo debe quedar con 30%.');
      ctx.equal(ctx.num(ventaRow.Mush_Pct_Venta), 20, 'Mush debe quedar con 20%.');
      ctx.equal(ctx.num(dist.Steve_Valor), 10000, 'La distribucion debe repartir 10000 a Steve.');
      ctx.equal(ctx.num(dist.Majo_Valor), 6000, 'La distribucion debe repartir 6000 a Majo.');
      ctx.equal(ctx.num(dist.Mush_Valor), 4000, 'La distribucion debe repartir 4000 a Mush.');
    }
  },
  {
    id: 'medio_pago_nuevo_y_rechazo_al_desactivar',
    title: 'Medio de pago nuevo que luego se desactiva y rechaza compras',
    tags: ['admin', 'medios-pago', 'compras'],
    run: async ctx => {
      await ctx.reset();

      const medio = 'TransferTest';
      await ctx.batch([
        batchStep('guardarMedioPagoQTAS', {
          medioPago: medio,
          nota: 'Creado por test',
          activo: true
        }),
        batchStep('registrarCompraQTAS', compraPayloadBase({
          proveedor: 'Proveedor Medio Test',
          medioPago: medio,
          lineas: [
            compraLinea('Producto', 'AcSup', 2, 'g', 18000, true, 'Compra con medio nuevo')
          ]
        })),
        batchStep('cambiarEstadoMedioPagoQTAS', {
          medioPago: medio,
          activo: false
        })
      ]);

      await ctx.expectError(
        'registrarCompraQTAS',
        [compraPayloadBase({
          proveedor: 'Proveedor Medio Test 2',
          medioPago: medio,
          lineas: [
            compraLinea('Gasto', 'Taxi', 1, 'und', 15000, false, 'Debe fallar')
          ]
        })],
        'no esta disponible'
      );

      const state = await snapshotLigero(ctx, {
        includeConfig: true
      });
      const medioRow = (state.configuracionAvanzada.mediosPago || []).find(row => row.medioPago === medio);
      ctx.assert(medioRow, 'El medio creado debe seguir existiendo en configuracion.');
      ctx.assert(
        medioRow.activo === false || String(medioRow.activo).toLowerCase() === 'false',
        'El medio debe quedar desactivado.'
      );
    }
  },
  {
    id: 'compra_eliminacion_reciente_reconstruye_costos',
    title: 'Eliminar compra reciente limpia cascada y recompone costos',
    tags: ['compras', 'costos', 'correccion'],
    run: async ctx => {
      await ctx.reset();

      const results = await ctx.batch([
        batchStep('registrarCompraQTAS', compraPayloadBase({
          fechaCompra: '2026-06-20',
          proveedor: 'Proveedor Delete Test 1',
          lineas: [
            compraLinea('Insumo', 'CajaDeleteTest', 10, 'und', 5000, true, 'Costo base')
          ]
        })),
        batchStep('registrarCompraQTAS', compraPayloadBase({
          fechaCompra: '2026-06-21',
          proveedor: 'Proveedor Delete Test 2',
          lineas: [
            compraLinea('Insumo', 'CajaDeleteTest', 10, 'und', 7000, true, 'Costo a corregir')
          ]
        }))
      ]);
      const compra2 = results[1];

      const eliminacion = await ctx.call('eliminarCompraRecienteQTAS', {
        compraId: compra2.compraId
      });

      ctx.equal(ctx.num(eliminacion.removed.compras), 1, 'Debe eliminarse una compra.');
      ctx.equal(ctx.num(eliminacion.removed.compraDetalle), 1, 'Debe eliminarse un detalle de compra.');

      let state = await snapshotLigero(ctx, {
        sheetNames: ['Compras', 'Compra_Detalle', 'Costos_Referencia'],
        includeCompras: true
      });
      let costos = ctx.sheetRows(state, 'Costos_Referencia')
        .filter(row => row.Item === 'CajaDeleteTest');

      ctx.equal(ctx.sheetRows(state, 'Compras').length, 1, 'Solo debe quedar una compra.');
      ctx.equal(costos.length, 1, 'El historico de costos debe volver a una sola fila vigente.');
      ctx.equal(ctx.num(costos[0].Costo_Unitario), 500, 'Debe conservarse el costo de la compra restante.');

      const compra3 = await ctx.call('registrarCompraQTAS', compraPayloadBase({
        fechaCompra: '2026-06-22',
        proveedor: 'Proveedor Delete Test 3',
        lineas: [
          compraLinea('Insumo', 'CajaDeleteTest', 4, 'und', 3600, true, 'Reingreso corregido')
        ]
      }));

      ctx.equal(ctx.num(compra3.compraId), ctx.num(compra2.compraId), 'La nueva compra debe reutilizar el ultimo ID disponible.');

      state = await snapshotLigero(ctx, {
        sheetNames: ['Compras', 'Costos_Referencia'],
        includeCompras: true
      });
      costos = ctx.sheetRows(state, 'Costos_Referencia')
        .filter(row => row.Item === 'CajaDeleteTest');

      ctx.assert(
        (state.costosVigentes || []).some(row => row.item === 'CajaDeleteTest' && ctx.num(row.costoUnitario) === 900),
        'El costo vigente debe reflejar la compra corregida.'
      );
      ctx.equal(costos.length, 2, 'El historico debe conservar la fila previa y la corregida.');
    }
  },
  {
    id: 'venta_eliminacion_reciente_reusa_id',
    title: 'Eliminar venta reciente limpia cascada y libera el ultimo ID',
    tags: ['ventas', 'correccion'],
    run: async ctx => {
      await ctx.reset();

      const results = await ctx.batch([
        batchStep('registrarVentaQTAS', ventaPayloadBase({
          cliente: { nombre: 'Cliente Delete Base' },
          pendienteEnvio: true,
          lineas: [
            ventaLinea('AcSup', 1, 'g', 20000)
          ],
          pagos: [
            pagoLinea('Efectivo', 20000, 'Pago base')
          ]
        })),
        batchStep('registrarVentaQTAS', ventaPayloadBase({
          cliente: { nombre: 'Cliente Delete Target' },
          pendienteEnvio: true,
          lineas: [
            ventaLinea('AcSup', 1, 'g', 22000)
          ]
        }))
      ]);
      const venta2 = results[1];

      const eliminacion = await ctx.call('eliminarVentaRecienteQTAS', {
        ventaId: venta2.ventaId
      });

      ctx.equal(ctx.num(eliminacion.removed.ventas), 1, 'Debe eliminarse una venta.');
      ctx.equal(ctx.num(eliminacion.removed.detalle), 1, 'Debe eliminarse el detalle asociado.');
      ctx.equal(ctx.num(eliminacion.removed.distribucionIngresos), 1, 'Debe eliminarse la distribucion de la venta borrada.');
      ctx.equal(ctx.num(eliminacion.removed.ventasEnvio), 1, 'Debe eliminarse el seguimiento de envio.');

      let state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas', 'Venta_Detalle', 'Pagos', 'Distribucion_Ingresos', 'Ventas_Envio', 'Venta_Detalle_Costos_Calc'],
        includeDashboard: true
      });
      ctx.equal(ctx.sheetRows(state, 'Ventas').length, 1, 'Solo debe quedar la venta base.');
      ctx.equal(
        ctx.sheetRows(state, 'Venta_Detalle').filter(row => ctx.num(row.Venta_ID) === ctx.num(venta2.ventaId)).length,
        0,
        'No deben quedar detalles de la venta eliminada.'
      );
      ctx.equal(
        ctx.sheetRows(state, 'Venta_Detalle_Costos_Calc').filter(row => ctx.num(row.Venta_ID) === ctx.num(venta2.ventaId)).length,
        0,
        'No debe quedar analitica de costos para la venta eliminada.'
      );

      const venta3 = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Delete Reuse' },
        lineas: [
          ventaLinea('AcSup', 1, 'g', 21000)
        ]
      }));

      ctx.equal(ctx.num(venta3.ventaId), ctx.num(venta2.ventaId), 'La siguiente venta debe reutilizar el ultimo ID disponible.');

      state = await snapshotLigero(ctx, {
        sheetNames: ['Ventas'],
        includeDashboard: true
      });
      ctx.equal(ctx.sheetRows(state, 'Ventas').length, 2, 'Deben quedar dos ventas activas tras reingresar la corregida.');
    }
  },
  {
    id: 'inventario_compra_directa_actualiza_snapshot',
    title: 'Compra directa alimenta inventario y snapshot',
    tags: ['inventario', 'compras', 'smoke'],
    run: async ctx => {
      await ctx.reset();

      const compra = await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Inventario Test',
        lineas: [
          compraLinea('Insumo', 'Alcohol', 100, 'g', 2000, true, 'Stock alcohol'),
          compraLinea('Insumo', 'Bolsa_Zip_Negra', 20, 'und', 6000, true, 'Stock bolsas micros')
        ]
      }));

      ctx.assert(compra.inventario && compra.inventario.ok, 'La compra debe sincronizar inventario.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Inventario_Movimientos', 'Inventario_Snapshot']
      });
      const movimientos = ctx.sheetRows(state, 'Inventario_Movimientos');
      const snapshot = ctx.sheetRows(state, 'Inventario_Snapshot');
      const alcohol = snapshot.find(row => row.Item === 'Alcohol' && row.Unidad === 'g');
      const bolsas = snapshot.find(row => row.Item === 'Bolsa_Zip_Negra' && row.Unidad === 'und');

      ctx.equal(movimientos.length, 2, 'La compra debe crear dos movimientos de inventario.');
      ctx.assert(alcohol, 'Alcohol debe existir en snapshot.');
      ctx.assert(bolsas, 'Bolsa_Zip_Negra debe existir en snapshot.');
      ctx.equal(ctx.num(alcohol.Stock_Actual), 100, 'Alcohol debe quedar con stock 100.');
      ctx.equal(ctx.num(bolsas.Stock_Actual), 20, 'Bolsa_Zip_Negra debe quedar con stock 20.');
    }
  },
  {
    id: 'inventario_produccion_y_venta_fabricada',
    title: 'Produccion consume materia prima y venta baja terminado',
    tags: ['inventario', 'produccion', 'smoke'],
    run: async ctx => {
      await ctx.reset();
      const prep = await ctx.batch([
        batchStep('guardarComponenteProductoQTAS', {
          producto: '100mg',
          unidadVenta: 'und',
          orden: 10,
          tipoComponente: 'Producto',
          itemComponente: 'AcMed',
          cantidadComponente: 0.1,
          unidadComponente: 'g',
          mermaPct: 0,
          nota: 'Receta test inventario',
          activo: true
        }),
        batchStep('guardarComponenteProductoQTAS', {
          producto: '100mg',
          unidadVenta: 'und',
          orden: 20,
          tipoComponente: 'Insumo',
          itemComponente: 'Capsulas',
          cantidadComponente: 1,
          unidadComponente: 'und',
          mermaPct: 0,
          nota: 'Receta test inventario',
          activo: true
        }),
        batchStep('guardarReglaCostoProductoQTAS', {
          producto: '100mg',
          unidadVenta: 'und',
          cantidadMin: 1,
          cantidadMax: 24,
          orden: 10,
          tipoComponente: 'Insumo',
          itemComponente: 'Bolsa_Zip_Negra',
          cantidadComponente: 1,
          unidadComponente: 'und',
          aplicacion: 'PorLinea',
          mermaPct: 0,
          nota: 'Micros <25: zip negra.',
          activo: true
        }),
        batchStep('guardarReglaCostoProductoQTAS', {
          producto: '100mg',
          unidadVenta: 'und',
          cantidadMin: 1,
          cantidadMax: 24,
          orden: 20,
          tipoComponente: 'Insumo',
          itemComponente: 'Bolsa_Papel_0_5lb',
          cantidadComponente: 1,
          unidadComponente: 'und',
          aplicacion: 'PorLinea',
          mermaPct: 0,
          nota: 'Micros <25: bolsa papel 0.5 lb.',
          activo: true
        }),
        batchStep('guardarReglaCostoProductoQTAS', {
          producto: '100mg',
          unidadVenta: 'und',
          cantidadMin: 1,
          cantidadMax: 24,
          orden: 30,
          tipoComponente: 'Insumo',
          itemComponente: 'Calca_Micros_Logo',
          cantidadComponente: 1,
          unidadComponente: 'und',
          aplicacion: 'PorLinea',
          mermaPct: 0,
          nota: 'Micros <25: calca logo.',
          activo: true
        })
      ]);
      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Produccion Test',
        lineas: [
          compraLinea('Producto', 'AcMed', 10, 'g', 140000, true, 'Base activa'),
          compraLinea('Insumo', 'Capsulas', 100, 'und', 1500, true, 'Capsulas vacias'),
          compraLinea('Insumo', 'Bolsa_Zip_Negra', 10, 'und', 3000, true, 'Zip negra micros'),
          compraLinea('Insumo', 'Bolsa_Papel_0_5lb', 10, 'und', 3000, true, 'Bolsa papel micros'),
          compraLinea('Insumo', 'Calca_Micros_Logo', 10, 'und', 5000, true, 'Calca logo micros')
        ]
      }));
      const produccion = await ctx.call('registrarProduccionQTAS', {
        fechaProduccion: '2026-06-20',
        producto: '100mg',
        unidad: 'und',
        cantidad: 10,
        comentarioProduccion: 'Lote automatizado'
      });
      const venta = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Inventario Fabricado' },
        lineas: [
          ventaLinea('100mg', 2, 'und', 2000)
        ]
      }));

      ctx.assert(Array.isArray(prep) && prep.length === 5, 'La preparacion de receta y packaging debe ejecutarse completa.');
      ctx.assert(produccion.ok, 'La produccion debe registrarse correctamente.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Inventario_Movimientos', 'Inventario_Snapshot', 'Producciones', 'Produccion_Detalle']
      });
      const snapshot = ctx.sheetRows(state, 'Inventario_Snapshot');
      const movimientosVenta = ctx.sheetRows(state, 'Inventario_Movimientos')
        .filter(row => ctx.num(row.Venta_ID) === ctx.num(venta.ventaId));
      const stock100 = snapshot.find(row => row.Item === '100mg' && row.Unidad === 'und');
      const stockAcMed = snapshot.find(row => row.Item === 'AcMed' && row.Unidad === 'g');
      const stockCapsulas = snapshot.find(row => row.Item === 'Capsulas' && row.Unidad === 'und');
      const stockZipNegra = snapshot.find(row => row.Item === 'Bolsa_Zip_Negra' && row.Unidad === 'und');
      const stockBolsaPapel = snapshot.find(row => row.Item === 'Bolsa_Papel_0_5lb' && row.Unidad === 'und');
      const stockCalcaMicros = snapshot.find(row => row.Item === 'Calca_Micros_Logo' && row.Unidad === 'und');

      ctx.equal(ctx.sheetRows(state, 'Producciones').length, 1, 'Debe existir un encabezado de produccion.');
      ctx.assert(ctx.sheetRows(state, 'Produccion_Detalle').length >= 3, 'La produccion debe materializar entradas y salidas.');
      ctx.equal(ctx.num(stock100.Stock_Actual), 8, 'El terminado 100mg debe quedar en 8.');
      ctx.equal(ctx.num(stockAcMed.Stock_Actual), 9, 'AcMed debe bajar por la produccion del lote.');
      ctx.equal(ctx.num(stockCapsulas.Stock_Actual), 90, 'Capsulas debe bajar por la produccion.');
      ctx.equal(ctx.num(stockZipNegra.Stock_Actual), 9, 'La venta debe consumir una zip negra por el pedido.');
      ctx.equal(ctx.num(stockBolsaPapel.Stock_Actual), 9, 'La venta debe consumir una bolsa papel 0.5 lb por el pedido.');
      ctx.equal(ctx.num(stockCalcaMicros.Stock_Actual), 9, 'La venta debe consumir una calca logo por el pedido.');
      ctx.assert(
        movimientosVenta.some(row => row.Item === 'Bolsa_Zip_Negra' && row.Operacion === 'Salida'),
        'La venta debe consumir Bolsa_Zip_Negra.'
      );
      ctx.assert(
        movimientosVenta.some(row => row.Item === 'Bolsa_Papel_0_5lb' && row.Operacion === 'Salida'),
        'La venta debe consumir Bolsa_Papel_0_5lb.'
      );
      ctx.assert(
        movimientosVenta.some(row => row.Item === 'Calca_Micros_Logo' && row.Operacion === 'Salida'),
        'La venta debe consumir Calca_Micros_Logo.'
      );
      ctx.assert(
        !movimientosVenta.some(row => row.Item === 'Bolsa_Barata' || row.Item === 'Calcas'),
        'La venta ya no debe consumir empaques o calcas genericas legacy.'
      );
      ctx.assert(
        !movimientosVenta.some(row => row.Item === 'Frasco_Capsulas' || row.Item === 'Calca_Micros_Instrucciones'),
        'Las ventas menores a 25 unidades no deben consumir frasco ni calca de instrucciones.'
      );
    }
  },
  {
    id: 'inventario_extracto_usa_receta_alineada',
    title: 'Produccion de extracto consume insumos especificos de la receta alineada',
    tags: ['inventario', 'produccion', 'extractos', 'smoke'],
    run: async ctx => {
      await ctx.reset();
      const preparacion = await ctx.batch([
        batchStep('guardarComponenteProductoQTAS', {
          producto: 'CordyExt',
          unidadVenta: 'und',
          orden: 10,
          tipoComponente: 'Producto',
          itemComponente: 'Cordy',
          cantidadComponente: 7.5,
          unidadComponente: 'g',
          mermaPct: 0,
          nota: 'Receta operativa de extracto',
          activo: true
        }),
        batchStep('guardarComponenteProductoQTAS', {
          producto: 'CordyExt',
          unidadVenta: 'und',
          orden: 20,
          tipoComponente: 'Insumo',
          itemComponente: 'Goteros',
          cantidadComponente: 1,
          unidadComponente: 'und',
          mermaPct: 0,
          nota: 'Receta operativa de extracto',
          activo: true
        }),
        batchStep('guardarComponenteProductoQTAS', {
          producto: 'CordyExt',
          unidadVenta: 'und',
          orden: 30,
          tipoComponente: 'Insumo',
          itemComponente: 'Alcohol',
          cantidadComponente: 20,
          unidadComponente: 'g',
          mermaPct: 0,
          nota: 'Receta operativa de extracto',
          activo: true
        }),
        batchStep('guardarComponenteProductoQTAS', {
          producto: 'CordyExt',
          unidadVenta: 'und',
          orden: 50,
          tipoComponente: 'Insumo',
          itemComponente: 'Bolsa_Papel_1lb',
          cantidadComponente: 1,
          unidadComponente: 'und',
          mermaPct: 0,
          nota: 'Receta operativa de extracto',
          activo: true
        }),
        batchStep('guardarComponenteProductoQTAS', {
          producto: 'CordyExt',
          unidadVenta: 'und',
          orden: 60,
          tipoComponente: 'Insumo',
          itemComponente: 'Calca_Cordy_Ext',
          cantidadComponente: 1,
          unidadComponente: 'und',
          mermaPct: 0,
          nota: 'Receta operativa de extracto',
          activo: true
        })
      ]);
      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Extracto Test',
        lineas: [
          compraLinea('Producto', 'Cordy', 100, 'g', 31000, true, 'Hongo base extracto'),
          compraLinea('Insumo', 'Alcohol', 200, 'g', 2526, true, 'Alcohol extracto'),
          compraLinea('Insumo', 'Goteros', 10, 'und', 10000, true, 'Goteros extracto'),
          compraLinea('Insumo', 'Bolsa_Papel_1lb', 10, 'und', 5000, true, 'Bolsa papel extracto'),
          compraLinea('Insumo', 'Calca_Cordy_Ext', 10, 'und', 5000, true, 'Calca extracto')
        ]
      }));
      const produccion = await ctx.call('registrarProduccionQTAS', {
        fechaProduccion: '2026-06-20',
        producto: 'CordyExt',
        unidad: 'und',
        cantidad: 2,
        comentarioProduccion: 'Lote extracto receta alineada'
      });
      ctx.assert(Array.isArray(preparacion) && preparacion.length === 5, 'La receta operativa del extracto debe prepararse completa.');
      ctx.assert(produccion.ok, 'La produccion del extracto debe registrarse correctamente.');

      const state = await snapshotLigero(ctx, {
        sheetNames: ['Inventario_Movimientos', 'Inventario_Snapshot', 'Producciones', 'Produccion_Detalle']
      });
      const snapshot = ctx.sheetRows(state, 'Inventario_Snapshot');
      const movimientos = ctx.sheetRows(state, 'Inventario_Movimientos')
        .filter(row => String(row.Produccion_ID || '') === String(produccion.produccionId || ''));
      const detalle = ctx.sheetRows(state, 'Produccion_Detalle')
        .filter(row => String(row.Produccion_ID || '') === String(produccion.produccionId || ''));

      const stockCordyExt = snapshot.find(row => row.Item === 'CordyExt' && row.Unidad === 'und');
      const stockCordy = snapshot.find(row => row.Item === 'Cordy' && row.Unidad === 'g');
      const stockAlcohol = snapshot.find(row => row.Item === 'Alcohol' && row.Unidad === 'g');
      const stockGoteros = snapshot.find(row => row.Item === 'Goteros' && row.Unidad === 'und');
      const stockBolsaPapel = snapshot.find(row => row.Item === 'Bolsa_Papel_1lb' && row.Unidad === 'und');
      const stockCalcaCordyExt = snapshot.find(row => row.Item === 'Calca_Cordy_Ext' && row.Unidad === 'und');

      ctx.equal(detalle.length, 6, 'La produccion debe materializar 1 entrada y 5 salidas inventariables.');
      ctx.assert(stockCordyExt, 'CordyExt debe existir en snapshot.');
      ctx.assert(stockCordy, 'Cordy debe existir en snapshot.');
      ctx.assert(stockAlcohol, 'Alcohol debe existir en snapshot.');
      ctx.assert(stockGoteros, 'Goteros debe existir en snapshot.');
      ctx.assert(stockBolsaPapel, 'Bolsa_Papel_1lb debe existir en snapshot.');
      ctx.assert(stockCalcaCordyExt, 'Calca_Cordy_Ext debe existir en snapshot.');

      ctx.equal(ctx.num(stockCordyExt.Stock_Actual), 2, 'CordyExt fabricado debe quedar en 2.');
      ctx.equal(ctx.num(stockCordy.Stock_Actual), 85, 'Cordy debe bajar 15 g por 2 extractos.');
      ctx.equal(ctx.num(stockAlcohol.Stock_Actual), 160, 'Alcohol debe bajar 40 g por 2 extractos.');
      ctx.equal(ctx.num(stockGoteros.Stock_Actual), 8, 'Goteros debe bajar 2 unidades.');
      ctx.equal(ctx.num(stockBolsaPapel.Stock_Actual), 8, 'Bolsa_Papel_1lb debe bajar 2 unidades.');
      ctx.equal(ctx.num(stockCalcaCordyExt.Stock_Actual), 8, 'Calca_Cordy_Ext debe bajar 2 unidades.');

      ctx.assert(
        movimientos.some(row => row.Item === 'Cordy' && row.Operacion === 'Salida'),
        'La produccion debe consumir Cordy como base.'
      );
      ctx.assert(
        movimientos.some(row => row.Item === 'Alcohol' && row.Operacion === 'Salida'),
        'La produccion debe consumir Alcohol.'
      );
      ctx.assert(
        movimientos.some(row => row.Item === 'Goteros' && row.Operacion === 'Salida'),
        'La produccion debe consumir Goteros.'
      );
      ctx.assert(
        movimientos.some(row => row.Item === 'Bolsa_Papel_1lb' && row.Operacion === 'Salida'),
        'La produccion debe consumir Bolsa_Papel_1lb.'
      );
      ctx.assert(
        movimientos.some(row => row.Item === 'Calca_Cordy_Ext' && row.Operacion === 'Salida'),
        'La produccion debe consumir Calca_Cordy_Ext.'
      );
      ctx.assert(
        !movimientos.some(row => row.Item === 'Bolsa_Barata' || row.Item === 'Calcas'),
        'La produccion del extracto no debe usar empaques o calcas genericas.'
      );
      ctx.assert(
        !movimientos.some(row => row.Item === 'Agua'),
        'Agua no debe generar movimiento porque quedo NoControlado.'
      );
    }
  }
];
