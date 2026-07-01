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

export const SCENARIOS = [
  {
    id: 'producto_y_precio_nuevo',
    title: 'Producto nuevo y cambio de precio desde configuracion',
    tags: ['admin', 'catalogo', 'precios'],
    run: async ctx => {
      await ctx.reset();

      const producto = 'ProdTestAuto';
      const unidad = 'und';
      await ctx.call('guardarProductoConfiguracionQTAS', {
        producto,
        unidad,
        nota: 'Creado por test',
        activo: true
      });

      await ctx.call('guardarCambioPrecioFrontendQTAS', {
        producto,
        unidad,
        precio: 12345,
        fechaDesde: '2026-06-20',
        nota: 'Precio de prueba'
      });

      const catalogo = await ctx.call('getCatalogoQTAS', '2026-06-21');
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
          pagoLinea('Efectivo', 20000, 'Pago total')
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
    tags: ['ventas', 'envios', 'sin-deuda'],
    run: async ctx => {
      await ctx.reset();

      const venta = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Envio Test' },
        pendienteEnvio: true,
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000)
        ],
        pagos: [
          pagoLinea('Efectivo', 20000, 'Pago total')
        ]
      }));

      ctx.equal(String(venta.estadoEnvio), 'Pendiente', 'La venta debe quedar marcada como pendiente de envio.');

      let state = await snapshotLigero(ctx, {
        includeDashboard: true
      });
      ctx.equal((state.dashboard.ventasPendientes || []).length, 0, 'La venta pagada no debe quedar como deuda.');
      ctx.equal((state.dashboard.enviosPendientes || []).length, 1, 'La venta debe aparecer en el panel de envios pendientes.');

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
      await ctx.call('guardarProductoConfiguracionQTAS', {
        producto,
        unidad: 'und',
        nota: 'Creado por test',
        activo: true
      });

      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Costo Test',
        lineas: [
          compraLinea('Producto', producto, 10, 'und', 100000, true, 'Costo incremental base')
        ]
      }));

      const venta = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Test' },
        lineas: [
          ventaLinea(producto, 1, 'und', 15000)
        ]
      }));

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
        sheetNames: ['Costo_Producto_Calc']
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
    id: 'costo_producto_por_componentes',
    title: 'Costo de producto compuesto y margen por venta',
    tags: ['costos', 'receta', 'ventas'],
    run: async ctx => {
      await ctx.reset();

      const producto = 'ProdCostoCompuesto';
      await ctx.call('guardarProductoConfiguracionQTAS', {
        producto,
        unidad: 'und',
        nota: 'Creado por test',
        activo: true
      });

      await ctx.call('guardarCambioPrecioFrontendQTAS', {
        producto,
        unidad: 'und',
        precio: 2000,
        fechaDesde: '2026-06-20',
        nota: 'Precio producto compuesto'
      });

      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Receta Test',
        lineas: [
          compraLinea('Insumo', 'CajaTest', 10, 'und', 5000, true, 'Caja para test')
        ]
      }));

      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Receta Test',
        lineas: [
          compraLinea('Insumo', 'EtiquetaTest', 20, 'und', 2000, true, 'Etiqueta para test')
        ]
      }));

      await ctx.call('guardarComponenteProductoQTAS', {
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
      });

      await ctx.call('guardarComponenteProductoQTAS', {
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
      });

      await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        cliente: { nombre: 'Cliente Test' },
        lineas: [
          ventaLinea(producto, 1, 'und', 2000)
        ]
      }));

      await ctx.call('reconstruirCostoProductoCalculadoQTAS', {
        fechaBase: '2026-06-20',
        silent: true
      });

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

      const configuracion = await ctx.call('guardarReglaDistribucionFrontendQTAS', {
        fechaDesde: '2026-07-01',
        steve: 50,
        majo: 30,
        mush: 20,
        nota: 'Regla automatizada'
      });

      ctx.assert(
        (configuracion.reglasDistribucion || []).some(row => row.fechaDesde === '2026-07-01' && ctx.num(row.steve) === 50),
        'La nueva regla debe quedar registrada.'
      );

      const venta = await ctx.call('registrarVentaQTAS', ventaPayloadBase({
        fechaVenta: '2026-07-02',
        cliente: { nombre: 'Cliente Regla Test' },
        lineas: [
          ventaLinea('AcSup', 1, 'g', 20000)
        ]
      }));

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
      await ctx.call('guardarMedioPagoQTAS', {
        medioPago: medio,
        nota: 'Creado por test',
        activo: true
      });

      await ctx.call('registrarCompraQTAS', compraPayloadBase({
        proveedor: 'Proveedor Medio Test',
        medioPago: medio,
        lineas: [
          compraLinea('Producto', 'AcSup', 2, 'g', 18000, true, 'Compra con medio nuevo')
        ]
      }));

      await ctx.call('cambiarEstadoMedioPagoQTAS', {
        medioPago: medio,
        activo: false
      });

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
  }
];
