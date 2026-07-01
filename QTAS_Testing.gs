function testPingQTAS() {
  const ss = SpreadsheetApp.getActive();
  return testSerializarValorQTAS_({
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    timezone: ss.getSpreadsheetTimeZone(),
    sheets: ss.getSheets().map(sheet => sheet.getName())
  });
}

function testLimpiarRegistrosPruebaQTAS(payload) {
  assertOperacionDestructivaPermitidaQTAS_('testLimpiarRegistrosPruebaQTAS');

  const settings = Object.assign({
    dryRun: false
  }, payload || {});
  const nombresClientePrueba = [
    'Cliente Test',
    'Cliente Deuda Test',
    'Cliente Pagado Test',
    'Cliente Multi Test',
    'Cliente Pago Pendiente',
    'Cliente Regla Test',
    'Cliente Envio Test'
  ];
  const proveedoresPrueba = [
    'Proveedor Test',
    'Proveedor Costo Test',
    'Proveedor Gasto Test',
    'Proveedor Mixto Test',
    'Proveedor Medio Test',
    'Proveedor Medio Test 2',
    'Proveedor Producto Invalido Test',
    'Proveedor Receta Test'
  ];
  const productosPrueba = ['ProdTestAuto', 'ProdCostoCompuesto'];
  const mediosPagoPrueba = ['TransferTest'];
  const notaReglaPrueba = 'Regla automatizada';

  return withScriptLock_('test limpiar registros prueba', () => {
    const ss = SpreadsheetApp.getActive();
    const configSheet = materializarConfigMediosPagoQTAS_();
    const sheets = {
      productos: ss.getSheetByName(QTAS.sheets.productos),
      precios: ss.getSheetByName(QTAS.sheets.precios),
      compras: ss.getSheetByName(QTAS.sheets.compras),
      compraDetalle: ss.getSheetByName(QTAS.sheets.compraDetalle),
      costosReferencia: ss.getSheetByName(QTAS.sheets.costosReferencia),
      productoComponentes: ss.getSheetByName(QTAS.sheets.productoComponentes),
      costoProductoCalculado: ss.getSheetByName(QTAS.sheets.costoProductoCalculado),
      ventaDetalleCostosCalculado: ss.getSheetByName(QTAS.sheets.ventaDetalleCostosCalculado),
      origenesFondosReglas: ss.getSheetByName(QTAS.sheets.origenesFondosReglas),
      compraOrigenesFondos: ss.getSheetByName(QTAS.sheets.compraOrigenesFondos),
      clientes: ss.getSheetByName(QTAS.sheets.clientes),
      ventas: ss.getSheetByName(QTAS.sheets.ventas),
      detalle: ss.getSheetByName(QTAS.sheets.detalle),
      pagos: ss.getSheetByName(QTAS.sheets.pagos),
      ventasEnvio: ss.getSheetByName(QTAS.sheets.ventasEnvio),
      distribucionReglas: ss.getSheetByName(QTAS.sheets.distribucionReglas),
      distribucionIngresos: ss.getSheetByName(QTAS.sheets.distribucionIngresos),
      config: configSheet
    };
    const headers = {};
    Object.keys(sheets).forEach(key => {
      headers[key] = QTAS.schemas[sheets[key] ? sheets[key].getName() : ''] || [];
    });

    const ventasAntes = leerObjetos_(sheets.ventas);
    const ventasPrueba = ventasAntes.filter(row => testEsVentaPruebaQTAS_(row, nombresClientePrueba));
    const ventaIdsPrueba = testConstruirSetQTAS_(ventasPrueba.map(row => texto_(row.Venta_ID)));
    const clienteIdsPrueba = testConstruirSetQTAS_(ventasPrueba.map(row => texto_(row.Cliente_ID)));

    const comprasAntes = leerObjetos_(sheets.compras);
    const comprasPrueba = comprasAntes.filter(row => testEsCompraPruebaQTAS_(row, proveedoresPrueba));
    const compraIdsPrueba = testConstruirSetQTAS_(comprasPrueba.map(row => texto_(row.Compra_ID)));

    const detalleAntes = leerObjetos_(sheets.detalle);
    const pagosAntes = leerObjetos_(sheets.pagos);
    const ventasEnvioAntes = sheets.ventasEnvio ? leerObjetos_(sheets.ventasEnvio) : [];
    const distribucionAntes = leerObjetos_(sheets.distribucionIngresos);
    const compraDetalleAntes = leerObjetos_(sheets.compraDetalle);
    const costosAntes = leerObjetos_(sheets.costosReferencia);
    const componentesAntes = sheets.productoComponentes ? leerObjetos_(sheets.productoComponentes) : [];
    const costoProductoAntes = sheets.costoProductoCalculado ? leerObjetos_(sheets.costoProductoCalculado) : [];
    const ventaDetalleCostosAntes = sheets.ventaDetalleCostosCalculado ? leerObjetos_(sheets.ventaDetalleCostosCalculado) : [];
    const origenesFondosReglasAntes = sheets.origenesFondosReglas ? leerObjetos_(sheets.origenesFondosReglas) : [];
    const compraOrigenesAntes = sheets.compraOrigenesFondos ? leerObjetos_(sheets.compraOrigenesFondos) : [];
    const clientesAntes = leerObjetos_(sheets.clientes);
    const productosAntes = leerObjetos_(sheets.productos);
    const preciosAntes = leerObjetos_(sheets.precios);
    const configAntes = leerObjetos_(sheets.config);
    const reglasAntes = leerObjetos_(sheets.distribucionReglas);

    const ventasDespues = ventasAntes.filter(row => !ventaIdsPrueba[texto_(row.Venta_ID)]);
    const detalleDespues = detalleAntes.filter(row => !ventaIdsPrueba[texto_(row.Venta_ID)]);
    const pagosDespues = pagosAntes.filter(row => !ventaIdsPrueba[texto_(row.Venta_ID)]);
    const ventasEnvioDespues = ventasEnvioAntes.filter(row => !ventaIdsPrueba[texto_(row.Venta_ID)]);
    const distribucionDespues = distribucionAntes.filter(row =>
      !ventaIdsPrueba[texto_(row.Venta_ID)] &&
      !testIncluyeValorQTAS_(nombresClientePrueba, row.Nombre)
    );
    const comprasDespues = comprasAntes.filter(row => !compraIdsPrueba[texto_(row.Compra_ID)]);
    const compraDetalleDespues = compraDetalleAntes.filter(row =>
      !compraIdsPrueba[texto_(row.Compra_ID)] &&
      !testIncluyeValorQTAS_(proveedoresPrueba, row.Proveedor)
    );
    const costosDespues = costosAntes.filter(row =>
      !compraIdsPrueba[texto_(row.Compra_ID)] &&
      !testIncluyeValorQTAS_(proveedoresPrueba, row.Proveedor)
    );
    const componentesDespues = componentesAntes.filter(row =>
      !testIncluyeValorQTAS_(productosPrueba, row.Producto_Estandar) &&
      texto_(row.Nota) !== 'Componente de prueba'
    );
    const costoProductoDespues = costoProductoAntes.filter(row =>
      !testIncluyeValorQTAS_(productosPrueba, row.Producto_Estandar)
    );
    const ventaDetalleCostosDespues = ventaDetalleCostosAntes.filter(row =>
      !testIncluyeValorQTAS_(productosPrueba, row.Producto_Estandar) &&
      !testIncluyeValorQTAS_(nombresClientePrueba, row.Nombre)
    );
    const origenesFondosReglasDespues = origenesFondosReglasAntes.filter(row =>
      texto_(row.Nota) !== 'Regla de fondos de prueba'
    );
    const compraOrigenesDespues = compraOrigenesAntes.filter(row =>
      !compraIdsPrueba[texto_(row.Compra_ID)] &&
      texto_(row.Fuente_Registro).indexOf('TEST-LEGACY') < 0
    );
    const clientesDespues = clientesAntes.filter(row =>
      !clienteIdsPrueba[texto_(row.Cliente_ID)] &&
      !testIncluyeValorQTAS_(nombresClientePrueba, row.Nombre)
    );
    const productosDespues = productosAntes.filter(row =>
      !(testIncluyeValorQTAS_(productosPrueba, row.Producto_Estandar) && texto_(row.Nota) === 'Creado por test')
    );
    const preciosDespues = preciosAntes.filter(row =>
      !(
        testIncluyeValorQTAS_(productosPrueba, row.Producto_Estandar) ||
        texto_(row.Nota) === 'Precio de prueba'
      )
    );
    const configDespues = configAntes.filter(row =>
      !(
        testIncluyeValorQTAS_(mediosPagoPrueba, row.Medio_Pago) ||
        texto_(row.Nota) === 'Creado por test'
      )
    );
    const reglasDespues = testNormalizarReglasTrasLimpiezaQTAS_(
      reglasAntes.filter(row => texto_(row.Nota) !== notaReglaPrueba)
    );

    const resumen = {
      ok: true,
      dryRun: Boolean(settings.dryRun),
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      removed: {
        ventas: ventasAntes.length - ventasDespues.length,
        detalle: detalleAntes.length - detalleDespues.length,
        pagos: pagosAntes.length - pagosDespues.length,
        ventasEnvio: ventasEnvioAntes.length - ventasEnvioDespues.length,
        distribucionIngresos: distribucionAntes.length - distribucionDespues.length,
        compras: comprasAntes.length - comprasDespues.length,
        compraDetalle: compraDetalleAntes.length - compraDetalleDespues.length,
        costosReferencia: costosAntes.length - costosDespues.length,
        productoComponentes: componentesAntes.length - componentesDespues.length,
        costoProductoCalculado: costoProductoAntes.length - costoProductoDespues.length,
        ventaDetalleCostosCalculado: ventaDetalleCostosAntes.length - ventaDetalleCostosDespues.length,
        origenesFondosReglas: origenesFondosReglasAntes.length - origenesFondosReglasDespues.length,
        compraOrigenesFondos: compraOrigenesAntes.length - compraOrigenesDespues.length,
        clientes: clientesAntes.length - clientesDespues.length,
        productos: productosAntes.length - productosDespues.length,
        precios: preciosAntes.length - preciosDespues.length,
        configMediosPago: configAntes.length - configDespues.length,
        distribucionReglas: reglasAntes.length - reglasDespues.length
      }
    };

    if (!settings.dryRun) {
      reemplazarObjetos_(sheets.ventas, headers.ventas, ventasDespues);
      reemplazarObjetos_(sheets.detalle, headers.detalle, detalleDespues);
      reemplazarObjetos_(sheets.pagos, headers.pagos, pagosDespues);
      if (sheets.ventasEnvio) {
        reemplazarObjetos_(sheets.ventasEnvio, headers.ventasEnvio, ventasEnvioDespues);
      }
      reemplazarObjetos_(sheets.distribucionIngresos, headers.distribucionIngresos, distribucionDespues);
      reemplazarObjetos_(sheets.compras, headers.compras, comprasDespues);
      reemplazarObjetos_(sheets.compraDetalle, headers.compraDetalle, compraDetalleDespues);
      reemplazarObjetos_(sheets.costosReferencia, headers.costosReferencia, costosDespues);
      if (sheets.productoComponentes) {
        reemplazarObjetos_(sheets.productoComponentes, headers.productoComponentes, componentesDespues);
      }
      if (sheets.costoProductoCalculado) {
        reemplazarObjetos_(sheets.costoProductoCalculado, headers.costoProductoCalculado, costoProductoDespues);
      }
      if (sheets.ventaDetalleCostosCalculado) {
        reemplazarObjetos_(sheets.ventaDetalleCostosCalculado, headers.ventaDetalleCostosCalculado, ventaDetalleCostosDespues);
      }
      if (sheets.origenesFondosReglas) {
        reemplazarObjetos_(sheets.origenesFondosReglas, headers.origenesFondosReglas, origenesFondosReglasDespues);
      }
      if (sheets.compraOrigenesFondos) {
        reemplazarObjetos_(sheets.compraOrigenesFondos, headers.compraOrigenesFondos, compraOrigenesDespues);
      }
      reemplazarObjetos_(sheets.clientes, headers.clientes, clientesDespues);
      reemplazarObjetos_(sheets.productos, headers.productos, productosDespues);
      reemplazarObjetos_(sheets.precios, headers.precios, preciosDespues);
      reemplazarObjetos_(sheets.config, headers.config, configDespues);
      reemplazarObjetos_(sheets.distribucionReglas, headers.distribucionReglas, reglasDespues);

      limpiarCachesEjecucionQTAS_();
      invalidarCacheDocumentoQTAS_('precios_referencia_memoria');
      invalidarCacheDocumentoQTAS_('distribucion_reglas_memoria');
    }

    return testSerializarValorQTAS_(resumen);
  });
}

function testResetEntornoQTAS(payload) {
  assertOperacionDestructivaPermitidaQTAS_('testResetEntornoQTAS');

  const settings = Object.assign({
    aplicarFormatos: false
  }, payload || {});

  return withScriptLock_('test reset entorno', () => {
    const ss = asegurarModeloOperativoQTAS_({
      aplicarFormatos: Boolean(settings.aplicarFormatos)
    });
    const configSheet = materializarConfigMediosPagoQTAS_();
    const sheetNames = testHojasSoportadasQTAS_()
      .concat([configSheet.getName()])
      .filter((sheetName, index, array) => array.indexOf(sheetName) === index);

    sheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      limpiarDatos_(sheet);
      limpiarCacheHeadersHojaQTAS_(sheet);
    });

    testResetearPropiedadesQTAS_('QTAS_SEQ_');
    limpiarCachesEjecucionQTAS_();
    invalidarCacheDocumentoQTAS_('precios_referencia_memoria');
    invalidarCacheDocumentoQTAS_('distribucion_reglas_memoria');

    sembrarProductosYPrecios_();
    sembrarConfig_();
    sembrarReglasDistribucion_();

    if (settings.aplicarFormatos) {
      aplicarFormatosModeloQTAS_(ss);
    }

    return testSnapshotQTAS({
      includeDashboard: true,
      includeCompras: true,
      includeConfig: true
    });
  });
}

function testSnapshotQTAS(payload) {
  const settings = Object.assign({
    includeDashboard: true,
    includeCompras: true,
    includeConfig: true,
    sheetNames: testHojasSoportadasQTAS_()
  }, payload || {});
  const ss = SpreadsheetApp.getActive();
  const sheets = {};

  (settings.sheetNames || []).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    sheets[sheetName] = sheet ? leerObjetos_(sheet) : [];
  });

  const snapshot = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheets: sheets
  };

  if (settings.includeDashboard !== false) {
    snapshot.dashboard = dashboardVentasConsistenteQTAS_();
  }

  if (settings.includeCompras !== false) {
    snapshot.comprasRecientes = listarComprasRecientesQTAS_();
    snapshot.costosVigentes = listarCostosVigentesQTAS_();
  }

  if (settings.includeConfig !== false) {
    snapshot.configuracionAvanzada = getConfiguracionAvanzadaQTAS();
  }

  return testSerializarValorQTAS_(snapshot);
}

function testHojasSoportadasQTAS_() {
  return [
    QTAS.sheets.productos,
    QTAS.sheets.precios,
    QTAS.sheets.compras,
    QTAS.sheets.compraDetalle,
    QTAS.sheets.costosReferencia,
    QTAS.sheets.productoComponentes,
    QTAS.sheets.costoProductoCalculado,
    QTAS.sheets.ventaDetalleCostosCalculado,
    QTAS.sheets.origenesFondosReglas,
    QTAS.sheets.compraOrigenesFondos,
    QTAS.sheets.clientes,
    QTAS.sheets.ventas,
    QTAS.sheets.detalle,
    QTAS.sheets.pagos,
    QTAS.sheets.ventasEnvio,
    QTAS.sheets.distribucionReglas,
    QTAS.sheets.distribucionIngresos,
    QTAS.sheets.config
  ];
}

function testResetearPropiedadesQTAS_(prefix) {
  const props = PropertiesService.getDocumentProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(key => {
    if (key.indexOf(prefix) === 0) {
      props.deleteProperty(key);
    }
  });
}

function testSerializarValorQTAS_(value) {
  if (value === null || value === undefined) return value;
  if (value instanceof Date) {
    return Utilities.formatDate(value, zonaHorariaQTAS_(), 'yyyy-MM-dd HH:mm:ss');
  }
  if (Array.isArray(value)) {
    return value.map(item => testSerializarValorQTAS_(item));
  }
  if (typeof value === 'object') {
    const output = {};
    Object.keys(value).forEach(key => {
      output[key] = testSerializarValorQTAS_(value[key]);
    });
    return output;
  }
  return value;
}

function testEsVentaPruebaQTAS_(row, nombresClientePrueba) {
  return texto_(row.Comentario_Venta) === 'Escenario automatizado' ||
    testIncluyeValorQTAS_(nombresClientePrueba, row.Nombre);
}

function testEsCompraPruebaQTAS_(row, proveedoresPrueba) {
  return texto_(row.Comentario_Compra) === 'Escenario automatizado' ||
    testIncluyeValorQTAS_(proveedoresPrueba, row.Proveedor);
}

function testIncluyeValorQTAS_(values, value) {
  return (values || []).indexOf(texto_(value)) >= 0;
}

function testConstruirSetQTAS_(values) {
  return (values || []).reduce((acc, value) => {
    const key = texto_(value);
    if (key) acc[key] = true;
    return acc;
  }, {});
}

function testNormalizarReglasTrasLimpiezaQTAS_(rows) {
  const activas = (rows || [])
    .map(row => Object.assign({}, row))
    .sort((a, b) =>
      resolverFechaOperacion_(a.Fecha_Desde, new Date()) - resolverFechaOperacion_(b.Fecha_Desde, new Date())
    );

  activas.forEach((row, index) => {
    row.Fecha_Hasta = index < activas.length - 1
      ? diaAnterior_(activas[index + 1].Fecha_Desde)
      : '';
  });

  return activas;
}
