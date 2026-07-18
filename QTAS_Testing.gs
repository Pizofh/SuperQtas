function testPingQTAS() {
  const ss = SpreadsheetApp.getActive();
  const fechaHistoricaEsperada = '2024-01-01 00:00:00';
  const fechaHistorica = fechaMomentoExactaQTAS_(fechaHistoricaEsperada);
  const fechaHistoricaNormalizada = Utilities.formatDate(
    fechaHistorica,
    zonaHorariaQTAS_(),
    'yyyy-MM-dd HH:mm:ss'
  );

  if (fechaHistoricaNormalizada !== fechaHistoricaEsperada) {
    throw new Error(
      `Fecha historica desplazada: ${fechaHistoricaEsperada} -> ${fechaHistoricaNormalizada}.`
    );
  }

  return testSerializarValorQTAS_({
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    timezone: ss.getSpreadsheetTimeZone(),
    historicalDateParsing: fechaHistoricaNormalizada,
    sheets: ss.getSheets().map(sheet => sheet.getName())
  });
}

function testAgruparReglasOrigenesFondosLegacyQTAS() {
  const reglas = agruparReglasOrigenesFondosQTAS_([
    {
      Regla_ID: 'CAJA-OLD-STEV',
      Origen_Fondos: 'Caja',
      Fecha_Desde: '2024-05-31',
      Fecha_Hasta: '2025-07-31',
      Aportante: 'Steve',
      Porcentaje: 40,
      Nota: 'Historico Caja'
    },
    {
      Regla_ID: 'CAJA-OLD-MAJO',
      Origen_Fondos: 'Caja',
      Fecha_Desde: '2024-05-31',
      Fecha_Hasta: '2025-07-31',
      Aportante: 'Majo',
      Porcentaje: 40,
      Nota: 'Historico Caja'
    },
    {
      Regla_ID: 'CAJA-OLD-MUSH',
      Origen_Fondos: 'Caja',
      Fecha_Desde: '2024-05-31',
      Fecha_Hasta: '2025-07-31',
      Aportante: 'Mush',
      Porcentaje: 20,
      Nota: 'Historico Caja'
    }
  ]);

  validarReglasOrigenesFondosQTAS_(reglas);

  const regla = reglas[0] || {};
  if (
    reglas.length !== 1 ||
    regla.reglaId !== 'CAJA-OLD' ||
    numero_(regla.steve) !== 40 ||
    numero_(regla.majo) !== 40 ||
    numero_(regla.mush) !== 20 ||
    (regla.reglaIds || []).length !== 3
  ) {
    throw new Error('No se pudo agrupar la regla historica de Caja 40/40/20.');
  }

  return testSerializarValorQTAS_({
    ok: true,
    regla: regla
  });
}

function habilitarOperacionesDestructivasQAQTAS() {
  PropertiesService.getScriptProperties().setProperty('QTAS_ALLOW_DESTRUCTIVE', 'true');
  return estadoOperacionesDestructivasQAQTAS();
}

function bloquearOperacionesDestructivasQAQTAS() {
  PropertiesService.getScriptProperties().deleteProperty('QTAS_ALLOW_DESTRUCTIVE');
  return estadoOperacionesDestructivasQAQTAS();
}

function estadoOperacionesDestructivasQAQTAS() {
  const ss = SpreadsheetApp.getActive();
  return testSerializarValorQTAS_({
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    allowDestructive: operacionesDestructivasPermitidasQTAS_()
  });
}

function testResetEntornoQTAS(payload) {
  assertOperacionDestructivaPermitidaQTAS_('testResetEntornoQTAS');

  const settings = Object.assign({
    aplicarFormatos: false,
    includeSnapshot: true,
    asegurarModelo: true
  }, payload || {});

  return withScriptLock_('test reset entorno', () => {
    establecerSincronizacionInventarioQTAS_(true);
    const ss = settings.asegurarModelo === false
      ? SpreadsheetApp.getActive()
      : asegurarModeloOperativoQTAS_({
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
    invalidarCacheDocumentoQTAS_('origenes_fondos_reglas_memoria');
    invalidarCacheDocumentoQTAS_('origenes_fondos_reglas_memoria_v2');

    sembrarProductosYPrecios_();
    sembrarConfig_();
    sembrarReglasDistribucion_();

    if (settings.aplicarFormatos) {
      aplicarFormatosModeloQTAS_(ss);
    }

    if (settings.includeSnapshot === false) {
      return testSerializarValorQTAS_({
        ok: true,
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        reset: true
      });
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

function testEjecutarLoteQTAS(payload) {
  const settings = payload || {};
  const steps = Array.isArray(settings.steps) ? settings.steps : [];

  return testSerializarValorQTAS_({
    ok: true,
    results: steps.map((step, index) => {
      const functionName = texto_(step && (step.functionName || step.fn));
      const parameters = Array.isArray(step && step.parameters) ? step.parameters : [];
      if (!functionName) {
        throw new Error(`Falta functionName en el paso ${index + 1}.`);
      }

      const target = testResolverFuncionLoteQTAS_(functionName);
      return target.apply(null, parameters);
    })
  });
}

function testHojasSoportadasQTAS_() {
  return [
    QTAS.sheets.productos,
    QTAS.sheets.precios,
    QTAS.sheets.compras,
    QTAS.sheets.compraDetalle,
    QTAS.sheets.costosReferencia,
    QTAS.sheets.inventarioControl,
    QTAS.sheets.producciones,
    QTAS.sheets.produccionDetalle,
    QTAS.sheets.inventarioMovimientos,
    QTAS.sheets.inventarioSnapshot,
    QTAS.sheets.productoComponentes,
    QTAS.sheets.productoReglasCosto,
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

function testResolverFuncionLoteQTAS_(functionName) {
  let target = null;

  try {
    if (typeof globalThis !== 'undefined' && globalThis) {
      target = globalThis[functionName];
    }
  } catch (error) {
    target = null;
  }

  if (typeof target !== 'function') {
    try {
      target = eval(functionName);
    } catch (error) {
      target = null;
    }
  }

  if (typeof target !== 'function') {
    throw new Error(`No existe la funcion ${functionName} para lote de testing.`);
  }

  return target;
}
