// Ejecutable desde el editor de QA cuando la credencial del harness no esta disponible.
// Solo crea una pestana temporal y una copia manual del libro de QA.
function testSeguridadHistoricosQTAS() {
  const ss = SpreadsheetApp.getActive();
  if (!ss || !/QA/i.test(ss.getName())) {
    throw new Error('Esta prueba solo se ejecuta en un libro de QA.');
  }
  const name = '__QTAS_SEGURIDAD_' + Utilities.getUuid().slice(0, 8);
  let sheet;
  let backupId = '';
  try {
    sheet = ss.insertSheet(name);
    const firstDate = new Date('2025-08-01T05:00:00.123Z');
    const secondDate = new Date('2026-07-01T05:00:00.456Z');
    const grid = [
      ['Venta_ID', 'Fecha_Base', 'Medio_Pago', 'Formula'],
      [1, firstDate, 'NequiMajo', '=40+2'],
      [9, new Date('2026-10-04T05:00:00Z'), 'NequiSteve', '=2+2'],
      [2, secondDate, 'Daviplata', '=30+3'],
      [9, new Date('2026-10-04T05:00:00Z'), 'NequiSteve', '=3+3']
    ];
    sheet.getRange(1, 1, grid.length, grid[0].length).setValues(grid);
    sheet.getRange(2, 2, 4, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(1, 1, grid.length, grid[0].length).createFilter()
      .setColumnFilterCriteria(3, SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(['NequiMajo', 'Daviplata']).build());
    SpreadsheetApp.flush();
    if (!sheet.isRowHiddenByFilter(2) || !sheet.isRowHiddenByFilter(4)) {
      throw new Error('La prueba no pudo reproducir el filtro de las filas historicas.');
    }
    const rows = leerObjetosConMeta_(sheet).filter(row => row.Venta_ID === 9);
    const plan = planificarEliminacionFilasQTAS_(sheet, grid[0], rows);
    if (ejecutarEliminacionFilasQTAS_(plan) !== 2) {
      throw new Error('La eliminacion selectiva no retiro exactamente las dos filas elegidas.');
    }
    SpreadsheetApp.flush();
    const values = sheet.getRange(2, 1, 2, 4).getValues();
    const formulas = sheet.getRange(2, 4, 2, 1).getFormulas();
    const formats = sheet.getRange(2, 2, 2, 1).getNumberFormats();
    if (sheet.getLastRow() !== 3 || values[0][0] !== 1 || values[1][0] !== 2 ||
        values[0][1].getTime() !== firstDate.getTime() ||
        values[1][1].getTime() !== secondDate.getTime() ||
        formulas[0][0] !== '=40+2' || formulas[1][0] !== '=30+3' ||
        formats.some(row => row[0] !== 'yyyy-mm-dd hh:mm:ss')) {
      throw new Error('Las filas historicas no conservaron fechas, formulas y formatos.');
    }
    ss.deleteSheet(sheet);
    sheet = null;
    const backup = crearBackupManualQTAS();
    backupId = backup.backupFileId;
    const metadata = JSON.parse(DriveApp.getFileById(backupId).getDescription());
    if (!backup.ok || !backup.created || !metadata.verified || metadata.kind !== 'manual') {
      throw new Error('La copia nativa de QA no quedo verificada.');
    }
    const result = { ok: true, filteredDeletion: true, preservedHistoricalDates: true,
      preservedFormulasAndFormats: true, verifiedNativeBackup: true,
      spreadsheetId: ss.getId(), temporarySheetRemoved: true, testBackupTrashed: true };
    console.log(JSON.stringify(result));
    return result;
  } finally {
    if (sheet) ss.deleteSheet(sheet);
    if (backupId) DriveApp.getFileById(backupId).setTrashed(true);
  }
}

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
    PropertiesService.getScriptProperties().deleteProperty('QTAS_POST_VENTA_QUEUE_V1');
    limpiarCachesEjecucionQTAS_();
    invalidarCacheDocumentoQTAS_('precios_referencia_memoria');
    invalidarCacheDocumentoQTAS_('distribucion_reglas_memoria');
    invalidarCacheDocumentoQTAS_('origenes_fondos_reglas_memoria');
    invalidarCacheDocumentoQTAS_('origenes_fondos_reglas_memoria_v2');
    invalidarCacheDocumentoQTAS_('inventario_dashboard_v1');

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
