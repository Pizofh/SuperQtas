function habilitarDestructivoTemporalQTAS() {
  PropertiesService.getScriptProperties().setProperty('QTAS_ALLOW_DESTRUCTIVE', 'true');
  Logger.log(PropertiesService.getScriptProperties().getProperty('QTAS_ALLOW_DESTRUCTIVE'));
}

function deshabilitarDestructivoTemporalQTAS() {
  PropertiesService.getScriptProperties().deleteProperty('QTAS_ALLOW_DESTRUCTIVE');
  Logger.log('QTAS_ALLOW_DESTRUCTIVE eliminado');
}
function asegurarModeloOperativoQTAS_(options) {
  const config = Object.assign({
    aplicarFormatos: false
  }, options || {});
  const ss = SpreadsheetApp.getActive();

  Object.keys(QTAS.schemas).forEach(sheetName => {
    if (sheetName === QTAS.sheets.config) return;
    if (sheetName === QTAS.sheets.ventasEnvio) {
      asegurarHojaVentasEnvioSeguimientoQTAS_(ss, { create: true, validate: true });
      return;
    }
    asegurarHojaModelo_(ss, sheetName, QTAS.schemas[sheetName]);
  });

  asegurarConfigMediosPagoQTAS_(ss);
  sembrarProductosYPrecios_();
  sembrarReglasDistribucion_();
  sembrarConfig_();
  asegurarControlesInventarioBaseQTAS_();

  if (config.aplicarFormatos) {
    aplicarFormatosModeloQTAS_(ss);
  }

  return ss;
}

function validarModeloSoloLecturaQTAS_(options) {
  const settings = Array.isArray(options)
    ? { sheetNames: options, validarConfig: true }
    : Object.assign({ validarConfig: true }, options || {});
  const ss = SpreadsheetApp.getActive();
  const targetSheets = (settings.sheetNames && settings.sheetNames.length
    ? settings.sheetNames
    : Object.keys(QTAS.schemas).filter(sheetName => sheetName !== QTAS.sheets.config)
  ).filter((sheetName, index, arr) =>
    sheetName &&
    sheetName !== QTAS.sheets.config &&
    arr.indexOf(sheetName) === index
  );

  targetSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      if (esHojaOpcionalQTAS_(sheetName)) {
        return;
      }
      throw new Error(`Falta la hoja ${sheetName}. Ejecuta "Crear / reparar modelo" manualmente.`);
    }

    if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[sheetName])) {
      throw new Error(`La hoja ${sheetName} no coincide con la estructura esperada. Ejecuta "Crear / reparar modelo" manualmente.`);
    }
  });

  if (settings.validarConfig) {
    const configSheet = obtenerHojaConfigQTAS_(ss, { create: false });
    if (!configSheet) {
      throw new Error(`Falta la hoja ${QTAS.sheets.config}. Ejecuta "Crear / reparar modelo" manualmente.`);
    }

    if (!esConfigMediosPagoCanonicoQTAS_(configSheet)) {
      throw new Error(`La hoja ${configSheet.getName()} no tiene una estructura valida. Ejecuta "Crear / reparar modelo" manualmente.`);
    }
  }

  return ss;
}

function asegurarConfigMediosPagoQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const sheet = obtenerHojaConfigQTAS_(spreadsheet, { create: true });
  const headers = getHeaders_(sheet);
  const expected = QTAS.schemas[QTAS.sheets.config];

  if (!headers.length) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    limpiarCacheHeadersHojaQTAS_(sheet);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, expected.length).setFontWeight('bold');
    return sheet;
  }

  if (esConfigMediosPagoCanonicoQTAS_(sheet)) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    return sheet;
  }

  sincronizarHeaders_(sheet, headers, expected);
  return sheet;
}

function sembrarProductosYPrecios_() {
  const ss = SpreadsheetApp.getActive();
  const productosSheet = ss.getSheetByName(QTAS.sheets.productos);
  const preciosSheet = ss.getSheetByName(QTAS.sheets.precios);
  let escribioPrecios = false;

  if (productosSheet.getLastRow() <= 1) {
    const seen = {};
    const rows = [];

    PRECIOS_INICIALES.forEach(item => {
      if (seen[item.producto]) return;
      seen[item.producto] = true;
      rows.push([
        item.producto,
        item.unidad,
        true,
        item.producto === '500mg' ? 'Producto creado, precio pendiente' : ''
      ]);
    });

    escribirFilas_(productosSheet, rows);
  }

  if (preciosSheet.getLastRow() <= 1) {
    const rows = construirFilasPreciosNormalizados_();
    escribirFilas_(preciosSheet, rows);
    escribioPrecios = true;
  }

  if (escribioPrecios) {
    invalidarCacheDocumentoQTAS_('precios_referencia_memoria');
  }
}

function sembrarConfig_() {
  const sheet = obtenerHojaConfigQTAS_(SpreadsheetApp.getActive(), { create: true });
  if (sheet.getLastRow() > 1) return;

  if (!esConfigMediosPagoCanonicoQTAS_(sheet)) return;

  escribirFilas_(sheet, MEDIOS_PAGO.map(item => [item, true, '']));
}

function applyFormatosConfigMediosPagoQTAS_(sheet) {
  if (!sheet) return;
  sheet.autoResizeColumns(1, Math.max(sheet.getLastColumn(), 3));
}

function sembrarReglasDistribucion_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(QTAS.sheets.distribucionReglas);
  if (sheet.getLastRow() > 1) return;

  const headers = QTAS.schemas[QTAS.sheets.distribucionReglas];
  const rows = DISTRIBUCION_REGLAS_INICIALES.map((regla, index) => filaDesdeHeaders_(headers, {
    Regla_ID: 'DIST-' + String(index + 1).padStart(4, '0'),
    Fecha_Desde: fecha_(regla.desde),
    Fecha_Hasta: regla.hasta ? fecha_(regla.hasta) : '',
    Steve_Pct: regla.steve,
    Majo_Pct: regla.majo,
    Mush_Pct: regla.mush,
    Activo: true,
    Nota: regla.nota
  }));

  escribirFilas_(sheet, rows);
  invalidarCacheDocumentoQTAS_('distribucion_reglas_memoria');
}
