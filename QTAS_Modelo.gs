function asegurarModeloOperativoQTAS_(options) {
  const config = Object.assign({
    aplicarFormatos: false
  }, options || {});
  const ss = SpreadsheetApp.getActive();

  Object.keys(QTAS.schemas).forEach(sheetName => {
    if (sheetName === QTAS.sheets.config) return;
    asegurarHojaModelo_(ss, sheetName, QTAS.schemas[sheetName]);
  });

  asegurarConfigMediosPagoQTAS_(ss);
  sembrarProductosYPrecios_();
  sembrarReglasDistribucion_();
  sembrarConfig_();

  if (config.aplicarFormatos) {
    aplicarFormatosModeloQTAS_(ss);
  }

  return ss;
}

function asegurarModeloCompletoQTAS_() {
  return asegurarModeloOperativoQTAS_({ aplicarFormatos: true });
}

function asegurarModeloCompletoQTAS() {
  return asegurarModeloCompletoQTAS_();
}

function diagnosticarModeloQTAS() {
  const ss = SpreadsheetApp.getActive();
  const expectedSheets = Object.keys(QTAS.schemas);
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());
  const missingSheets = [];
  const missingOptionalSheets = [];
  const headerIssues = [];

  expectedSheets.forEach(sheetName => {
    if (sheetName === QTAS.sheets.config) return;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      if (esHojaOpcionalQTAS_(sheetName)) {
        missingOptionalSheets.push(sheetName);
      } else {
        missingSheets.push(sheetName);
      }
      return;
    }

    if (sheetName === QTAS.sheets.precios && esHojaPreciosLegacy_(sheet)) {
      return;
    }

    const actualHeaders = getHeaders_(sheet);
    const expectedHeaders = QTAS.schemas[sheetName];
    if (!headersIguales_(actualHeaders, expectedHeaders)) {
      headerIssues.push({
        sheetName: sheetName,
        actualHeaders: actualHeaders,
        expectedHeaders: expectedHeaders
      });
    }
  });

  const configSheet = obtenerHojaConfigQTAS_(ss, { create: false });
  const configStatus = {
    present: Boolean(configSheet),
    sheetName: configSheet ? configSheet.getName() : '',
    valid: false,
    mode: 'missing'
  };

  if (configSheet) {
    if (esConfigMediosPagoCanonicoQTAS_(configSheet)) {
      configStatus.valid = true;
      configStatus.mode = 'canonical';
    } else if (esConfigLegacyQTAS_(configSheet)) {
      configStatus.valid = true;
      configStatus.mode = 'legacy';
    } else {
      configStatus.mode = 'invalid';
    }
  }

  return {
    ok: missingSheets.length === 0 && headerIssues.length === 0 && configStatus.valid,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    missingSheets: missingSheets,
    missingOptionalSheets: missingOptionalSheets,
    headerIssues: headerIssues,
    config: configStatus,
    extraSheets: existingSheets.filter(sheetName =>
      expectedSheets.indexOf(sheetName) === -1 &&
      sheetName !== QTAS.sheets.configLegacy
    ),
    expectedSheets: expectedSheets,
    existingSheets: existingSheets,
    suggestedAction: missingSheets.length || headerIssues.length || !configStatus.valid
      ? 'Run asegurarModeloCompletoQTAS_() on this spreadsheet before using it.'
      : 'Model looks consistent.'
  };
}

function snapshotEstructuraModeloQTAS() {
  const ss = SpreadsheetApp.getActive();
  const expectedSheets = Object.keys(QTAS.schemas);
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());
  const configSheet = obtenerHojaConfigQTAS_(ss, { create: false });
  const configHeadersEsperados = QTAS.schemas[QTAS.sheets.config];
  const configHeadersActuales = configSheet ? getHeaders_(configSheet) : [];
  let configMode = 'missing';
  let configValida = false;

  if (configSheet) {
    if (esConfigMediosPagoCanonicoQTAS_(configSheet)) {
      configMode = 'canonical';
      configValida = true;
    } else if (esConfigLegacyQTAS_(configSheet)) {
      configMode = 'legacy';
      configValida = true;
    } else {
      configMode = 'invalid';
    }
  }

  const hojasCanonicas = expectedSheets
    .filter(sheetName => sheetName !== QTAS.sheets.config)
    .map(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      const expectedHeaders = QTAS.schemas[sheetName];

      if (!sheet) {
        return {
          logicalName: sheetName,
          sheetName: sheetName,
          exists: false,
          optional: esHojaOpcionalQTAS_(sheetName),
          legacyPriceSheet: false,
          headerOk: false,
          actualHeaders: [],
          expectedHeaders: expectedHeaders,
          lastRow: 0,
          dataRows: 0
        };
      }

      const actualHeaders = getHeaders_(sheet);
      const legacyPriceSheet = sheetName === QTAS.sheets.precios && esHojaPreciosLegacy_(sheet);
      return {
        logicalName: sheetName,
        sheetName: sheet.getName(),
        exists: true,
        optional: esHojaOpcionalQTAS_(sheetName),
        legacyPriceSheet: legacyPriceSheet,
        headerOk: legacyPriceSheet || headersIguales_(actualHeaders, expectedHeaders),
        actualHeaders: actualHeaders,
        expectedHeaders: expectedHeaders,
        lastRow: sheet.getLastRow(),
        dataRows: Math.max(sheet.getLastRow() - 1, 0)
      };
    });

  hojasCanonicas.push({
    logicalName: QTAS.sheets.config,
    sheetName: configSheet ? configSheet.getName() : QTAS.sheets.config,
    exists: Boolean(configSheet),
    optional: false,
    legacyPriceSheet: false,
    headerOk: configValida,
    actualHeaders: configHeadersActuales,
    expectedHeaders: configMode === 'legacy'
      ? ['Tipo', 'Valor', 'Activo']
      : configHeadersEsperados,
    lastRow: configSheet ? configSheet.getLastRow() : 0,
    dataRows: configSheet ? Math.max(configSheet.getLastRow() - 1, 0) : 0,
    mode: configMode
  });

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheets: hojasCanonicas,
    extraSheets: existingSheets.filter(sheetName =>
      expectedSheets.indexOf(sheetName) === -1 &&
      sheetName !== QTAS.sheets.configLegacy
    ),
    allSheets: existingSheets
  };
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

    if (sheetName === QTAS.sheets.precios && esHojaPreciosLegacy_(sheet)) {
      return;
    }

    if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[sheetName])) {
      throw new Error(`La hoja ${sheetName} no coincide con la estructura esperada. Ejecuta "Crear / reparar modelo" manualmente.`);
    }
  });

  if (settings.validarConfig) {
    const configSheet = obtenerHojaConfigQTAS_(ss, { create: false });
    if (!configSheet) {
      throw new Error(`Falta la hoja ${QTAS.sheets.config} o ${QTAS.sheets.configLegacy}. Ejecuta "Crear / reparar modelo" manualmente.`);
    }

    if (!esConfigMediosPagoCanonicoQTAS_(configSheet) && !esConfigLegacyQTAS_(configSheet)) {
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

  if (esConfigMediosPagoCanonicoQTAS_(sheet) || esConfigLegacyQTAS_(sheet)) {
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
