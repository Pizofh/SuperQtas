function aplicarVistaOperadorQTAS() {
  const ss = SpreadsheetApp.getActive();
  const layout = layoutLibroQTAS_();
  const activeSheet = ss.getActiveSheet();
  const activeName = activeSheet ? activeSheet.getName() : '';
  const activeIsBackend = layout.backendHidden.indexOf(activeName) >= 0;
  const fallbackSheet = ss.getSheetByName(layout.operatorVisible[0]) || ss.getSheets()[0];

  if (activeIsBackend && fallbackSheet) {
    ss.setActiveSheet(fallbackSheet);
  }

  let visibles = 0;
  let ocultas = 0;

  ss.getSheets().forEach(sheet => {
    const category = categoriaHojaLibroQTAS_(sheet.getName(), layout);
    aplicarColorPestanaLibroQTAS_(sheet, category, layout.colors);

    if (category === 'backend') {
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
      ocultas++;
      return;
    }

    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
    visibles++;
  });

  const result = {
    ok: true,
    visibles: visibles,
    ocultas: ocultas,
    operatorVisible: layout.operatorVisible.slice(),
    analysisVisible: layout.analysisVisible.slice(),
    backendHidden: layout.backendHidden.slice()
  };

  maybeAlert_(
    `Vista operador aplicada. Visibles=${visibles}, ocultas=${ocultas}. ` +
    'Las hojas operativas quedaron resaltadas y el backend quedo oculto.'
  );
  return result;
}

function mostrarTodasLasHojasQTAS() {
  const ss = SpreadsheetApp.getActive();
  const layout = layoutLibroQTAS_();

  ss.getSheets().forEach(sheet => {
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
    aplicarColorPestanaLibroQTAS_(
      sheet,
      categoriaHojaLibroQTAS_(sheet.getName(), layout),
      layout.colors
    );
  });

  const result = {
    ok: true,
    totalSheets: ss.getSheets().length
  };

  maybeAlert_('Todas las hojas quedaron visibles.');
  return result;
}

function prepararLibroPsyloScibioQTAS(payload) {
  const settings = Object.assign({
    borrarHojasLegacy: false,
    dryRunHojasLegacy: true,
    forceSheetNames: [],
    fullHistorico: false
  }, payload || {});

  const catalogoOperativo = alinearCatalogoOperativoPsyloScibioQTAS_();
  const packaging = ajustarPackagingPsyloScibioQTAS();
  const inventarioCanonico = canonizarInventarioPsyloScibioQTAS({ silent: true });
  const costoProducto = reconstruirCostoProductoCalculadoQTAS({ silent: true });
  const ventaDetalleCostos = settings.fullHistorico === true
    ? reconstruirVentaDetalleCostosCalculadoQTAS({ silent: true })
    : {
      ok: true,
      skipped: true,
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0,
      reason: 'Se omitio la reconstruccion historica completa para reducir riesgo de timeout.'
    };
  const hojasLegacy = depurarHojasNoOficialesQTAS({
    dryRun: settings.borrarHojasLegacy === true ? false : settings.dryRunHojasLegacy !== false,
    forceSheetNames: settings.forceSheetNames
  });

  return {
    ok: true,
    catalogoOperativo: catalogoOperativo,
    packaging: packaging,
    inventarioCanonico: inventarioCanonico,
    costoProducto: costoProducto,
    ventaDetalleCostos: ventaDetalleCostos,
    hojasLegacy: hojasLegacy,
    assumptions: unirUnicos_(
      []
        .concat(catalogoOperativo && catalogoOperativo.assumptions || [])
        .concat(packaging && packaging.assumptions || [])
        .concat(inventarioCanonico && inventarioCanonico.caveats || [])
        .concat([
          settings.fullHistorico === true
            ? 'Se recalculo el historico completo de Venta_Detalle_Costos_Calc.'
            : 'La reconstruccion completa de Venta_Detalle_Costos_Calc queda opcional para evitar ejecuciones demasiado largas.',
          'Las hojas oficiales del ERP se conservan aunque hoy tengan una sola fila.',
          'Las hojas no oficiales con datos y nombres no reconocidos se dejan en revision manual por seguridad.'
        ])
    )
  };
}

function prepararLibroPsyloScibioQTAS_Log() {
  const result = prepararLibroPsyloScibioQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function prepararLibroPsyloScibioConBorradoQTAS_Log() {
  const result = prepararLibroPsyloScibioQTAS({
    borrarHojasLegacy: true,
    dryRunHojasLegacy: false
  });
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function prepararLibroPsyloScibioFullHistoricoQTAS_Log() {
  const result = prepararLibroPsyloScibioQTAS({
    fullHistorico: true
  });
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function alinearCatalogoOperativoPsyloScibioQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const productosPlan = {
    vino: {
      activo: false,
      nota: 'Producto legacy pausado; fuera del catalogo operativo actual.'
    },
    chocordy: {
      activo: false,
      nota: 'Producto pausado; fuera del catalogo operativo actual.'
    },
    shiiext: {
      activo: true
    },
    tin: {
      activo: true
    }
  };
  const componentesPlan = {
    chocordy: false,
    shiiext: true
  };
  const reglasPlan = {
    chocordy: false,
    shiiext: true
  };

  const productosSheet = ss.getSheetByName(QTAS.sheets.productos);
  const productosHeaders = getHeaders_(productosSheet);
  const productosRows = leerObjetosConMeta_(productosSheet);
  const productosResult = alinearProductosOperativosPsyloScibioQTAS_(
    productosSheet,
    productosHeaders,
    productosRows,
    productosPlan
  );

  const componentesSheet = ss.getSheetByName(QTAS.sheets.productoComponentes);
  const componentesHeaders = getHeaders_(componentesSheet);
  const componentesRows = leerObjetosConMeta_(componentesSheet);
  const componentesResult = alinearActivosPorProductoPsyloScibioQTAS_(
    componentesSheet,
    componentesHeaders,
    componentesRows,
    componentesPlan
  );

  const reglasSheet = ss.getSheetByName(QTAS.sheets.productoReglasCosto);
  const reglasHeaders = getHeaders_(reglasSheet);
  const reglasRows = leerObjetosConMeta_(reglasSheet);
  const reglasResult = alinearActivosPorProductoPsyloScibioQTAS_(
    reglasSheet,
    reglasHeaders,
    reglasRows,
    reglasPlan
  );

  return {
    ok: true,
    productosActualizados: productosResult.updated,
    componentesActualizados: componentesResult.updated,
    reglasActualizadas: reglasResult.updated,
    productosInactivos: ['Vino', 'Chocordy'],
    productosActivos: ['ShiiExt', 'Tin'],
    assumptions: [
      'Vino y Chocordy quedan fuera del catalogo operativo actual, pero su historico se conserva.',
      'ShiiExt queda activo y controlado como producto fabricado.',
      'Tin se conserva activo con su receta/costo actual aunque siga siendo un producto poco frecuente.'
    ]
  };
}

function alinearProductosOperativosPsyloScibioQTAS_(sheet, headers, rows, plan) {
  const configuracion = plan || {};
  const vistos = {};
  let updated = 0;

  (rows || []).forEach(row => {
    const key = normalizarClaveTexto_(row.Producto_Estandar);
    const regla = key ? configuracion[key] : null;
    if (!regla) return;

    vistos[key] = true;
    const activoDeseado = regla.activo !== false;
    const notaActual = texto_(row.Nota);
    const notaDeseada = regla.nota
      ? (notaActual || texto_(regla.nota))
      : notaActual;
    const activoActual = estaActivo_(row.Activo);

    if (activoActual === activoDeseado && notaActual === notaDeseada) {
      return;
    }

    actualizarFilaObjeto_(sheet, row.__rowNumber, headers, Object.assign({}, row, {
      Activo: activoDeseado,
      Nota: notaDeseada
    }));
    updated++;
  });

  return {
    updated: updated,
    missing: Object.keys(configuracion)
      .filter(key => !vistos[key])
      .sort((a, b) => a.localeCompare(b))
  };
}

function alinearActivosPorProductoPsyloScibioQTAS_(sheet, headers, rows, plan) {
  const configuracion = plan || {};
  let updated = 0;

  (rows || []).forEach(row => {
    const key = normalizarClaveTexto_(row.Producto_Estandar);
    if (!key || configuracion[key] === undefined) return;

    const activoDeseado = configuracion[key] === true;
    const activoActual = estaActivo_(row.Activo);
    if (activoActual === activoDeseado) return;

    actualizarFilaObjeto_(sheet, row.__rowNumber, headers, Object.assign({}, row, {
      Activo: activoDeseado
    }));
    updated++;
  });

  return {
    updated: updated
  };
}

function depurarHojasNoOficialesQTAS(payload) {
  const settings = Object.assign({
    dryRun: true,
    forceSheetNames: []
  }, payload || {});
  if (settings.dryRun !== true) {
    assertOperacionDestructivaPermitidaQTAS_('depurar hojas no oficiales');
  }

  const ss = SpreadsheetApp.getActive();
  const officialMap = nombresHojasOficialesQTAS_();
  const forceMap = {};
  const autoCandidates = [];
  const manualReview = [];
  const deleted = [];

  (settings.forceSheetNames || []).forEach(name => {
    const key = normalizarClaveTexto_(name);
    if (key) forceMap[key] = true;
  });

  ss.getSheets().forEach(sheet => {
    const audit = auditarHojaNoOficialQTAS_(sheet, officialMap, forceMap);
    if (audit.status === 'official') return;
    if (audit.status === 'auto_candidate') {
      autoCandidates.push(audit);
      return;
    }
    manualReview.push(audit);
  });

  if (settings.dryRun !== true) {
    autoCandidates.forEach(item => {
      const sheet = ss.getSheetByName(item.name);
      if (!sheet) return;
      ss.deleteSheet(sheet);
      deleted.push(item.name);
    });
  }

  return {
    ok: true,
    dryRun: settings.dryRun !== false,
    officialSheets: Object.keys(officialMap).sort((a, b) => a.localeCompare(b)),
    autoCandidates: autoCandidates.map(item => ({
      name: item.name,
      rows: item.rows,
      columns: item.columns,
      hidden: item.hidden,
      reasons: item.reasons
    })),
    manualReview: manualReview.map(item => ({
      name: item.name,
      rows: item.rows,
      columns: item.columns,
      hidden: item.hidden,
      reasons: item.reasons
    })),
    deleted: deleted
  };
}

function depurarHojasNoOficialesQTAS_Log() {
  const result = depurarHojasNoOficialesQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function depurarHojasNoOficialesQTAS_Borrar_Log() {
  const result = depurarHojasNoOficialesQTAS({ dryRun: false });
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function exportarLibroTSVQTAS() {
  const ss = SpreadsheetApp.getActive();
  const sourceFile = DriveApp.getFileById(ss.getId());
  const rootFolder = asegurarCarpetaExportQTAS_(sourceFile);
  const layout = layoutLibroQTAS_();
  const timezone = zonaHorariaQTAS_();
  const stampFile = Utilities.formatDate(new Date(), timezone, 'yyyyMMdd_HHmmss');
  const stampText = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');
  const exportFolder = rootFolder.createFolder(`${sanitizarNombreArchivoQTAS_(ss.getName())}__TSV__${stampFile}`);
  const sheets = ss.getSheets();
  const summaryRows = [];

  sheets.forEach((sheet, index) => {
    const values = sheet.getDataRange().getDisplayValues();
    const tsv = values.map(row =>
      row.map(textoPlanoTSVQTAS_).join('\t')
    ).join('\n');
    const fileName = [
      String(index + 1).padStart(2, '0'),
      sanitizarNombreArchivoQTAS_(sheet.getName())
    ].join('_') + '.tsv';

    exportFolder.createFile(fileName, tsv, MimeType.PLAIN_TEXT);
    summaryRows.push({
      index: index + 1,
      name: sheet.getName(),
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden(),
      category: categoriaHojaLibroQTAS_(sheet.getName(), layout)
    });
  });

  exportFolder.createFile(
    '00_RESUMEN_EXPORT.md',
    construirResumenExportLibroQTAS_(ss, summaryRows, exportFolder, stampText),
    MimeType.PLAIN_TEXT
  );

  const result = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    folderId: exportFolder.getId(),
    folderName: exportFolder.getName(),
    folderUrl: exportFolder.getUrl(),
    filesCreated: summaryRows.length + 1
  };

  maybeAlert_(
    `Export listo. Carpeta: ${exportFolder.getName()}. ` +
    `Archivos: ${result.filesCreated}. URL: ${exportFolder.getUrl()}`
  );
  return result;
}

function layoutLibroQTAS_() {
  return {
    operatorVisible: [  
      QTAS.sheets.ventas,
      QTAS.sheets.compras,
      QTAS.sheets.pagos,
      QTAS.sheets.ventasEnvio,
      QTAS.sheets.producciones,
      QTAS.sheets.inventarioSnapshot
    ],
    analysisVisible: [
      QTAS.sheets.detalle,
      QTAS.sheets.compraDetalle,
      QTAS.sheets.produccionDetalle,
      QTAS.sheets.inventarioMovimientos,
      QTAS.sheets.ventaDetalleCostosCalculado,
      QTAS.sheets.distribucionIngresos,
      QTAS.sheets.compraOrigenesFondos
    ],
    backendHidden: [
      QTAS.sheets.clientes,
      QTAS.sheets.precios,
      QTAS.sheets.productos,
      QTAS.sheets.config,
      QTAS.sheets.productoComponentes,
      QTAS.sheets.productoReglasCosto,
      QTAS.sheets.distribucionReglas,
      QTAS.sheets.origenesFondosReglas,
      QTAS.sheets.inventarioControl,
      QTAS.sheets.costosReferencia,
      QTAS.sheets.costoProductoCalculado
    ],
    colors: {
      operator: '#11b319',
      analysis: '#d90606',
      backend: '#ffffff',
      other: '#8B5E3C'
    }
  };
}

function categoriaHojaLibroQTAS_(sheetName, layout) {
  const rules = layout || layoutLibroQTAS_();
  const name = texto_(sheetName);

  if (rules.operatorVisible.indexOf(name) >= 0) return 'operator';
  if (rules.analysisVisible.indexOf(name) >= 0) return 'analysis';
  if (rules.backendHidden.indexOf(name) >= 0) return 'backend';
  return 'other';
}

function aplicarColorPestanaLibroQTAS_(sheet, category, colors) {
  if (!sheet) return;

  const palette = colors || layoutLibroQTAS_().colors;
  const color = palette[category] || palette.other;
  sheet.setTabColor(color);
}

function asegurarCarpetaExportQTAS_(sourceFile) {
  const props = PropertiesService.getScriptProperties();
  const configuredId = texto_(props.getProperty('QTAS_EXPORT_FOLDER_ID'));

  if (configuredId) {
    try {
      return DriveApp.getFolderById(configuredId);
    } catch (error) {
      props.deleteProperty('QTAS_EXPORT_FOLDER_ID');
    }
  }

  const file = sourceFile || DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folderName = 'QTAS_Exports';
  const existing = parent.getFoldersByName(folderName);
  const folder = existing.hasNext()
    ? existing.next()
    : parent.createFolder(folderName);

  props.setProperty('QTAS_EXPORT_FOLDER_ID', folder.getId());
  return folder;
}

function construirResumenExportLibroQTAS_(ss, rows, folder, stampText) {
  const sheetRows = rows || [];
  const lines = [
    '# Export QTAS para revision',
    '',
    `- Libro: ${ss.getName()}`,
    `- Spreadsheet ID: ${ss.getId()}`,
    `- Fecha export: ${stampText}`,
    '- Formato recomendado: tsv',
    '- Extension archivos: .tsv',
    `- Carpeta export: ${folder.getName()}`,
    `- URL carpeta: ${folder.getUrl()}`,
    '',
    '## Hojas',
    '',
    '| # | Hoja | Categoria | Filas | Columnas | Oculta |',
    '| --- | --- | --- | ---: | ---: | --- |'
  ];

  sheetRows.forEach(row => {
    lines.push(
      `| ${row.index} | ${row.name} | ${etiquetaCategoriaLibroQTAS_(row.category)} | ` +
      `${row.rows} | ${row.columns} | ${row.hidden ? 'Si' : 'No'} |`
    );
  });

  lines.push('');
  lines.push('## Nota');
  lines.push('');
  lines.push('- `operator`: hojas pensadas para consulta operativa.');
  lines.push('- `analysis`: hojas utiles para BI, conciliacion y analitica.');
  lines.push('- `backend`: hojas ocultables para reducir ruido del libro.');

  return lines.join('\n');
}

function nombresHojasOficialesQTAS_() {
  return Object.keys(QTAS.schemas).reduce((acc, sheetName) => {
    const name = texto_(sheetName);
    if (name) acc[name] = true;
    return acc;
  }, {});
}

function auditarHojaNoOficialQTAS_(sheet, officialMap, forceMap) {
  const name = sheet ? sheet.getName() : '';
  const normalized = normalizarClaveTexto_(name);
  const rows = sheet ? sheet.getLastRow() : 0;
  const columns = sheet ? sheet.getLastColumn() : 0;
  const hidden = sheet ? sheet.isSheetHidden() : false;
  const reasons = [];

  if (!sheet || !name) {
    return {
      status: 'manual_review',
      name: name,
      rows: rows,
      columns: columns,
      hidden: hidden,
      reasons: ['Hoja invalida o sin nombre.']
    };
  }

  if (officialMap && officialMap[name]) {
    return {
      status: 'official',
      name: name,
      rows: rows,
      columns: columns,
      hidden: hidden,
      reasons: ['Hoja oficial del ERP.']
    };
  }

  if (forceMap && forceMap[normalized]) {
    reasons.push('Incluida manualmente en forceSheetNames.');
    return {
      status: 'auto_candidate',
      name: name,
      rows: rows,
      columns: columns,
      hidden: hidden,
      reasons: reasons
    };
  }

  if (esNombreHojaLegacyQTAS_(name)) {
    reasons.push('Nombre compatible con copia, hoja temporal o legacy.');
  }
  if (rows <= 1) {
    reasons.push('Sin datos operativos; solo encabezado o vacia.');
  }
  if (columns <= 1) {
    reasons.push('Estructura minima no operativa.');
  }

  if (reasons.length && (esNombreHojaLegacyQTAS_(name) || rows <= 1 || columns <= 1)) {
    return {
      status: 'auto_candidate',
      name: name,
      rows: rows,
      columns: columns,
      hidden: hidden,
      reasons: reasons
    };
  }

  return {
    status: 'manual_review',
    name: name,
    rows: rows,
    columns: columns,
    hidden: hidden,
    reasons: ['No es oficial, pero no cumple criterios seguros de borrado automatico.']
  };
}

function esNombreHojaLegacyQTAS_(name) {
  const value = texto_(name);
  return [
    /^copia de /i,
    /^copy of /i,
    /^hoja\s*\d+$/i,
    /^sheet\s*\d+$/i,
    /^tmp[_\s-]/i,
    /^temp[_\s-]/i,
    /^backup[_\s-]/i,
    /^old[_\s-]/i,
    /^legacy[_\s-]/i,
    /_old$/i,
    /_legacy$/i,
    /_backup$/i
  ].some(pattern => pattern.test(value));
}

function etiquetaCategoriaLibroQTAS_(category) {
  if (category === 'operator') return 'operator';
  if (category === 'analysis') return 'analysis';
  if (category === 'backend') return 'backend';
  return 'other';
}

function sanitizarNombreArchivoQTAS_(value) {
  return texto_(value)
    .replace(/[\\/:*?"<>|#%&{}$!'@+=`~]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80) || 'sin_nombre';
}

function textoPlanoTSVQTAS_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/\r?\n/g, ' ')
    .replace(/\t/g, ' ')
    .trim();
}
