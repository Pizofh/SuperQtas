function exportarLibroParaCodexQTAS() {
  return exportarLibroTextoPlanoQTAS({
    formato: 'tsv',
    comprimir: true,
    incluirOcultas: true,
    soloCanonicas: false
  });
}

function exportarLibroParaCodexQTAS_Log() {
  const result = exportarLibroParaCodexQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function exportarLibroCanonicoParaCodexQTAS() {
  return exportarLibroTextoPlanoQTAS({
    formato: 'tsv',
    comprimir: true,
    incluirOcultas: false,
    soloCanonicas: true
  });
}

function exportarLibroCanonicoParaCodexQTAS_Log() {
  const result = exportarLibroCanonicoParaCodexQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function exportarLibroTextoPlanoQTAS(payload) {
  return withScriptLock_('exportar libro texto plano', () =>
    exportarLibroTextoPlanoQTAS_(payload)
  );
}

function exportarLibroTextoPlanoQTAS_(payload) {
  const settings = Object.assign({
    formato: 'tsv',
    comprimir: true,
    incluirOcultas: true,
    soloCanonicas: false
  }, payload || {});
  const formato = normalizarFormatoExportQTAS_(settings.formato);
  const ss = SpreadsheetApp.getActive();
  const sourceFile = DriveApp.getFileById(ss.getId());
  const rootFolder = asegurarCarpetaExportsQTAS_(sourceFile);
  const timestamp = Utilities.formatDate(new Date(), zonaHorariaQTAS_(), 'yyyyMMdd_HHmmss');
  const exportFolder = rootFolder.createFolder(
    construirNombreExportLibroQTAS_(ss.getName(), timestamp)
  );
  const sheets = listarHojasExportQTAS_(ss, settings);
  const extension = formato === 'md' ? 'md' : formato;
  const blobs = [];
  const resumen = [];

  sheets.forEach((sheet, index) => {
    const matrix = leerMatrizDisplayHojaQTAS_(sheet);
    const content = serializarMatrizExportQTAS_(matrix, formato);
    const fileName = construirNombreArchivoHojaExportQTAS_(sheet.getName(), index + 1, extension);
    const blob = Utilities.newBlob(content, MimeType.PLAIN_TEXT, fileName);

    exportFolder.createFile(blob);
    blobs.push(blob);
    resumen.push({
      orden: index + 1,
      hoja: sheet.getName(),
      filas: matrix.length,
      columnas: matrix.length ? matrix[0].length : 0,
      oculta: hojaOcultaQTAS_(sheet)
    });
  });

  const summaryName = '00_RESUMEN_EXPORT.md';
  const summaryContent = construirResumenExportQTAS_(ss, settings, resumen, extension);
  const summaryBlob = Utilities.newBlob(summaryContent, MimeType.PLAIN_TEXT, summaryName);
  exportFolder.createFile(summaryBlob);
  blobs.unshift(summaryBlob);

  let zipFile = null;
  if (settings.comprimir !== false) {
    const zipName = `${construirNombreExportLibroQTAS_(ss.getName(), timestamp)}.zip`;
    const zipBlob = Utilities.zip(blobs, zipName);
    zipFile = exportFolder.createFile(zipBlob);
  }

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    formato: formato,
    extension: extension,
    hojasExportadas: resumen.length,
    folderId: exportFolder.getId(),
    folderName: exportFolder.getName(),
    folderUrl: exportFolder.getUrl(),
    zipFileId: zipFile ? zipFile.getId() : '',
    zipFileName: zipFile ? zipFile.getName() : '',
    zipFileUrl: zipFile ? zipFile.getUrl() : '',
    resumen: resumen
  };
}

function normalizarFormatoExportQTAS_(value) {
  const key = normalizarClaveTexto_(value);
  if (['csv', 'md', 'markdown', 'tsv'].includes(key)) {
    return key === 'markdown' ? 'md' : key;
  }
  return 'tsv';
}

function asegurarCarpetaExportsQTAS_(sourceFile) {
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
  const parentFolders = file.getParents();
  const parent = parentFolders.hasNext()
    ? parentFolders.next()
    : DriveApp.getRootFolder();
  const folderName = 'QTAS_Exports';
  const existingFolders = parent.getFoldersByName(folderName);
  const folder = existingFolders.hasNext()
    ? existingFolders.next()
    : parent.createFolder(folderName);

  props.setProperty('QTAS_EXPORT_FOLDER_ID', folder.getId());
  return folder;
}

function listarHojasExportQTAS_(ss, settings) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const soloCanonicas = Boolean(settings && settings.soloCanonicas);
  const incluirOcultas = settings && settings.incluirOcultas !== false;
  const canonicalSheetNames = obtenerNombresHojasCanonicasQTAS_();

  return spreadsheet.getSheets()
    .filter(sheet => incluirOcultas || !hojaOcultaQTAS_(sheet))
    .filter(sheet => !soloCanonicas || canonicalSheetNames.indexOf(sheet.getName()) >= 0);
}

function obtenerNombresHojasCanonicasQTAS_() {
  const canonical = [
    QTAS.sheets.productos,
    QTAS.sheets.precios,
    QTAS.sheets.compras,
    QTAS.sheets.compraDetalle,
    QTAS.sheets.costosReferencia,
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
  return [...new Set(canonical.map(texto_).filter(Boolean))];
}

function hojaOcultaQTAS_(sheet) {
  if (!sheet) return false;
  try {
    return sheet.isSheetHidden();
  } catch (error) {
    return false;
  }
}

function construirNombreExportLibroQTAS_(spreadsheetName, timestamp) {
  const base = [
    'QTAS_EXPORT',
    texto_(spreadsheetName),
    texto_(timestamp)
  ].join('__');
  return sanitizarNombreArchivoExportQTAS_(base);
}

function construirNombreArchivoHojaExportQTAS_(sheetName, index, extension) {
  const prefix = String(Math.max(1, numero_(index))).padStart(2, '0');
  const name = `${prefix}_${texto_(sheetName)}`;
  return `${sanitizarNombreArchivoExportQTAS_(name)}.${texto_(extension) || 'tsv'}`;
}

function sanitizarNombreArchivoExportQTAS_(value) {
  return texto_(value)
    .replace(/[\\\/:*?"<>|#%&{}$!'@+=`~]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 120) || 'archivo';
}

function leerMatrizDisplayHojaQTAS_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
}

function serializarMatrizExportQTAS_(matrix, formato) {
  const values = matrix || [];
  if (!values.length) return '';

  if (formato === 'csv') {
    return values.map(row =>
      row.map(value => serializarCeldaCsvQTAS_(value)).join(',')
    ).join('\n');
  }

  if (formato === 'md') {
    const normalized = values.map(row => row.map(value => sanitizarCeldaTextoPlanoQTAS_(value)));
    const headers = normalized[0] || [];
    const separator = headers.map(() => '---');
    const body = normalized.slice(1);
    const rows = [
      `| ${headers.join(' | ')} |`,
      `| ${separator.join(' | ')} |`
    ].concat(body.map(row => `| ${row.join(' | ')} |`));
    return rows.join('\n');
  }

  return values.map(row =>
    row.map(value => sanitizarCeldaTextoPlanoQTAS_(value)).join('\t')
  ).join('\n');
}

function sanitizarCeldaTextoPlanoQTAS_(value) {
  return texto_(value)
    .replace(/\r\n/g, '\\n')
    .replace(/[\r\n]/g, '\\n')
    .replace(/\t/g, '    ')
    .replace(/\|/g, '\\|');
}

function serializarCeldaCsvQTAS_(value) {
  const text = sanitizarCeldaTextoPlanoQTAS_(value).replace(/"/g, '""');
  return `"${text}"`;
}

function construirResumenExportQTAS_(ss, settings, resumen, extension) {
  const lines = [
    '# Export QTAS para Codex',
    '',
    `- Libro: ${texto_(ss && ss.getName())}`,
    `- Spreadsheet ID: ${texto_(ss && ss.getId())}`,
    `- Fecha export: ${Utilities.formatDate(new Date(), zonaHorariaQTAS_(), 'yyyy-MM-dd HH:mm:ss')}`,
    `- Formato recomendado: ${texto_(settings && settings.formato) || 'tsv'}`,
    `- Extension archivos: .${texto_(extension) || 'tsv'}`,
    `- Solo canonicas: ${Boolean(settings && settings.soloCanonicas)}`,
    `- Incluye ocultas: ${Boolean(!settings || settings.incluirOcultas !== false)}`,
    '',
    '## Hojas',
    '',
    '| # | Hoja | Filas | Columnas | Oculta |',
    '| --- | --- | ---: | ---: | --- |'
  ];

  (resumen || []).forEach(item => {
    lines.push(
      `| ${numero_(item.orden)} | ${texto_(item.hoja)} | ${numero_(item.filas)} | ` +
      `${numero_(item.columnas)} | ${item.oculta ? 'Si' : 'No'} |`
    );
  });

  lines.push('');
  lines.push('## Nota');
  lines.push('');
  lines.push(
    '- Para pasar esto a Codex/ChatGPT, lo ideal es usar los `.tsv` y este resumen `.md`.'
  );
  lines.push(
    '- `tsv` suele ser mas comodo que `csv` porque evita choques con comas dentro de las celdas.'
  );

  return lines.join('\n');
}
