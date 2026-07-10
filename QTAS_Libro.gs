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
      QTAS.sheets.productos,
      QTAS.sheets.precios,
      QTAS.sheets.clientes,
      QTAS.sheets.ventas,
      QTAS.sheets.compras,
      QTAS.sheets.pagos,
      QTAS.sheets.ventasEnvio,
      QTAS.sheets.costosReferencia,
      QTAS.sheets.costoProductoCalculado
    ],
    analysisVisible: [
      QTAS.sheets.detalle,
      QTAS.sheets.compraDetalle,
      QTAS.sheets.ventaDetalleCostosCalculado,
      QTAS.sheets.distribucionIngresos,
      QTAS.sheets.compraOrigenesFondos
    ],
    backendHidden: [
      QTAS.sheets.config,
      QTAS.sheets.productoComponentes,
      QTAS.sheets.productoReglasCosto,
      QTAS.sheets.distribucionReglas,
      QTAS.sheets.origenesFondosReglas
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
