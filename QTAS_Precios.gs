function esHojaPreciosLegacy_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return false;
  const headers = getHeaders_(sheet);
  return texto_(headers[1]).toLowerCase() === 'precio';
}

function leerPreciosConfigurados_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.precios);

  if (!sheet || sheet.getLastRow() < 2) {
    return construirPreciosNormalizadosBase_();
  }

  if (headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.precios])) {
    return leerObjetos_(sheet)
      .map(row => ({
        precioId: texto_(row.Precio_ID),
        producto: texto_(row.Producto_Estandar),
        precio: numero_(row.Precio),
        unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
        desde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
        hasta: row.Fecha_Hasta
          ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date())
          : null,
        activo: estaActivo_(row.Activo),
        nota: texto_(row.Nota)
      }))
      .filter(row => row.producto && row.precio > 0);
  }

  if (esHojaPreciosLegacy_(sheet)) {
    return convertirPreciosLegacyANormalizados_(sheet);
  }

  return construirPreciosNormalizadosBase_();
}

function construirPreciosNormalizadosBase_() {
  return PRECIOS_INICIALES
    .filter(item => numero_(item.precio) > 0)
    .map((item, index) => ({
      precioId: 'PREC-' + String(index + 1).padStart(4, '0'),
      producto: item.producto,
      precio: numero_(item.precio),
      unidad: item.unidad,
      desde: resolverFechaOperacion_(item.desde, new Date()),
      hasta: item.hasta ? resolverFechaOperacion_(item.hasta, item.desde) : null,
      activo: true,
      nota: 'Precio base'
    }));
}

function convertirPreciosLegacyANormalizados_(sheet) {
  const values = sheet.getDataRange().getValues();
  const hasLegacyHeader = texto_(values[0][1]).toLowerCase() === 'precio';
  const body = values.slice(hasLegacyHeader ? 1 : 0);
  const legacyMap = {};

  body.forEach(row => {
    const producto = texto_(row[0]);
    const precio = numero_(row[1]);
    const unidad = normalizarUnidadCanonicaQTAS_(row[3] || row[2]);

    if (!producto || precio <= 0) return;

    legacyMap[producto] = {
      precio,
      unidad: unidad || 'und'
    };
  });

  const rows = construirPreciosNormalizadosBase_().map(item => ({
    precioId: item.precioId,
    producto: item.producto,
    precio: item.precio,
    unidad: item.unidad,
    desde: item.desde,
    hasta: item.hasta,
    activo: item.activo,
    nota: item.nota
  }));

  Object.keys(legacyMap).forEach(producto => {
    const sameProduct = rows
      .filter(item => item.producto === producto)
      .sort((a, b) => a.desde - b.desde);

    if (sameProduct.length) {
      const latest = sameProduct[sameProduct.length - 1];
      latest.precio = legacyMap[producto].precio;
      if (legacyMap[producto].unidad) latest.unidad = legacyMap[producto].unidad;
      latest.nota = unirUnicos_([latest.nota, 'Normalizado desde hoja legacy']);
      return;
    }

    rows.push({
      precioId: '',
      producto,
      precio: legacyMap[producto].precio,
      unidad: legacyMap[producto].unidad,
      desde: resolverFechaOperacion_('2024-01-01', new Date()),
      hasta: null,
      activo: true,
      nota: 'Creado desde hoja legacy'
    });
  });

  return rows
    .filter(item => item.producto && item.precio > 0)
    .sort((a, b) => {
      if (a.producto !== b.producto) {
        return a.producto.localeCompare(b.producto);
      }
      return a.desde - b.desde;
    })
    .map((item, index) => ({
      precioId: item.precioId || 'PREC-' + String(index + 1).padStart(4, '0'),
      producto: item.producto,
      precio: item.precio,
      unidad: item.unidad,
      desde: item.desde,
      hasta: item.hasta,
      activo: item.activo,
      nota: item.nota
    }));
}

function construirFilasPreciosNormalizados_() {
  return leerPreciosConfigurados_().map(item => [
    item.precioId,
    item.producto,
    item.precio,
    item.unidad,
    item.desde,
    item.hasta || '',
    item.activo,
    item.nota
  ]);
}

function cargarPreciosReferenciaEnMemoria_() {
  const namespace = 'precios_referencia_memoria';

  return obtenerMemoEjecucionQTAS_(`cache:${namespace}`, () => {
    const cached = leerCacheDocumentoQTAS_(namespace);
    if (cached && Array.isArray(cached.rows)) {
      return cached.rows;
    }

    const rows = leerPreciosConfigurados_()
      .filter(item => item.activo)
      .map(item => ({
        producto: texto_(item.producto),
        unidad: texto_(item.unidad),
        precio: numero_(item.precio),
        desde: fechaInput_(item.desde),
        hasta: item.hasta ? fechaInput_(item.hasta) : ''
      }));

    guardarCacheDocumentoQTAS_(namespace, { rows: rows }, 300);
    return rows;
  });
}

function obtenerPrecioVigenteDesdeCache_(preciosCache, producto, unidad, fechaConsulta) {
  const baseDate = fechaInput_(resolverFechaOperacion_(fechaConsulta, new Date()));
  const productName = texto_(producto);
  const saleUnit = normalizarUnidadCanonicaQTAS_(unidad);

  const matches = preciosCache.filter(item => {
    if (item.producto !== productName) return false;
    if (normalizarUnidadCanonicaQTAS_(item.unidad) !== saleUnit) return false;
    const desde = fechaInput_(item.desde);
    const hasta = item.hasta ? fechaInput_(item.hasta) : '';
    return baseDate >= desde && (!hasta || baseDate <= hasta);
  });

  if (!matches.length) return 0;

  matches.sort((a, b) => b.desde.localeCompare(a.desde));
  return matches[0].precio;
}

function normalizarPreciosReferenciaInterno_() {
  assertOperacionDestructivaPermitidaQTAS_('normalizarPreciosReferenciaInterno_');
  asegurarModeloOperativoQTAS_();

  const ss = SpreadsheetApp.getActive();
  const headers = QTAS.schemas[QTAS.sheets.precios];
  const sheet = asegurarHojaModelo_(ss, QTAS.sheets.precios, headers);
  const needsBackup = sheet.getLastRow() > 1 &&
    !headersIguales_(getHeaders_(sheet), headers);

  let snapshotName = '';
  if (needsBackup) {
    snapshotName = snapshotSheet_(sheet, 'Precios_Referencia_Backup');
  }

  const rows = construirFilasPreciosNormalizados_();
  const currentMaxColumns = sheet.getMaxColumns();

  if (currentMaxColumns < headers.length) {
    sheet.insertColumnsAfter(currentMaxColumns, headers.length - currentMaxColumns);
  } else if (currentMaxColumns > headers.length) {
    sheet.deleteColumns(headers.length + 1, currentMaxColumns - headers.length);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  limpiarCacheHeadersHojaQTAS_(sheet);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  escribirFilas_(sheet, rows);
  sheet.autoResizeColumns(1, headers.length);

  const message = [
    `Precios normalizados: ${rows.length} fila(s).`,
    snapshotName ? `Backup: ${snapshotName}.` : ''
  ].filter(Boolean).join(' ');

  maybeAlert_(message);
  invalidarCacheDocumentoQTAS_('precios_referencia_memoria');

  return {
    ok: true,
    rows: rows.length,
    backup: snapshotName
  };
}

function agregarCambioPrecioQTAS(producto, unidad, nuevoPrecio, fechaDesde, nota) {
  return withScriptLock_('agregar cambio de precio', () => {
    asegurarModeloOperativoQTAS_();

    const ss = SpreadsheetApp.getActive();
    const headers = QTAS.schemas[QTAS.sheets.precios];
    const sheet = asegurarHojaModelo_(ss, QTAS.sheets.precios, headers);

    if (!headersIguales_(getHeaders_(sheet), headers)) {
      normalizarPreciosReferenciaInterno_();
    }

    const rows = leerObjetos_(sheet);
    const colHasta = headers.indexOf('Fecha_Hasta') + 1;
    const desde = resolverFechaOperacion_(fechaDesde || new Date(), new Date());
    const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);

    rows.forEach((row, index) => {
      const sameProduct = texto_(row.Producto_Estandar) === texto_(producto);
      const sameUnit = normalizarUnidadCanonicaQTAS_(row.Unidad) === unidadCanonica;
      const noEndDate = !row.Fecha_Hasta;

      if (sameProduct && sameUnit && noEndDate && estaActivo_(row.Activo)) {
        sheet.getRange(index + 2, colHasta).setValue(diaAnterior_(desde));
      }
    });

    const nextId = 'PREC-' + String(Math.max(sheet.getLastRow(), 1)).padStart(4, '0');

    escribirFilas_(sheet, [[
      nextId,
      texto_(producto),
      numero_(nuevoPrecio),
      unidadCanonica,
      desde,
      '',
      true,
      nota || 'Cambio manual de precio'
    ]]);
    invalidarCacheDocumentoQTAS_('precios_referencia_memoria');
  });
}
