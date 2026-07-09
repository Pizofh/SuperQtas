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

  throw new Error(
    `La hoja ${QTAS.sheets.precios} no coincide con la estructura esperada. ` +
    'Corrige la hoja antes de continuar.'
  );
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

function agregarCambioPrecioQTAS(producto, unidad, nuevoPrecio, fechaDesde, nota) {
  return withScriptLock_('agregar cambio de precio', () => {
    asegurarModeloOperativoQTAS_();

    const ss = SpreadsheetApp.getActive();
    const headers = QTAS.schemas[QTAS.sheets.precios];
    const sheet = asegurarHojaModelo_(ss, QTAS.sheets.precios, headers);

    if (!headersIguales_(getHeaders_(sheet), headers)) {
      throw new Error(
        `La hoja ${QTAS.sheets.precios} no coincide con la estructura esperada. ` +
        'Corrige la hoja antes de registrar cambios de precio.'
      );
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
