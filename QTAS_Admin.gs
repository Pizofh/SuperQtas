function getConfiguracionAvanzadaQTAS() {
  validarModeloSoloLecturaQTAS_();

  return {
    hoy: fechaInput_(new Date()),
    productos: listarProductosQTAS_(),
    precios: listarPreciosQTAS_(),
    mediosPago: listarMediosPagoQTAS_(),
    reglasDistribucion: listarReglasDistribucionQTAS_(),
    componentesProducto: listarComponentesProductoQTAS_(),
    reglasCostoProducto: listarReglasCostoProductoQTAS_()
  };
}

function guardarProductoConfiguracionQTAS(payload) {
  return withScriptLock_('guardar producto configuracion', () => {
    asegurarModeloOperativoQTAS_();

    const producto = texto_(payload && payload.producto);
    const productoOriginal = texto_(payload && payload.productoOriginal);
    const unidad = normalizarUnidadCanonicaQTAS_(payload && payload.unidad);
    const nota = texto_(payload && payload.nota);
    const activo = payload && payload.activo !== false;

    if (!producto) throw new Error('Falta el nombre del producto.');
    if (!unidad) throw new Error('Falta la unidad del producto.');

    if (productoOriginal && normalizarClaveTexto_(productoOriginal) !== normalizarClaveTexto_(producto)) {
      throw new Error('Renombrar productos desde este panel aun no esta habilitado.');
    }

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.productos);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const existente = rows.find(row =>
      normalizarClaveTexto_(row.Producto_Estandar) === normalizarClaveTexto_(producto)
    );

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, {
        Producto_Estandar: producto,
        Unidad_Default: unidad,
        Activo: activo,
        Nota: nota
      }));
    } else {
      escribirFilas_(sheet, [filaDesdeHeaders_(headers, {
        Producto_Estandar: producto,
        Unidad_Default: unidad,
        Activo: activo,
        Nota: nota
      })]);
    }

    return getConfiguracionAvanzadaQTAS();
  });
}

function cambiarEstadoProductoQTAS(payload) {
  return withScriptLock_('cambiar estado producto', () => {
    asegurarModeloOperativoQTAS_();

    const producto = texto_(payload && payload.producto);
    const activo = Boolean(payload && payload.activo);
    if (!producto) throw new Error('Falta el producto.');

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.productos);
    const headers = getHeaders_(sheet);
    const existente = leerObjetosConMeta_(sheet).find(row =>
      normalizarClaveTexto_(row.Producto_Estandar) === normalizarClaveTexto_(producto)
    );

    if (!existente) throw new Error('No se encontro el producto.');

    actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, {
      Activo: activo
    }));

    return getConfiguracionAvanzadaQTAS();
  });
}

function guardarCambioPrecioFrontendQTAS(payload) {
  return withScriptLock_('guardar cambio precio frontend', () => {
    asegurarModeloOperativoQTAS_();

    const producto = texto_(payload && payload.producto);
    const unidad = normalizarUnidadCanonicaQTAS_(payload && payload.unidad);
    const precio = redondear_(numero_(payload && payload.precio));
    const fechaDesde = resolverFechaOperacion_(payload && payload.fechaDesde, new Date());
    const nota = texto_(payload && payload.nota) || 'Cambio desde configuracion avanzada';

    if (!producto) throw new Error('Falta el producto.');
    if (!unidad) throw new Error('Falta la unidad.');
    if (precio <= 0) throw new Error('El precio debe ser mayor a cero.');

    const existentes = leerPreciosConfigurados_()
      .filter(item =>
        normalizarClaveTexto_(item.producto) === normalizarClaveTexto_(producto) &&
        normalizarClaveTexto_(item.unidad) === normalizarClaveTexto_(unidad)
      )
      .sort((a, b) => b.desde - a.desde);

    if (existentes.length && fechaDesde <= resolverFechaOperacion_(existentes[0].desde, new Date())) {
      throw new Error('La nueva fecha debe ser posterior al ultimo cambio registrado para ese producto.');
    }

    agregarCambioPrecioQTAS(producto, unidad, precio, fechaDesde, nota);
    return getConfiguracionAvanzadaQTAS();
  });
}

function guardarMedioPagoQTAS(payload) {
  return withScriptLock_('guardar medio pago', () => {
    asegurarModeloOperativoQTAS_();

    const medioPago = texto_(payload && payload.medioPago);
    const medioOriginal = texto_(payload && payload.medioOriginal);
    const nota = texto_(payload && payload.nota);
    const activo = payload && payload.activo !== false;

    if (!medioPago) throw new Error('Falta el medio de pago.');
    if (medioOriginal && normalizarClaveTexto_(medioOriginal) !== normalizarClaveTexto_(medioPago)) {
      throw new Error('Renombrar medios de pago desde este panel aun no esta habilitado.');
    }

    const sheet = materializarConfigMediosPagoQTAS_();
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const existente = rows.find(row =>
      normalizarClaveTexto_(row.Medio_Pago) === normalizarClaveTexto_(medioPago)
    );

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, {
        Medio_Pago: medioPago,
        Activo: activo,
        Nota: nota
      }));
    } else {
      escribirFilas_(sheet, [filaDesdeHeaders_(headers, {
        Medio_Pago: medioPago,
        Activo: activo,
        Nota: nota
      })]);
    }

    applyFormatosConfigMediosPagoQTAS_(sheet);
    return getConfiguracionAvanzadaQTAS();
  });
}

function cambiarEstadoMedioPagoQTAS(payload) {
  return withScriptLock_('cambiar estado medio pago', () => {
    asegurarModeloOperativoQTAS_();

    const medioPago = texto_(payload && payload.medioPago);
    const activo = Boolean(payload && payload.activo);
    if (!medioPago) throw new Error('Falta el medio de pago.');

    const sheet = materializarConfigMediosPagoQTAS_();
    const headers = getHeaders_(sheet);
    const existente = leerObjetosConMeta_(sheet).find(row =>
      normalizarClaveTexto_(row.Medio_Pago) === normalizarClaveTexto_(medioPago)
    );

    if (!existente) throw new Error('No se encontro el medio de pago.');

    actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, {
      Activo: activo
    }));

    return getConfiguracionAvanzadaQTAS();
  });
}

function guardarReglaDistribucionFrontendQTAS(payload) {
  return withScriptLock_('guardar regla distribucion frontend', () => {
    asegurarModeloOperativoQTAS_();

    const fechaDesde = resolverFechaOperacion_(payload && payload.fechaDesde, new Date());
    const steve = redondear_(numero_(payload && payload.steve));
    const majo = redondear_(numero_(payload && payload.majo));
    const mush = redondear_(numero_(payload && payload.mush));
    const nota = texto_(payload && payload.nota) || 'Nueva regla desde configuracion avanzada';

    const total = redondear_(steve + majo + mush);
    if (Math.abs(total - 100) > 0.01) {
      throw new Error('La regla debe sumar 100%.');
    }

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.distribucionReglas);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet)
      .map(row => ({
        rowNumber: row.__rowNumber,
        reglaId: texto_(row.Regla_ID),
        desde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
        hasta: row.Fecha_Hasta ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date()) : null,
        steve: redondear_(numero_(row.Steve_Pct)),
        majo: redondear_(numero_(row.Majo_Pct)),
        mush: redondear_(numero_(row.Mush_Pct)),
        activo: estaActivo_(row.Activo),
        nota: texto_(row.Nota),
        raw: row
      }))
      .filter(row => row.reglaId && row.activo)
      .sort((a, b) => a.desde - b.desde);

    const ultima = rows.length ? rows[rows.length - 1] : null;
    if (ultima && fechaDesde <= ultima.desde) {
      throw new Error('La nueva regla debe empezar despues de la ultima regla vigente.');
    }

    const reglasProyectadas = rows.map(row => ({
      reglaId: row.reglaId,
      desde: row.desde,
      hasta: row.hasta,
      steve: row.steve,
      majo: row.majo,
      mush: row.mush,
      activo: row.activo,
      nota: row.nota
    }));

    if (reglasProyectadas.length) {
      reglasProyectadas[reglasProyectadas.length - 1].hasta = diaAnterior_(fechaDesde);
    }

    reglasProyectadas.push({
      reglaId: siguienteIdConPrefijo_(sheet, 'Regla_ID', 'DIST-', 4),
      desde: fechaDesde,
      hasta: null,
      steve: steve,
      majo: majo,
      mush: mush,
      activo: true,
      nota: nota
    });

    validarReglasDistribucionQTAS_(reglasProyectadas);

    if (ultima) {
      actualizarFilaObjeto_(sheet, ultima.rowNumber, headers, Object.assign({}, ultima.raw, {
        Fecha_Hasta: diaAnterior_(fechaDesde)
      }));
    }

    const nueva = reglasProyectadas[reglasProyectadas.length - 1];
    escribirFilas_(sheet, [filaDesdeHeaders_(headers, {
      Regla_ID: nueva.reglaId,
      Fecha_Desde: nueva.desde,
      Fecha_Hasta: '',
      Steve_Pct: nueva.steve,
      Majo_Pct: nueva.majo,
      Mush_Pct: nueva.mush,
      Activo: true,
      Nota: nueva.nota
    })]);
    invalidarCacheDocumentoQTAS_('distribucion_reglas_memoria');

    return getConfiguracionAvanzadaQTAS();
  });
}

function listarProductosQTAS_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(QTAS.sheets.productos);
  return leerObjetos_(sheet)
    .map(row => ({
      producto: texto_(row.Producto_Estandar),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad_Default),
      activo: estaActivo_(row.Activo),
      nota: texto_(row.Nota)
    }))
    .filter(row => row.producto)
    .sort((a, b) => a.producto.localeCompare(b.producto));
}

function listarPreciosQTAS_() {
  return leerPreciosConfigurados_()
    .map(row => ({
      precioId: texto_(row.precioId),
      producto: texto_(row.producto),
      unidad: texto_(row.unidad),
      precio: redondear_(numero_(row.precio)),
      fechaDesde: fechaInput_(row.desde),
      fechaHasta: row.hasta ? fechaInput_(row.hasta) : '',
      activo: estaActivo_(row.activo),
      nota: texto_(row.nota)
    }))
    .sort((a, b) => {
      if (a.producto !== b.producto) return a.producto.localeCompare(b.producto);
      return b.fechaDesde.localeCompare(a.fechaDesde);
    });
}

function listarMediosPagoQTAS_() {
  return leerMediosPagoConfiguradosQTAS_()
    .map(row => ({
      medioPago: texto_(row.medioPago),
      activo: estaActivo_(row.activo),
      nota: texto_(row.nota)
    }))
    .filter(row => row.medioPago)
    .sort((a, b) => a.medioPago.localeCompare(b.medioPago));
}

function listarReglasDistribucionQTAS_() {
  return leerReglasDistribucionQTAS_()
    .map(row => ({
      reglaId: texto_(row.reglaId),
      fechaDesde: fechaInput_(row.desde),
      fechaHasta: row.hasta ? fechaInput_(row.hasta) : '',
      steve: redondear_(numero_(row.steve)),
      majo: redondear_(numero_(row.majo)),
      mush: redondear_(numero_(row.mush)),
      activo: estaActivo_(row.activo),
      nota: texto_(row.nota)
    }))
    .sort((a, b) => b.fechaDesde.localeCompare(a.fechaDesde));
}

function materializarConfigMediosPagoQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const headers = QTAS.schemas[QTAS.sheets.config];
  const medios = leerMediosPagoConfiguradosQTAS_();
  let sheet = ss.getSheetByName(QTAS.sheets.config);
  const legacy = ss.getSheetByName(QTAS.sheets.configLegacy);

  if (!sheet && legacy) {
    legacy.setName(QTAS.sheets.config);
    sheet = legacy;
  }

  if (!sheet) {
    sheet = asegurarHojaModelo_(ss, QTAS.sheets.config, headers);
  }

  if (!headersIguales_(getHeaders_(sheet), headers)) {
    assertOperacionDestructivaPermitidaQTAS_('reconstruir Config_MediosPago');
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    limpiarCacheHeadersHojaQTAS_(sheet);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    escribirFilas_(sheet, medios.map(item => [
      item.medioPago,
      item.activo,
      item.nota || ''
    ]));
  }

  applyFormatosConfigMediosPagoQTAS_(sheet);
  return sheet;
}
