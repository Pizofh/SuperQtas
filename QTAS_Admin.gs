function getConfiguracionAvanzadaQTAS() {
  validarModeloSoloLecturaQTAS_();

  return {
    hoy: fechaInput_(new Date()),
    productos: listarProductosQTAS_(),
    precios: listarPreciosQTAS_(),
    mediosPago: listarMediosPagoQTAS_(),
    reglasDistribucion: listarReglasDistribucionQTAS_(),
    reglasOrigenesFondos: listarReglasOrigenesFondosQTAS_(),
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

function guardarReglaOrigenFondosFrontendQTAS(payload) {
  return withScriptLock_('guardar regla origen fondos frontend', () => {
    asegurarModeloOperativoQTAS_();

    const origenFondos = texto_(payload && payload.origenFondos);
    const fechaDesde = resolverFechaOperacion_(payload && payload.fechaDesde, new Date());
    const steve = redondear_(numero_(payload && payload.steve));
    const majo = redondear_(numero_(payload && payload.majo));
    const mush = redondear_(numero_(payload && payload.mush));
    const nota = texto_(payload && payload.nota) || 'Nueva regla de origen de fondos';

    if (!origenFondos) {
      throw new Error('Falta el origen de fondos.');
    }

    const total = redondear_(steve + majo + mush);
    if (Math.abs(total - 100) > 0.01) {
      throw new Error('La regla debe sumar 100%.');
    }

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.origenesFondosReglas);
    const headers = getHeaders_(sheet);
    const reglas = leerReglasOrigenesFondosQTAS_()
      .filter(row => normalizarClaveTexto_(row.origenFondos) === normalizarClaveTexto_(origenFondos))
      .sort((a, b) => a.desde - b.desde);

    const ultima = reglas.length ? reglas[reglas.length - 1] : null;
    if (ultima && fechaDesde <= ultima.desde) {
      throw new Error('La nueva regla debe empezar despues de la ultima regla registrada para ese origen.');
    }

    if (ultima) {
      leerObjetosConMeta_(sheet)
        .filter(row => texto_(row.Regla_ID) === texto_(ultima.reglaId))
        .forEach(row => {
          actualizarFilaObjeto_(sheet, row.__rowNumber, headers, Object.assign({}, row, {
            Fecha_Hasta: diaAnterior_(fechaDesde)
          }));
        });
    }

    const reglaId = siguienteIdConPrefijo_(sheet, 'Regla_ID', 'ORG-', 4);
    const aportantes = construirAportantesOrigenFondosQTAS_({
      steve: steve,
      majo: majo,
      mush: mush
    });

    escribirFilas_(sheet, aportantes.map(item => filaDesdeHeaders_(headers, {
      Regla_ID: reglaId,
      Origen_Fondos: origenFondos,
      Fecha_Desde: fechaDesde,
      Fecha_Hasta: '',
      Aportante: item.aportante,
      Porcentaje: item.porcentaje,
      Nota: nota
    })));

    invalidarCacheDocumentoQTAS_('origenes_fondos_reglas_memoria');
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

function listarReglasOrigenesFondosQTAS_() {
  return leerReglasOrigenesFondosQTAS_()
    .map(row => ({
      reglaId: texto_(row.reglaId),
      origenFondos: texto_(row.origenFondos),
      fechaDesde: fechaInput_(row.desde),
      fechaHasta: row.hasta ? fechaInput_(row.hasta) : '',
      steve: redondear_(numero_(row.steve)),
      majo: redondear_(numero_(row.majo)),
      mush: redondear_(numero_(row.mush)),
      activo: true,
      nota: texto_(row.nota)
    }))
    .sort((a, b) => {
      const origen = a.origenFondos.localeCompare(b.origenFondos, undefined, { sensitivity: 'base' });
      if (origen !== 0) return origen;
      return b.fechaDesde.localeCompare(a.fechaDesde);
    });
}

function materializarConfigMediosPagoQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const headers = QTAS.schemas[QTAS.sheets.config];
  const medios = leerMediosPagoConfiguradosQTAS_();
  let sheet = ss.getSheetByName(QTAS.sheets.config);

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

function auditarIntegridadFinancieraQTAS() {
  const ss = validarHojasIntegridadFinancieraQTAS_();
  const plan = construirPlanIntegridadFinancieraQTAS_(ss);
  Logger.log(JSON.stringify(plan.resumen, null, 2));
  return plan.resumen;
}

function corregirVentasHistoricasConfirmadasQTAS() {
  assertOperacionDestructivaPermitidaQTAS_('corregir ventas historicas confirmadas');

  try {
    return withScriptLock_('corregir ventas historicas confirmadas', () => {
      const ss = validarModeloSoloLecturaQTAS_({
        sheetNames: [
          QTAS.sheets.ventas,
          QTAS.sheets.detalle,
          QTAS.sheets.pagos,
          QTAS.sheets.distribucionIngresos,
          QTAS.sheets.ventaDetalleCostosCalculado
        ],
        validarConfig: false
      });
      const ventasSheet = ss.getSheetByName(QTAS.sheets.ventas);
      const detalleSheet = ss.getSheetByName(QTAS.sheets.detalle);
      const pagosSheet = ss.getSheetByName(QTAS.sheets.pagos);
      const distribucionSheet = ss.getSheetByName(QTAS.sheets.distribucionIngresos);
      const ventasHeaders = getHeaders_(ventasSheet);
      const detalleHeaders = getHeaders_(detalleSheet);
      const pagosHeaders = getHeaders_(pagosSheet);
      const distribucionHeaders = getHeaders_(distribucionSheet);
      const ventas = leerObjetosConMeta_(ventasSheet);
      const detalle = leerObjetosConMeta_(detalleSheet);
      const pagos = leerObjetosConMeta_(pagosSheet);
      const distribuciones = leerObjetosConMeta_(distribucionSheet);
      const correcciones = construirCorreccionesVentasHistoricasConfirmadasQTAS_();
      const ventasPorId = indexarObjetosPorCampoQTAS_(ventas, 'Venta_ID');
      const detallePorId = indexarObjetosPorCampoQTAS_(detalle, 'Detalle_ID');
      const faltantes = correcciones
        .filter(item => !detallePorId[item.detalleId] || !ventasPorId[String(item.ventaId)])
        .map(item => item.detalleId);

      if (faltantes.length) {
        throw new Error(`No se encontraron los detalles esperados: ${faltantes.join(', ')}.`);
      }

      const ventaIdsAfectadas = {};
      const detalleActualizado = [];
      correcciones.forEach(item => {
        const existente = detallePorId[item.detalleId];
        if (numero_(existente.Venta_ID) !== numero_(item.ventaId)) {
          throw new Error(`El detalle ${item.detalleId} no pertenece a la venta ${item.ventaId}.`);
        }

        const actualizado = Object.assign({}, existente, item.values);
        actualizarFilaObjeto_(detalleSheet, existente.__rowNumber, detalleHeaders, actualizado);
        Object.assign(existente, actualizado);
        ventaIdsAfectadas[String(item.ventaId)] = true;
        detalleActualizado.push(existente);
      });

      const pagosActualizados = corregirPagosVentasHistoricasConfirmadasQTAS_({
        pagosSheet: pagosSheet,
        pagosHeaders: pagosHeaders,
        pagos: pagos,
        ventasPorId: ventasPorId
      });
      pagosActualizados.ventaIds.forEach(ventaId => {
        ventaIdsAfectadas[String(ventaId)] = true;
      });

      const ventasActualizadas = actualizarResumenVentasHistoricasConfirmadasQTAS_({
        ventasSheet: ventasSheet,
        ventasHeaders: ventasHeaders,
        ventasPorId: ventasPorId,
        detalle: detalle,
        pagos: pagos,
        ventaIds: Object.keys(ventaIdsAfectadas)
      });

      const distribucion = sincronizarDistribucionesVentasHistoricasConfirmadasQTAS_({
        distribucionSheet: distribucionSheet,
        distribucionHeaders: distribucionHeaders,
        distribuciones: distribuciones,
        ventasPorId: ventasPorId,
        pagos: pagos,
        ventaIds: Object.keys(ventaIdsAfectadas)
      });

      const primeraVentaAcMed = detallePorId['DET-000001-01'];
      const detallesAnalitica = detalleActualizado.slice();
      if (
        primeraVentaAcMed &&
        !detallesAnalitica.some(row => texto_(row.Detalle_ID) === 'DET-000001-01')
      ) {
        detallesAnalitica.push(primeraVentaAcMed);
      }

      limpiarCachesEjecucionQTAS_();
      const analitica = sincronizarVentaDetalleCostosLoteQTAS_(detallesAnalitica, {
        ss: ss,
        ahora: new Date()
      });
      const filaAcMed = leerObjetos_(
        ss.getSheetByName(QTAS.sheets.ventaDetalleCostosCalculado)
      ).find(row => texto_(row.Detalle_ID) === 'DET-000001-01');
      const verificaciones = verificarVentasHistoricasConfirmadasQTAS_({
        ventasPorId: ventasPorId,
        detalle: detalle,
        pagos: pagos,
        ventaIds: Object.keys(ventaIdsAfectadas)
      });
      const costoInicialAcMed = redondear_(numero_(filaAcMed && filaAcMed.Costo_Unitario_Usado));
      const result = {
        ok: verificaciones.ok && analitica.ok === true && costoInicialAcMed > 0,
        ventasActualizadas: ventasActualizadas,
        detallesActualizados: detalleActualizado.length,
        pagosActualizados: pagosActualizados.updated,
        pagosCreados: pagosActualizados.inserted,
        distribucion: distribucion,
        analitica: analitica,
        costoInicialAcMed: costoInicialAcMed,
        verificaciones: verificaciones,
        noModificadas: [2342, 2345]
      };

      guardarEstadoReparacionIntegridadFinancieraQTAS_({
        ok: result.ok,
        active: false,
        pending: !result.ok,
        status: result.ok ? 'Correcciones historicas aplicadas' : 'Correcciones aplicadas con pendientes',
        result: result,
        completedAt: new Date().toISOString()
      });
      Logger.log(JSON.stringify(result, null, 2));
      return result;
    });
  } finally {
    PropertiesService.getScriptProperties().deleteProperty('QTAS_ALLOW_DESTRUCTIVE');
  }
}

function corregirVentasHistoricasConfirmadasQTAS_Log() {
  const result = corregirVentasHistoricasConfirmadasQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function construirCorreccionesVentasHistoricasConfirmadasQTAS_() {
  return [
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000172-01', 172, 'Choco', 1, 'und', 20000, 20000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000172-02', 172, '200mg', 5, 'und', 3000, 15000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000235-01', 235, 'Lm', 40, 'g', 1500, 40000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000501-03', 501, '200mg', 10, 'und', 3000, 30000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000514-01', 514, 'Shii', 45, 'g', 1500, 30000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000552-01', 552, 'LmExt', 2, 'und', 50000, 90000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000633-01', 633, 'Shii', 45, 'g', 1500, 45000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000688-01', 688, 'Gano', 15, 'g', 1500, 20000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000688-02', 688, 'Shii', 100, 'g', 1500, 85000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000697-02', 697, 'Lm', 15, 'g', 1500, 25000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000697-03', 697, 'Cordy', 1, 'g', 2500, 3000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-000786-01', 786, 'AcSup', 3, 'g', 20000, 60000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-001245-01', 1245, '300mg', 25, 'und', 4000, 100000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-001282-01', 1282, 'AcMed', 12, 'g', 12000, 144000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-001369-02', 1369, 'Choco', 1, 'und', 20000, 20000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-001857-01', 1857, 'AcSup', 15, 'g', 20000, 255000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-002221-01', 2221, 'Lm', 20, 'g', 1500, 30000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-002221-02', 2221, 'AcAlt', 20, 'g', 18000, 168000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-002412-01', 2412, 'Lm', 30, 'g', 1500, 45000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-002464-01', 2464, 'AcSup', 15, 'g', 20000, 225000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-002464-02', 2464, 'Lm', 50, 'g', 1500, 45000),
    crearCorreccionDetalleVentaHistoricaQTAS_('DET-002591-01', 2591, '200mg', 10, 'und', 3000, 30000)
  ];
}

function crearCorreccionDetalleVentaHistoricaQTAS_(
  detalleId,
  ventaId,
  producto,
  cantidad,
  unidad,
  precioLista,
  subtotalNeto
) {
  const precioVendido = cantidad > 0 ? subtotalNeto / cantidad : 0;
  return {
    detalleId: detalleId,
    ventaId: ventaId,
    values: {
      Producto_Estandar: producto,
      Cantidad: cantidad,
      Unidad: unidad,
      Precio_Lista: precioLista,
      Precio_Vendido_Unitario: precioVendido,
      Descuento_Linea: 0,
      Subtotal_Bruto: subtotalNeto,
      Subtotal_Neto: subtotalNeto
    }
  };
}

function corregirPagosVentasHistoricasConfirmadasQTAS_(context) {
  const montosCorregidos = {
    '786': 60000,
    '1245': 100000
  };
  const ventaIdsSinPago = [2639, 2641, 2642];
  let updated = 0;
  const nuevos = [];

  Object.keys(montosCorregidos).forEach(ventaId => {
    const activos = context.pagos.filter(row =>
      texto_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro)
    );
    if (activos.length !== 1) {
      throw new Error(`La venta ${ventaId} debe tener exactamente un pago activo para corregirlo.`);
    }
    const pago = activos[0];
    const actualizado = Object.assign({}, pago, {
      Monto_Pago: montosCorregidos[ventaId]
    });
    actualizarFilaObjeto_(
      context.pagosSheet,
      pago.__rowNumber,
      context.pagosHeaders,
      actualizado
    );
    Object.assign(pago, actualizado);
    updated++;
  });

  ventaIdsSinPago.forEach(ventaId => {
    const venta = context.ventasPorId[String(ventaId)];
    if (!venta) throw new Error(`No se encontro la venta ${ventaId}.`);
    const activos = context.pagos.filter(row =>
      numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro)
    );
    if (activos.length) return;

    const pago = {
      Pago_ID: siguientePagoIdVentaQTAS_(context.pagos.concat(nuevos), ventaId),
      Venta_ID: ventaId,
      Fecha_Pago: venta.Fecha_Venta,
      Medio_Pago: 'Otro',
      Monto_Pago: redondear_(numero_(venta.Total_Pagado) || numero_(venta.Total_Venta)),
      Comentario_Pago: 'Pago historico conciliado; medio original no disponible.',
      Regla_Distribucion_Pago_ID: texto_(venta.Regla_Distribucion_Venta_ID),
      Steve_Pct_Pago: numero_(venta.Steve_Pct_Venta),
      Majo_Pct_Pago: numero_(venta.Majo_Pct_Venta),
      Mush_Pct_Pago: numero_(venta.Mush_Pct_Venta),
      Estado_Registro: QTAS.status.registro.activo
    };
    nuevos.push(pago);
  });

  if (nuevos.length) {
    escribirFilas_(
      context.pagosSheet,
      nuevos.map(row => filaDesdeHeaders_(context.pagosHeaders, row))
    );
    context.pagos.push.apply(context.pagos, nuevos);
  }

  return {
    updated: updated,
    inserted: nuevos.length,
    ventaIds: Object.keys(montosCorregidos).map(Number).concat(ventaIdsSinPago)
  };
}

function actualizarResumenVentasHistoricasConfirmadasQTAS_(context) {
  let updated = 0;
  (context.ventaIds || []).forEach(ventaId => {
    const venta = context.ventasPorId[String(ventaId)];
    if (!venta) throw new Error(`No se encontro la venta ${ventaId}.`);
    const lineas = context.detalle.filter(row =>
      texto_(row.Venta_ID) === String(ventaId) && !esRegistroAnulado_(row.Estado_Registro)
    );
    const pagos = context.pagos.filter(row =>
      texto_(row.Venta_ID) === String(ventaId) && !esRegistroAnulado_(row.Estado_Registro)
    );
    const totalVenta = redondear_(sumar_(lineas.map(row => row.Subtotal_Neto)));
    const totalPagado = redondear_(sumar_(pagos.map(row => row.Monto_Pago)));
    if (totalPagado > totalVenta + 0.01) {
      throw new Error(`La venta ${ventaId} quedaria con pagos mayores al total.`);
    }
    const saldo = redondear_(Math.max(totalVenta - totalPagado, 0));
    const actualizado = Object.assign({}, venta, {
      Productos_Resumen: resumenProductos_(lineas),
      Total_Venta: totalVenta,
      Total_Pagado: totalPagado,
      Saldo: saldo,
      Estado_Pago: obtenerEstadoPago_(totalVenta, totalPagado, venta.Estado_Registro)
    });
    actualizarFilaObjeto_(
      context.ventasSheet,
      venta.__rowNumber,
      context.ventasHeaders,
      actualizado
    );
    Object.assign(venta, actualizado);
    updated++;
  });
  return updated;
}

function sincronizarDistribucionesVentasHistoricasConfirmadasQTAS_(context) {
  const ventaIds = {};
  (context.ventaIds || []).forEach(ventaId => {
    ventaIds[String(ventaId)] = true;
  });
  const existentes = context.distribuciones.filter(row => ventaIds[texto_(row.Venta_ID)]);
  const existentesPorId = indexarObjetosPorCampoQTAS_(existentes, 'Distribucion_ID');
  const objetivosPorId = {};
  const nuevos = [];
  let updated = 0;
  let cancelled = 0;

  Object.keys(ventaIds).forEach(ventaId => {
    const venta = context.ventasPorId[ventaId];
    const pagos = context.pagos.filter(row =>
      texto_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro)
    );
    const objetivos = [construirFilaDistribucionVentaQTAS_(venta)]
      .concat(pagos.map(pago => construirFilaDistribucionPagoQTAS_(pago, venta)));

    objetivos.forEach(objetivo => {
      const id = texto_(objetivo.Distribucion_ID);
      objetivosPorId[id] = true;
      const existente = existentesPorId[id];
      if (!existente) {
        nuevos.push(objetivo);
        return;
      }
      actualizarFilaObjeto_(
        context.distribucionSheet,
        existente.__rowNumber,
        context.distribucionHeaders,
        Object.assign({}, existente, objetivo)
      );
      updated++;
    });
  });

  existentes.forEach(existente => {
    const id = texto_(existente.Distribucion_ID);
    if (objetivosPorId[id]) return;
    actualizarFilaObjeto_(
      context.distribucionSheet,
      existente.__rowNumber,
      context.distribucionHeaders,
      Object.assign({}, existente, {
        Monto_Base: 0,
        Steve_Valor: 0,
        Majo_Valor: 0,
        Mush_Valor: 0,
        Estado_Registro: QTAS.status.registro.anulado
      })
    );
    cancelled++;
  });

  if (nuevos.length) {
    escribirFilas_(
      context.distribucionSheet,
      nuevos.map(row => filaDesdeHeaders_(context.distribucionHeaders, row))
    );
  }

  return {
    updated: updated,
    inserted: nuevos.length,
    cancelled: cancelled
  };
}

function verificarVentasHistoricasConfirmadasQTAS_(context) {
  const totals = [];
  const payments = [];
  (context.ventaIds || []).forEach(ventaId => {
    const venta = context.ventasPorId[String(ventaId)];
    const totalDetalle = redondear_(sumar_(context.detalle
      .filter(row => texto_(row.Venta_ID) === String(ventaId) && !esRegistroAnulado_(row.Estado_Registro))
      .map(row => row.Subtotal_Neto)));
    const totalPagos = redondear_(sumar_(context.pagos
      .filter(row => texto_(row.Venta_ID) === String(ventaId) && !esRegistroAnulado_(row.Estado_Registro))
      .map(row => row.Monto_Pago)));
    if (Math.abs(numero_(venta.Total_Venta) - totalDetalle) > 0.01) {
      totals.push(Number(ventaId));
    }
    if (Math.abs(numero_(venta.Total_Pagado) - totalPagos) > 0.01) {
      payments.push(Number(ventaId));
    }
  });
  return {
    ok: totals.length === 0 && payments.length === 0,
    ventasDescuadradas: totals,
    pagosDescuadrados: payments
  };
}

function indexarObjetosPorCampoQTAS_(rows, field) {
  return (rows || []).reduce((acc, row) => {
    const key = texto_(row && row[field]);
    if (key) acc[key] = row;
    return acc;
  }, {});
}

function repararIntegridadFinancieraQTAS() {
  assertOperacionDestructivaPermitidaQTAS_('reparar integridad financiera historica');

  return withScriptLock_('reparar integridad financiera historica', () => {
    const estadoActivo = leerEstadoReparacionIntegridadFinancieraQTAS_();
    if (estadoActivo && estadoActivo.active) {
      return estadoActivo;
    }

    const ss = validarHojasIntegridadFinancieraQTAS_();
    const plan = construirPlanIntegridadFinancieraQTAS_(ss);
    if (plan.resumen.ok === true) {
      limpiarTriggersReparacionIntegridadFinancieraQTAS_();
      PropertiesService.getScriptProperties().deleteProperty('QTAS_ALLOW_DESTRUCTIVE');
      const clean = {
        ok: true,
        aplicado: false,
        active: false,
        pending: false,
        status: 'Sin cambios pendientes.',
        verification: resumirVerificacionIntegridadFinancieraQTAS_(plan.resumen),
        checkedAt: new Date().toISOString()
      };
      guardarEstadoReparacionIntegridadFinancieraQTAS_(clean);
      Logger.log(JSON.stringify(clean, null, 2));
      return clean;
    }

    const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
    const fondosSheet = ss.getSheetByName(QTAS.sheets.compraOrigenesFondos);
    const costosSheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
    const ventaDetalleSheet = ss.getSheetByName(QTAS.sheets.detalle);

    if (plan.correccionesCosto.length || plan.correccionesImpactoCosto.length) {
      sobrescribirObjetosHojaQTAS_(
        detalleSheet,
        getHeaders_(detalleSheet),
        plan.compraDetalleProyectado
      );
    }

    if (plan.correccionesFondos.length) {
      sobrescribirObjetosHojaQTAS_(
        fondosSheet,
        getHeaders_(fondosSheet),
        plan.fondosProyectados
      );
    }

    if (plan.correccionesCostoDerivado.length || plan.correccionesCoberturaPackaging.length) {
      sobrescribirObjetosHojaQTAS_(
        costosSheet,
        getHeaders_(costosSheet),
        plan.costosReferenciaProyectados
      );
    }

    if (plan.correccionesVentaDetalle.length) {
      sobrescribirObjetosHojaQTAS_(
        ventaDetalleSheet,
        getHeaders_(ventaDetalleSheet),
        plan.ventaDetalleProyectado
      );
    }

    limpiarCachesEjecucionQTAS_();
    const costos = reconstruirCostosReferenciaDesdeFuentesQTAS_({ ss: ss });
    const costoProducto = reconstruirCostoProductoCalculadoInternoQTAS_({
      ss: ss,
      fechaBase: new Date(),
      ahora: new Date()
    });
    limpiarCachesEjecucionQTAS_();

    const cambios = {
      compraDetalle: plan.correccionesCosto.length,
      comprasSinCostoUnitarioFalso: plan.correccionesImpactoCosto.length,
      costosReferenciaDerivados: plan.correccionesCostoDerivado.length,
      costosPackagingConHistoria: plan.correccionesCoberturaPackaging.length,
      compraOrigenesFondos: plan.correccionesFondos.length,
      comprasConFondosCorregidos: plan.resumen.compras.fondosDescuadrados,
      ventaDetalleCanonico: plan.correccionesVentaDetalle.length
    };
    const ventaDetalleCostos = iniciarReparacionAnaliticaIntegridadFinancieraQTAS_(
      ss,
      cambios
    );
    const result = {
      ok: true,
      aplicado: true,
      pending: true,
      cambios: cambios,
      reconstrucciones: {
        costosReferencia: costos,
        costoProducto: costoProducto,
        ventaDetalleCostos: ventaDetalleCostos
      }
    };

    Logger.log(JSON.stringify(result, null, 2));
    return result;
  });
}

function continuarReparacionIntegridadFinancieraQTAS() {
  try {
    return withScriptLock_('continuar reparacion integridad financiera', () => {
      let state = leerEstadoReparacionIntegridadFinancieraQTAS_();
      const incompleto = state && numero_(state.cursor) < numero_(state.total);
      if (state && !state.active && state.status === 'Error' && incompleto) {
        state = Object.assign({}, state, {
          ok: true,
          active: true,
          pending: true,
          status: 'Reanudando',
          error: '',
          resumedAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        });
        guardarEstadoReparacionIntegridadFinancieraQTAS_(state);
      }

      if (!state || !state.active) {
        limpiarTriggersReparacionIntegridadFinancieraQTAS_();
        return state || {
          ok: true,
          active: false,
          status: 'Sin proceso pendiente.'
        };
      }

      const ss = SpreadsheetApp.openById(state.spreadsheetId);
      const detalle = leerObjetos_(ss.getSheetByName(QTAS.sheets.detalle))
        .filter(row => texto_(row.Detalle_ID));
      const total = detalle.length;
      const cursor = Math.max(0, Math.min(numero_(state.cursor), total));
      const batchSize = Math.max(100, numero_(state.batchSize) || 500);
      const end = Math.min(cursor + batchSize, total);
      const ids = {};

      detalle.slice(cursor, end).forEach(row => {
        const id = texto_(row.Detalle_ID);
        if (id) ids[id] = true;
      });

      const sync = sincronizarAnaliticaIntegridadFinancieraQTAS_(ss, ids);
      const updated = Object.assign({}, state, {
        ok: true,
        active: end < total,
        status: end < total ? 'Procesando' : 'Verificando',
        cursor: end,
        total: total,
        processed: end,
        remaining: Math.max(total - end, 0),
        progressPct: total > 0 ? redondear_(end * 100 / total) : 100,
        lastBatch: sync,
        updatedAt: new Date().toISOString()
      });

      if (end < total) {
        guardarEstadoReparacionIntegridadFinancieraQTAS_(updated);
        programarContinuacionReparacionIntegridadFinancieraQTAS_();
        Logger.log(JSON.stringify(updated, null, 2));
        return updated;
      }

      limpiarCachesEjecucionQTAS_();
      const verification = construirPlanIntegridadFinancieraQTAS_(ss).resumen;
      const completed = Object.assign({}, updated, {
        ok: verification.ok === true,
        active: false,
        pending: false,
        status: verification.ok === true ? 'Completado' : 'Completado con pendientes',
        verification: resumirVerificacionIntegridadFinancieraQTAS_(verification),
        completedAt: new Date().toISOString()
      });
      guardarEstadoReparacionIntegridadFinancieraQTAS_(completed);
      limpiarTriggersReparacionIntegridadFinancieraQTAS_();
      PropertiesService.getScriptProperties().deleteProperty('QTAS_ALLOW_DESTRUCTIVE');
      Logger.log(JSON.stringify(completed, null, 2));
      return completed;
    });
  } catch (error) {
    const state = leerEstadoReparacionIntegridadFinancieraQTAS_() || {};
    if (normalizarClaveTexto_(error.message).indexOf('lock timeout') >= 0) {
      const waiting = Object.assign({}, state, {
        ok: true,
        active: true,
        pending: true,
        status: 'Esperando disponibilidad',
        lastLockContentionAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
      guardarEstadoReparacionIntegridadFinancieraQTAS_(waiting);
      programarContinuacionReparacionIntegridadFinancieraQTAS_();
      Logger.log(JSON.stringify(waiting, null, 2));
      return waiting;
    }

    const failed = Object.assign({}, state, {
      ok: false,
      active: false,
      pending: false,
      status: 'Error',
      error: error.message,
      failedAt: new Date().toISOString()
    });
    guardarEstadoReparacionIntegridadFinancieraQTAS_(failed);
    limpiarTriggersReparacionIntegridadFinancieraQTAS_();
    PropertiesService.getScriptProperties().deleteProperty('QTAS_ALLOW_DESTRUCTIVE');
    Logger.log(JSON.stringify(failed, null, 2));
    throw error;
  }
}

function estadoReparacionIntegridadFinancieraQTAS() {
  const state = leerEstadoReparacionIntegridadFinancieraQTAS_() || {
    ok: true,
    active: false,
    pending: false,
    status: 'Sin ejecucion registrada.'
  };
  Logger.log(JSON.stringify(state, null, 2));
  return state;
}

function iniciarReparacionAnaliticaIntegridadFinancieraQTAS_(ss, cambios) {
  const detalle = leerObjetos_(ss.getSheetByName(QTAS.sheets.detalle))
    .filter(row => texto_(row.Detalle_ID));
  const state = {
    ok: true,
    active: true,
    pending: true,
    status: 'Programado',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    cursor: 0,
    total: detalle.length,
    processed: 0,
    remaining: detalle.length,
    progressPct: 0,
    batchSize: 500,
    cambios: cambios || {},
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  guardarEstadoReparacionIntegridadFinancieraQTAS_(state);
  programarContinuacionReparacionIntegridadFinancieraQTAS_();
  return state;
}

function leerEstadoReparacionIntegridadFinancieraQTAS_() {
  const raw = PropertiesService.getScriptProperties().getProperty(
    'QTAS_FINANCIAL_REPAIR_STATE'
  );
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function guardarEstadoReparacionIntegridadFinancieraQTAS_(state) {
  PropertiesService.getScriptProperties().setProperty(
    'QTAS_FINANCIAL_REPAIR_STATE',
    JSON.stringify(state || {})
  );
}

function programarContinuacionReparacionIntegridadFinancieraQTAS_() {
  limpiarTriggersReparacionIntegridadFinancieraQTAS_();
  ScriptApp.newTrigger('continuarReparacionIntegridadFinancieraQTAS')
    .timeBased()
    .after(15000)
    .create();
}

function limpiarTriggersReparacionIntegridadFinancieraQTAS_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(trigger =>
      trigger.getHandlerFunction() === 'continuarReparacionIntegridadFinancieraQTAS'
    );
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  return triggers.length;
}

function resumirVerificacionIntegridadFinancieraQTAS_(verification) {
  return {
    ok: verification.ok === true,
    costoUnitarioInconsistente: numero_(verification.compras.costoUnitarioInconsistente),
    costosAgregadosComoUnitarios: numero_(verification.compras.costosAgregadosComoUnitarios),
    packagingSinCoberturaHistorica: numero_(verification.costos.packagingSinCoberturaHistorica),
    fondosDescuadrados: numero_(verification.compras.fondosDescuadrados),
    ventasDescuadradas: numero_(verification.ventas.totalesDescuadrados),
    comprasDescuadradas: numero_(verification.compras.totalesDescuadrados),
    ventaDetalleNoCanonico: numero_(verification.ventas.detalleNoCanonico),
    pagosDescuadrados: numero_(verification.ventas.pagosDescuadrados),
    costosAnaliticosExtremos: numero_(verification.ventas.costosAnaliticosExtremos),
    analiticaStale: numero_(verification.ventas.analiticaStale)
  };
}

function validarHojasIntegridadFinancieraQTAS_() {
  return validarModeloSoloLecturaQTAS_({
    sheetNames: [
      QTAS.sheets.compras,
      QTAS.sheets.compraDetalle,
      QTAS.sheets.compraOrigenesFondos,
      QTAS.sheets.costosReferencia,
      QTAS.sheets.costoProductoCalculado,
      QTAS.sheets.productoComponentes,
      QTAS.sheets.productoReglasCosto,
      QTAS.sheets.ventas,
      QTAS.sheets.detalle,
      QTAS.sheets.pagos,
      QTAS.sheets.ventaDetalleCostosCalculado
    ],
    validarConfig: false
  });
}

function construirPlanIntegridadFinancieraQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const compras = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.compras));
  const compraDetalle = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.compraDetalle));
  const fondos = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.compraOrigenesFondos));
  const costos = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.costosReferencia));
  const costoProducto = leerObjetos_(
    spreadsheet.getSheetByName(QTAS.sheets.costoProductoCalculado)
  );
  const ventas = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.ventas));
  const ventaDetalle = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.detalle));
  const pagos = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.pagos));
  const ventaDetalleCostos = leerObjetos_(
    spreadsheet.getSheetByName(QTAS.sheets.ventaDetalleCostosCalculado)
  );
  const compraDetalleProyectado = compraDetalle.map(row => Object.assign({}, row));
  const correccionesCosto = corregirCostoUnitarioCompraProyectadoQTAS_(compraDetalleProyectado);
  const correccionesImpactoCosto = corregirImpactoCostoAgregadoProyectadoQTAS_(
    compraDetalleProyectado
  );
  const planCostosDerivados = corregirCostosReferenciaDerivadosQTAS_(costos, correccionesCosto);
  const correccionesCoberturaPackaging = corregirCoberturaHistoricaPackagingQTAS_(
    planCostosDerivados.rows
  );
  const planFondos = corregirFondosCompraProyectadosQTAS_(compras, compraDetalle, fondos);
  const ventaDetalleProyectado = ventaDetalle.map(row => Object.assign({}, row));
  const correccionesVentaDetalle = normalizarVentaDetalleCanonicoProyectadoQTAS_(
    ventaDetalleProyectado
  );
  const periodosCosto = construirPeriodosCostoCorregidosQTAS_(
    correccionesCosto.concat(planCostosDerivados.correcciones),
    costos
  );
  const detalleIdsAnalitica = seleccionarDetalleAnaliticaAfectadoQTAS_(
    spreadsheet,
    ventaDetalle,
    periodosCosto
  );
  const ventaIds = {};
  const compraIds = {};
  const detalleVentaIds = {};

  ventas.forEach(row => {
    const id = texto_(row.Venta_ID);
    if (id) ventaIds[id] = row;
  });
  compras.forEach(row => {
    const id = texto_(row.Compra_ID);
    if (id) compraIds[id] = row;
  });
  ventaDetalle.forEach(row => {
    const id = texto_(row.Detalle_ID);
    if (id) detalleVentaIds[id] = true;
  });

  const ventasTotales = resumirDiferenciasTotalesQTAS_(
    ventas,
    ventaDetalle,
    'Venta_ID',
    'Total_Venta',
    'Subtotal_Neto'
  );
  const comprasTotales = resumirDiferenciasTotalesQTAS_(
    compras,
    compraDetalle,
    'Compra_ID',
    'Total_Compra',
    'Costo_Total_Linea'
  );
  const pagosDescuadrados = auditarPagosVentasQTAS_(ventas, pagos);
  const ventasAltas = ventas
    .filter(row =>
      !esRegistroAnulado_(row.Estado_Registro) &&
      numero_(row.Total_Venta) > 1000000
    )
    .map(row => ({
      ventaId: numero_(row.Venta_ID),
      fecha: fechaInput_(row.Fecha_Venta),
      cliente: texto_(row.Nombre),
      productos: texto_(row.Productos_Resumen),
      total: redondear_(numero_(row.Total_Venta))
    }))
    .sort((a, b) => b.total - a.total);
  const preciosAtipicos = auditarPreciosVentaAtipicosQTAS_(ventaDetalle);
  const costosAnaliticosExtremos = auditarCostosAnaliticosExtremosQTAS_(
    ventaDetalleCostos,
    costoProducto
  );
  const staleAnalitica = ventaDetalleCostos.filter(row =>
    texto_(row.Detalle_ID) && !detalleVentaIds[texto_(row.Detalle_ID)]
  );
  if (
    correccionesImpactoCosto.length ||
    correccionesCoberturaPackaging.length ||
    correccionesVentaDetalle.length ||
    costosAnaliticosExtremos.total
  ) {
    ventaDetalle.forEach(row => {
      const id = texto_(row.Detalle_ID);
      if (id) detalleIdsAnalitica[id] = true;
    });
  }

  return {
    compraDetalleProyectado: compraDetalleProyectado,
    ventaDetalleProyectado: ventaDetalleProyectado,
    costosReferenciaProyectados: planCostosDerivados.rows,
    fondosProyectados: planFondos.rows,
    correccionesCosto: correccionesCosto,
    correccionesImpactoCosto: correccionesImpactoCosto,
    correccionesCostoDerivado: planCostosDerivados.correcciones,
    correccionesCoberturaPackaging: correccionesCoberturaPackaging,
    correccionesFondos: planFondos.correcciones,
    correccionesVentaDetalle: correccionesVentaDetalle,
    detalleIdsAnalitica: detalleIdsAnalitica,
    resumen: {
      ok: correccionesCosto.length === 0 &&
        correccionesImpactoCosto.length === 0 &&
        planCostosDerivados.correcciones.length === 0 &&
        correccionesCoberturaPackaging.length === 0 &&
        planFondos.compras.length === 0 &&
        correccionesVentaDetalle.length === 0 &&
        costosAnaliticosExtremos.total === 0 &&
        pagosDescuadrados.length === 0 &&
        ventasTotales.length === 0 &&
        comprasTotales.length === 0 &&
        staleAnalitica.length === 0,
      compras: {
        filas: compras.length,
        detalle: compraDetalle.length,
        totalesDescuadrados: comprasTotales.length,
        costoUnitarioInconsistente: correccionesCosto.length,
        costosAgregadosComoUnitarios: correccionesImpactoCosto.length,
        costosReferenciaDerivados: planCostosDerivados.correcciones.length,
        fondosDescuadrados: planFondos.compras.length,
        correccionesCosto: correccionesCosto.map(item => ({
          compraDetalleId: item.compraDetalleId,
          compraId: item.compraId,
          item: item.item,
          cantidadAntes: item.cantidadAntes,
          cantidad: item.cantidad,
          costoTotal: item.costoTotal,
          costoUnitarioAntes: item.antes,
          costoUnitarioCorrecto: item.despues
        })),
        correccionesCostoDerivado: planCostosDerivados.correcciones.map(item => ({
          costoId: item.costoId,
          item: item.item,
          fechaDesde: fechaInput_(item.fechaCompra),
          costoUnitarioAntes: item.antes,
          costoUnitarioCorrecto: item.despues,
          fuenteId: item.compraDetalleId
        })),
        correccionesFondos: planFondos.compras,
        correccionesImpactoCosto: correccionesImpactoCosto
      },
      costos: {
        packagingSinCoberturaHistorica: correccionesCoberturaPackaging.length,
        correccionesCoberturaPackaging: correccionesCoberturaPackaging
      },
      ventas: {
        filas: ventas.length,
        detalle: ventaDetalle.length,
        totalesDescuadrados: ventasTotales.length,
        mayoresAUnMillon: ventasAltas.length,
        preciosAtipicos: preciosAtipicos.length,
        detalleNoCanonico: correccionesVentaDetalle.length,
        pagosDescuadrados: pagosDescuadrados.length,
        costosAnaliticosExtremos: costosAnaliticosExtremos.total,
        analiticaARecalcular: Object.keys(detalleIdsAnalitica).length,
        analiticaStale: staleAnalitica.length
      },
      pendientesRevisionManual: {
        ventasMayoresAUnMillon: ventasAltas,
        preciosVentaAtipicos: preciosAtipicos,
        ventaDetalleNoCanonico: correccionesVentaDetalle,
        costosAnaliticosExtremos: costosAnaliticosExtremos.rows,
        pagosDescuadrados: pagosDescuadrados,
        ventasTotalesDescuadrados: ventasTotales,
        comprasTotalesDescuadrados: comprasTotales,
        analiticaStale: staleAnalitica.map(row => ({
          detalleCostoId: texto_(row.Detalle_Costo_ID),
          detalleId: texto_(row.Detalle_ID),
          ventaId: numero_(row.Venta_ID)
        }))
      }
    }
  };
}

function corregirCostoUnitarioCompraProyectadoQTAS_(rows) {
  const correcciones = [];
  const cantidadesHistoricas = {
    'COMDET-000031-01': 2000,
    'COMDET-000083-01': 5000,
    'COMDET-000088-01': 105,
    'COMDET-000116-01': 100
  };

  (rows || []).forEach(row => {
    const compraDetalleId = texto_(row.Compra_Detalle_ID);
    const cantidadAntes = numero_(row.Cantidad);
    const cantidadHistorica = numero_(cantidadesHistoricas[compraDetalleId]);
    const cantidad = cantidadHistorica > 0 ? cantidadHistorica : cantidadAntes;
    const costoTotal = numero_(row.Costo_Total_Linea);
    if (cantidad <= 0 || costoTotal <= 0 || esRegistroAnulado_(row.Estado_Registro)) return;

    const esperado = redondear_(costoTotal / cantidad);
    const actual = redondear_(numero_(row.Costo_Unitario));
    const cambiaCantidad = Math.abs(cantidadAntes - cantidad) > 0.0001;
    if (!cambiaCantidad && Math.abs(actual - esperado) <= 0.01) return;

    correcciones.push({
      compraDetalleId: compraDetalleId,
      compraId: numero_(row.Compra_ID),
      fechaCompra: row.Fecha_Compra,
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
      item: texto_(row.Item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      cantidadAntes: redondear_(cantidadAntes),
      cantidad: redondear_(cantidad),
      costoTotal: redondear_(costoTotal),
      antes: actual,
      despues: esperado
    });
    row.Cantidad = cantidad;
    row.Costo_Unitario = esperado;
    if (cambiaCantidad) {
      row.Comentario_Linea = unirUnicos_([
        texto_(row.Comentario_Linea),
        `Cantidad historica confirmada: ${cantidad} ${texto_(row.Unidad)}.`
      ]);
    }
  });

  return correcciones;
}

function corregirImpactoCostoAgregadoProyectadoQTAS_(rows) {
  const correcciones = [];
  const idsAgregados = {
    'COMDET-000038-01': 'Compra agregada de materia prima; no representa una unidad terminada de Choco.'
  };

  (rows || []).forEach(row => {
    const compraDetalleId = texto_(row.Compra_Detalle_ID);
    const motivo = idsAgregados[compraDetalleId];
    const impactaCosto = row.Impacta_Costo === true ||
      String(row.Impacta_Costo).toLowerCase() === 'true';
    if (!motivo || !impactaCosto) return;

    row.Impacta_Costo = false;
    row.Comentario_Linea = unirUnicos_([texto_(row.Comentario_Linea), motivo]);
    correcciones.push({
      compraDetalleId: compraDetalleId,
      compraId: numero_(row.Compra_ID),
      item: texto_(row.Item),
      costoUnitarioDescartado: redondear_(numero_(row.Costo_Unitario)),
      motivo: motivo
    });
  });

  return correcciones;
}

function corregirCoberturaHistoricaPackagingQTAS_(rows) {
  const correcciones = [];
  const fechaBase = resolverFechaOperacion_('2024-01-01', new Date());

  (rows || []).forEach(row => {
    const esPlaceholder = numero_(row.Compra_ID) === 0 &&
      normalizarClaveTexto_(row.Fuente_Tipo) === normalizarClaveTexto_('Directo') &&
      normalizarClaveTexto_(row.Proveedor) === normalizarClaveTexto_('Placeholder packaging Psylo Scibio');
    if (!esPlaceholder || fechaTextoPlanoQTAS_(row.Fecha_Desde) === '2024-01-01') return;

    correcciones.push({
      costoId: texto_(row.Costo_ID),
      item: texto_(row.Item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      fechaDesdeAntes: fechaTextoPlanoQTAS_(row.Fecha_Desde),
      fechaDesde: '2024-01-01',
      costoUnitario: redondear_(numero_(row.Costo_Unitario))
    });
    row.Fecha_Desde = fechaBase;
    row.Fecha_Hasta = '';
  });

  return correcciones;
}

function normalizarVentaDetalleCanonicoProyectadoQTAS_(rows) {
  const correcciones = [];

  (rows || []).forEach(row => {
    if (!texto_(row.Detalle_ID)) return;
    const contexto = resolverContextoCostoVentaLegacyQTAS_(row);
    const productoAntes = texto_(row.Producto_Estandar);
    const unidadAntes = normalizarUnidadCanonicaQTAS_(row.Unidad);
    if (productoAntes === contexto.producto && unidadAntes === contexto.unidad) return;

    row.Producto_Estandar = contexto.producto;
    row.Unidad = contexto.unidad;
    correcciones.push({
      detalleId: texto_(row.Detalle_ID),
      ventaId: numero_(row.Venta_ID),
      productoAntes: productoAntes,
      producto: contexto.producto,
      unidadAntes: unidadAntes,
      unidad: contexto.unidad,
      cantidad: redondear_(numero_(row.Cantidad))
    });
  });

  return correcciones;
}

function corregirCostosReferenciaDerivadosQTAS_(costos, correccionesCompra) {
  const rows = (costos || []).map(row => Object.assign({}, row));
  const correcciones = [];

  (correccionesCompra || []).forEach(correccionCompra => {
    const keyCompra = claveCostoHistoricoQTAS_(
      correccionCompra.tipoItem,
      correccionCompra.item,
      correccionCompra.unidad
    );

    rows.forEach(row => {
      const esBackfill = normalizarClaveTexto_(row.Fuente_Tipo) ===
        normalizarClaveTexto_('AjusteHistorico') &&
        normalizarClaveTexto_(row.Fuente_ID).indexOf('backfillinicial') >= 0;
      const mismaClave = claveCostoHistoricoQTAS_(row.Tipo_Item, row.Item, row.Unidad) === keyCompra;
      const mismoCostoErrado = Math.abs(
        redondear_(numero_(row.Costo_Unitario)) - correccionCompra.antes
      ) <= 0.01;
      if (!esBackfill || !mismaClave || !mismoCostoErrado) return;

      const antes = redondear_(numero_(row.Costo_Unitario));
      row.Costo_Unitario = correccionCompra.despues;
      row.Nota = unirUnicos_([
        texto_(row.Nota),
        `Costo inicial corregido desde ${correccionCompra.compraDetalleId}: ${correccionCompra.despues}.`
      ]);
      correcciones.push({
        costoId: texto_(row.Costo_ID),
        compraDetalleId: texto_(row.Fuente_ID),
        compraId: 0,
        fechaCompra: row.Fecha_Desde,
        tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
        item: texto_(row.Item),
        unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
        antes: antes,
        despues: correccionCompra.despues
      });
    });
  });

  return {
    rows: rows,
    correcciones: correcciones
  };
}

function corregirFondosCompraProyectadosQTAS_(compras, detalle, fondos) {
  const comprasPorId = {};
  const detallePorId = {};
  const indicesPorCompra = {};
  const rows = (fondos || []).map(row => Object.assign({}, row));
  const correcciones = [];
  const comprasCorregidas = [];

  (compras || []).forEach(row => {
    const key = texto_(row.Compra_ID);
    if (key) comprasPorId[key] = row;
  });
  (detalle || []).forEach(row => {
    const key = texto_(row.Compra_Detalle_ID);
    if (key) detallePorId[key] = row;
  });
  rows.forEach((row, index) => {
    const key = texto_(row.Compra_ID);
    if (!key) return;
    if (!indicesPorCompra[key]) indicesPorCompra[key] = [];
    indicesPorCompra[key].push(index);
  });

  Object.keys(indicesPorCompra).forEach(compraKey => {
    const compra = comprasPorId[compraKey];
    const indices = indicesPorCompra[compraKey];
    if (!compra || !indices.length || esRegistroAnulado_(compra.Estado_Registro)) return;

    const totalCompra = redondear_(numero_(compra.Total_Compra));
    const totalAntes = redondear_(sumar_(indices.map(index => rows[index].Monto_Asignado)));
    if (totalCompra <= 0 || Math.abs(totalAntes - totalCompra) <= 0.01) return;

    const pesosAportantes = {};
    indices.forEach(index => {
      const row = rows[index];
      const aportante = normalizarAportanteOrigenFondosQTAS_(row.Aportante) || texto_(row.Aportante);
      if (!aportante) return;
      pesosAportantes[aportante] = Math.max(
        numero_(pesosAportantes[aportante]),
        numero_(row.Monto_Asignado)
      );
    });
    const aportantes = Object.keys(pesosAportantes)
      .filter(aportante => numero_(pesosAportantes[aportante]) > 0)
      .sort((a, b) => ordenAportanteOrigenFondosQTAS_(a) - ordenAportanteOrigenFondosQTAS_(b));
    const totalPesos = sumar_(aportantes.map(aportante => pesosAportantes[aportante]));
    if (!aportantes.length || totalPesos <= 0) return;

    const metasAportantes = distribuirTotalPorPesosQTAS_(
      totalCompra,
      aportantes.map(aportante => ({
        key: aportante,
        peso: pesosAportantes[aportante]
      }))
    );

    aportantes.forEach(aportante => {
      const indicesAportante = indices.filter(index =>
        (normalizarAportanteOrigenFondosQTAS_(rows[index].Aportante) || texto_(rows[index].Aportante)) === aportante
      );
      const asignaciones = distribuirTotalPorPesosQTAS_(
        metasAportantes[aportante],
        indicesAportante.map(index => ({
          key: String(index),
          peso: numero_(detallePorId[texto_(rows[index].Compra_Detalle_ID)] &&
            detallePorId[texto_(rows[index].Compra_Detalle_ID)].Costo_Total_Linea)
        }))
      );

      indicesAportante.forEach(index => {
        const row = rows[index];
        const detalleRow = detallePorId[texto_(row.Compra_Detalle_ID)];
        const totalLinea = numero_(detalleRow && detalleRow.Costo_Total_Linea);
        const antes = redondear_(numero_(row.Monto_Asignado));
        const despues = redondear_(numero_(asignaciones[String(index)]));
        row.Monto_Asignado = despues;
        row.Porcentaje = totalLinea > 0
          ? redondear_(despues * 100 / totalLinea)
          : 0;
        row.Nota = unirUnicos_([
          texto_(row.Nota),
          'Integridad financiera historica corregida.'
        ]);
        correcciones.push({
          compraOrigenId: texto_(row.Compra_Origen_ID),
          compraId: numero_(row.Compra_ID),
          aportante: aportante,
          antes: antes,
          despues: despues
        });
      });
    });

    const totalDespues = redondear_(sumar_(indices.map(index => rows[index].Monto_Asignado)));
    comprasCorregidas.push({
      compraId: numero_(compraKey),
      totalCompra: totalCompra,
      fondosAntes: totalAntes,
      fondosProyectados: totalDespues,
      filas: indices.length
    });
  });

  return {
    rows: rows,
    correcciones: correcciones,
    compras: comprasCorregidas
  };
}

function distribuirTotalPorPesosQTAS_(total, entries) {
  const validas = (entries || []).filter(entry => texto_(entry.key) && numero_(entry.peso) > 0);
  const totalPesos = sumar_(validas.map(entry => entry.peso));
  const result = {};
  let asignado = 0;

  validas.forEach((entry, index) => {
    const monto = index === validas.length - 1
      ? redondear_(numero_(total) - asignado)
      : redondear_(numero_(total) * numero_(entry.peso) / totalPesos);
    result[texto_(entry.key)] = monto;
    asignado = redondear_(asignado + monto);
  });

  return result;
}

function construirPeriodosCostoCorregidosQTAS_(correcciones, costos) {
  const periodos = {};

  (correcciones || []).forEach(correccion => {
    const key = claveCostoHistoricoQTAS_(
      correccion.tipoItem,
      correccion.item,
      correccion.unidad
    );
    if (!key) return;

    const candidatas = (costos || [])
      .filter(row => claveCostoHistoricoQTAS_(row.Tipo_Item, row.Item, row.Unidad) === key)
      .sort((a, b) =>
        resolverFechaOperacion_(a.Fecha_Desde, new Date()) -
        resolverFechaOperacion_(b.Fecha_Desde, new Date())
      );
    const exacta = candidatas.find(row =>
      texto_(row.Fuente_ID) === correccion.compraDetalleId
    ) || candidatas.find(row =>
      numero_(row.Compra_ID) === correccion.compraId &&
      fechaInput_(row.Fecha_Desde) === fechaInput_(correccion.fechaCompra)
    );
    const desde = exacta
      ? resolverFechaOperacion_(exacta.Fecha_Desde, correccion.fechaCompra)
      : resolverFechaOperacion_(correccion.fechaCompra, new Date());
    const siguiente = candidatas.find(row =>
      resolverFechaOperacion_(row.Fecha_Desde, new Date()) > desde
    );
    const hasta = exacta && exacta.Fecha_Hasta
      ? resolverFechaOperacion_(exacta.Fecha_Hasta, desde)
      : (siguiente ? diaAnterior_(siguiente.Fecha_Desde) : null);

    if (!periodos[key]) periodos[key] = [];
    periodos[key].push({ desde: desde, hasta: hasta });
  });

  return periodos;
}

function seleccionarDetalleAnaliticaAfectadoQTAS_(ss, detalleRows, periodosPorClave) {
  const keysAfectadas = Object.keys(periodosPorClave || {});
  if (!keysAfectadas.length) return {};

  const componentes = leerComponentesProductoActivosQTAS_();
  const reglas = leerReglasCostoProductoActivasQTAS_();
  const memo = {};
  const seleccionados = {};

  (detalleRows || []).forEach(row => {
    const contexto = resolverContextoCostoVentaLegacyQTAS_(row);
    const dependencias = construirDependenciasCostoProductoQTAS_(
      contexto.producto,
      contexto.unidad,
      componentes,
      reglas,
      memo,
      {}
    );
    const fechaVenta = valorFechaVentaCanonicaQTAS_(row, new Date());
    const afectado = keysAfectadas.some(key =>
      dependencias[key] &&
      periodosPorClave[key].some(periodo =>
        fechaVenta >= periodo.desde && (!periodo.hasta || fechaVenta <= periodo.hasta)
      )
    );
    if (afectado && texto_(row.Detalle_ID)) {
      seleccionados[texto_(row.Detalle_ID)] = true;
    }
  });

  return seleccionados;
}

function construirDependenciasCostoProductoQTAS_(producto, unidad, componentes, reglas, memo, stack) {
  const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);
  const memoKey = normalizarClaveProductoQTAS_(producto, unidadCanonica);
  if (memo[memoKey]) return memo[memoKey];

  const dependencias = {};
  const directa = claveCostoHistoricoQTAS_('Producto', producto, unidadCanonica);
  if (directa) dependencias[directa] = true;
  if (stack[memoKey]) return dependencias;

  const nextStack = Object.assign({}, stack);
  nextStack[memoKey] = true;
  const partidas = (componentes || []).concat(reglas || []).filter(row =>
    esMismaClaveProductoQTAS_(
      row.producto,
      row.unidadVenta,
      producto,
      unidadCanonica
    )
  );

  partidas.forEach(row => {
    const key = claveCostoHistoricoQTAS_(
      row.tipoComponente,
      row.itemComponente,
      row.unidadComponente
    );
    if (key) dependencias[key] = true;

    if (normalizarClaveTexto_(row.tipoComponente) === normalizarClaveTexto_('Producto')) {
      const nested = construirDependenciasCostoProductoQTAS_(
        row.itemComponente,
        row.unidadComponente,
        componentes,
        reglas,
        memo,
        nextStack
      );
      Object.keys(nested).forEach(nestedKey => {
        dependencias[nestedKey] = true;
      });
    }
  });

  memo[memoKey] = dependencias;
  return dependencias;
}

function sincronizarAnaliticaIntegridadFinancieraQTAS_(ss, detalleIds) {
  const detalleSheet = ss.getSheetByName(QTAS.sheets.detalle);
  const outputSheet = ss.getSheetByName(QTAS.sheets.ventaDetalleCostosCalculado);
  const detalle = leerObjetos_(detalleSheet).filter(row => texto_(row.Detalle_ID));
  const existentes = leerObjetos_(outputSheet);
  const porDetalle = {};
  const idsFuente = {};
  let recalculadas = 0;
  let creadas = 0;
  const recalcularTodo = detalleIds === null || detalleIds === undefined;

  existentes.forEach(row => {
    const id = texto_(row.Detalle_ID);
    if (id && !porDetalle[id]) porDetalle[id] = row;
  });
  detalle.forEach(row => {
    idsFuente[texto_(row.Detalle_ID)] = true;
  });

  const contexto = construirContextoAnaliticaCostosQTAS_({
    ss: ss,
    ahora: new Date()
  });
  const rows = detalle.map(row => {
    const id = texto_(row.Detalle_ID);
    if (recalcularTodo || detalleIds[id] || !porDetalle[id]) {
      if (porDetalle[id]) recalculadas++;
      else creadas++;
      return construirFilaVentaDetalleCostoCalculadoQTAS_(row, contexto);
    }
    return porDetalle[id];
  });
  const stale = existentes.filter(row =>
    texto_(row.Detalle_ID) && !idsFuente[texto_(row.Detalle_ID)]
  ).length;

  sobrescribirObjetosHojaQTAS_(outputSheet, getHeaders_(outputSheet), rows);
  return {
    ok: true,
    rows: rows.length,
    recalculadas: recalculadas,
    creadas: creadas,
    stale: stale
  };
}

function resumirDiferenciasTotalesQTAS_(cabeceras, detalles, idHeader, totalHeader, lineTotalHeader) {
  const sumas = {};
  (detalles || []).forEach(row => {
    const key = texto_(row[idHeader]);
    if (!key || esRegistroAnulado_(row.Estado_Registro)) return;
    sumas[key] = redondear_(numero_(sumas[key]) + numero_(row[lineTotalHeader]));
  });

  return (cabeceras || [])
    .filter(row => !esRegistroAnulado_(row.Estado_Registro))
    .map(row => {
      const key = texto_(row[idHeader]);
      const total = redondear_(numero_(row[totalHeader]));
      const detalle = redondear_(numero_(sumas[key]));
      return {
        id: numero_(key) || key,
        total: total,
        detalle: detalle,
        diferencia: redondear_(total - detalle)
      };
    })
    .filter(row => Math.abs(row.diferencia) > 0.01);
}

function auditarPagosVentasQTAS_(ventas, pagos) {
  const pagosPorVenta = {};
  (pagos || []).forEach(row => {
    if (esRegistroAnulado_(row.Estado_Registro)) return;
    const key = texto_(row.Venta_ID);
    pagosPorVenta[key] = redondear_(numero_(pagosPorVenta[key]) + numero_(row.Monto_Pago));
  });

  return (ventas || [])
    .filter(row => !esRegistroAnulado_(row.Estado_Registro))
    .map(row => {
      const key = texto_(row.Venta_ID);
      const registrado = redondear_(numero_(row.Total_Pagado));
      const comprobado = redondear_(numero_(pagosPorVenta[key]));
      return {
        ventaId: numero_(key),
        fecha: fechaInput_(row.Fecha_Venta),
        cliente: texto_(row.Nombre),
        totalVenta: redondear_(numero_(row.Total_Venta)),
        totalPagadoVenta: registrado,
        totalPagos: comprobado,
        diferencia: redondear_(registrado - comprobado)
      };
    })
    .filter(row => Math.abs(row.diferencia) > 0.01);
}

function auditarPreciosVentaAtipicosQTAS_(detalle) {
  const preciosPorProducto = {};
  (detalle || []).forEach(row => {
    const key = normalizarClaveTexto_(row.Producto_Estandar);
    const precio = numero_(row.Precio_Vendido_Unitario);
    if (!key || precio <= 0 || esRegistroAnulado_(row.Estado_Registro)) return;
    if (!preciosPorProducto[key]) preciosPorProducto[key] = [];
    preciosPorProducto[key].push(precio);
  });

  const medianas = {};
  Object.keys(preciosPorProducto).forEach(key => {
    const values = preciosPorProducto[key].slice().sort((a, b) => a - b);
    if (values.length < 5) return;
    const middle = Math.floor(values.length / 2);
    medianas[key] = values.length % 2
      ? values[middle]
      : (values[middle - 1] + values[middle]) / 2;
  });

  return (detalle || [])
    .map(row => {
      const key = normalizarClaveTexto_(row.Producto_Estandar);
      const mediana = numero_(medianas[key]);
      const precio = numero_(row.Precio_Vendido_Unitario);
      const ratio = mediana > 0 ? precio / mediana : 0;
      return {
        detalleId: texto_(row.Detalle_ID),
        ventaId: numero_(row.Venta_ID),
        fecha: fechaInput_(row.Fecha_Venta),
        producto: texto_(row.Producto_Estandar),
        cantidad: redondear_(numero_(row.Cantidad)),
        unidad: texto_(row.Unidad),
        precioUnitario: redondear_(precio),
        medianaProducto: redondear_(mediana),
        vecesMediana: redondear_(ratio),
        descuento: redondear_(numero_(row.Descuento_Linea)),
        subtotalNeto: redondear_(numero_(row.Subtotal_Neto))
      };
    })
    .filter(row => row.vecesMediana >= 5)
    .sort((a, b) => b.vecesMediana - a.vecesMediana)
    .slice(0, 30);
}

function auditarCostosAnaliticosExtremosQTAS_(rows, costosProducto) {
  const costoActualPorProducto = {};
  (costosProducto || []).forEach(row => {
    const key = normalizarClaveProductoQTAS_(row.Producto_Estandar, row.Unidad_Venta);
    const costo = numero_(row.Costo_Unitario_Total);
    if (key && costo > 0) costoActualPorProducto[key] = costo;
  });

  const atipicos = (rows || [])
    .map(row => ({
      detalleId: texto_(row.Detalle_ID),
      ventaId: numero_(row.Venta_ID),
      fecha: fechaInput_(row.Fecha_Venta),
      producto: texto_(row.Producto_Estandar),
      cantidad: redondear_(numero_(row.Cantidad)),
      unidad: texto_(row.Unidad),
      costoUnitario: redondear_(numero_(row.Costo_Unitario_Usado)),
      costoTotal: redondear_(numero_(row.Costo_Total_Estimado)),
      subtotalNeto: redondear_(numero_(row.Subtotal_Neto)),
      actualizadoEn: texto_(row.Actualizado_En),
      costoActualReferencia: redondear_(numero_(costoActualPorProducto[
        normalizarClaveProductoQTAS_(row.Producto_Estandar, row.Unidad)
      ]))
    }))
    .map(row => Object.assign({}, row, {
      vecesCostoActual: row.costoActualReferencia > 0
        ? redondear_(row.costoUnitario / row.costoActualReferencia)
        : 0,
      vecesVenta: row.subtotalNeto > 0
        ? redondear_(row.costoTotal / row.subtotalNeto)
        : 0
    }))
    .filter(row =>
      row.costoUnitario >= 1000000 ||
      row.vecesVenta >= 5
    )
    .sort((a, b) => {
      const severidadA = Math.max(a.vecesCostoActual, a.vecesVenta);
      const severidadB = Math.max(b.vecesCostoActual, b.vecesVenta);
      return severidadB - severidadA;
    });

  return {
    total: atipicos.length,
    rows: atipicos.slice(0, 50)
  };
}
