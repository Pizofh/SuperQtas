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

function repararIntegridadFinancieraQTAS() {
  assertOperacionDestructivaPermitidaQTAS_('reparar integridad financiera historica');

  return withScriptLock_('reparar integridad financiera historica', () => {
    const ss = validarHojasIntegridadFinancieraQTAS_();
    const plan = construirPlanIntegridadFinancieraQTAS_(ss);
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
    const ventaDetalleCostos = sincronizarAnaliticaIntegridadFinancieraQTAS_(
      ss,
      plan.detalleIdsAnalitica
    );
    limpiarCachesEjecucionQTAS_();

    const verificacion = construirPlanIntegridadFinancieraQTAS_(ss).resumen;
    const result = {
      ok: verificacion.ok === true,
      aplicado: true,
      cambios: {
        compraDetalle: plan.correccionesCosto.length,
        comprasSinCostoUnitarioFalso: plan.correccionesImpactoCosto.length,
        costosReferenciaDerivados: plan.correccionesCostoDerivado.length,
        costosPackagingConHistoria: plan.correccionesCoberturaPackaging.length,
        compraOrigenesFondos: plan.correccionesFondos.length,
        comprasConFondosCorregidos: plan.resumen.compras.fondosDescuadrados,
        ventaDetalleCanonico: plan.correccionesVentaDetalle.length,
        analiticaVentaRecalculada: ventaDetalleCostos.recalculadas,
        analiticaVentaCreada: ventaDetalleCostos.creadas,
        analiticaVentaStaleEliminada: ventaDetalleCostos.stale
      },
      reconstrucciones: {
        costosReferencia: costos,
        costoProducto: costoProducto,
        ventaDetalleCostos: ventaDetalleCostos
      },
      pendientesRevisionManual: verificacion.pendientesRevisionManual,
      verificacion: {
        costoUnitarioInconsistente: verificacion.compras.costoUnitarioInconsistente,
        costosAgregadosComoUnitarios: verificacion.compras.costosAgregadosComoUnitarios,
        packagingSinCoberturaHistorica: verificacion.costos.packagingSinCoberturaHistorica,
        fondosDescuadrados: verificacion.compras.fondosDescuadrados,
        ventasDescuadradas: verificacion.ventas.totalesDescuadrados,
        comprasDescuadradas: verificacion.compras.totalesDescuadrados,
        ventaDetalleNoCanonico: verificacion.ventas.detalleNoCanonico,
        costosAnaliticosExtremos: verificacion.ventas.costosAnaliticosExtremos
      }
    };

    Logger.log(JSON.stringify(result, null, 2));
    return result;
  });
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
  const costosAnaliticosExtremos = auditarCostosAnaliticosExtremosQTAS_(ventaDetalleCostos);
  const staleAnalitica = ventaDetalleCostos.filter(row =>
    texto_(row.Detalle_ID) && !detalleVentaIds[texto_(row.Detalle_ID)]
  );
  if (
    correccionesImpactoCosto.length ||
    correccionesCoberturaPackaging.length ||
    correccionesVentaDetalle.length ||
    costosAnaliticosExtremos.length
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
        costosAnaliticosExtremos.length === 0 &&
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
        costosAnaliticosExtremos: costosAnaliticosExtremos.length,
        analiticaARecalcular: Object.keys(detalleIdsAnalitica).length,
        analiticaStale: staleAnalitica.length
      },
      pendientesRevisionManual: {
        ventasMayoresAUnMillon: ventasAltas,
        preciosVentaAtipicos: preciosAtipicos,
        ventaDetalleNoCanonico: correccionesVentaDetalle,
        costosAnaliticosExtremos: costosAnaliticosExtremos,
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

function auditarCostosAnaliticosExtremosQTAS_(rows) {
  return (rows || [])
    .filter(row => numero_(row.Costo_Unitario_Usado) >= 1000000)
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
      actualizadoEn: texto_(row.Actualizado_En)
    }))
    .sort((a, b) => b.costoUnitario - a.costoUnitario)
    .slice(0, 50);
}
