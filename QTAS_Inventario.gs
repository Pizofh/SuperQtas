function getDashboardInventarioQTAS() {
  asegurarModeloOperativoQTAS_();
  asegurarControlesInventarioBaseQTAS_();

  const ss = SpreadsheetApp.getActive();
  const movimientosRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioMovimientos);
  const snapshotRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioSnapshot);

  if (movimientosRef.ok && snapshotRef.ok && movimientosRef.sheet.getLastRow() <= 1 && snapshotRef.sheet.getLastRow() <= 1) {
    reconstruirInventarioInternoQTAS_({ ss: ss });
  }

  const controles = listarControlesInventarioQTAS_();
  const stock = listarSnapshotInventarioQTAS_();
  const alertas = stock
    .filter(row => row.estadoStock !== 'OK')
    .sort((a, b) => {
      const prioridad = prioridadEstadoStockInventarioQTAS_(a.estadoStock) - prioridadEstadoStockInventarioQTAS_(b.estadoStock);
      if (prioridad !== 0) return prioridad;
      return a.item.localeCompare(b.item);
    });
  const movimientosRecientes = listarMovimientosInventarioRecientesQTAS_();
  const produccionesRecientes = listarProduccionesRecientesQTAS_();
  const fabricados = controles
    .filter(row =>
      row.activo &&
      row.tipoItem === 'Producto' &&
      row.modoStock === 'Fabricado'
    )
    .sort((a, b) => a.item.localeCompare(b.item))
    .map(row => ({
      controlId: row.controlId,
      producto: row.item,
      unidad: row.unidad,
      stockMinimo: row.stockMinimo,
      stockObjetivo: row.stockObjetivo,
      nota: row.nota
    }));

  return {
    ok: true,
    hoy: fechaInput_(new Date()),
    resumen: {
      itemsActivos: stock.length,
      alertas: alertas.length,
      sinStock: stock.filter(row => row.estadoStock === 'Sin stock').length,
      fabricados: fabricados.length,
      controlesPendientes: controles.filter(row =>
        row.activo &&
        row.modoStock !== 'NoControlado' &&
        numero_(row.stockMinimo) <= 0 &&
        numero_(row.stockObjetivo) <= 0
      ).length
    },
    stockInicialPendiente: stock.some(row => numero_(row.stockActual) < 0),
    alertas: alertas.slice(0, 40),
    stock: stock,
    movimientosRecientes: movimientosRecientes,
    produccionesRecientes: produccionesRecientes,
    productosFabricados: fabricados
  };
}

function guardarControlInventarioQTAS(payload) {
  return withScriptLock_('guardar control inventario', () => {
    asegurarModeloOperativoQTAS_();
    asegurarControlesInventarioBaseQTAS_();

    const tipoItem = normalizarTipoCompraItemQTAS_(payload && payload.tipoItem);
    const item = texto_(payload && payload.item);
    const unidad = normalizarUnidadCanonicaQTAS_(payload && payload.unidad);
    const modoStock = normalizarModoStockInventarioQTAS_(payload && payload.modoStock);
    const stockMinimo = redondear_(Math.max(0, numero_(payload && payload.stockMinimo)));
    const stockObjetivo = redondear_(Math.max(0, numero_(payload && payload.stockObjetivo)));
    const activo = !payload || payload.activo !== false;
    const nota = texto_(payload && payload.nota);
    const controlIdBuscado = texto_(payload && payload.controlId);

    if (!item) throw new Error('Falta el item.');
    if (!unidad) throw new Error('Falta la unidad.');
    if (!modoStock) throw new Error('Falta el modo de stock.');

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.inventarioControl);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const existente = controlIdBuscado
      ? rows.find(row => texto_(row.Control_ID) === controlIdBuscado)
      : rows.find(row =>
        normalizarClaveTexto_(row.Tipo_Item) === normalizarClaveTexto_(tipoItem) &&
        normalizarClaveTexto_(row.Item) === normalizarClaveTexto_(item) &&
        normalizarClaveTexto_(row.Unidad) === normalizarClaveTexto_(unidad)
      );

    const normalizado = {
      Control_ID: existente
        ? texto_(existente.Control_ID)
        : siguienteIdConPrefijo_(sheet, 'Control_ID', 'INVCTRL-', 4),
      Tipo_Item: tipoItem,
      Item: item,
      Unidad: unidad,
      Modo_Stock: modoStock,
      Stock_Minimo: stockMinimo,
      Stock_Objetivo: stockObjetivo,
      Activo: activo,
      Nota: nota
    };

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, normalizado));
    } else {
      escribirFilas_(sheet, [filaDesdeHeaders_(headers, normalizado)]);
    }

    reconstruirSnapshotInventarioQTAS_({ ss: ss });

    return {
      ok: true,
      controlId: normalizado.Control_ID,
      dashboard: getDashboardInventarioQTAS()
    };
  });
}

function registrarProduccionQTAS(payload) {
  return withScriptLock_('registrar produccion', () => {
    asegurarModeloOperativoQTAS_();
    asegurarControlesInventarioBaseQTAS_();

    const producto = texto_(payload && payload.producto);
    const unidad = normalizarUnidadCanonicaQTAS_(payload && payload.unidad);
    const cantidad = redondear_(Math.max(0, numero_(payload && payload.cantidad)));
    const comentario = texto_(payload && payload.comentarioProduccion);
    const ahora = new Date();
    const fechaProduccionBase = resolverFechaOperacion_(payload && payload.fechaProduccion, ahora);
    const fechaProduccion = combinarFechaYHora_(fechaProduccionBase, ahora);

    if (!producto) throw new Error('Falta el producto fabricado.');
    if (!unidad) throw new Error('Falta la unidad del producto.');
    if (cantidad <= 0) throw new Error('La cantidad fabricada debe ser mayor a cero.');

    const control = obtenerControlInventarioEfectivoQTAS_('Producto', producto, unidad, construirIndiceControlesInventarioQTAS_(listarControlesInventarioQTAS_()));
    if (control.modoStock !== 'Fabricado') {
      throw new Error(`El producto ${producto} no esta marcado como Fabricado en Inventario_Control.`);
    }

    const componentes = leerComponentesProductoActivosQTAS_();
    const reglas = leerReglasCostoProductoActivasQTAS_();
    const receta = componentes.filter(row =>
      esMismaClaveProductoQTAS_(row.producto, row.unidadVenta, producto, unidad)
    );

    if (!receta.length) {
      throw new Error(`No existe receta activa para ${producto} (${unidad}).`);
    }

    const ss = SpreadsheetApp.getActive();
    const produccionesSheet = ss.getSheetByName(QTAS.sheets.producciones);
    const produccionDetalleSheet = ss.getSheetByName(QTAS.sheets.produccionDetalle);
    const produccionesHeaders = getHeaders_(produccionesSheet);
    const detalleHeaders = getHeaders_(produccionDetalleSheet);
    const produccionId = siguienteIdConPrefijoPersistenteQTAS_(
      'produccion_id',
      produccionesSheet,
      'Produccion_ID',
      'PRD-',
      6
    );

    const detallesMaterializados = construirDetalleProduccionMaterializadoQTAS_({
      produccionId: produccionId,
      producto: producto,
      unidad: unidad,
      cantidad: cantidad,
      fechaProduccion: fechaProduccionBase,
      comentario: comentario,
      componentes: componentes,
      reglas: reglas
    });

    escribirFilas_(produccionesSheet, [filaDesdeHeaders_(produccionesHeaders, {
      Produccion_ID: produccionId,
      Fecha_Produccion: fechaProduccion,
      Producto_Estandar: producto,
      Cantidad_Producida: cantidad,
      Unidad: unidad,
      Comentario_Produccion: comentario,
      Estado_Registro: QTAS.status.registro.activo
    })]);

    escribirFilas_(
      produccionDetalleSheet,
      detallesMaterializados.map(row => filaDesdeHeaders_(detalleHeaders, row))
    );

    const inventario = sincronizarInventarioDesdeProduccionDetalleQTAS_({
      ss: ss,
      produccionId: produccionId,
      fechaProduccion: fechaProduccionBase,
      detalleRows: detallesMaterializados
    });

    return {
      ok: true,
      produccionId: produccionId,
      producto: producto,
      cantidad: cantidad,
      unidad: unidad,
      detalleLineas: detallesMaterializados.length,
      inventario: inventario,
      dashboard: getDashboardInventarioQTAS()
    };
  });
}

function reconstruirInventarioInternoQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();

  asegurarControlesInventarioBaseQTAS_();

  const movimientosRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioMovimientos);
  const snapshotRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioSnapshot);
  if (!movimientosRef.ok || !snapshotRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: unirUnicos_([movimientosRef.reason, snapshotRef.reason]),
      movimientos: 0,
      snapshot: 0
    };
  }

  const movimientos = []
    .concat(construirMovimientosInventarioDesdeComprasQTAS_({ ss: ss }))
    .concat(construirMovimientosInventarioDesdeProduccionesQTAS_({ ss: ss }))
    .concat(construirMovimientosInventarioDesdeVentasQTAS_({ ss: ss }))
    .sort((a, b) => {
      const fechaA = resolverFechaOperacion_(a.Fecha_Movimiento, new Date());
      const fechaB = resolverFechaOperacion_(b.Fecha_Movimiento, new Date());
      if (fechaA.getTime() !== fechaB.getTime()) return fechaA - fechaB;
      if (texto_(a.Fuente_Tipo) !== texto_(b.Fuente_Tipo)) {
        return texto_(a.Fuente_Tipo).localeCompare(texto_(b.Fuente_Tipo));
      }
      return texto_(a.Fuente_ID).localeCompare(texto_(b.Fuente_ID));
    });

  sobrescribirObjetosHojaQTAS_(movimientosRef.sheet, movimientosRef.headers, asignarIdsMovimientosInventarioQTAS_(movimientos));
  const snapshotRows = reconstruirSnapshotInventarioQTAS_({ ss: ss });
  limpiarCachesEjecucionQTAS_();

  return {
    ok: true,
    skipped: false,
    movimientos: movimientos.length,
    snapshot: snapshotRows.length
  };
}

function sincronizarInventarioDesdeCompraQTAS_(context) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, context || {});
  const ss = settings.ss || SpreadsheetApp.getActive();

  asegurarControlesInventarioBaseQTAS_();

  const movimientosRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioMovimientos);
  if (!movimientosRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: movimientosRef.reason,
      movimientos: 0
    };
  }

  const controlsIndex = construirIndiceControlesInventarioQTAS_(listarControlesInventarioQTAS_());
  const movimientos = construirMovimientosInventarioCompraDetalleQTAS_(
    settings.lineas || [],
    {
      compraId: settings.compraId,
      fechaCompra: settings.fechaCompra,
      controlsIndex: controlsIndex
    }
  );

  if (movimientos.length) {
    appendMovimientosInventarioQTAS_(movimientosRef.sheet, movimientosRef.headers, movimientos);
    reconstruirSnapshotInventarioQTAS_({ ss: ss });
  }

  return {
    ok: true,
    skipped: false,
    movimientos: movimientos.length
  };
}

function sincronizarInventarioDesdeVentaQTAS_(context) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, context || {});
  const ss = settings.ss || SpreadsheetApp.getActive();

  asegurarControlesInventarioBaseQTAS_();

  const movimientosRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioMovimientos);
  if (!movimientosRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: movimientosRef.reason,
      movimientos: 0
    };
  }

  const movimientos = construirMovimientosInventarioVentaLoteQTAS_(settings.detalleRows || [], {
    ss: ss
  });

  if (movimientos.length) {
    appendMovimientosInventarioQTAS_(movimientosRef.sheet, movimientosRef.headers, movimientos);
    reconstruirSnapshotInventarioQTAS_({ ss: ss });
  }

  return {
    ok: true,
    skipped: false,
    movimientos: movimientos.length
  };
}

function sincronizarInventarioDesdeProduccionDetalleQTAS_(context) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, context || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const movimientosRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioMovimientos);
  if (!movimientosRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: movimientosRef.reason,
      movimientos: 0
    };
  }

  const movimientos = construirMovimientosInventarioDesdeProduccionDetalleRowsQTAS_(
    settings.detalleRows || [],
    settings.fechaProduccion
  );

  if (movimientos.length) {
    appendMovimientosInventarioQTAS_(movimientosRef.sheet, movimientosRef.headers, movimientos);
    reconstruirSnapshotInventarioQTAS_({ ss: ss });
  }

  return {
    ok: true,
    skipped: false,
    movimientos: movimientos.length
  };
}

function asegurarControlesInventarioBaseQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.inventarioControl);
  if (!sheet) return [];

  const headers = getHeaders_(sheet);
  const existentes = leerObjetosConMeta_(sheet);
  const index = construirIndiceControlesInventarioQTAS_(
    existentes.map(row => normalizarControlInventarioQTAS_(row))
  );
  const candidatos = construirCandidatosControlInventarioQTAS_(ss);
  const nuevos = [];
  let nextId = siguienteIdConPrefijo_(sheet, 'Control_ID', 'INVCTRL-', 4);

  candidatos.forEach(candidato => {
    const key = claveControlInventarioQTAS_(candidato.tipoItem, candidato.item, candidato.unidad);
    if (!key || index[key]) return;

    nuevos.push(filaDesdeHeaders_(headers, {
      Control_ID: nextId,
      Tipo_Item: candidato.tipoItem,
      Item: candidato.item,
      Unidad: candidato.unidad,
      Modo_Stock: obtenerModoStockPorDefectoInventarioQTAS_(candidato.tipoItem, candidato.item, candidato.unidad),
      Stock_Minimo: 0,
      Stock_Objetivo: 0,
      Activo: true,
      Nota: ''
    }));
    index[key] = true;
    nextId = siguienteIdConPrefijoDesdeValorQTAS_(nextId, 'INVCTRL-', 4);
  });

  if (nuevos.length) {
    escribirFilas_(sheet, nuevos);
  }

  return listarControlesInventarioQTAS_();
}

function listarControlesInventarioQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const ref = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioControl);
  if (!ref.ok) return [];

  return leerObjetos_(ref.sheet)
    .map(row => normalizarControlInventarioQTAS_(row))
    .filter(row => row.item && row.unidad)
    .sort((a, b) => {
      if (a.tipoItem !== b.tipoItem) return a.tipoItem.localeCompare(b.tipoItem);
      if (a.item !== b.item) return a.item.localeCompare(b.item);
      return a.unidad.localeCompare(b.unidad);
    });
}

function listarSnapshotInventarioQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const ref = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioSnapshot);
  if (!ref.ok) return [];

  return leerObjetos_(ref.sheet)
    .map(row => ({
      inventarioId: texto_(row.Inventario_ID),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
      item: texto_(row.Item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      modoStock: normalizarModoStockInventarioQTAS_(row.Modo_Stock),
      entradas: redondear_(numero_(row.Entradas)),
      salidas: redondear_(numero_(row.Salidas)),
      stockActual: redondear_(numero_(row.Stock_Actual)),
      stockMinimo: redondear_(numero_(row.Stock_Minimo)),
      stockObjetivo: redondear_(numero_(row.Stock_Objetivo)),
      estadoStock: texto_(row.Estado_Stock) || 'OK',
      ultimoMovimiento: row.Ultimo_Movimiento ? fechaTextoPlanoQTAS_(row.Ultimo_Movimiento) : '',
      activo: estaActivo_(row.Activo),
      nota: texto_(row.Nota)
    }))
    .filter(row => row.item && row.unidad && row.activo)
    .sort((a, b) => {
      const prioridad = prioridadEstadoStockInventarioQTAS_(a.estadoStock) - prioridadEstadoStockInventarioQTAS_(b.estadoStock);
      if (prioridad !== 0) return prioridad;
      if (a.tipoItem !== b.tipoItem) return a.tipoItem.localeCompare(b.tipoItem);
      return a.item.localeCompare(b.item);
    });
}

function listarMovimientosInventarioRecientesQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const ref = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioMovimientos);
  if (!ref.ok) return [];

  return leerObjetos_(ref.sheet)
    .map(row => ({
      movimientoId: texto_(row.Movimiento_ID),
      fechaMovimiento: row.Fecha_Movimiento ? fechaInput_(row.Fecha_Movimiento) : '',
      fuenteTipo: texto_(row.Fuente_Tipo),
      fuenteId: texto_(row.Fuente_ID),
      operacion: texto_(row.Operacion),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
      item: texto_(row.Item),
      cantidad: redondear_(numero_(row.Cantidad)),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      cantidadSignada: redondear_(numero_(row.Cantidad_Signada)),
      compraId: numero_(row.Compra_ID),
      ventaId: numero_(row.Venta_ID),
      produccionId: texto_(row.Produccion_ID),
      detalleId: texto_(row.Detalle_ID),
      nota: texto_(row.Nota)
    }))
    .filter(row => row.movimientoId)
    .sort((a, b) => {
      const fechaA = a.fechaMovimiento ? fecha_(a.fechaMovimiento) : new Date(0);
      const fechaB = b.fechaMovimiento ? fecha_(b.fechaMovimiento) : new Date(0);
      if (fechaA.getTime() !== fechaB.getTime()) return fechaB - fechaA;
      return rowSortDescQTAS_(a.movimientoId, b.movimientoId);
    })
    .slice(0, 30);
}

function listarProduccionesRecientesQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const ref = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.producciones);
  if (!ref.ok) return [];

  return leerObjetos_(ref.sheet)
    .filter(row => texto_(row.Produccion_ID) && !esRegistroAnulado_(row.Estado_Registro))
    .map(row => ({
      produccionId: texto_(row.Produccion_ID),
      fechaProduccion: fechaInput_(row.Fecha_Produccion),
      producto: texto_(row.Producto_Estandar),
      cantidad: redondear_(numero_(row.Cantidad_Producida)),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      comentario: texto_(row.Comentario_Produccion)
    }))
    .sort((a, b) => {
      const fechaA = fecha_(a.fechaProduccion);
      const fechaB = fecha_(b.fechaProduccion);
      if (fechaA.getTime() !== fechaB.getTime()) return fechaB - fechaA;
      return rowSortDescQTAS_(a.produccionId, b.produccionId);
    })
    .slice(0, 20);
}

function reconstruirSnapshotInventarioQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const snapshotRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.inventarioSnapshot);
  if (!snapshotRef.ok) return [];

  const controles = listarControlesInventarioQTAS_();
  const controlsIndex = construirIndiceControlesInventarioQTAS_(controles);
  const movimientos = settings.movimientos || leerObjetos_(ss.getSheetByName(QTAS.sheets.inventarioMovimientos));
  const rows = construirSnapshotInventarioQTAS_(movimientos, controlsIndex);

  sobrescribirObjetosHojaQTAS_(snapshotRef.sheet, snapshotRef.headers, rows);
  return rows;
}

function construirSnapshotInventarioQTAS_(movimientosRaw, controlsIndex) {
  const movimientos = (movimientosRaw || [])
    .map(row => ({
      fechaMovimiento: row.Fecha_Movimiento || row.fechaMovimiento || '',
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item || row.tipoItem),
      item: texto_(row.Item || row.item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad || row.unidad),
      cantidadSignada: redondear_(numero_(row.Cantidad_Signada || row.cantidadSignada))
    }))
    .filter(row => row.item && row.unidad && row.tipoItem);
  const keys = {};

  Object.keys(controlsIndex || {}).forEach(key => {
    const control = controlsIndex[key];
    if (!control || control.activo !== true) return;
    if (control.modoStock === 'NoControlado') return;
    keys[key] = true;
  });

  return Object.keys(keys)
    .map(key => construirFilaSnapshotInventarioQTAS_(key, movimientos, controlsIndex))
    .filter(Boolean)
    .sort((a, b) => {
      if (texto_(a.Estado_Stock) !== texto_(b.Estado_Stock)) {
        return prioridadEstadoStockInventarioQTAS_(a.Estado_Stock) - prioridadEstadoStockInventarioQTAS_(b.Estado_Stock);
      }
      if (texto_(a.Tipo_Item) !== texto_(b.Tipo_Item)) {
        return texto_(a.Tipo_Item).localeCompare(texto_(b.Tipo_Item));
      }
      return texto_(a.Item).localeCompare(texto_(b.Item));
    });
}

function construirFilaSnapshotInventarioQTAS_(key, movimientos, controlsIndex) {
  const parts = texto_(key).split('|');
  if (parts.length < 3) return null;

  const tipoItem = parts[0];
  const item = parts[1];
  const unidad = parts[2];
  const control = obtenerControlInventarioEfectivoQTAS_(tipoItem, item, unidad, controlsIndex);
  if (control.modoStock === 'NoControlado') return null;

  const rows = (movimientos || []).filter(row =>
    claveControlInventarioQTAS_(row.tipoItem, row.item, row.unidad) === key
  );
  const entradas = redondear_(sumar_(rows.filter(row => numero_(row.cantidadSignada) > 0).map(row => row.cantidadSignada)));
  const salidas = redondear_(Math.abs(sumar_(rows.filter(row => numero_(row.cantidadSignada) < 0).map(row => row.cantidadSignada))));
  const stockActual = redondear_(entradas - salidas);
  const ultimaFecha = rows.length
    ? rows
      .map(row => resolverFechaOperacion_(row.fechaMovimiento, new Date(0)))
      .sort((a, b) => b - a)[0]
    : '';

  return {
    Inventario_ID: `STK-${normalizarClaveProductoQTAS_(item, unidad)}`.slice(0, 99),
    Tipo_Item: tipoItem,
    Item: item,
    Unidad: unidad,
    Modo_Stock: control.modoStock,
    Entradas: entradas,
    Salidas: salidas,
    Stock_Actual: stockActual,
    Stock_Minimo: control.stockMinimo,
    Stock_Objetivo: control.stockObjetivo,
    Estado_Stock: clasificarEstadoStockInventarioQTAS_(stockActual, control.stockMinimo, control.stockObjetivo),
    Ultimo_Movimiento: ultimaFecha || '',
    Activo: control.activo,
    Nota: control.nota
  };
}

function construirMovimientosInventarioDesdeComprasQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
  const controlsIndex = construirIndiceControlesInventarioQTAS_(listarControlesInventarioQTAS_());

  return leerObjetos_(sheet)
    .filter(row => texto_(row.Compra_Detalle_ID) && !esRegistroAnulado_(row.Estado_Registro))
    .reduce((acc, row) => acc.concat(
      construirMovimientosInventarioCompraDetalleQTAS_([row], {
        compraId: numero_(row.Compra_ID),
        fechaCompra: valorFechaCompraCanonicaQTAS_(row, new Date()),
        controlsIndex: controlsIndex
      })
    ), []);
}

function construirMovimientosInventarioCompraDetalleQTAS_(lineas, context) {
  const controlsIndex = context && context.controlsIndex
    ? context.controlsIndex
    : construirIndiceControlesInventarioQTAS_(listarControlesInventarioQTAS_());

  return (lineas || [])
    .filter(linea => !esRegistroAnulado_(linea.Estado_Registro))
    .filter(linea => normalizarTipoCompraItemQTAS_(linea.Tipo_Item) !== 'Gasto')
    .filter(linea => numero_(linea.Cantidad) > 0)
    .map(linea => {
      const tipoItem = normalizarTipoCompraItemQTAS_(linea.Tipo_Item);
      const item = texto_(linea.Item);
      const unidad = normalizarUnidadCanonicaQTAS_(linea.Unidad);
      const control = obtenerControlInventarioEfectivoQTAS_(tipoItem, item, unidad, controlsIndex);
      if (!control.activo || control.modoStock === 'NoControlado') return null;

      return crearMovimientoInventarioQTAS_({
        fechaMovimiento: context && context.fechaCompra ? context.fechaCompra : valorFechaCompraCanonicaQTAS_(linea, new Date()),
        fuenteTipo: 'Compra',
        fuenteId: texto_(linea.Compra_Detalle_ID),
        operacion: 'Entrada',
        tipoItem: tipoItem,
        item: item,
        cantidad: numero_(linea.Cantidad),
        unidad: unidad,
        compraId: context && context.compraId ? context.compraId : numero_(linea.Compra_ID),
        detalleId: texto_(linea.Compra_Detalle_ID),
        nota: texto_(linea.Comentario_Linea)
      });
    })
    .filter(Boolean);
}

function construirMovimientosInventarioDesdeProduccionesQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const detalleSheet = ss.getSheetByName(QTAS.sheets.produccionDetalle);

  return construirMovimientosInventarioDesdeProduccionDetalleRowsQTAS_(
    leerObjetos_(detalleSheet).filter(row => texto_(row.Produccion_Detalle_ID)),
    null
  );
}

function construirMovimientosInventarioDesdeProduccionDetalleRowsQTAS_(rows, fallbackDate) {
  return (rows || [])
    .map(row => {
      const cantidad = redondear_(Math.abs(numero_(row.Cantidad)));
      if (cantidad <= 0) return null;

      return crearMovimientoInventarioQTAS_({
        fechaMovimiento: row.Fecha_Produccion || fallbackDate || new Date(),
        fuenteTipo: 'Produccion',
        fuenteId: texto_(row.Produccion_Detalle_ID),
        operacion: texto_(row.Operacion) || 'Salida',
        tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
        item: texto_(row.Item),
        cantidad: cantidad,
        unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
        produccionId: texto_(row.Produccion_ID),
        detalleId: texto_(row.Produccion_Detalle_ID),
        nota: texto_(row.Nota)
      });
    })
    .filter(Boolean);
}

function construirMovimientosInventarioDesdeVentasQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const detalleSheet = ss.getSheetByName(QTAS.sheets.detalle);
  return construirMovimientosInventarioVentaLoteQTAS_(
    leerObjetos_(detalleSheet).filter(row => texto_(row.Detalle_ID) && !esRegistroAnulado_(row.Estado_Registro)),
    { ss: ss }
  );
}

function construirMovimientosInventarioVentaLoteQTAS_(detalleRows, payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const componentes = settings.componentes || leerComponentesProductoActivosQTAS_();
  const reglas = settings.reglas || leerReglasCostoProductoActivasQTAS_();
  const controlsIndex = construirIndiceControlesInventarioQTAS_(listarControlesInventarioQTAS_());

  return (detalleRows || []).reduce((acc, row) => acc.concat(
    construirMovimientosInventarioVentaDesdeDetalleQTAS_(row, {
      componentes: componentes,
      reglas: reglas,
      controlsIndex: controlsIndex
    })
  ), []);
}

function construirMovimientosInventarioVentaDesdeDetalleQTAS_(row, context) {
  const detalleId = texto_(row && row.Detalle_ID);
  const producto = texto_(row && row.Producto_Estandar);
  const unidad = normalizarUnidadCanonicaQTAS_(row && row.Unidad);
  const cantidadVenta = redondear_(Math.max(0, numero_(row && row.Cantidad)));
  const fechaVenta = valorFechaVentaCanonicaQTAS_(row, new Date());
  if (!detalleId || !producto || !unidad || cantidadVenta <= 0 || esRegistroAnulado_(row && row.Estado_Registro)) {
    return [];
  }

  const control = obtenerControlInventarioEfectivoQTAS_('Producto', producto, unidad, context.controlsIndex);
  const movimientos = [];

  if (control.activo && control.modoStock !== 'NoControlado') {
    if (control.modoStock === 'Directo' || control.modoStock === 'Fabricado') {
      movimientos.push(crearMovimientoInventarioQTAS_({
        fechaMovimiento: fechaVenta,
        fuenteTipo: 'Venta',
        fuenteId: detalleId,
        operacion: 'Salida',
        tipoItem: 'Producto',
        item: producto,
        cantidad: cantidadVenta,
        unidad: unidad,
        ventaId: numero_(row && row.Venta_ID),
        detalleId: detalleId,
        nota: 'Salida por venta'
      }));
    } else if (control.modoStock === 'PorRecetaVenta') {
      movimientos.push.apply(movimientos, construirMovimientosBaseRecetaVentaQTAS_({
        producto: producto,
        unidad: unidad,
        cantidadVenta: cantidadVenta,
        fechaBase: fechaVenta,
        ventaId: numero_(row && row.Venta_ID),
        detalleId: detalleId,
        componentes: context.componentes,
        controlesIndex: context.controlsIndex
      }));
    }

    movimientos.push.apply(movimientos, construirMovimientosReglasVentaInventarioQTAS_({
      producto: producto,
      unidad: unidad,
      cantidadVenta: cantidadVenta,
      fechaBase: fechaVenta,
      ventaId: numero_(row && row.Venta_ID),
      detalleId: detalleId,
      reglas: context.reglas,
      controlesIndex: context.controlsIndex,
      componentes: context.componentes
    }));
  }

  return movimientos;
}

function construirMovimientosBaseRecetaVentaQTAS_(context) {
  const componentesProducto = (context.componentes || [])
    .filter(row =>
      esMismaClaveProductoQTAS_(row.producto, row.unidadVenta, context.producto, context.unidad)
    );

  return componentesProducto.reduce((acc, row) => acc.concat(
    construirMovimientosSalidaDesdeComponenteInventarioQTAS_(row, {
      multiplicador: context.cantidadVenta,
      fechaBase: context.fechaBase,
      fuenteTipo: 'Venta',
      fuenteId: context.detalleId,
      ventaId: context.ventaId,
      detalleId: context.detalleId,
      controlesIndex: context.controlesIndex,
      componentes: context.componentes,
      stack: [normalizarClaveProductoQTAS_(context.producto, context.unidad)],
      nota: 'Consumo por receta de venta'
    })
  ), []);
}

function construirMovimientosReglasVentaInventarioQTAS_(context) {
  const reglasAplicadas = seleccionarReglasCostoProductoQTAS_(
    context.producto,
    context.unidad,
    context.fechaBase,
    context.cantidadVenta,
    context.reglas || []
  );

  return reglasAplicadas.reduce((acc, row) => acc.concat(
    construirMovimientosSalidaDesdeComponenteInventarioQTAS_(row, {
      multiplicador: row.aplicacion === 'PorLinea' ? 1 : context.cantidadVenta,
      fechaBase: context.fechaBase,
      fuenteTipo: 'Venta',
      fuenteId: context.detalleId,
      ventaId: context.ventaId,
      detalleId: context.detalleId,
      controlesIndex: context.controlesIndex,
      componentes: context.componentes,
      stack: [normalizarClaveProductoQTAS_(context.producto, context.unidad)],
      nota: `Consumo por regla ${row.aplicacion || 'PorUnidad'}`
    })
  ), []);
}

function construirMovimientosSalidaDesdeComponenteInventarioQTAS_(row, context) {
  const cantidadAplicada = redondear_(
    numero_(row.cantidadComponente) *
    numero_(context.multiplicador) *
    (1 + Math.max(0, numero_(row.mermaPct)) / 100)
  );
  if (cantidadAplicada <= 0) return [];

  return expandirSalidaInventarioStockeableQTAS_({
    tipoItem: row.tipoComponente,
    item: row.itemComponente,
    cantidad: cantidadAplicada,
    unidad: row.unidadComponente,
    fechaBase: context.fechaBase,
    fuenteTipo: context.fuenteTipo,
    fuenteId: context.fuenteId,
    ventaId: context.ventaId || 0,
    detalleId: context.detalleId || '',
    produccionId: context.produccionId || '',
    controlesIndex: context.controlesIndex,
    componentes: context.componentes,
    stack: (context.stack || []).slice(),
    nota: context.nota
  });
}

function expandirSalidaInventarioStockeableQTAS_(context) {
  const tipoItem = normalizarTipoCompraItemQTAS_(context.tipoItem);
  const item = texto_(context.item);
  const unidad = normalizarUnidadCanonicaQTAS_(context.unidad);
  const cantidad = redondear_(Math.max(0, numero_(context.cantidad)));
  if (!item || !unidad || cantidad <= 0 || tipoItem === 'Gasto') return [];

  const control = obtenerControlInventarioEfectivoQTAS_(tipoItem, item, unidad, context.controlesIndex);
  if (!control.activo || control.modoStock === 'NoControlado') return [];

  if (tipoItem !== 'Producto' || control.modoStock === 'Directo' || control.modoStock === 'Fabricado') {
    return [crearMovimientoInventarioQTAS_({
      fechaMovimiento: context.fechaBase,
      fuenteTipo: context.fuenteTipo,
      fuenteId: context.fuenteId,
      operacion: 'Salida',
      tipoItem: tipoItem,
      item: item,
      cantidad: cantidad,
      unidad: unidad,
      ventaId: context.ventaId,
      detalleId: context.detalleId,
      produccionId: context.produccionId,
      nota: context.nota
    })];
  }

  const stack = (context.stack || []).slice();
  const currentKey = normalizarClaveProductoQTAS_(item, unidad);
  if (stack.indexOf(currentKey) >= 0) {
    return [crearMovimientoInventarioQTAS_({
      fechaMovimiento: context.fechaBase,
      fuenteTipo: context.fuenteTipo,
      fuenteId: context.fuenteId,
      operacion: 'Salida',
      tipoItem: tipoItem,
      item: item,
      cantidad: cantidad,
      unidad: unidad,
      ventaId: context.ventaId,
      detalleId: context.detalleId,
      produccionId: context.produccionId,
      nota: unirUnicos_([context.nota, 'Fallback por recursion'])
    })];
  }

  const componentes = (context.componentes || leerComponentesProductoActivosQTAS_())
    .filter(row => esMismaClaveProductoQTAS_(row.producto, row.unidadVenta, item, unidad));
  if (!componentes.length) {
    return [crearMovimientoInventarioQTAS_({
      fechaMovimiento: context.fechaBase,
      fuenteTipo: context.fuenteTipo,
      fuenteId: context.fuenteId,
      operacion: 'Salida',
      tipoItem: tipoItem,
      item: item,
      cantidad: cantidad,
      unidad: unidad,
      ventaId: context.ventaId,
      detalleId: context.detalleId,
      produccionId: context.produccionId,
      nota: unirUnicos_([context.nota, 'Fallback sin receta activa'])
    })];
  }

  return componentes.reduce((acc, row) => acc.concat(
    construirMovimientosSalidaDesdeComponenteInventarioQTAS_(row, {
      multiplicador: cantidad,
      fechaBase: context.fechaBase,
      fuenteTipo: context.fuenteTipo,
      fuenteId: context.fuenteId,
      ventaId: context.ventaId,
      detalleId: context.detalleId,
      produccionId: context.produccionId,
      controlesIndex: context.controlesIndex,
      componentes: context.componentes,
      stack: stack.concat([currentKey]),
      nota: context.nota
    })
  ), []);
}

function construirDetalleProduccionMaterializadoQTAS_(context) {
  const componentes = (context.componentes || [])
    .filter(row =>
      esMismaClaveProductoQTAS_(row.producto, row.unidadVenta, context.producto, context.unidad)
    );
  const controlsIndex = construirIndiceControlesInventarioQTAS_(listarControlesInventarioQTAS_());
  const rows = [{
    Produccion_Detalle_ID: produccionDetalleIdQTAS_(context.produccionId, 1),
    Produccion_ID: context.produccionId,
    Fecha_Produccion: context.fechaProduccion,
    Operacion: 'Entrada',
    Tipo_Item: 'Producto',
    Item: context.producto,
    Cantidad: context.cantidad,
    Unidad: context.unidad,
    Nota: unirUnicos_([context.comentario, 'Ingreso de producto terminado'])
  }];

  let correlativo = 2;
  componentes.forEach(row => {
    const movimientos = construirMovimientosSalidaDesdeComponenteInventarioQTAS_(row, {
      multiplicador: context.cantidad,
      fechaBase: context.fechaProduccion,
      fuenteTipo: 'Produccion',
      fuenteId: context.produccionId,
      produccionId: context.produccionId,
      controlesIndex: controlsIndex,
      componentes: context.componentes,
      stack: [normalizarClaveProductoQTAS_(context.producto, context.unidad)],
      nota: 'Consumo por produccion'
    });

    movimientos.forEach(mov => {
      rows.push({
        Produccion_Detalle_ID: produccionDetalleIdQTAS_(context.produccionId, correlativo++),
        Produccion_ID: context.produccionId,
        Fecha_Produccion: context.fechaProduccion,
        Operacion: 'Salida',
        Tipo_Item: mov.Tipo_Item,
        Item: mov.Item,
        Cantidad: Math.abs(numero_(mov.Cantidad)),
        Unidad: mov.Unidad,
        Nota: mov.Nota
      });
    });
  });

  return rows;
}

function appendMovimientosInventarioQTAS_(sheet, headers, rows) {
  if (!sheet || !headers || !headers.length || !rows || !rows.length) return;

  let nextId = siguienteIdConPrefijo_(sheet, 'Movimiento_ID', 'INV-', 6);
  const objects = rows.map(row => {
    const withId = Object.assign({}, row, {
      Movimiento_ID: nextId
    });
    nextId = siguienteIdConPrefijoDesdeValorQTAS_(nextId, 'INV-', 6);
    return withId;
  });

  escribirFilas_(sheet, objects.map(obj => filaDesdeHeaders_(headers, obj)));
}

function asignarIdsMovimientosInventarioQTAS_(rows) {
  let currentId = 'INV-000001';
  return (rows || []).map((row, index) => {
    if (index === 0) {
      currentId = 'INV-000001';
    } else {
      currentId = siguienteIdConPrefijoDesdeValorQTAS_(currentId, 'INV-', 6);
    }
    return Object.assign({}, row, {
      Movimiento_ID: currentId
    });
  });
}

function construirCandidatosControlInventarioQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const mapa = {};
  const add = function(tipoItem, item, unidad) {
    const tipo = normalizarTipoCompraItemQTAS_(tipoItem);
    const nombre = texto_(item);
    const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);
    if (!nombre || !unidadCanonica || tipo === 'Gasto') return;
    mapa[claveControlInventarioQTAS_(tipo, nombre, unidadCanonica)] = {
      tipoItem: tipo,
      item: nombre,
      unidad: unidadCanonica
    };
  };

  leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.productos))
    .filter(row => texto_(row.Producto_Estandar))
    .filter(row => estaActivo_(row.Activo))
    .forEach(row => add('Producto', row.Producto_Estandar, row.Unidad_Default));

  leerComponentesProductoActivosQTAS_().forEach(row => {
    add(row.tipoComponente, row.itemComponente, row.unidadComponente);
  });

  leerReglasCostoProductoActivasQTAS_().forEach(row => {
    add(row.tipoComponente, row.itemComponente, row.unidadComponente);
  });

  return Object.keys(mapa).map(key => mapa[key]);
}

function construirIndiceControlesInventarioQTAS_(rows) {
  return (rows || []).reduce((acc, row) => {
    const normalized = normalizarControlInventarioQTAS_(row);
    const key = claveControlInventarioQTAS_(normalized.tipoItem, normalized.item, normalized.unidad);
    if (key && !acc[key]) {
      acc[key] = normalized;
    }
    return acc;
  }, {});
}

function obtenerControlInventarioEfectivoQTAS_(tipoItem, item, unidad, controlsIndex) {
  const key = claveControlInventarioQTAS_(tipoItem, item, unidad);
  const existente = controlsIndex && controlsIndex[key] ? controlsIndex[key] : null;
  if (existente) return existente;

  const tipo = normalizarTipoCompraItemQTAS_(tipoItem);
  const nombre = texto_(item);
  const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);

  return {
    controlId: '',
    tipoItem: tipo,
    item: nombre,
    unidad: unidadCanonica,
    modoStock: obtenerModoStockPorDefectoInventarioQTAS_(tipo, nombre, unidadCanonica),
    stockMinimo: 0,
    stockObjetivo: 0,
    activo: tipo !== 'Gasto',
    nota: ''
  };
}

function normalizarControlInventarioQTAS_(row) {
  return {
    controlId: texto_(row.Control_ID || row.controlId),
    tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item || row.tipoItem),
    item: texto_(row.Item || row.item),
    unidad: normalizarUnidadCanonicaQTAS_(row.Unidad || row.unidad),
    modoStock: normalizarModoStockInventarioQTAS_(row.Modo_Stock || row.modoStock),
    stockMinimo: redondear_(Math.max(0, numero_(row.Stock_Minimo || row.stockMinimo))),
    stockObjetivo: redondear_(Math.max(0, numero_(row.Stock_Objetivo || row.stockObjetivo))),
    activo: estaActivo_(row.Activo !== undefined ? row.Activo : row.activo),
    nota: texto_(row.Nota || row.nota)
  };
}

function normalizarModoStockInventarioQTAS_(value) {
  const key = normalizarClaveTexto_(value);
  if (!key) return 'Directo';
  if (key === normalizarClaveTexto_('Directo')) return 'Directo';
  if (key === normalizarClaveTexto_('Fabricado')) return 'Fabricado';
  if (key === normalizarClaveTexto_('PorRecetaVenta') || key === normalizarClaveTexto_('Por receta venta')) {
    return 'PorRecetaVenta';
  }
  if (key === normalizarClaveTexto_('NoControlado') || key === normalizarClaveTexto_('No controlado')) {
    return 'NoControlado';
  }
  return 'Directo';
}

function obtenerModoStockPorDefectoInventarioQTAS_(tipoItem, item, unidad) {
  const tipo = normalizarTipoCompraItemQTAS_(tipoItem);
  const producto = texto_(item);
  const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);

  if (tipo === 'Gasto') return 'NoControlado';
  if (tipo === 'Insumo') return 'Directo';
  if (PRODUCTOS_RECETA_VENTA_INVENTARIO_QTAS.some(nombre =>
    normalizarClaveTexto_(nombre) === normalizarClaveTexto_(producto)
  )) {
    return 'PorRecetaVenta';
  }
  if (PRODUCTOS_FABRICADOS_INVENTARIO_QTAS.some(nombre =>
    normalizarClaveTexto_(nombre) === normalizarClaveTexto_(producto)
  )) {
    return 'Fabricado';
  }
  return unidadCanonica ? 'Directo' : 'NoControlado';
}

function claveControlInventarioQTAS_(tipoItem, item, unidad) {
  return [
    normalizarTipoCompraItemQTAS_(tipoItem),
    texto_(item),
    normalizarUnidadCanonicaQTAS_(unidad)
  ].join('|');
}

function crearMovimientoInventarioQTAS_(context) {
  const operacion = texto_(context.operacion) === 'Entrada' ? 'Entrada' : 'Salida';
  const cantidad = redondear_(Math.abs(numero_(context.cantidad)));
  const cantidadSignada = operacion === 'Entrada' ? cantidad : redondear_(-cantidad);

  return {
    Movimiento_ID: '',
    Fecha_Movimiento: resolverFechaOperacion_(context.fechaMovimiento, new Date()),
    Fuente_Tipo: texto_(context.fuenteTipo),
    Fuente_ID: texto_(context.fuenteId),
    Operacion: operacion,
    Tipo_Item: normalizarTipoCompraItemQTAS_(context.tipoItem),
    Item: texto_(context.item),
    Cantidad: cantidad,
    Unidad: normalizarUnidadCanonicaQTAS_(context.unidad),
    Cantidad_Signada: cantidadSignada,
    Compra_ID: numero_(context.compraId),
    Venta_ID: numero_(context.ventaId),
    Produccion_ID: texto_(context.produccionId),
    Detalle_ID: texto_(context.detalleId),
    Nota: texto_(context.nota)
  };
}

function clasificarEstadoStockInventarioQTAS_(stockActual, stockMinimo, stockObjetivo) {
  const stock = redondear_(numero_(stockActual));
  const minimo = redondear_(Math.max(0, numero_(stockMinimo)));
  const objetivo = redondear_(Math.max(0, numero_(stockObjetivo)));

  if (stock <= 0.009) return 'Sin stock';
  if (minimo > 0 && stock <= minimo + 0.009) return 'Bajo minimo';
  if (objetivo > 0 && stock < objetivo - 0.009) return 'Por reponer';
  return 'OK';
}

function prioridadEstadoStockInventarioQTAS_(estado) {
  const key = texto_(estado);
  if (key === 'Sin stock') return 0;
  if (key === 'Bajo minimo') return 1;
  if (key === 'Por reponer') return 2;
  return 3;
}

function produccionDetalleIdQTAS_(produccionId, correlativo) {
  return [
    'PRDDET',
    texto_(produccionId).replace(/[^A-Za-z0-9]+/g, ''),
    String(correlativo).padStart(3, '0')
  ].join('-').slice(0, 99);
}

function rowSortDescQTAS_(a, b) {
  return texto_(b).localeCompare(texto_(a));
}
