function guardarComponenteProductoQTAS(payload) {
  return withScriptLock_('guardar componente producto', () => {
    asegurarModeloOperativoQTAS_();

    const producto = texto_(payload && payload.producto);
    const unidadVenta = normalizarUnidadCanonicaQTAS_(payload && payload.unidadVenta);
    const tipoComponente = normalizarTipoCompraItemQTAS_(payload && payload.tipoComponente);
    const itemComponente = texto_(payload && payload.itemComponente);
    const medidaComponente = normalizarCantidadUnidadQTAS_(
      payload && payload.cantidadComponente,
      payload && payload.unidadComponente
    );
    const cantidadComponente = medidaComponente.cantidad;
    const unidadComponente = medidaComponente.unidad;
    const orden = Math.max(1, Math.floor(numero_(payload && payload.orden) || 1));
    const mermaPct = redondear_(Math.max(0, numero_(payload && payload.mermaPct)));
    const nota = texto_(payload && payload.nota);
    const activo = !payload || payload.activo !== false;
    const componenteIdBuscado = texto_(payload && payload.componenteId);

    if (!producto) throw new Error('Falta el producto.');
    if (!unidadVenta) throw new Error('Falta la unidad de venta.');
    if (!itemComponente) throw new Error('Falta el item del componente.');
    if (!unidadComponente) throw new Error('Falta la unidad del componente.');
    if (cantidadComponente <= 0) throw new Error('La cantidad del componente debe ser mayor a cero.');
    validarProductoCanonicoExistenteQTAS_(SpreadsheetApp.getActive(), producto);

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.productoComponentes);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const existente = componenteIdBuscado
      ? rows.find(row => texto_(row.Componente_ID) === componenteIdBuscado)
      : rows.find(row =>
        normalizarClaveTexto_(row.Producto_Estandar) === normalizarClaveTexto_(producto) &&
        normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad_Venta)) ===
          normalizarClaveTexto_(unidadVenta) &&
        normalizarClaveTexto_(row.Tipo_Componente) === normalizarClaveTexto_(tipoComponente) &&
        normalizarClaveTexto_(row.Item_Componente) === normalizarClaveTexto_(itemComponente) &&
        numero_(row.Orden) === orden
      );

    const normalizado = {
      Componente_ID: existente
        ? texto_(existente.Componente_ID)
        : siguienteIdConPrefijo_(sheet, 'Componente_ID', 'RCP-', 4),
      Producto_Estandar: producto,
      Unidad_Venta: unidadVenta,
      Orden: orden,
      Tipo_Componente: tipoComponente,
      Item_Componente: itemComponente,
      Cantidad_Componente: cantidadComponente,
      Unidad_Componente: unidadComponente,
      Merma_Pct: mermaPct,
      Activo: activo,
      Nota: nota
    };

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, normalizado));
    } else {
      escribirFilas_(sheet, [filaDesdeHeaders_(headers, normalizado)]);
    }

    return {
      ok: true,
      componenteId: normalizado.Componente_ID,
      componentes: listarComponentesProductoQTAS_()
    };
  });
}

function guardarReglaCostoProductoQTAS(payload) {
  return withScriptLock_('guardar regla costo producto', () => {
    asegurarModeloOperativoQTAS_();

    const producto = texto_(payload && payload.producto);
    const unidadVenta = normalizarUnidadCanonicaQTAS_(payload && payload.unidadVenta);
    const tipoComponente = normalizarTipoCompraItemQTAS_(payload && payload.tipoComponente);
    const itemComponente = texto_(payload && payload.itemComponente);
    const medidaComponente = normalizarCantidadUnidadQTAS_(
      payload && payload.cantidadComponente,
      payload && payload.unidadComponente
    );
    const cantidadComponente = medidaComponente.cantidad;
    const unidadComponente = medidaComponente.unidad;
    const orden = Math.max(1, Math.floor(numero_(payload && payload.orden) || 1));
    const aplicacion = normalizarAplicacionReglaCostoQTAS_(payload && payload.aplicacion);
    const mermaPct = redondear_(Math.max(0, numero_(payload && payload.mermaPct)));
    const cantidadMin = texto_(payload && payload.cantidadMin) === ''
      ? ''
      : redondear_(Math.max(0, numero_(payload && payload.cantidadMin)));
    const cantidadMax = texto_(payload && payload.cantidadMax) === ''
      ? ''
      : redondear_(Math.max(0, numero_(payload && payload.cantidadMax)));
    const fechaDesde = payload && payload.fechaDesde
      ? resolverFechaOperacion_(payload.fechaDesde, new Date())
      : '';
    const fechaHasta = payload && payload.fechaHasta
      ? resolverFechaOperacion_(payload.fechaHasta, new Date())
      : '';
    const nota = texto_(payload && payload.nota);
    const activo = !payload || payload.activo !== false;
    const reglaIdBuscada = texto_(payload && payload.reglaId);

    if (!producto) throw new Error('Falta el producto.');
    if (!unidadVenta) throw new Error('Falta la unidad de venta.');
    if (!itemComponente) throw new Error('Falta el item del componente.');
    if (!unidadComponente) throw new Error('Falta la unidad del componente.');
    if (cantidadComponente <= 0) throw new Error('La cantidad del componente debe ser mayor a cero.');
    if (cantidadMax !== '' && cantidadMax < numero_(cantidadMin)) {
      throw new Error('Cantidad_Max no puede ser menor que Cantidad_Min.');
    }
    if (fechaDesde && fechaHasta && resolverFechaOperacion_(fechaHasta, fechaDesde) < fechaDesde) {
      throw new Error('Fecha_Hasta no puede ser anterior a Fecha_Desde.');
    }
    validarProductoCanonicoExistenteQTAS_(SpreadsheetApp.getActive(), producto);

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.productoReglasCosto);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const existente = reglaIdBuscada
      ? rows.find(row => texto_(row.Regla_Costo_ID) === reglaIdBuscada)
      : rows.find(row =>
        normalizarClaveTexto_(row.Producto_Estandar) === normalizarClaveTexto_(producto) &&
        normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad_Venta)) ===
          normalizarClaveTexto_(unidadVenta) &&
        numero_(row.Orden) === orden &&
        normalizarClaveTexto_(row.Tipo_Componente) === normalizarClaveTexto_(tipoComponente) &&
        normalizarClaveTexto_(row.Item_Componente) === normalizarClaveTexto_(itemComponente) &&
        normalizarClaveTexto_(row.Aplicacion) === normalizarClaveTexto_(aplicacion) &&
        fechaTextoPlanoQTAS_(row.Fecha_Desde) === fechaTextoPlanoQTAS_(fechaDesde) &&
        fechaTextoPlanoQTAS_(row.Fecha_Hasta) === fechaTextoPlanoQTAS_(fechaHasta) &&
        texto_(row.Cantidad_Min) === texto_(cantidadMin) &&
        texto_(row.Cantidad_Max) === texto_(cantidadMax)
      );

    const normalizado = {
      Regla_Costo_ID: existente
        ? texto_(existente.Regla_Costo_ID)
        : siguienteIdConPrefijo_(sheet, 'Regla_Costo_ID', 'RCR-', 4),
      Producto_Estandar: producto,
      Unidad_Venta: unidadVenta,
      Fecha_Desde: fechaDesde,
      Fecha_Hasta: fechaHasta,
      Cantidad_Min: cantidadMin,
      Cantidad_Max: cantidadMax,
      Orden: orden,
      Tipo_Componente: tipoComponente,
      Item_Componente: itemComponente,
      Cantidad_Componente: cantidadComponente,
      Unidad_Componente: unidadComponente,
      Aplicacion: aplicacion,
      Merma_Pct: mermaPct,
      Activo: activo,
      Nota: nota
    };

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, normalizado));
    } else {
      escribirFilas_(sheet, [filaDesdeHeaders_(headers, normalizado)]);
    }

    return {
      ok: true,
      reglaId: normalizado.Regla_Costo_ID,
      reglas: listarReglasCostoProductoQTAS_()
    };
  });
}

function listarComponentesProductoQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.productoComponentes);
  if (!sheet) return [];

  return leerObjetos_(sheet)
    .map(row => {
      const medidaComponente = normalizarCantidadUnidadQTAS_(row.Cantidad_Componente, row.Unidad_Componente);
      return {
        componenteId: texto_(row.Componente_ID),
        producto: texto_(row.Producto_Estandar),
        unidadVenta: normalizarUnidadCanonicaQTAS_(row.Unidad_Venta),
        orden: numero_(row.Orden),
        tipoComponente: normalizarTipoCompraItemQTAS_(row.Tipo_Componente),
        itemComponente: texto_(row.Item_Componente),
        cantidadComponente: medidaComponente.cantidad,
        unidadComponente: medidaComponente.unidad,
        mermaPct: redondear_(numero_(row.Merma_Pct)),
        activo: estaActivo_(row.Activo),
        nota: texto_(row.Nota)
      };
    })
    .filter(row => row.componenteId || row.producto || row.itemComponente)
    .sort((a, b) => {
      if (a.producto !== b.producto) return a.producto.localeCompare(b.producto);
      if (a.unidadVenta !== b.unidadVenta) return a.unidadVenta.localeCompare(b.unidadVenta);
      if (a.orden !== b.orden) return a.orden - b.orden;
      return a.itemComponente.localeCompare(b.itemComponente);
    });
}

function listarReglasCostoProductoQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.productoReglasCosto);
  if (!sheet) return [];

  return leerObjetos_(sheet)
    .map(row => {
      const medidaComponente = normalizarCantidadUnidadQTAS_(row.Cantidad_Componente, row.Unidad_Componente);
      return {
        reglaId: texto_(row.Regla_Costo_ID),
        producto: texto_(row.Producto_Estandar),
        unidadVenta: normalizarUnidadCanonicaQTAS_(row.Unidad_Venta),
        fechaDesde: row.Fecha_Desde ? fechaTextoPlanoQTAS_(row.Fecha_Desde) : '',
        fechaHasta: row.Fecha_Hasta ? fechaTextoPlanoQTAS_(row.Fecha_Hasta) : '',
        cantidadMin: texto_(row.Cantidad_Min) === '' ? '' : redondear_(numero_(row.Cantidad_Min)),
        cantidadMax: texto_(row.Cantidad_Max) === '' ? '' : redondear_(numero_(row.Cantidad_Max)),
        orden: numero_(row.Orden),
        tipoComponente: normalizarTipoCompraItemQTAS_(row.Tipo_Componente),
        itemComponente: texto_(row.Item_Componente),
        cantidadComponente: medidaComponente.cantidad,
        unidadComponente: medidaComponente.unidad,
        aplicacion: normalizarAplicacionReglaCostoQTAS_(row.Aplicacion),
        mermaPct: redondear_(numero_(row.Merma_Pct)),
        activo: estaActivo_(row.Activo),
        nota: texto_(row.Nota)
      };
    })
    .filter(row => row.reglaId || row.producto || row.itemComponente)
    .sort((a, b) => {
      if (a.producto !== b.producto) return a.producto.localeCompare(b.producto);
      if (a.unidadVenta !== b.unidadVenta) return a.unidadVenta.localeCompare(b.unidadVenta);
      if (a.orden !== b.orden) return a.orden - b.orden;
      if (a.aplicacion !== b.aplicacion) return a.aplicacion.localeCompare(b.aplicacion);
      return a.itemComponente.localeCompare(b.itemComponente);
    });
}

function reconstruirCostoProductoCalculadoQTAS(payload) {
  const settings = Object.assign({
    fechaBase: new Date(),
    silent: false
  }, payload || {});
  asegurarModeloOperativoQTAS_();
  const result = reconstruirCostoProductoCalculadoInternoQTAS_(settings);

  if (!settings.silent) {
    maybeAlert_(
      result.skipped
        ? `Costo_Producto_Calc omitida: ${texto_(result.reason)}`
        : `Costo_Producto_Calc sincronizada con ${result.rows} fila(s). ` +
          `Nuevas=${result.inserted}, actualizadas=${result.updated}.`
    );
  }

  return result;
}

function reconstruirVentaDetalleCostosCalculadoQTAS(payload) {
  const settings = Object.assign({
    silent: false
  }, payload || {});
  asegurarModeloOperativoQTAS_();
  const result = reconstruirVentaDetalleCostosCalculadoInternoQTAS_(settings);

  if (!settings.silent) {
    maybeAlert_(
      result.skipped
        ? `Venta_Detalle_Costos_Calc omitida: ${texto_(result.reason)}`
        : `Venta_Detalle_Costos_Calc sincronizada con ${result.rows} fila(s). ` +
          `Nuevas=${result.inserted}, actualizadas=${result.updated}.`
    );
  }

  return result;
}

function reconstruirAnaliticaCostosQTAS(payload) {
  const settings = Object.assign({
    fechaBase: new Date(),
    silent: false
  }, payload || {});
  asegurarModeloOperativoQTAS_();

  const ss = SpreadsheetApp.getActive();
  const ahora = new Date();
  const costosCache = cargarCostosEnMemoria_();
  const componentes = leerComponentesProductoActivosQTAS_();
  const reglas = leerReglasCostoProductoActivasQTAS_();
  const shared = {
    ss: ss,
    fechaBase: settings.fechaBase,
    ahora: ahora,
    costosCache: costosCache,
    componentes: componentes,
    reglas: reglas
  };

  const costoProducto = reconstruirCostoProductoCalculadoInternoQTAS_(shared);
  const ventaDetalle = reconstruirVentaDetalleCostosCalculadoInternoQTAS_(shared);

  if (!settings.silent) {
    maybeAlert_(
      `Analitica de costos sincronizada. ` +
      `Costo_Producto_Calc=${costoProducto.rows}, ` +
      `Venta_Detalle_Costos_Calc=${ventaDetalle.rows}.`
    );
  }

  return {
    ok: true,
    costoProducto: costoProducto,
    ventaDetalle: ventaDetalle
  };
}

function reconstruirAnaliticaCostosQTAS_Log() {
  const result = reconstruirAnaliticaCostosQTAS({ silent: true });
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function reconstruirCostoProductoCalculadoInternoQTAS_(payload) {
  const settings = Object.assign({
    fechaBase: new Date(),
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const productosRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.productos);
  const outputRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.costoProductoCalculado);
  if (!productosRef.ok || !outputRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: unirUnicos_([productosRef.reason, outputRef.reason]),
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0,
      fechaBase: fechaInput_(settings.fechaBase)
    };
  }

  const contexto = construirContextoAnaliticaCostosQTAS_(settings);
  const rows = leerObjetos_(productosRef.sheet)
    .filter(row => estaActivo_(row.Activo))
    .map(row => construirFilaCostoProductoCalculadoQTAS_(row, contexto));
  const upsert = upsertObjetosPorIdQTAS_(
    outputRef.sheet,
    outputRef.headers,
    rows,
    'Costo_Producto_ID'
  );

  return {
    ok: true,
    skipped: false,
    rows: rows.length,
    inserted: upsert.inserted,
    updated: upsert.updated,
    stale: upsert.stale,
    fechaBase: fechaInput_(contexto.fechaBase)
  };
}

function reconstruirVentaDetalleCostosCalculadoInternoQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const detalleRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.detalle);
  const outputRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.ventaDetalleCostosCalculado);
  if (!detalleRef.ok || !outputRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: unirUnicos_([detalleRef.reason, outputRef.reason]),
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0
    };
  }

  const contexto = construirContextoAnaliticaCostosQTAS_(settings);
  const rows = leerObjetos_(detalleRef.sheet)
    .filter(row => texto_(row.Detalle_ID))
    .map(row => construirFilaVentaDetalleCostoCalculadoQTAS_(row, contexto));
  const upsert = upsertObjetosPorIdQTAS_(
    outputRef.sheet,
    outputRef.headers,
    rows,
    'Detalle_Costo_ID'
  );

  return {
    ok: true,
    skipped: false,
    rows: rows.length,
    inserted: upsert.inserted,
    updated: upsert.updated,
    stale: upsert.stale
  };
}

function sincronizarVentaDetalleCostosLoteQTAS_(detalleRows, payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const rowsFuente = (detalleRows || []).filter(row => texto_(row && row.Detalle_ID));
  if (!rowsFuente.length) {
    return {
      ok: true,
      skipped: false,
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0
    };
  }

  const ss = settings.ss || SpreadsheetApp.getActive();
  const outputRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.ventaDetalleCostosCalculado);
  if (!outputRef.ok) {
    return {
      ok: false,
      skipped: true,
      reason: outputRef.reason,
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0
    };
  }

  const contexto = construirContextoAnaliticaCostosQTAS_(settings);
  const rows = rowsFuente.map(row => construirFilaVentaDetalleCostoCalculadoQTAS_(row, contexto));
  const upsert = upsertObjetosLotePorIdQTAS_(
    outputRef.sheet,
    outputRef.headers,
    rows,
    'Detalle_Costo_ID'
  );

  return {
    ok: true,
    skipped: false,
    rows: rows.length,
    inserted: upsert.inserted,
    updated: upsert.updated,
    stale: upsert.stale
  };
}

function construirContextoAnaliticaCostosQTAS_(payload) {
  const settings = Object.assign({
    fechaBase: new Date(),
    ahora: new Date()
  }, payload || {});
  return {
    fechaBase: resolverFechaOperacion_(settings.fechaBase, new Date()),
    ahora: settings.ahora instanceof Date ? settings.ahora : new Date(),
    costosCache: settings.costosCache || cargarCostosEnMemoria_(),
    componentes: settings.componentes || leerComponentesProductoActivosQTAS_(),
    reglas: settings.reglas || leerReglasCostoProductoActivasQTAS_()
  };
}

function construirFilaCostoProductoCalculadoQTAS_(row, context) {
  const producto = texto_(row && row.Producto_Estandar);
  const unidadVenta = normalizarUnidadCanonicaQTAS_(row && row.Unidad_Default);
  const costo = calcularCostoProductoEnFechaQTAS_(producto, unidadVenta, context.fechaBase, {
    cantidadVenta: 1,
    incluirReglasPorLinea: false,
    incluirReglasPorUnidad: true,
    costosCache: context.costosCache,
    componentes: context.componentes,
    reglas: context.reglas
  });

  return {
    Costo_Producto_ID: `CP-${normalizarClaveProductoQTAS_(producto, unidadVenta)}`,
    Fecha_Calculo: context.ahora,
    Producto_Estandar: producto,
    Unidad_Venta: unidadVenta,
    Metodo_Costo: costo.metodoCosto,
    Costo_Unitario_Total: costo.costoUnitarioTotal,
    Costo_Unitario_Componentes: costo.costoUnitarioComponentes,
    Componentes_Activos: costo.componentesActivos,
    Componentes_Con_Costo: costo.componentesConCosto,
    Componentes_Sin_Costo: costo.componentesSinCosto,
    Cobertura_Costo_Pct: costo.coberturaCostoPct,
    Estado_Costo: costo.estadoCosto,
    Nota: costo.nota
  };
}

function construirFilaVentaDetalleCostoCalculadoQTAS_(row, context) {
  const fechaVenta = valorFechaVentaCanonicaQTAS_(row, context.ahora);
  const contextoCosto = resolverContextoCostoVentaLegacyQTAS_(row);
  const producto = contextoCosto.producto;
  const unidad = contextoCosto.unidad;
  const cantidad = contextoCosto.cantidad;
  const subtotalNeto = redondear_(numero_(row.Subtotal_Neto));
  let costo = contextoCosto.excluirCosto
    ? construirResultadoCostoNoCosteableQTAS_(contextoCosto.motivoExclusion)
    : calcularCostoProductoEnFechaQTAS_(producto, unidad, fechaVenta, {
      cantidadVenta: cantidad,
      incluirReglasPorLinea: true,
      incluirReglasPorUnidad: true,
      costosCache: context.costosCache,
      componentes: context.componentes,
      reglas: context.reglas
    });
  if (contextoCosto.notaNormalizacion) {
    costo = Object.assign({}, costo, {
      nota: unirUnicos_([contextoCosto.notaNormalizacion, costo.nota])
    });
  }
  const costoTotal = redondear_(costo.costoTotalLineaUsado);
  const margenBruto = redondear_(subtotalNeto - costoTotal);
  const margenPct = subtotalNeto > 0
    ? redondear_(margenBruto * 100 / subtotalNeto)
    : 0;

  return {
    Detalle_Costo_ID: `VDC-${texto_(row.Detalle_ID)}`,
    Detalle_ID: texto_(row.Detalle_ID),
    Venta_ID: numero_(row.Venta_ID),
    Fecha_Venta: fechaVenta,
    Cliente_ID: texto_(row.Cliente_ID),
    Nombre: texto_(row.Nombre),
    Producto_Estandar: producto,
    Cantidad: cantidad,
    Unidad: unidad,
    Subtotal_Neto: subtotalNeto,
    Costo_Unitario_Usado: costo.costoUnitarioTotal,
    Costo_Total_Estimado: costoTotal,
    Margen_Bruto_Estimado: margenBruto,
    Margen_Porcentaje_Estimado: margenPct,
    Metodo_Costo: costo.metodoCosto,
    Componentes_Con_Costo: costo.componentesConCosto,
    Componentes_Sin_Costo: costo.componentesSinCosto,
    Cobertura_Costo_Pct: costo.coberturaCostoPct,
    Estado_Costo: costo.estadoCosto,
    Actualizado_En: context.ahora,
    Nota: costo.nota
  };
}

function resolverHojaCanonicaOperativaQTAS_(ss, sheetName) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return {
      ok: false,
      sheet: null,
      headers: [],
      reason: `Falta la hoja ${sheetName}.`
    };
  }

  const headers = getHeaders_(sheet);
  if (!headersIguales_(headers, QTAS.schemas[sheetName])) {
    return {
      ok: false,
      sheet: sheet,
      headers: headers,
      reason: `La hoja ${sheetName} no coincide con la estructura esperada.`
    };
  }

  return {
    ok: true,
    sheet: sheet,
    headers: headers,
    reason: ''
  };
}

function resolverContextoCostoVentaLegacyQTAS_(row) {
  const productoOriginal = texto_(row && row.Producto_Estandar);
  const unidadOriginal = normalizarUnidadCanonicaQTAS_(row && row.Unidad);
  const cantidad = redondear_(numero_(row && row.Cantidad));
  const aliasMap = {
    excordy: 'CordyExt'
  };
  const productosGramos = [
    'AcAlt',
    'AcMed',
    'AcSup',
    'ColaDP',
    'ColaDPPow',
    'Cordy',
    'CordyPow',
    'Gano',
    'GanoPow',
    'Lm',
    'LmPow',
    'Shii',
    'ShiiPow'
  ];
  const productosUnidades = [
    '50mg',
    '100mg',
    '150mg',
    '200mg',
    '300mg',
    '500mg',
    'Choco',
    'Chocordy',
    'ColaDPExt',
    'CordyExt',
    'GanoExt',
    'LmExt',
    'ShiiExt',
    'Tin'
  ];

  let producto = productoOriginal;
  let unidad = unidadOriginal;
  const notas = [];
  const productoKeyOriginal = normalizarClaveTexto_(productoOriginal);
  const alias = aliasMap[productoKeyOriginal];

  if (alias && alias !== productoOriginal) {
    producto = alias;
    notas.push(`Producto legacy normalizado de ${productoOriginal} a ${producto}.`);
  }

  if (productoPerteneceAListaQTAS_(producto, productosGramos) && unidad !== 'g') {
    unidad = 'g';
    notas.push(`Unidad legacy normalizada a g para ${producto}.`);
  } else if (productoPerteneceAListaQTAS_(producto, productosUnidades) && unidad !== 'und') {
    unidad = 'und';
    notas.push(`Unidad legacy normalizada a und para ${producto}.`);
  }

  const motivoExclusion = esProductoNoCosteableAnaliticaQTAS_(producto)
    ? `Producto legacy no costeable (${producto}); excluido de margen/costo automatizado.`
    : '';

  return {
    producto: producto,
    unidad: unidad,
    cantidad: cantidad,
    excluirCosto: Boolean(motivoExclusion),
    motivoExclusion: motivoExclusion,
    notaNormalizacion: notas.join(' ')
  };
}

function productoPerteneceAListaQTAS_(producto, values) {
  const key = normalizarClaveTexto_(producto);
  return (values || []).some(value => normalizarClaveTexto_(value) === key);
}

function esProductoNoCosteableAnaliticaQTAS_(producto) {
  return productoPerteneceAListaQTAS_(producto, ['Feria', 'Vino', 'Kit']);
}

function construirResultadoCostoNoCosteableQTAS_(nota) {
  return {
    costoUnitarioTotal: 0,
    costoUnitarioComponentes: 0,
    costoTotalLineaUsado: 0,
    componentesActivos: 0,
    componentesConCosto: 0,
    componentesSinCosto: 0,
    coberturaCostoPct: 0,
    estadoCosto: 'No costeable',
    metodoCosto: 'Excluido analitica',
    nota: texto_(nota)
  };
}

function calcularCostoProductoEnFechaQTAS_(producto, unidadVenta, fechaBase, options) {
  const settings = Object.assign({
    cantidadVenta: 1,
    incluirReglasPorLinea: true,
    incluirReglasPorUnidad: true,
    stack: []
  }, options || {});
  const unidadVentaCanonica = normalizarUnidadCanonicaQTAS_(unidadVenta);
  const cantidadVenta = Math.max(0, redondear_(numero_(settings.cantidadVenta)));
  if (cantidadVenta <= 0) {
    return {
      costoUnitarioTotal: 0,
      costoUnitarioComponentes: 0,
      costoTotalLineaUsado: 0,
      componentesActivos: 0,
      componentesConCosto: 0,
      componentesSinCosto: 0,
      coberturaCostoPct: 0,
      estadoCosto: 'Sin cantidad',
      metodoCosto: 'Sin costo',
      nota: 'La cantidad de venta es cero.'
    };
  }
  if (esProductoNoCosteableAnaliticaQTAS_(producto)) {
    return construirResultadoCostoNoCosteableQTAS_(
      `Producto legacy no costeable (${producto}); excluido de margen/costo automatizado.`
    );
  }
  const todosLosComponentes = settings.componentes || leerComponentesProductoActivosQTAS_();
  const todasLasReglas = settings.reglas || leerReglasCostoProductoActivasQTAS_();
  const componentes = todosLosComponentes.filter(row =>
    esMismaClaveProductoQTAS_(row.producto, row.unidadVenta, producto, unidadVentaCanonica)
  );
  const reglasSeleccionadas = seleccionarReglasCostoProductoQTAS_(
    producto,
    unidadVentaCanonica,
    fechaBase,
    cantidadVenta,
    todasLasReglas
  );
  const reglasAplicadas = reglasSeleccionadas.filter(row =>
    !(
      (row.aplicacion === 'PorLinea' && settings.incluirReglasPorLinea !== true) ||
      (row.aplicacion === 'PorUnidad' && settings.incluirReglasPorUnidad !== true)
    )
  );
  const reglasPorLineaOmitidas = reglasSeleccionadas.filter(row =>
    row.aplicacion === 'PorLinea' && settings.incluirReglasPorLinea !== true
  ).length;
  const costosCache = settings.costosCache || cargarCostosEnMemoria_();
  const directCost = redondear_(obtenerCostoVigenteDesdeCache_(
    costosCache,
    producto,
    unidadVentaCanonica,
    fechaBase,
    'Producto'
  ));
  const partidaMultiplicadoraBase = cantidadVenta > 0 ? cantidadVenta : 1;
  const stack = (settings.stack || []).slice();

  const partidas = [];

  componentes.forEach(row => {
    partidas.push(construirDesgloseComponenteCostoQTAS_({
      productoActual: producto,
      unidadVentaActual: unidadVentaCanonica,
      fechaBase: fechaBase,
      costosCache: costosCache,
      componentes: todosLosComponentes,
      reglas: todasLasReglas,
      stack: stack,
      tipoComponente: row.tipoComponente,
      itemComponente: row.itemComponente,
      cantidadComponente: row.cantidadComponente,
      unidadComponente: row.unidadComponente,
      mermaPct: row.mermaPct,
      multiplicador: partidaMultiplicadoraBase,
      origen: 'Receta',
      aplicacion: 'PorUnidad'
    }));
  });

  reglasAplicadas.forEach(row => {
    partidas.push(construirDesgloseComponenteCostoQTAS_({
      productoActual: producto,
      unidadVentaActual: unidadVentaCanonica,
      fechaBase: fechaBase,
      costosCache: costosCache,
      componentes: todosLosComponentes,
      reglas: todasLasReglas,
      stack: stack,
      tipoComponente: row.tipoComponente,
      itemComponente: row.itemComponente,
      cantidadComponente: row.cantidadComponente,
      unidadComponente: row.unidadComponente,
      mermaPct: row.mermaPct,
      multiplicador: row.aplicacion === 'PorLinea' ? 1 : partidaMultiplicadoraBase,
      origen: 'Regla',
      aplicacion: row.aplicacion
    }));
  });

  if (!partidas.length) {
    if (directCost > 0) {
      const costoTotalDirecto = redondear_(cantidadVenta > 0 ? cantidadVenta * directCost : 0);
      return {
        costoUnitarioTotal: directCost,
        costoUnitarioComponentes: 0,
        costoTotalLineaUsado: costoTotalDirecto,
        componentesActivos: 0,
        componentesConCosto: 0,
        componentesSinCosto: 0,
        coberturaCostoPct: 100,
        estadoCosto: 'Directo',
        metodoCosto: 'Costo directo',
        nota: 'Sin receta activa; se usa costo directo vigente.'
      };
    }

    return {
      costoUnitarioTotal: 0,
      costoUnitarioComponentes: 0,
      costoTotalLineaUsado: 0,
      componentesActivos: 0,
      componentesConCosto: 0,
      componentesSinCosto: 0,
      coberturaCostoPct: 0,
      estadoCosto: 'Sin receta',
      metodoCosto: 'Sin costo',
      nota: 'No existe receta activa ni costo directo vigente.'
    };
  }

  const componentesActivos = partidas.length;
  const componentesConCosto = partidas.filter(item => item.tieneCosto).length;
  const componentesSinCosto = Math.max(componentesActivos - componentesConCosto, 0);
  const costoTotalComponentes = redondear_(sumar_(partidas.map(item => item.costoTotal)));
  const costoUnitarioComponentes = cantidadVenta > 0
    ? redondear_(costoTotalComponentes / cantidadVenta)
    : costoTotalComponentes;
  const coberturaCostoPct = componentesActivos
    ? redondear_(componentesConCosto * 100 / componentesActivos)
    : 0;
  const hayReglas = reglasAplicadas.length > 0 || reglasPorLineaOmitidas > 0;

  if (componentesSinCosto === 0) {
    return {
      costoUnitarioTotal: costoUnitarioComponentes,
      costoUnitarioComponentes: costoUnitarioComponentes,
      costoTotalLineaUsado: costoTotalComponentes,
      componentesActivos: componentesActivos,
      componentesConCosto: componentesConCosto,
      componentesSinCosto: componentesSinCosto,
      coberturaCostoPct: coberturaCostoPct,
      estadoCosto: 'Completo',
      metodoCosto: hayReglas ? 'Receta + reglas' : 'Receta',
      nota: construirNotaCostoProductoQTAS_(reglasAplicadas.length, reglasPorLineaOmitidas, false, componentesSinCosto)
    };
  }

  if (directCost > 0) {
    return {
      costoUnitarioTotal: directCost,
      costoUnitarioComponentes: costoUnitarioComponentes,
      costoTotalLineaUsado: redondear_(cantidadVenta > 0 ? cantidadVenta * directCost : 0),
      componentesActivos: componentesActivos,
      componentesConCosto: componentesConCosto,
      componentesSinCosto: componentesSinCosto,
      coberturaCostoPct: coberturaCostoPct,
      estadoCosto: 'Fallback directo',
      metodoCosto: 'Costo directo fallback',
      nota: construirNotaCostoProductoQTAS_(reglasAplicadas.length, reglasPorLineaOmitidas, true, componentesSinCosto)
    };
  }

  return {
    costoUnitarioTotal: costoUnitarioComponentes,
    costoUnitarioComponentes: costoUnitarioComponentes,
    costoTotalLineaUsado: costoTotalComponentes,
    componentesActivos: componentesActivos,
    componentesConCosto: componentesConCosto,
    componentesSinCosto: componentesSinCosto,
    coberturaCostoPct: coberturaCostoPct,
    estadoCosto: componentesConCosto > 0 ? 'Parcial' : 'Sin costo',
    metodoCosto: hayReglas ? 'Receta + reglas parcial' : 'Receta parcial',
    nota: construirNotaCostoProductoQTAS_(reglasAplicadas.length, reglasPorLineaOmitidas, false, componentesSinCosto)
  };
}

function leerComponentesProductoActivosQTAS_() {
  return listarComponentesProductoQTAS_()
    .filter(row => row.activo)
    .filter(row => row.producto && row.itemComponente && row.unidadVenta && row.unidadComponente);
}

function leerReglasCostoProductoActivasQTAS_() {
  return listarReglasCostoProductoQTAS_()
    .filter(row => row.activo)
    .filter(row => row.producto && row.itemComponente && row.unidadVenta && row.unidadComponente);
}

function validarProductoCanonicoExistenteQTAS_(ss, producto) {
  const sheet = (ss || SpreadsheetApp.getActive()).getSheetByName(QTAS.sheets.productos);
  const existe = leerObjetos_(sheet).some(row =>
    normalizarClaveTexto_(row.Producto_Estandar) === normalizarClaveTexto_(producto)
  );
  if (!existe) {
    throw new Error(`El producto ${producto} no existe en Productos.`);
  }
}

function normalizarClaveProductoQTAS_(producto, unidad) {
  return [
    texto_(producto),
    texto_(unidad)
  ]
    .join('-')
    .replace(/[^A-Za-z0-9_-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80);
}

function normalizarAplicacionReglaCostoQTAS_(value) {
  const key = normalizarClaveTexto_(value);
  if (!key) return 'PorUnidad';
  if (['por linea', 'linea', 'por pedido', 'pedido'].includes(key)) return 'PorLinea';
  if (['por unidad', 'unidad', 'unitario', 'por und'].includes(key)) return 'PorUnidad';
  return key === 'porlinea' ? 'PorLinea' : 'PorUnidad';
}

function seleccionarReglasCostoProductoQTAS_(producto, unidadVenta, fechaBase, cantidadVenta, reglas) {
  return (reglas || [])
    .filter(row =>
      esMismaClaveProductoQTAS_(row.producto, row.unidadVenta, producto, unidadVenta) &&
      reglaFechaAplicaCostoQTAS_(row, fechaBase) &&
      reglaCantidadAplicaCostoQTAS_(row, cantidadVenta)
    )
    .sort((a, b) => {
      if (a.orden !== b.orden) return a.orden - b.orden;
      if (a.aplicacion !== b.aplicacion) return a.aplicacion.localeCompare(b.aplicacion);
      return a.itemComponente.localeCompare(b.itemComponente);
    });
}

function reglaFechaAplicaCostoQTAS_(row, fechaBase) {
  const fecha = resolverFechaOperacion_(fechaBase, new Date());
  const desde = row && row.fechaDesde ? resolverFechaOperacion_(row.fechaDesde, fecha) : null;
  const hasta = row && row.fechaHasta ? resolverFechaOperacion_(row.fechaHasta, fecha) : null;
  if (desde && fecha < desde) return false;
  if (hasta && fecha > hasta) return false;
  return true;
}

function reglaCantidadAplicaCostoQTAS_(row, cantidadVenta) {
  const cantidad = Math.max(0, numero_(cantidadVenta));
  const min = row && row.cantidadMin !== '' ? numero_(row.cantidadMin) : 0;
  const hasMax = row && row.cantidadMax !== '';
  const max = hasMax ? numero_(row.cantidadMax) : 0;
  if (cantidad < min) return false;
  if (hasMax && cantidad > max) return false;
  return true;
}

function construirDesgloseComponenteCostoQTAS_(context) {
  const costoUnitarioBase = redondear_(resolverCostoUnitarioComponenteQTAS_(context));
  const factorMerma = 1 + Math.max(0, numero_(context.mermaPct)) / 100;
  const cantidadAplicada = redondear_(numero_(context.cantidadComponente) * numero_(context.multiplicador) * factorMerma);
  return {
    origen: texto_(context.origen),
    aplicacion: texto_(context.aplicacion),
    itemComponente: texto_(context.itemComponente),
    tipoComponente: normalizarTipoCompraItemQTAS_(context.tipoComponente),
    cantidadAplicada: cantidadAplicada,
    unidadComponente: normalizarUnidadCanonicaQTAS_(context.unidadComponente),
    costoUnitarioBase: costoUnitarioBase,
    costoTotal: redondear_(cantidadAplicada * costoUnitarioBase),
    tieneCosto: costoUnitarioBase > 0
  };
}

function resolverCostoUnitarioComponenteQTAS_(context) {
  const tipoComponente = normalizarTipoCompraItemQTAS_(context.tipoComponente);
  const itemComponente = texto_(context.itemComponente);
  const unidadComponente = normalizarUnidadCanonicaQTAS_(context.unidadComponente);
  const fechaBase = resolverFechaOperacion_(context.fechaBase, new Date());
  const costosCache = context.costosCache || [];
  const costoDirecto = redondear_(obtenerCostoVigenteDesdeCache_(
    costosCache,
    itemComponente,
    unidadComponente,
    fechaBase,
    tipoComponente
  ));

  if (normalizarClaveTexto_(tipoComponente) !== normalizarClaveTexto_('Producto')) {
    return costoDirecto;
  }

  if (esMismaClaveProductoQTAS_(
    itemComponente,
    unidadComponente,
    context.productoActual,
    context.unidadVentaActual
  )) {
    return costoDirecto;
  }

  const componentKey = normalizarClaveProductoQTAS_(itemComponente, unidadComponente);
  const currentKey = normalizarClaveProductoQTAS_(context.productoActual, context.unidadVentaActual);
  const stack = (context.stack || []).slice();
  if (stack.indexOf(componentKey) >= 0) {
    return costoDirecto;
  }

  const costoRecursivo = calcularCostoProductoEnFechaQTAS_(itemComponente, unidadComponente, fechaBase, {
    cantidadVenta: 1,
    incluirReglasPorLinea: false,
    incluirReglasPorUnidad: true,
    costosCache: costosCache,
    componentes: context.componentes || leerComponentesProductoActivosQTAS_(),
    reglas: context.reglas || leerReglasCostoProductoActivasQTAS_(),
    stack: stack.concat([currentKey])
  });

  return costoRecursivo && costoRecursivo.costoUnitarioTotal > 0
    ? redondear_(costoRecursivo.costoUnitarioTotal)
    : costoDirecto;
}

function construirNotaCostoProductoQTAS_(reglasAplicadas, reglasPorLineaOmitidas, usoFallbackDirecto, componentesSinCosto) {
  const notas = [];
  if (reglasAplicadas > 0) {
    notas.push(`Se aplicaron ${reglasAplicadas} regla(s) variables.`);
  }
  if (reglasPorLineaOmitidas > 0) {
    notas.push(`Se omitieron ${reglasPorLineaOmitidas} regla(s) PorLinea en este calculo de referencia.`);
  }
  if (usoFallbackDirecto) {
    notas.push(`Receta incompleta; se uso costo directo vigente. Faltan ${componentesSinCosto} componente(s).`);
  } else if (componentesSinCosto > 0) {
    notas.push(`Faltan ${componentesSinCosto} componente(s) sin costo vigente.`);
  } else if (!notas.length) {
    notas.push('Costo calculado desde receta activa.');
  }
  return notas.join(' ');
}

function esMismaClaveProductoQTAS_(productoA, unidadA, productoB, unidadB) {
  return normalizarClaveTexto_(productoA) === normalizarClaveTexto_(productoB) &&
    normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(unidadA)) ===
      normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(unidadB));
}

function upsertObjetosPorIdQTAS_(sheet, headers, objects, idHeader) {
  const rowsExistentes = leerObjetosConMeta_(sheet);
  const existentesPorId = {};
  rowsExistentes.forEach(row => {
    const id = texto_(row[idHeader]);
    if (!id || existentesPorId[id]) return;
    existentesPorId[id] = true;
  });

  const sincronizados = [];
  const incomingIds = {};
  let inserted = 0;
  let updated = 0;

  (objects || []).forEach(obj => {
    const id = texto_(obj && obj[idHeader]);
    if (!id) return;
    if (incomingIds[id]) return;
    incomingIds[id] = true;
    sincronizados.push(filaDesdeHeaders_(headers, obj));

    if (existentesPorId[id]) {
      updated++;
      return;
    }
    inserted++;
  });

  const bodyRowsExistentes = Math.max((sheet ? sheet.getLastRow() : 1) - 1, 0);
  if (sincronizados.length) {
    escribirFilasDesdeFilaQTAS_(sheet, 2, sincronizados);
  }
  if (bodyRowsExistentes > sincronizados.length) {
    sheet
      .getRange(2 + sincronizados.length, 1, bodyRowsExistentes - sincronizados.length, headers.length)
      .clearContent();
  }

  const stale = rowsExistentes.filter(row => {
    const id = texto_(row[idHeader]);
    return id && !incomingIds[id];
  }).length;

  return {
    inserted: inserted,
    updated: updated,
    stale: stale
  };
}

function upsertObjetosLotePorIdQTAS_(sheet, headers, objects, idHeader) {
  const rowsExistentes = leerObjetosConMeta_(sheet);
  const existentesPorId = {};
  rowsExistentes.forEach(row => {
    const id = texto_(row[idHeader]);
    if (!id || existentesPorId[id]) return;
    existentesPorId[id] = row;
  });

  const nuevos = [];
  const procesados = {};
  let inserted = 0;
  let updated = 0;

  (objects || []).forEach(obj => {
    const id = texto_(obj && obj[idHeader]);
    if (!id || procesados[id]) return;
    procesados[id] = true;

    const existente = existentesPorId[id];
    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, obj));
      updated++;
      return;
    }

    nuevos.push(filaDesdeHeaders_(headers, obj));
    inserted++;
  });

  if (nuevos.length) {
    escribirFilas_(sheet, nuevos);
  }

  return {
    inserted: inserted,
    updated: updated,
    stale: 0
  };
}

function sincronizarPlantillaCosteoBaseQTAS() {
  return withScriptLock_('sincronizar plantilla costeo base', () =>
    sincronizarPlantillaCosteoBaseQTAS_()
  );
}

function sincronizarPlantillaCosteoBaseQTAS_Log() {
  const result = sincronizarPlantillaCosteoBaseQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function sembrarCostosDirectosBaseQTAS(payload) {
  return withScriptLock_('sembrar costos directos base', () => {
    asegurarModeloOperativoQTAS_({ aplicarFormatos: true });

    const settings = Object.assign({
      fechaDesde: new Date(),
      proveedor: 'Referencia directa inicial'
    }, payload || {});
    const fechaDesde = resolverFechaOperacion_(settings.fechaDesde, new Date());
    const proveedor = texto_(settings.proveedor) || 'Referencia directa inicial';
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const seeds = construirCostosDirectosBaseSUPERQTAS_()
      .concat(construirCostosDirectosDerivadosPowSUPERQTAS_(fechaDesde));
    let actualizados = 0;

    seeds.forEach(seed => {
      actualizados += upsertCostoReferenciaHistoricoQTAS_({
        sheet: sheet,
        headers: headers,
        rows: rows,
        fechaDesde: fechaDesde,
        proveedor: proveedor,
        compraId: 0,
        fuenteTipo: 'Directo',
        fuenteId: construirFuenteDirectaCostoQTAS_(seed),
        comentario: texto_(seed.nota),
        nota: texto_(seed.referencia),
        linea: {
          Tipo_Item: seed.tipoItem,
          Item: seed.item,
          Unidad: seed.unidad,
          Costo_Unitario: seed.costoUnitario,
          Comentario_Linea: seed.nota
        }
      });
    });
    actualizados += asegurarCoberturaVigenteCostosSembradosQTAS_(
      sheet,
      headers,
      rows,
      seeds,
      fechaDesde,
      proveedor
    );
    const reparacionHistorica = repararCoberturaHistoricaCostosQTAS_({
      fechaDesde: fechaDesde,
      proveedor: proveedor,
      silent: true
    });
    actualizados += numero_(reparacionHistorica && reparacionHistorica.actualizados);

    limpiarCachesEjecucionQTAS_();

    return {
      ok: true,
      fechaDesde: fechaInput_(fechaDesde),
      proveedor: proveedor,
      costosSembrados: seeds.map(seed => ({
        item: seed.item,
        tipoItem: seed.tipoItem,
        unidad: seed.unidad,
        costoUnitario: seed.costoUnitario,
        referencia: seed.referencia
      })),
      actualizados: actualizados,
      reparacionHistorica: reparacionHistorica,
      cobertura: resumenCoberturaPlantillaCostoQTAS_(
        construirRecetasBaseSUPERQTAS_(),
        construirReglasCostoSUPERQTAS_()
      )
    };
  });
}

function sembrarCostosDirectosBaseQTAS_Log() {
  const result = sembrarCostosDirectosBaseQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function sembrarCostosDirectosBaseQTAS_Historico() {
  return sembrarCostosDirectosBaseQTAS({
    fechaDesde: '2024-01-01',
    proveedor: 'Referencia directa historica'
  });
}

function sembrarCostosDirectosBaseQTAS_Historico_Log() {
  const result = sembrarCostosDirectosBaseQTAS_Historico();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function repararCoberturaHistoricaCostosQTAS(payload) {
  return withScriptLock_('reparar cobertura historica costos', () =>
    repararCoberturaHistoricaCostosQTAS_(payload)
  );
}

function repararCoberturaHistoricaCostosQTAS_Log() {
  const result = repararCoberturaHistoricaCostosQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function sincronizarPlantillaYAnaliticaCostosQTAS() {
  return withScriptLock_('sincronizar plantilla y analitica costos', () => {
    const plantilla = sincronizarPlantillaCosteoBaseQTAS_();
    const analitica = reconstruirAnaliticaCostosQTAS({ silent: true });
    return {
      ok: true,
      plantilla: plantilla,
      analitica: analitica
    };
  });
}

function sincronizarPlantillaYAnaliticaCostosQTAS_Log() {
  const result = sincronizarPlantillaYAnaliticaCostosQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function sincronizarPlantillaCosteoBaseQTAS_() {
  asegurarModeloOperativoQTAS_({ aplicarFormatos: true });

  const ss = SpreadsheetApp.getActive();
  const componentesSheet = ss.getSheetByName(QTAS.sheets.productoComponentes);
  const reglasSheet = ss.getSheetByName(QTAS.sheets.productoReglasCosto);
  const componentesHeaders = getHeaders_(componentesSheet);
  const reglasHeaders = getHeaders_(reglasSheet);
  const productosDisponibles = {};
  leerObjetos_(ss.getSheetByName(QTAS.sheets.productos)).forEach(row => {
    const key = normalizarClaveTexto_(row.Producto_Estandar);
    if (key) productosDisponibles[key] = true;
  });

  const componentesPlantillaBruta = construirRecetasBaseSUPERQTAS_();
  const reglasPlantillaBruta = construirReglasCostoSUPERQTAS_();
  const productosFaltantes = {};
  const componentesPlantilla = componentesPlantillaBruta.filter(row => {
    const existe = productosDisponibles[normalizarClaveTexto_(row.Producto_Estandar)] === true;
    if (!existe) productosFaltantes[texto_(row.Producto_Estandar)] = true;
    return existe;
  });
  const reglasPlantilla = reglasPlantillaBruta.filter(row => {
    const existe = productosDisponibles[normalizarClaveTexto_(row.Producto_Estandar)] === true;
    if (!existe) productosFaltantes[texto_(row.Producto_Estandar)] = true;
    return existe;
  });

  const componentesResult = upsertPlantillaComponentesQTAS_(
    componentesSheet,
    componentesHeaders,
    componentesPlantilla
  );
  const reglasResult = upsertPlantillaReglasCostoQTAS_(
    reglasSheet,
    reglasHeaders,
    reglasPlantilla
  );
  const cobertura = resumenCoberturaPlantillaCostoQTAS_(componentesPlantilla, reglasPlantilla);

  return {
    ok: true,
    componentes: componentesResult,
    reglas: reglasResult,
    cobertura: cobertura,
    assumptions: [
      'Cordy se incluyo en la familia de bolsa cara junto con CordyPow.',
      'Las reglas PorLinea se usan para empaques que dependen del pedido; las PorUnidad para piezas individuales.',
      'Alcohol en extractos queda sembrado como placeholder en g para que luego ajustes la cantidad real.'
    ],
    productosOmitidos: Object.keys(productosFaltantes).sort((a, b) => a.localeCompare(b))
  };
}

function construirCostosDirectosBaseSUPERQTAS_() {
  return [
    crearCostoDirectoBaseQTAS_(
      'Alcohol',
      'Insumo',
      'g',
      48000 / 3800,
      'Alcohol base para extractos.',
      '48000 / 3800 g'
    ),
    crearCostoDirectoBaseQTAS_(
      'Bolsa_Barata',
      'Insumo',
      'und',
      300,
      'Bolsa barata unitaria.',
      'Valor unitario manual'
    ),
    crearCostoDirectoBaseQTAS_(
      'Bolsa_Media',
      'Insumo',
      'und',
      800,
      'Bolsa media unitaria.',
      'Valor unitario manual'
    ),
    crearCostoDirectoBaseQTAS_(
      'Bolsa_Cara',
      'Insumo',
      'und',
      1100,
      'Bolsa cara unitaria.',
      'Valor unitario manual'
    ),
    crearCostoDirectoBaseQTAS_(
      'Capsulas',
      'Insumo',
      'und',
      15000 / 1000,
      'Capsulas unitarias.',
      '15000 / 1000 und'
    ),
    crearCostoDirectoBaseQTAS_(
      'Chocolate',
      'Insumo',
      'g',
      174000 / 2500,
      'Chocolate por gramo.',
      '174000 / 2500 g'
    ),
    crearCostoDirectoBaseQTAS_(
      'Agua',
      'Insumo',
      'g',
      3000 / 5000,
      'Agua por gramo para referencias futuras.',
      '3000 / 5000 g'
    ),
    crearCostoDirectoBaseQTAS_(
      'Agua',
      'Insumo',
      'und',
      120,
      'Aprox 200 g por tinto desde bidon de 5L.',
      '3000 / 25 tintos aprox'
    ),
    crearCostoDirectoBaseQTAS_(
      'Cafe',
      'Insumo',
      'g',
      35000 / 453.59237,
      'Cafe por gramo a partir de 1 libra.',
      '35000 / 453.59237 g'
    ),
    crearCostoDirectoBaseQTAS_(
      'Cafe',
      'Insumo',
      'und',
      35000 / 45,
      'Aprox 45 tintos por libra usando ~10 g por tinto.',
      '35000 / 45 tintos aprox'
    ),
    crearCostoDirectoBaseQTAS_(
      'Vasos',
      'Insumo',
      'und',
      7000 / 50,
      'Vaso unitario.',
      '7000 / 50 und'
    ),
    crearCostoDirectoBaseQTAS_(
      'Mano de obra',
      'Gasto',
      'und',
      1000,
      'Placeholder razonable mientras no se costea formalmente el trabajo propio.',
      'Valor manual provisional'
    )
  ];
}

function construirCostosDirectosDerivadosPowSUPERQTAS_(fechaBase) {
  const fecha = resolverFechaOperacion_(fechaBase, new Date());
  const historicos = leerCostosHistoricosQTAS_();
  const costoCache = cargarCostosEnMemoria_();

  return [
    { target: 'ColaDPPow', source: 'ColaDP' },
    { target: 'CordyPow', source: 'Cordy' },
    { target: 'GanoPow', source: 'Gano' },
    { target: 'ShiiPow', source: 'Shii' }
  ]
    .map(pair => {
      const costoTarget = redondear_(obtenerCostoVigenteDesdeCache_(
        costoCache,
        pair.target,
        'g',
        fecha,
        'Producto'
      ));
      if (costoTarget > 0) return null;

      const costoSource = redondear_(
        obtenerCostoVigenteDesdeCache_(costoCache, pair.source, 'g', fecha, 'Producto') ||
        obtenerCostoMasRecienteHistoricoQTAS_(historicos, pair.source, 'g', 'Producto')
      );
      if (costoSource <= 0) return null;

      return crearCostoDirectoBaseQTAS_(
        pair.target,
        'Producto',
        'g',
        costoSource,
        `Costo provisional copiado desde ${pair.source} mientras llega costo propio.`,
        `Derivado de ${pair.source} (${costoSource}/g)`
      );
    })
    .filter(Boolean);
}

function asegurarCoberturaVigenteCostosSembradosQTAS_(sheet, headers, rows, seeds, fechaMinima, proveedor) {
  const hoy = resolverFechaOperacion_(new Date(), new Date());
  let actualizados = 0;

  (seeds || []).forEach(seed => {
    const existentes = (rows || [])
      .filter(row =>
        normalizarClaveTexto_(row.Tipo_Item) === normalizarClaveTexto_(seed.tipoItem) &&
        normalizarClaveTexto_(row.Item) === normalizarClaveTexto_(seed.item) &&
        normalizarClaveTexto_(row.Unidad) === normalizarClaveTexto_(seed.unidad) &&
        estaActivo_(row.Activo)
      )
      .map(row => ({
        raw: row,
        desde: resolverFechaOperacion_(row.Fecha_Desde, hoy),
        hasta: row.Fecha_Hasta ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || hoy) : null
      }))
      .sort((a, b) => a.desde - b.desde);

    const vigenteHoy = existentes.some(row =>
      hoy.getTime() >= row.desde.getTime() &&
      (!row.hasta || hoy.getTime() <= row.hasta.getTime()) &&
      numero_(row.raw && row.raw.Costo_Unitario) > 0
    );
    if (vigenteHoy) return;

    const pasados = existentes.filter(row => row.desde.getTime() <= hoy.getTime());
    const ultimoPasado = pasados.length ? pasados[pasados.length - 1] : null;
    const fechaRelleno = ultimoPasado
      ? diaSiguiente_(ultimoPasado.hasta || ultimoPasado.desde)
      : resolverFechaOperacion_(fechaMinima, hoy);
    const fechaDesdeRelleno = fechaRelleno.getTime() > hoy.getTime() ? hoy : fechaRelleno;

    actualizados += upsertCostoReferenciaHistoricoQTAS_({
      sheet: sheet,
      headers: headers,
      rows: rows,
      fechaDesde: fechaDesdeRelleno,
      proveedor: proveedor,
      compraId: 0,
      fuenteTipo: 'Directo',
      fuenteId: `${construirFuenteDirectaCostoQTAS_(seed)}-vigente`,
      comentario: unirUnicos_([
        texto_(seed.nota),
        'Relleno automatico de vigencia actual'
      ]),
      nota: texto_(seed.referencia),
      linea: {
        Tipo_Item: seed.tipoItem,
        Item: seed.item,
        Unidad: seed.unidad,
        Costo_Unitario: seed.costoUnitario,
        Comentario_Linea: texto_(seed.nota)
      }
    });
  });

  return actualizados;
}

function repararCoberturaHistoricaCostosQTAS_(payload) {
  const settings = Object.assign({
    fechaDesde: '2024-01-01',
    proveedor: 'Ajuste historico automatico',
    silent: false
  }, payload || {});
  asegurarModeloOperativoQTAS_({ aplicarFormatos: true });

  const fechaDesde = resolverFechaOperacion_(settings.fechaDesde, new Date());
  const hoy = resolverFechaOperacion_(new Date(), new Date());
  const proveedor = texto_(settings.proveedor) || 'Ajuste historico automatico';
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
  const headers = getHeaders_(sheet);
  const rows = leerObjetosConMeta_(sheet);
  const componentes = leerComponentesProductoActivosQTAS_();
  const reglas = leerReglasCostoProductoActivasQTAS_();
  const itemsObjetivo = construirItemsObjetivoCoberturaHistoricaQTAS_(componentes, reglas);
  const ajustesAplicados = [];
  const itemsSinFuente = [];
  let actualizados = 0;

  itemsObjetivo.forEach(item => {
    const plan = construirPlanCoberturaHistoricaItemQTAS_(rows, item, fechaDesde, hoy);
    if (!plan.tieneFuente) {
      itemsSinFuente.push(`${item.item} (${item.unidad}, ${item.tipoItem})`);
      return;
    }

    plan.ajustes.forEach((ajuste, index) => {
      actualizados += upsertCostoReferenciaHistoricoQTAS_({
        sheet: sheet,
        headers: headers,
        rows: rows,
        fechaDesde: ajuste.fechaDesde,
        proveedor: proveedor,
        compraId: 0,
        fuenteTipo: 'AjusteHistorico',
        fuenteId: construirFuenteAjusteHistoricoCostoQTAS_(item, ajuste.tipo, ajuste.fechaDesde, index),
        comentario: ajuste.comentario,
        nota: ajuste.referencia,
        linea: {
          Tipo_Item: item.tipoItem,
          Item: item.item,
          Unidad: item.unidad,
          Costo_Unitario: ajuste.costoUnitario,
          Comentario_Linea: ajuste.comentario
        }
      });
      ajustesAplicados.push({
        item: item.item,
        tipoItem: item.tipoItem,
        unidad: item.unidad,
        tipoAjuste: ajuste.tipo,
        fechaDesde: fechaInput_(ajuste.fechaDesde),
        costoUnitario: ajuste.costoUnitario,
        referencia: ajuste.referencia
      });
    });
  });

  limpiarCachesEjecucionQTAS_();

  const result = {
    ok: true,
    fechaDesde: fechaInput_(fechaDesde),
    proveedor: proveedor,
    itemsRevisados: itemsObjetivo.length,
    actualizados: actualizados,
    ajustes: ajustesAplicados,
    itemsSinFuenteHistorica: itemsSinFuente.sort((a, b) => a.localeCompare(b)),
    cobertura: resumenCoberturaPlantillaCostoQTAS_(componentes, reglas)
  };

  if (!settings.silent) {
    maybeAlert_(
      `Cobertura historica revisada. ` +
      `Items=${itemsObjetivo.length}, ajustes=${actualizados}, sin fuente=${itemsSinFuente.length}.`
    );
  }

  return result;
}

function construirItemsObjetivoCoberturaHistoricaQTAS_(componentes, reglas) {
  const mapa = {};

  (componentes || []).forEach(row => {
    registrarItemObjetivoCoberturaHistoricaQTAS_(mapa, {
      tipoItem: row.tipoComponente || row.Tipo_Componente,
      item: row.itemComponente || row.Item_Componente,
      unidad: row.unidadComponente || row.Unidad_Componente
    });
  });

  (reglas || []).forEach(row => {
    registrarItemObjetivoCoberturaHistoricaQTAS_(mapa, {
      tipoItem: row.tipoComponente || row.Tipo_Componente,
      item: row.itemComponente || row.Item_Componente,
      unidad: row.unidadComponente || row.Unidad_Componente
    });
  });

  return Object.keys(mapa)
    .map(key => mapa[key])
    .sort((a, b) => {
      if (a.tipoItem !== b.tipoItem) return a.tipoItem.localeCompare(b.tipoItem);
      if (a.item !== b.item) return a.item.localeCompare(b.item);
      return a.unidad.localeCompare(b.unidad);
    });
}

function registrarItemObjetivoCoberturaHistoricaQTAS_(mapa, rawItem) {
  const item = texto_(rawItem && rawItem.item);
  const unidad = normalizarUnidadCanonicaQTAS_(rawItem && rawItem.unidad);
  const tipoItem = normalizarTipoCompraItemQTAS_(rawItem && rawItem.tipoItem);
  if (!item || !unidad || !tipoItem) return;

  const key = [
    normalizarClaveTexto_(tipoItem),
    normalizarClaveTexto_(item),
    normalizarClaveTexto_(unidad)
  ].join('|');
  if (mapa[key]) return;

  mapa[key] = {
    tipoItem: tipoItem,
    item: item,
    unidad: unidad
  };
}

function construirPlanCoberturaHistoricaItemQTAS_(rows, item, fechaDesdeMinima, hoy) {
  const existentes = leerCostosHistoricosPorClaveDesdeRowsQTAS_(rows, item);
  if (!existentes.length) {
    return {
      tieneFuente: false,
      ajustes: []
    };
  }

  const ajustes = [];
  const fechaDesde = resolverFechaOperacion_(fechaDesdeMinima, hoy || new Date());
  const fechaActual = resolverFechaOperacion_(hoy, new Date());
  const primero = existentes[0];

  if (primero.desde.getTime() > fechaDesde.getTime()) {
    ajustes.push({
      tipo: 'BackfillInicial',
      fechaDesde: fechaDesde,
      costoUnitario: primero.costoUnitario,
      comentario: (
        `Piso historico automatico usando el primer costo conocido ` +
        `desde ${fechaInput_(primero.desde)}.`
      ),
      referencia: `${primero.item} ${primero.unidad} | costo base ${primero.costoUnitario}`
    });
  }

  existentes.forEach((actual, index) => {
    const siguiente = existentes[index + 1];
    if (siguiente && actual.hasta) {
      const gapStart = diaSiguiente_(actual.hasta);
      const gapEnd = diaAnterior_(siguiente.desde);
      if (gapStart.getTime() <= gapEnd.getTime()) {
        ajustes.push({
          tipo: 'HuecoIntermedio',
          fechaDesde: gapStart,
          costoUnitario: actual.costoUnitario,
          comentario: (
            `Hueco historico automatico entre ${fechaInput_(gapStart)} ` +
            `y ${fechaInput_(gapEnd)} usando el ultimo costo conocido.`
          ),
          referencia: `${actual.item} ${actual.unidad} | costo base ${actual.costoUnitario}`
        });
      }
    }
  });

  const ultimo = existentes[existentes.length - 1];
  if (ultimo.hasta && ultimo.hasta.getTime() < fechaActual.getTime()) {
    ajustes.push({
      tipo: 'ExtensionVigente',
      fechaDesde: diaSiguiente_(ultimo.hasta),
      costoUnitario: ultimo.costoUnitario,
      comentario: (
        `Extension automatica de vigencia actual usando el ultimo costo conocido ` +
        `desde ${fechaInput_(ultimo.desde)}.`
      ),
      referencia: `${ultimo.item} ${ultimo.unidad} | costo base ${ultimo.costoUnitario}`
    });
  }

  return {
    tieneFuente: true,
    ajustes: ajustes
  };
}

function leerCostosHistoricosPorClaveDesdeRowsQTAS_(rows, item) {
  const tipoItemKey = normalizarClaveTexto_(item && item.tipoItem);
  const itemKey = normalizarClaveTexto_(item && item.item);
  const unidadKey = normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(item && item.unidad));

  return (rows || [])
    .filter(row =>
      normalizarClaveTexto_(row.Tipo_Item) === tipoItemKey &&
      normalizarClaveTexto_(row.Item) === itemKey &&
      normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad)) === unidadKey &&
      estaActivo_(row.Activo) &&
      numero_(row.Costo_Unitario) > 0
    )
    .map(row => ({
      raw: row,
      item: texto_(row.Item),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      costoUnitario: redondear_(numero_(row.Costo_Unitario)),
      desde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
      hasta: row.Fecha_Hasta
        ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date())
        : null
    }))
    .sort((a, b) => {
      if (a.desde.getTime() !== b.desde.getTime()) return a.desde - b.desde;
      return a.costoUnitario - b.costoUnitario;
    });
}

function construirFuenteAjusteHistoricoCostoQTAS_(item, tipoAjuste, fechaDesde, index) {
  const fecha = fechaInput_(fechaDesde);
  const base = [
    'AUTO',
    normalizarClaveTexto_(item && item.tipoItem),
    normalizarClaveProductoQTAS_(item && item.item, item && item.unidad),
    normalizarClaveTexto_(tipoAjuste),
    fecha,
    String((index || 0) + 1)
  ]
    .join('-')
    .replace(/[^A-Za-z0-9_-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return base.slice(0, 99);
}

function obtenerCostoMasRecienteHistoricoQTAS_(rows, item, unidad, tipoItem) {
  return (rows || [])
    .filter(row =>
      normalizarClaveTexto_(row.item) === normalizarClaveTexto_(item) &&
      normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.unidad)) ===
        normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(unidad)) &&
      normalizarClaveTexto_(row.tipoItem) === normalizarClaveTexto_(tipoItem) &&
      numero_(row.costoUnitario) > 0
    )
    .sort((a, b) => b.fechaDesde - a.fechaDesde)
    .map(row => redondear_(numero_(row.costoUnitario)))[0] || 0;
}

function crearCostoDirectoBaseQTAS_(item, tipoItem, unidad, costoUnitario, nota, referencia) {
  return {
    item: texto_(item),
    tipoItem: normalizarTipoCompraItemQTAS_(tipoItem),
    unidad: normalizarUnidadCanonicaQTAS_(unidad),
    costoUnitario: redondear_(numero_(costoUnitario)),
    nota: texto_(nota),
    referencia: texto_(referencia)
  };
}

function construirFuenteDirectaCostoQTAS_(seed) {
  return [
    'DIR',
    normalizarClaveProductoQTAS_(seed && seed.item, seed && seed.unidad)
  ].join('-').slice(0, 99);
}

function construirRecetasBaseSUPERQTAS_() {
  const rows = [];

  [
    { producto: '50mg', cantidadAcMed: 0.05 },
    { producto: '100mg', cantidadAcMed: 0.10 },
    { producto: '150mg', cantidadAcMed: 0.15 },
    { producto: '200mg', cantidadAcMed: 0.20 },
    { producto: '300mg', cantidadAcMed: 0.30 },
    { producto: '500mg', cantidadAcMed: 0.50 }
  ].forEach(item => {
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      10,
      'Producto',
      'AcMed',
      item.cantidadAcMed,
      'g',
      'Base dosificada desde AcMed.'
    ));
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      20,
      'Insumo',
      'Capsulas',
      1,
      'und',
      'Una capsula por unidad vendida.'
    ));
  });

  [
    'AcAlt',
    'AcMed',
    'AcSup',
    'ColaDP',
    'ColaDPPow',
    'Cordy',
    'CordyPow',
    'Gano',
    'GanoPow',
    'Lm',
    'LmPow',
    'Shii',
    'ShiiPow'
  ].forEach(producto => {
    rows.push(crearFilaComponentePlantillaQTAS_(
      producto,
      'g',
      10,
      'Producto',
      producto,
      1,
      'g',
      'Producto base por gramo.'
    ));
  });

  rows.push(crearFilaComponentePlantillaQTAS_(
    'Choco',
    'und',
    10,
    'Insumo',
    'Chocolate',
    7,
    'g',
    'Base de chocolate por unidad.'
  ));
  rows.push(crearFilaComponentePlantillaQTAS_(
    'Choco',
    'und',
    20,
    'Producto',
    'AcSup',
    0.7,
    'g',
    'Aporte de AcSup por unidad.'
  ));

  rows.push(crearFilaComponentePlantillaQTAS_(
    'Chocordy',
    'und',
    10,
    'Insumo',
    'Chocolate',
    85,
    'g',
    'Base de chocolate por unidad.'
  ));
  rows.push(crearFilaComponentePlantillaQTAS_(
    'Chocordy',
    'und',
    20,
    'Producto',
    'Cordy',
    7,
    'g',
    'Aporte de Cordy por unidad.'
  ));

  [
    { producto: 'ColaDPExt', base: 'ColaDP' },
    { producto: 'CordyExt', base: 'Cordy' },
    { producto: 'GanoExt', base: 'Gano' },
    { producto: 'LmExt', base: 'Lm' },
    { producto: 'ShiiExt', base: 'Shii' }
  ].forEach(item => {
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      10,
      'Producto',
      item.base,
      7.5,
      'g',
      'Base del hongo por gotero.'
    ));
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      20,
      'Insumo',
      'Goteros',
      1,
      'und',
      'Un gotero por unidad.'
    ));
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      30,
      'Insumo',
      'Alcohol',
      1,
      'g',
      'Placeholder en g para que ajustes la formula real despues.'
    ));
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      40,
      'Gasto',
      'Mano de obra',
      1,
      'und',
      'Unidad base de mano de obra.'
    ));
    rows.push(crearFilaComponentePlantillaQTAS_(
      item.producto,
      'und',
      50,
      'Insumo',
      'Calcas',
      1,
      'und',
      'Una calca por gotero.'
    ));
  });

  rows.push(crearFilaComponentePlantillaQTAS_(
    'Tin',
    'und',
    10,
    'Insumo',
    'Vasos',
    1,
    'und',
    'Plantilla simple para Tin.'
  ));
  rows.push(crearFilaComponentePlantillaQTAS_(
    'Tin',
    'und',
    20,
    'Insumo',
    'Cafe',
    1,
    'und',
    'Plantilla simple para Tin.'
  ));
  rows.push(crearFilaComponentePlantillaQTAS_(
    'Tin',
    'und',
    30,
    'Insumo',
    'Agua',
    1,
    'und',
    'Plantilla simple para Tin.'
  ));

  return rows;
}

function construirReglasCostoSUPERQTAS_() {
  const rows = [];

  [
    '50mg',
    '100mg',
    '150mg',
    '200mg',
    '300mg',
    '500mg'
  ].forEach(producto => {
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'und',
      1,
      24,
      10,
      'Insumo',
      'Bolsa_Barata',
      2,
      'und',
      'PorLinea',
      'Menos de 25 unidades: 2 bolsas baratas.'
    ));
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'und',
      1,
      24,
      20,
      'Insumo',
      'Calcas',
      1,
      'und',
      'PorLinea',
      'Menos de 25 unidades: 1 calca.'
    ));
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'und',
      25,
      '',
      10,
      'Insumo',
      'Bolsa_Barata',
      1,
      'und',
      'PorLinea',
      '25 unidades o mas: 1 bolsa barata.'
    ));
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'und',
      25,
      '',
      20,
      'Insumo',
      'Calcas',
      2,
      'und',
      'PorLinea',
      '25 unidades o mas: 2 calcas.'
    ));
  });

  [
    'AcAlt',
    'AcMed',
    'AcSup'
  ].forEach(producto => {
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'g',
      '',
      '',
      10,
      'Insumo',
      'Bolsa_Barata',
      2,
      'und',
      'PorLinea',
      'Dos bolsas baratas por pedido.'
    ));
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'g',
      '',
      '',
      20,
      'Insumo',
      'Calcas',
      1,
      'und',
      'PorLinea',
      'Una calca por pedido.'
    ));
  });

  [
    'ColaDP',
    'ColaDPPow',
    'Cordy',
    'CordyPow',
    'Gano',
    'GanoPow',
    'Lm',
    'LmPow',
    'Shii',
    'ShiiPow'
  ].forEach(producto => {
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'g',
      '',
      '',
      10,
      'Insumo',
      'Bolsa_Cara',
      1,
      'und',
      'PorLinea',
      'Una bolsa cara por pedido.'
    ));
    rows.push(crearFilaReglaCostoPlantillaQTAS_(
      producto,
      'g',
      '',
      '',
      20,
      'Insumo',
      'Calcas',
      1,
      'und',
      'PorLinea',
      'Una calca por pedido.'
    ));
  });

  rows.push(crearFilaReglaCostoPlantillaQTAS_(
    'Choco',
    'und',
    '',
    '',
    10,
    'Insumo',
    'Bolsa_Barata',
    1,
    'und',
    'PorUnidad',
    'Una bolsa barata por unidad.'
  ));
  rows.push(crearFilaReglaCostoPlantillaQTAS_(
    'Choco',
    'und',
    '',
    '',
    20,
    'Insumo',
    'Bolsa_Media',
    1,
    'und',
    'PorUnidad',
    'Una bolsa media por unidad.'
  ));
  rows.push(crearFilaReglaCostoPlantillaQTAS_(
    'Choco',
    'und',
    '',
    '',
    30,
    'Insumo',
    'Calcas',
    1,
    'und',
    'PorUnidad',
    'Una calca por unidad.'
  ));

  rows.push(crearFilaReglaCostoPlantillaQTAS_(
    'Chocordy',
    'und',
    '',
    '',
    10,
    'Insumo',
    'Bolsa_Cara',
    1,
    'und',
    'PorUnidad',
    'Una bolsa cara por unidad.'
  ));
  rows.push(crearFilaReglaCostoPlantillaQTAS_(
    'Chocordy',
    'und',
    '',
    '',
    20,
    'Insumo',
    'Calcas',
    1,
    'und',
    'PorUnidad',
    'Una calca por unidad.'
  ));

  return rows;
}

function crearFilaComponentePlantillaQTAS_(
  producto,
  unidadVenta,
  orden,
  tipoComponente,
  itemComponente,
  cantidadComponente,
  unidadComponente,
  nota
) {
  const medida = normalizarCantidadUnidadQTAS_(cantidadComponente, unidadComponente);
  return {
    Producto_Estandar: texto_(producto),
    Unidad_Venta: normalizarUnidadCanonicaQTAS_(unidadVenta),
    Orden: Math.max(1, Math.floor(numero_(orden) || 1)),
    Tipo_Componente: normalizarTipoCompraItemQTAS_(tipoComponente),
    Item_Componente: texto_(itemComponente),
    Cantidad_Componente: medida.cantidad,
    Unidad_Componente: medida.unidad,
    Merma_Pct: 0,
    Activo: true,
    Nota: texto_(nota)
  };
}

function crearFilaReglaCostoPlantillaQTAS_(
  producto,
  unidadVenta,
  cantidadMin,
  cantidadMax,
  orden,
  tipoComponente,
  itemComponente,
  cantidadComponente,
  unidadComponente,
  aplicacion,
  nota
) {
  const medida = normalizarCantidadUnidadQTAS_(cantidadComponente, unidadComponente);
  return {
    Producto_Estandar: texto_(producto),
    Unidad_Venta: normalizarUnidadCanonicaQTAS_(unidadVenta),
    Fecha_Desde: '',
    Fecha_Hasta: '',
    Cantidad_Min: cantidadMin === '' ? '' : redondear_(Math.max(0, numero_(cantidadMin))),
    Cantidad_Max: cantidadMax === '' ? '' : redondear_(Math.max(0, numero_(cantidadMax))),
    Orden: Math.max(1, Math.floor(numero_(orden) || 1)),
    Tipo_Componente: normalizarTipoCompraItemQTAS_(tipoComponente),
    Item_Componente: texto_(itemComponente),
    Cantidad_Componente: medida.cantidad,
    Unidad_Componente: medida.unidad,
    Aplicacion: normalizarAplicacionReglaCostoQTAS_(aplicacion),
    Merma_Pct: 0,
    Activo: true,
    Nota: texto_(nota)
  };
}

function upsertPlantillaComponentesQTAS_(sheet, headers, templateRows) {
  const existentes = leerObjetosConMeta_(sheet);
  const porClave = {};
  existentes.forEach(row => {
    const clave = claveComponentePlantillaQTAS_(row);
    if (!clave || porClave[clave]) return;
    porClave[clave] = row;
  });

  let nextId = siguienteIdConPrefijo_(sheet, 'Componente_ID', 'RCP-', 4);
  let inserted = 0;
  let updated = 0;
  const nuevos = [];

  (templateRows || []).forEach(row => {
    const clave = claveComponentePlantillaQTAS_(row);
    if (!clave) return;

    const existente = porClave[clave];
    const normalizado = Object.assign({}, row, {
      Componente_ID: existente ? texto_(existente.Componente_ID) : nextId
    });

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, normalizado));
      updated++;
    } else {
      nuevos.push(filaDesdeHeaders_(headers, normalizado));
      inserted++;
      nextId = siguienteIdConPrefijoDesdeValorQTAS_(nextId, 'RCP-', 4);
    }
  });

  if (nuevos.length) {
    escribirFilas_(sheet, nuevos);
  }

  return {
    inserted: inserted,
    updated: updated,
    templateRows: (templateRows || []).length
  };
}

function upsertPlantillaReglasCostoQTAS_(sheet, headers, templateRows) {
  const existentes = leerObjetosConMeta_(sheet);
  const porClave = {};
  existentes.forEach(row => {
    const clave = claveReglaCostoPlantillaQTAS_(row);
    if (!clave || porClave[clave]) return;
    porClave[clave] = row;
  });

  let nextId = siguienteIdConPrefijo_(sheet, 'Regla_Costo_ID', 'RCR-', 4);
  let inserted = 0;
  let updated = 0;
  const nuevos = [];

  (templateRows || []).forEach(row => {
    const clave = claveReglaCostoPlantillaQTAS_(row);
    if (!clave) return;

    const existente = porClave[clave];
    const normalizado = Object.assign({}, row, {
      Regla_Costo_ID: existente ? texto_(existente.Regla_Costo_ID) : nextId
    });

    if (existente) {
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, Object.assign({}, existente, normalizado));
      updated++;
    } else {
      nuevos.push(filaDesdeHeaders_(headers, normalizado));
      inserted++;
      nextId = siguienteIdConPrefijoDesdeValorQTAS_(nextId, 'RCR-', 4);
    }
  });

  if (nuevos.length) {
    escribirFilas_(sheet, nuevos);
  }

  return {
    inserted: inserted,
    updated: updated,
    templateRows: (templateRows || []).length
  };
}

function claveComponentePlantillaQTAS_(row) {
  return [
    normalizarClaveTexto_(row.Producto_Estandar || row.producto),
    normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad_Venta || row.unidadVenta)),
    String(Math.max(1, Math.floor(numero_(row.Orden || row.orden) || 1))),
    normalizarClaveTexto_(row.Tipo_Componente || row.tipoComponente),
    normalizarClaveTexto_(row.Item_Componente || row.itemComponente)
  ].join('|');
}

function claveReglaCostoPlantillaQTAS_(row) {
  return [
    normalizarClaveTexto_(row.Producto_Estandar || row.producto),
    normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad_Venta || row.unidadVenta)),
    fechaTextoPlanoQTAS_(row.Fecha_Desde || row.fechaDesde),
    fechaTextoPlanoQTAS_(row.Fecha_Hasta || row.fechaHasta),
    texto_(row.Cantidad_Min || row.cantidadMin),
    texto_(row.Cantidad_Max || row.cantidadMax),
    String(Math.max(1, Math.floor(numero_(row.Orden || row.orden) || 1))),
    normalizarClaveTexto_(row.Tipo_Componente || row.tipoComponente),
    normalizarClaveTexto_(row.Item_Componente || row.itemComponente),
    normalizarClaveTexto_(row.Aplicacion || row.aplicacion)
  ].join('|');
}

function siguienteIdConPrefijoDesdeValorQTAS_(currentId, prefix, padLength) {
  const match = texto_(currentId).match(/(\d+)$/);
  const next = match ? Number(match[1]) + 1 : 1;
  return `${texto_(prefix)}${String(next).padStart(padLength || 4, '0')}`;
}

function resumenCoberturaPlantillaCostoQTAS_(componentes, reglas) {
  const fechaBase = new Date();
  const costosCache = cargarCostosEnMemoria_();
  const items = {};

  (componentes || []).forEach(row => {
    if (!texto_(row.Item_Componente) || !normalizarUnidadCanonicaQTAS_(row.Unidad_Componente)) return;
    const key = [
      texto_(row.Tipo_Componente),
      texto_(row.Item_Componente),
      normalizarUnidadCanonicaQTAS_(row.Unidad_Componente)
    ].join('|');
    items[key] = {
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Componente),
      item: texto_(row.Item_Componente),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad_Componente)
    };
  });

  (reglas || []).forEach(row => {
    if (!texto_(row.Item_Componente) || !normalizarUnidadCanonicaQTAS_(row.Unidad_Componente)) return;
    const key = [
      texto_(row.Tipo_Componente),
      texto_(row.Item_Componente),
      normalizarUnidadCanonicaQTAS_(row.Unidad_Componente)
    ].join('|');
    items[key] = {
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Componente),
      item: texto_(row.Item_Componente),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad_Componente)
    };
  });

  const componentesFaltantes = Object.keys(items)
    .map(key => items[key])
    .filter(item => {
      if (item.tipoItem === 'Producto') {
        const costo = calcularCostoProductoEnFechaQTAS_(item.item, item.unidad, fechaBase, {
          cantidadVenta: 1,
          incluirReglasPorLinea: false,
          incluirReglasPorUnidad: true,
          costosCache: costosCache,
          componentes: leerComponentesProductoActivosQTAS_(),
          reglas: leerReglasCostoProductoActivasQTAS_()
        });
        return numero_(costo.costoUnitarioTotal) <= 0;
      }

      return numero_(obtenerCostoVigenteDesdeCache_(
        costosCache,
        item.item,
        item.unidad,
        fechaBase,
        item.tipoItem
      )) <= 0;
    })
    .map(item => `${item.item} (${item.unidad}, ${item.tipoItem})`)
    .sort((a, b) => a.localeCompare(b));

  return {
    componentesTemplate: (componentes || []).length,
    reglasTemplate: (reglas || []).length,
    itemsSinCostoHoy: componentesFaltantes
  };
}
