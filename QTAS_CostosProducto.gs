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
    const costoProducto = sincronizarCostoProductoDesdeProductosQTAS_([producto], {
      ss: ss,
      ahora: new Date()
    });

    return {
      ok: true,
      componenteId: normalizado.Componente_ID,
      componentes: listarComponentesProductoQTAS_(),
      costoProducto: costoProducto
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
    const costoProducto = sincronizarCostoProductoDesdeProductosQTAS_([producto], {
      ss: ss,
      ahora: new Date()
    });

    return {
      ok: true,
      reglaId: normalizado.Regla_Costo_ID,
      reglas: listarReglasCostoProductoQTAS_(),
      costoProducto: costoProducto
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

function sincronizarCostoProductoDesdeProductosQTAS_(productos, payload) {
  const productosObjetivo = unirProductosUnicosQTAS_(productos);
  if (!productosObjetivo.length) {
    return {
      ok: true,
      skipped: false,
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0
    };
  }

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
    .filter(row => productoPerteneceAListaQTAS_(row.Producto_Estandar, productosObjetivo))
    .map(row => construirFilaCostoProductoCalculadoQTAS_(row, contexto));
  const upsert = upsertObjetosLotePorIdQTAS_(
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

function unirProductosUnicosQTAS_(productos) {
  const mapa = {};
  (productos || []).forEach(producto => {
    const key = normalizarClaveTexto_(producto);
    if (key) {
      mapa[key] = true;
    }
  });
  return Object.keys(mapa);
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

function alinearRecetasExtractosBaseQTAS() {
  return withScriptLock_('alinear recetas extractos base', () => {
    asegurarModeloOperativoQTAS_();

    const extractos = [
      { producto: 'ColaDPExt', base: 'ColaDP', calcaExt: 'Calca_ColaDP_Ext' },
      { producto: 'CordyExt', base: 'Cordy', calcaExt: 'Calca_Cordy_Ext' },
      { producto: 'GanoExt', base: 'Gano', calcaExt: 'Calca_Gano_Ext' },
      { producto: 'LmExt', base: 'Lm', calcaExt: 'Calca_Lm_Ext' },
      { producto: 'ShiiExt', base: 'Shii', calcaExt: 'Calca_Shii_Ext' }
    ];
    const componentesComunes = [
      {
        orden: 20,
        tipoComponente: 'Insumo',
        itemComponente: 'Goteros',
        cantidadComponente: 1,
        unidadComponente: 'und',
        nota: 'Un gotero por unidad.'
      },
      {
        orden: 30,
        tipoComponente: 'Insumo',
        itemComponente: 'Alcohol',
        cantidadComponente: 20,
        unidadComponente: 'g',
        nota: 'Supuesto operativo: 400 ml de alcohol por lote de 20 extracciones.'
      },
      {
        orden: 35,
        tipoComponente: 'Insumo',
        itemComponente: 'Agua',
        cantidadComponente: 60,
        unidadComponente: 'g',
        nota: 'Supuesto operativo: 400 ml iniciales + 800 ml adicionales por lote de 20 extracciones.'
      },
      {
        orden: 40,
        tipoComponente: 'Gasto',
        itemComponente: 'Mano de obra',
        cantidadComponente: 1,
        unidadComponente: 'und',
        nota: 'Unidad base de mano de obra.'
      },
      {
        orden: 50,
        tipoComponente: 'Insumo',
        itemComponente: 'Bolsa_Papel_1lb',
        cantidadComponente: 1,
        unidadComponente: 'und',
        nota: 'Una bolsa de papel de 1 lb por unidad.'
      }
    ];
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.productoComponentes);
    const headers = getHeaders_(sheet);
    const rows = leerObjetos_(sheet);
    const productosObjetivo = {};
    const existentesPorClave = {};
    let nextId = siguienteIdConPrefijo_(sheet, 'Componente_ID', 'RCP-', 4);

    extractos.forEach(item => {
      productosObjetivo[normalizarClaveTexto_(item.producto)] = true;
      validarProductoCanonicoExistenteQTAS_(ss, item.producto);
    });

    rows.forEach(row => {
      const key = claveRecetaComponenteQTAS_(row);
      if (!key || existentesPorClave[key]) return;
      existentesPorClave[key] = row;
    });

    const deseados = [];
    extractos.forEach(item => {
      deseados.push(
        crearComponenteAlineadoExtractoQTAS_(item.producto, 'und', 10, 'Producto', item.base, 7.5, 'g',
          'Supuesto operativo: 150 g de hongo seco por lote de 20 extracciones.')
      );

      componentesComunes.forEach(componente => {
        deseados.push(
          crearComponenteAlineadoExtractoQTAS_(
            item.producto,
            'und',
            componente.orden,
            componente.tipoComponente,
            componente.itemComponente,
            componente.cantidadComponente,
            componente.unidadComponente,
            componente.nota
          )
        );
      });

      deseados.push(
        crearComponenteAlineadoExtractoQTAS_(
          item.producto,
          'und',
          60,
          'Insumo',
          item.calcaExt,
          1,
          'und',
          'Una calca especifica de extracto por unidad.'
        )
      );
    });

    const deseadosConId = deseados.map(row => {
      const key = claveRecetaComponenteQTAS_(row);
      const existente = key ? existentesPorClave[key] : null;
      const withId = Object.assign({}, row, {
        Componente_ID: existente
          ? texto_(existente.Componente_ID)
          : nextId
      });
      if (!existente) {
        nextId = siguienteIdConPrefijoDesdeValorQTAS_(nextId, 'RCP-', 4);
      }
      return withId;
    });

    const conservados = rows
      .filter(row => !productosObjetivo[normalizarClaveTexto_(row.Producto_Estandar)])
      .map(row => ({
        Componente_ID: texto_(row.Componente_ID),
        Producto_Estandar: texto_(row.Producto_Estandar),
        Unidad_Venta: normalizarUnidadCanonicaQTAS_(row.Unidad_Venta),
        Orden: Math.max(1, Math.floor(numero_(row.Orden) || 1)),
        Tipo_Componente: normalizarTipoCompraItemQTAS_(row.Tipo_Componente),
        Item_Componente: texto_(row.Item_Componente),
        Cantidad_Componente: redondear_(numero_(row.Cantidad_Componente)),
        Unidad_Componente: normalizarUnidadCanonicaQTAS_(row.Unidad_Componente),
        Merma_Pct: redondear_(numero_(row.Merma_Pct)),
        Activo: estaActivo_(row.Activo),
        Nota: texto_(row.Nota)
      }));

    const finalRows = conservados
      .concat(deseadosConId)
      .sort((a, b) => {
        if (a.Producto_Estandar !== b.Producto_Estandar) return a.Producto_Estandar.localeCompare(b.Producto_Estandar);
        if (a.Unidad_Venta !== b.Unidad_Venta) return a.Unidad_Venta.localeCompare(b.Unidad_Venta);
        if (numero_(a.Orden) !== numero_(b.Orden)) return numero_(a.Orden) - numero_(b.Orden);
        return texto_(a.Item_Componente).localeCompare(texto_(b.Item_Componente));
      });

    sobrescribirObjetosHojaQTAS_(sheet, headers, finalRows);

    const costoProducto = sincronizarCostoProductoDesdeProductosQTAS_(
      extractos.map(item => item.producto),
      {
        ss: ss,
        ahora: new Date()
      }
    );

    return {
      ok: true,
      recetasAlineadas: extractos.map(item => item.producto),
      componentesObjetivo: deseadosConId.length,
      productosActualizados: extractos.length,
      costoProducto: costoProducto,
      assumptions: [
        'Se asume un lote de 20 extractos terminados de 50 ml.',
        'Se usa 150 g de hongo seco por lote, equivalente a 7.5 g por unidad.',
        'Se usa 400 ml de alcohol por lote, equivalente a 20 g por unidad.',
        'Se usa agua total aproximada de 1200 ml por lote, equivalente a 60 g por unidad.',
        'Cada extracto usa bolsa de papel de 1 lb y una calca especifica por producto.'
      ]
    };
  });
}

function crearComponenteAlineadoExtractoQTAS_(
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

function claveRecetaComponenteQTAS_(row) {
  return [
    normalizarClaveTexto_(row.Producto_Estandar),
    normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad_Venta)),
    String(Math.max(1, Math.floor(numero_(row.Orden) || 1))),
    normalizarClaveTexto_(row.Tipo_Componente),
    normalizarClaveTexto_(row.Item_Componente)
  ].join('|');
}

function ajustarPackagingPsyloScibioQTAS() {
  asegurarModeloOperativoQTAS_({ aplicarFormatos: true });

  const recetas = alinearRecetasExtractosBaseQTAS();
  const costos = sembrarCostosPackagingPsyloScibioQTAS();
  const reglas = alinearReglasPackagingPsyloScibioQTAS();
  const productosCosto = Array.from(new Set(
    []
      .concat(recetas && recetas.recetasAlineadas || [])
      .concat(reglas && reglas.productosAfectados || [])
  ));
  const costoProducto = productosCosto.length
    ? sincronizarCostoProductoDesdeProductosQTAS_(productosCosto, {
      ss: SpreadsheetApp.getActive(),
      ahora: new Date()
    })
    : {
      ok: true,
      rows: 0,
      inserted: 0,
      updated: 0,
      stale: 0
    };

  limpiarCachesEjecucionQTAS_();

  return {
    ok: true,
    recetas: recetas,
    costos: costos,
    reglas: reglas,
    costoProducto: costoProducto,
    assumptions: unirUnicos_(
      []
        .concat(recetas && recetas.assumptions || [])
        .concat(costos && costos.assumptions || [])
        .concat(reglas && reglas.assumptions || [])
        .concat([
          'AcAlt, AcMed, AcSup y Choco no se tocan en este ajuste porque falta confirmar la separacion final de empaque por tamano.',
          'El borrado de hojas legacy se deja en un script aparte con dryRun por seguridad.'
        ])
    )
  };
}

function ajustarPackagingPsyloScibioQTAS_Log() {
  const result = ajustarPackagingPsyloScibioQTAS();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function sembrarCostosPackagingPsyloScibioQTAS() {
  return withScriptLock_('sembrar costos packaging psylo scibio', () => {
    asegurarModeloOperativoQTAS_({ aplicarFormatos: true });

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
    const headers = getHeaders_(sheet);
    const rows = leerObjetosConMeta_(sheet);
    const fechaDesde = resolverFechaOperacion_(new Date(), new Date());
    const proveedor = 'Placeholder packaging Psylo Scibio';
    const costosCache = cargarCostosEnMemoria_();
    const seeds = crearSeedsCostosPackagingPsyloScibioQTAS_(costosCache, fechaDesde);
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

    limpiarCachesEjecucionQTAS_();

    return {
      ok: true,
      fechaDesde: fechaInput_(fechaDesde),
      proveedor: proveedor,
      actualizados: actualizados,
      costosSembrados: seeds.map(seed => ({
        item: seed.item,
        tipoItem: seed.tipoItem,
        unidad: seed.unidad,
        costoUnitario: seed.costoUnitario,
        referencia: seed.referencia
      })),
      assumptions: [
        'Bolsa_Papel_0_5lb y Bolsa_Papel_1lb se siembran como placeholder usando la referencia vigente de Bolsa_Barata.',
        'Bolsa_Kraft_Zip_Mediana, Bolsa_Zip_Negra y Bolsa_Zip_Plateada se siembran usando la referencia vigente de Bolsa_Media.',
        'Bolsa_Kraft_Zip_Grande y Frasco_Capsulas se siembran usando la referencia vigente de Bolsa_Cara.',
        'Las calcas especificas se siembran con un costo unitario estimado a partir de Calcas genericas dividido entre el stock inicial de calcas contabilizadas.'
      ]
    };
  });
}

function alinearReglasPackagingPsyloScibioQTAS() {
  return withScriptLock_('alinear reglas packaging psylo scibio', () => {
    asegurarModeloOperativoQTAS_();

    const plan = construirReglasPackagingPsyloScibioQTAS_();
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.productoReglasCosto);
    const headers = getHeaders_(sheet);
    const rows = leerObjetos_(sheet);
    const productosObjetivo = {};
    const existentesPorClave = {};
    let nextId = siguienteIdConPrefijo_(sheet, 'Regla_Costo_ID', 'RCR-', 4);

    plan.productosObjetivo.forEach(producto => {
      productosObjetivo[normalizarClaveTexto_(producto)] = true;
      validarProductoCanonicoExistenteQTAS_(ss, producto);
    });

    rows.forEach(row => {
      const key = claveReglaPackagingPsyloScibioQTAS_(row);
      if (!key || existentesPorClave[key]) return;
      existentesPorClave[key] = row;
    });

    const deseadasConId = plan.rows.map(row => {
      const key = claveReglaPackagingPsyloScibioQTAS_(row);
      const existente = key ? existentesPorClave[key] : null;
      const withId = Object.assign({}, row, {
        Regla_Costo_ID: existente
          ? texto_(existente.Regla_Costo_ID)
          : nextId
      });
      if (!existente) {
        nextId = siguienteIdConPrefijoDesdeValorQTAS_(nextId, 'RCR-', 4);
      }
      return withId;
    });

    const conservadas = rows
      .filter(row => !productosObjetivo[normalizarClaveTexto_(row.Producto_Estandar)])
      .map(row => ({
        Regla_Costo_ID: texto_(row.Regla_Costo_ID),
        Producto_Estandar: texto_(row.Producto_Estandar),
        Unidad_Venta: normalizarUnidadCanonicaQTAS_(row.Unidad_Venta),
        Fecha_Desde: row.Fecha_Desde ? resolverFechaOperacion_(row.Fecha_Desde, new Date()) : '',
        Fecha_Hasta: row.Fecha_Hasta ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date()) : '',
        Cantidad_Min: texto_(row.Cantidad_Min) === '' ? '' : redondear_(numero_(row.Cantidad_Min)),
        Cantidad_Max: texto_(row.Cantidad_Max) === '' ? '' : redondear_(numero_(row.Cantidad_Max)),
        Orden: Math.max(1, Math.floor(numero_(row.Orden) || 1)),
        Tipo_Componente: normalizarTipoCompraItemQTAS_(row.Tipo_Componente),
        Item_Componente: texto_(row.Item_Componente),
        Cantidad_Componente: redondear_(numero_(row.Cantidad_Componente)),
        Unidad_Componente: normalizarUnidadCanonicaQTAS_(row.Unidad_Componente),
        Aplicacion: normalizarAplicacionReglaCostoQTAS_(row.Aplicacion),
        Merma_Pct: redondear_(numero_(row.Merma_Pct)),
        Activo: estaActivo_(row.Activo),
        Nota: texto_(row.Nota)
      }));

    const finalRows = conservadas
      .concat(deseadasConId)
      .sort((a, b) => {
        if (a.Producto_Estandar !== b.Producto_Estandar) return a.Producto_Estandar.localeCompare(b.Producto_Estandar);
        if (a.Unidad_Venta !== b.Unidad_Venta) return a.Unidad_Venta.localeCompare(b.Unidad_Venta);
        if (numero_(a.Cantidad_Min) !== numero_(b.Cantidad_Min)) return numero_(a.Cantidad_Min) - numero_(b.Cantidad_Min);
        if (texto_(a.Cantidad_Max) !== texto_(b.Cantidad_Max)) return texto_(a.Cantidad_Max).localeCompare(texto_(b.Cantidad_Max));
        if (numero_(a.Orden) !== numero_(b.Orden)) return numero_(a.Orden) - numero_(b.Orden);
        return texto_(a.Item_Componente).localeCompare(texto_(b.Item_Componente));
      });

    sobrescribirObjetosHojaQTAS_(sheet, headers, finalRows);

    return {
      ok: true,
      reglasAlineadas: deseadasConId.length,
      productosAfectados: plan.productosObjetivo.slice(),
      assumptions: plan.assumptions.slice()
    };
  });
}

function construirReglasPackagingPsyloScibioQTAS_() {
  const rows = [];
  const productosObjetivo = [];
  const assumptions = [
    'Micros de 1 a 24 unidades consumen zip negra + bolsa papel 0.5 lb + calca logo.',
    'Micros de 25 o mas siempre consumen bolsa papel 0.5 lb + calca logo; 100mg, 200mg y 300mg ademas consumen frasco y calca de instrucciones.',
    'Polvos y hongos en gramos consumen bolsa kraft con zip mediana hasta 50 g y grande por encima de 50 g, mas su calca especifica de bolsa.'
  ];
  const micros = ['50mg', '100mg', '150mg', '200mg', '300mg', '500mg'];
  const microsConFrasco = {
    '100mg': true,
    '200mg': true,
    '300mg': true
  };
  const polvos = [
    { producto: 'ColaDP', calca: 'Calca_ColaDP_Bolsa' },
    { producto: 'ColaDPPow', calca: 'Calca_ColaDP_Bolsa' },
    { producto: 'Cordy', calca: 'Calca_Cordy_Bolsa' },
    { producto: 'CordyPow', calca: 'Calca_Cordy_Bolsa' },
    { producto: 'Gano', calca: 'Calca_Gano_Bolsa' },
    { producto: 'GanoPow', calca: 'Calca_Gano_Bolsa' },
    { producto: 'Lm', calca: 'Calca_Lm_Bolsa' },
    { producto: 'LmPow', calca: 'Calca_Lm_Bolsa' },
    { producto: 'Shii', calca: 'Calca_Shii_Bolsa' },
    { producto: 'ShiiPow', calca: 'Calca_Shii_Bolsa' }
  ];

  micros.forEach(producto => {
    productosObjetivo.push(producto);
    rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 1, 24, 10, 'Insumo', 'Bolsa_Zip_Negra', 1, 'und', 'PorLinea', 'Micros <25: una zip negra.'));
    rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 1, 24, 20, 'Insumo', 'Bolsa_Papel_0_5lb', 1, 'und', 'PorLinea', 'Micros <25: una bolsa papel 0.5 lb.'));
    rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 1, 24, 30, 'Insumo', 'Calca_Micros_Logo', 1, 'und', 'PorLinea', 'Micros <25: una calca logo.'));

    rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 25, '', 20, 'Insumo', 'Bolsa_Papel_0_5lb', 1, 'und', 'PorLinea', 'Micros >=25: una bolsa papel 0.5 lb.'));
    rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 25, '', 30, 'Insumo', 'Calca_Micros_Logo', 1, 'und', 'PorLinea', 'Micros >=25: una calca logo.'));

    if (microsConFrasco[producto]) {
      rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 25, '', 10, 'Insumo', 'Frasco_Capsulas', 1, 'und', 'PorLinea', 'Micros >=25: un frasco por pedido.'));
      rows.push(crearReglaPackagingPsyloScibioQTAS_(producto, 'und', 25, '', 40, 'Insumo', 'Calca_Micros_Instrucciones', 1, 'und', 'PorLinea', 'Micros >=25: una calca de instrucciones.'));
    }
  });

  polvos.forEach(item => {
    productosObjetivo.push(item.producto);
    rows.push(crearReglaPackagingPsyloScibioQTAS_(item.producto, 'g', '', 50, 10, 'Insumo', 'Bolsa_Kraft_Zip_Mediana', 1, 'und', 'PorLinea', 'Hasta 50 g: bolsa kraft zip mediana.'));
    rows.push(crearReglaPackagingPsyloScibioQTAS_(item.producto, 'g', '', 50, 20, 'Insumo', item.calca, 1, 'und', 'PorLinea', 'Hasta 50 g: una calca especifica de bolsa.'));
    rows.push(crearReglaPackagingPsyloScibioQTAS_(item.producto, 'g', 50.0001, '', 10, 'Insumo', 'Bolsa_Kraft_Zip_Grande', 1, 'und', 'PorLinea', 'Mas de 50 g: bolsa kraft zip grande.'));
    rows.push(crearReglaPackagingPsyloScibioQTAS_(item.producto, 'g', 50.0001, '', 20, 'Insumo', item.calca, 1, 'und', 'PorLinea', 'Mas de 50 g: una calca especifica de bolsa.'));
  });

  return {
    rows: rows,
    productosObjetivo: productosObjetivo.filter((value, index, array) => array.indexOf(value) === index),
    assumptions: assumptions
  };
}

function crearReglaPackagingPsyloScibioQTAS_(
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
    Cantidad_Min: cantidadMin === '' ? '' : redondear_(numero_(cantidadMin)),
    Cantidad_Max: cantidadMax === '' ? '' : redondear_(numero_(cantidadMax)),
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

function claveReglaPackagingPsyloScibioQTAS_(row) {
  return [
    normalizarClaveTexto_(row.Producto_Estandar),
    normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad_Venta)),
    fechaTextoPlanoQTAS_(row.Fecha_Desde),
    fechaTextoPlanoQTAS_(row.Fecha_Hasta),
    texto_(row.Cantidad_Min),
    texto_(row.Cantidad_Max),
    String(Math.max(1, Math.floor(numero_(row.Orden) || 1))),
    normalizarClaveTexto_(row.Tipo_Componente),
    normalizarClaveTexto_(row.Item_Componente),
    normalizarClaveTexto_(row.Aplicacion)
  ].join('|');
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

function crearSeedsCostosPackagingPsyloScibioQTAS_(costosCache, fechaBase) {
  const costoCalca = estimarCostoUnitarioCalcaPsyloScibioQTAS_(costosCache);
  const costoBolsaBarata = resolverCostoPlaceholderPackagingPsyloScibioQTAS_(costosCache, 'Bolsa_Barata', 'und', 'Insumo', 300);
  const costoBolsaMedia = resolverCostoPlaceholderPackagingPsyloScibioQTAS_(costosCache, 'Bolsa_Media', 'und', 'Insumo', 800);
  const costoBolsaCara = resolverCostoPlaceholderPackagingPsyloScibioQTAS_(costosCache, 'Bolsa_Cara', 'und', 'Insumo', 1100);
  const seeds = [
    crearCostoDirectoBaseQTAS_('Bolsa_Papel_0_5lb', 'Insumo', 'und', costoBolsaBarata, 'Placeholder derivado de Bolsa_Barata.', `Placeholder ${costoBolsaBarata}/und desde Bolsa_Barata`),
    crearCostoDirectoBaseQTAS_('Bolsa_Papel_1lb', 'Insumo', 'und', costoBolsaBarata, 'Placeholder derivado de Bolsa_Barata.', `Placeholder ${costoBolsaBarata}/und desde Bolsa_Barata`),
    crearCostoDirectoBaseQTAS_('Bolsa_Kraft_Zip_Mediana', 'Insumo', 'und', costoBolsaMedia, 'Placeholder derivado de Bolsa_Media.', `Placeholder ${costoBolsaMedia}/und desde Bolsa_Media`),
    crearCostoDirectoBaseQTAS_('Bolsa_Zip_Negra', 'Insumo', 'und', costoBolsaMedia, 'Placeholder derivado de Bolsa_Media.', `Placeholder ${costoBolsaMedia}/und desde Bolsa_Media`),
    crearCostoDirectoBaseQTAS_('Bolsa_Zip_Plateada', 'Insumo', 'und', costoBolsaMedia, 'Placeholder derivado de Bolsa_Media.', `Placeholder ${costoBolsaMedia}/und desde Bolsa_Media`),
    crearCostoDirectoBaseQTAS_('Bolsa_Kraft_Zip_Grande', 'Insumo', 'und', costoBolsaCara, 'Placeholder derivado de Bolsa_Cara.', `Placeholder ${costoBolsaCara}/und desde Bolsa_Cara`),
    crearCostoDirectoBaseQTAS_('Frasco_Capsulas', 'Insumo', 'und', costoBolsaCara, 'Placeholder derivado de Bolsa_Cara.', `Placeholder ${costoBolsaCara}/und desde Bolsa_Cara`)
  ];

  [
    'Calca_Choco',
    'Calca_ColaDP_Bolsa',
    'Calca_ColaDP_Ext',
    'Calca_Cordy_Bolsa',
    'Calca_Cordy_Ext',
    'Calca_Gano_Bolsa',
    'Calca_Gano_Ext',
    'Calca_Lm_Bolsa',
    'Calca_Lm_Ext',
    'Calca_Micros_Logo',
    'Calca_Micros_Instrucciones',
    'Calca_Shii_Bolsa',
    'Calca_Shii_Ext'
  ].forEach(item => {
    seeds.push(crearCostoDirectoBaseQTAS_(
      item,
      'Insumo',
      'und',
      costoCalca,
      'Placeholder estimado a partir de Calcas genericas y el conteo inicial de calcas.',
      `Placeholder ${costoCalca}/und para calca especifica`
    ));
  });

  return seeds.filter(seed =>
    numero_(seed.costoUnitario) > 0 &&
    texto_(seed.item) &&
    texto_(seed.unidad) &&
    fechaBase
  );
}

function resolverCostoPlaceholderPackagingPsyloScibioQTAS_(costosCache, item, unidad, tipoItem, fallback) {
  const directo = redondear_(obtenerCostoVigenteDesdeCache_(
    costosCache || cargarCostosEnMemoria_(),
    item,
    unidad,
    new Date(),
    tipoItem
  ));
  return directo > 0 ? directo : redondear_(numero_(fallback));
}

function estimarCostoUnitarioCalcaPsyloScibioQTAS_(costosCache) {
  const costoGenerico = redondear_(obtenerCostoVigenteDesdeCache_(
    costosCache || cargarCostosEnMemoria_(),
    'Calcas',
    'und',
    new Date(),
    'Insumo'
  ));
  const totalCalcas = agruparCantidadStockInicialPsyloScibioQTAS_(/^Calca_/);
  if (costoGenerico > 0 && totalCalcas > 0) {
    return redondear_(costoGenerico / totalCalcas);
  }
  return 150;
}

function agruparCantidadStockInicialPsyloScibioQTAS_(pattern) {
  const matcher = pattern instanceof RegExp ? pattern : null;
  const snapshot = construirStockInicialPsyloScibioQTAS_() || {};
  const rows = Array.isArray(snapshot)
    ? snapshot
    : (Array.isArray(snapshot.stock) ? snapshot.stock : []);

  return redondear_(sumar_(
    rows
      .filter(row =>
        row &&
        row.tipoItem === 'Insumo' &&
        row.unidad === 'und' &&
        matcher &&
        matcher.test(texto_(row.item))
      )
      .map(row => numero_(row.cantidad))
  ));
}

