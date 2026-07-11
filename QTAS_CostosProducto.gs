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
  const resumen = [];

  extractos.forEach(item => {
    resumen.push(guardarComponenteProductoQTAS({
      producto: item.producto,
      unidadVenta: 'und',
      orden: 10,
      tipoComponente: 'Producto',
      itemComponente: item.base,
      cantidadComponente: 7.5,
      unidadComponente: 'g',
      mermaPct: 0,
      nota: 'Supuesto operativo: 150 g de hongo seco por lote de 20 extracciones.',
      activo: true
    }));

    componentesComunes.forEach(componente => {
      resumen.push(guardarComponenteProductoQTAS({
        producto: item.producto,
        unidadVenta: 'und',
        orden: componente.orden,
        tipoComponente: componente.tipoComponente,
        itemComponente: componente.itemComponente,
        cantidadComponente: componente.cantidadComponente,
        unidadComponente: componente.unidadComponente,
        mermaPct: 0,
        nota: componente.nota,
        activo: true
      }));
    });

    resumen.push(guardarComponenteProductoQTAS({
      producto: item.producto,
      unidadVenta: 'und',
      orden: 60,
      tipoComponente: 'Insumo',
      itemComponente: item.calcaExt,
      cantidadComponente: 1,
      unidadComponente: 'und',
      mermaPct: 0,
      nota: 'Una calca especifica de extracto por unidad.',
      activo: true
    }));
  });

  return {
    ok: true,
    recetasAlineadas: extractos.map(item => item.producto),
    operaciones: resumen.length,
    assumptions: [
      'Se asume un lote de 20 extractos terminados de 50 ml.',
      'Se usa 150 g de hongo seco por lote, equivalente a 7.5 g por unidad.',
      'Se usa 400 ml de alcohol por lote, equivalente a 20 g por unidad.',
      'Se usa agua total aproximada de 1200 ml por lote, equivalente a 60 g por unidad.',
      'Cada extracto usa bolsa de papel de 1 lb y una calca especifica por producto.'
    ]
  };
}

