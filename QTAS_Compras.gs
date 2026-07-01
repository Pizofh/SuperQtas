function getCatalogoComprasQTAS() {
  validarModeloSoloLecturaQTAS_();

  const ss = SpreadsheetApp.getActive();
  const productos = leerProductosActivosCompraQTAS_(ss);

  const mediosPago = leerMediosPagoConfiguradosQTAS_()
    .filter(row => row.activo)
    .map(row => row.medioPago);

  const costosVigentes = listarCostosVigentesQTAS_();
  const comprasRecientes = listarComprasRecientesQTAS_();
  const proveedores = leerObjetos_(ss.getSheetByName(QTAS.sheets.compras))
    .map(row => texto_(row.Proveedor))
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index)
    .sort((a, b) => a.localeCompare(b));

  return {
    hoy: fechaInput_(new Date()),
    mediosPago: mediosPago,
    productos: productos,
    itemsSugeridos: construirItemsSugeridosComprasQTAS_(productos, costosVigentes),
    costosVigentes: costosVigentes,
    comprasRecientes: comprasRecientes,
    proveedores: proveedores
  };
}

function registrarCompraQTAS(payload) {
  return withScriptLock_('registrar compra', () => {
    asegurarModeloOperativoQTAS_();
    validarCompraQTAS_(payload);

    const mediosActivos = leerMediosPagoConfiguradosQTAS_()
      .filter(row => row.activo)
      .map(row => normalizarClaveTexto_(row.medioPago));
    const medioPago = texto_(payload.medioPago);

    if (!mediosActivos.includes(normalizarClaveTexto_(medioPago))) {
      throw new Error('El medio de pago no esta disponible en la configuracion actual.');
    }

    const ss = SpreadsheetApp.getActive();
    const comprasSheet = ss.getSheetByName(QTAS.sheets.compras);
    const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
    const comprasHeaders = getHeaders_(comprasSheet);
    const detalleHeaders = getHeaders_(detalleSheet);

    const compraId = siguienteIdNumerico_(comprasSheet, 'Compra_ID');
    const ahora = new Date();
    const fechaCompraBase = resolverFechaOperacion_(payload.fechaCompra, ahora);
    const fechaCompra = combinarFechaYHora_(fechaCompraBase, ahora);
    const proveedor = texto_(payload.proveedor);
    const comentarioCompra = texto_(payload.comentarioCompra);
    const productosActivosIndex = construirIndiceProductosCompraQTAS_(leerProductosActivosCompraQTAS_(ss));

    const lineasPreparadas = (payload.lineas || []).map((linea, index) =>
      prepararLineaCompraQTAS_({
        compraId: compraId,
        fechaCompra: fechaCompra,
        proveedor: proveedor,
        linea: linea,
        index: index,
        productosActivosIndex: productosActivosIndex
      })
    );

    const totalCompra = redondear_(sumar_(lineasPreparadas.map(item => item.Costo_Total_Linea)));
    const itemsResumen = resumenItemsCompraQTAS_(lineasPreparadas);

    escribirFilas_(comprasSheet, [filaDesdeHeaders_(comprasHeaders, {
      Compra_ID: compraId,
      Fecha_Compra: fechaCompra,
      Proveedor: proveedor,
      Items_Resumen: itemsResumen,
      Total_Compra: totalCompra,
      Medio_Pago: medioPago,
      Comentario_Compra: comentarioCompra,
      Estado_Registro: QTAS.status.registro.activo
    })]);

    escribirFilas_(
      detalleSheet,
      lineasPreparadas.map(row => filaDesdeHeaders_(detalleHeaders, row))
    );

    const costosActualizados = actualizarCostosDesdeCompraQTAS_({
      compraId: compraId,
      proveedor: proveedor,
      fechaCompra: fechaCompraBase,
      lineas: lineasPreparadas
    });

    let costoProductoCalculado = null;
    try {
      costoProductoCalculado = reconstruirCostoProductoCalculadoInternoQTAS_({
        ss: ss,
        fechaBase: new Date(),
        ahora: ahora
      });
    } catch (error) {
      Logger.log(`No se pudo refrescar Costo_Producto_Calc tras Compra ${compraId}: ${error.message}`);
      costoProductoCalculado = {
        ok: false,
        skipped: true,
        reason: error.message,
        rows: 0,
        inserted: 0,
        updated: 0,
        stale: 0,
        fechaBase: fechaInput_(new Date())
      };
    }

    return {
      ok: true,
      compraId: compraId,
      totalCompra: totalCompra,
      lineas: lineasPreparadas.length,
      costosActualizados: costosActualizados,
      itemsResumen: itemsResumen,
      costoProductoCalculado: costoProductoCalculado
    };
  });
}

function prepararLineaCompraQTAS_(context) {
  const linea = context.linea || {};
  const tipoItem = normalizarTipoCompraItemQTAS_(linea.tipoItem);
  const referenciaProducto = resolverProductoCanonicoCompraQTAS_(
    context && context.productosActivosIndex,
    linea.item
  );
  const item = referenciaProducto ? referenciaProducto.item : texto_(linea.item);
  const unidadBase = texto_(linea.unidad) || (referenciaProducto ? texto_(referenciaProducto.unidad) : 'und') || 'und';
  const medida = normalizarCantidadUnidadQTAS_(linea.cantidad, unidadBase);
  const cantidad = medida.cantidad;
  const unidad = medida.unidad || normalizarUnidadCanonicaQTAS_(unidadBase) || 'und';
  const costoTotalLinea = redondear_(numero_(linea.costoTotalLinea));
  const costoUnitario = cantidad > 0 ? redondear_(costoTotalLinea / cantidad) : 0;
  const impactaCosto = linea.impactaCosto === undefined
    ? tipoItem !== 'Gasto'
    : Boolean(linea.impactaCosto);

  if (!item) {
    throw new Error(`Falta el item en la linea ${context.index + 1}.`);
  }
  if (cantidad <= 0) {
    throw new Error(`La cantidad debe ser mayor a cero en la linea ${context.index + 1}.`);
  }
  if (costoTotalLinea <= 0) {
    throw new Error(`El costo total debe ser mayor a cero en la linea ${context.index + 1}.`);
  }
  if (impactaCosto && costoUnitario <= 0) {
    throw new Error(`No se pudo calcular costo unitario en la linea ${context.index + 1}.`);
  }
  if (impactaCosto && tipoItem === 'Producto' && !referenciaProducto) {
    throw new Error(
      `La linea ${context.index + 1} impacta costo como Producto, ` +
      'pero el item no coincide con un producto canonico existente.'
    );
  }
  if (
    impactaCosto &&
    tipoItem === 'Producto' &&
    referenciaProducto &&
    normalizarClaveTexto_(unidad) !== normalizarClaveTexto_(referenciaProducto.unidad)
  ) {
    throw new Error(
      `La linea ${context.index + 1} debe usar la unidad ${referenciaProducto.unidad} ` +
      `para impactar costo del producto ${referenciaProducto.item}.`
    );
  }

  return {
    Compra_Detalle_ID: compraDetalleIdQTAS_(context.compraId, context.index + 1),
    Compra_ID: context.compraId,
    Fecha_Compra: context.fechaCompra,
    Proveedor: context.proveedor,
    Tipo_Item: tipoItem,
    Item: item,
    Cantidad: cantidad,
    Unidad: unidad,
    Costo_Total_Linea: costoTotalLinea,
    Costo_Unitario: costoUnitario,
    Impacta_Costo: impactaCosto,
    Comentario_Linea: texto_(linea.comentarioLinea),
    Estado_Registro: QTAS.status.registro.activo
  };
}

function validarCompraQTAS_(payload) {
  if (!payload) throw new Error('Compra vacia.');
  if (!texto_(payload.proveedor)) throw new Error('Falta el proveedor.');
  if (!texto_(payload.medioPago)) throw new Error('Falta el medio de pago.');
  if (!payload.lineas || !payload.lineas.length) {
    throw new Error('La compra debe tener al menos una linea.');
  }
}

function resumenItemsCompraQTAS_(lineas) {
  return (lineas || []).map(linea => {
    const item = texto_(linea.Item);
    const cantidad = formatearCantidad_(linea.Cantidad);
    const unidad = texto_(linea.Unidad);
    return `${item} ${cantidad}${unidad}`;
  }).join(' + ');
}

function compraDetalleIdQTAS_(compraId, lineNumber) {
  return `COMDET-${String(compraId).padStart(6, '0')}-${String(lineNumber).padStart(2, '0')}`;
}

function normalizarTipoCompraItemQTAS_(value) {
  const tipo = texto_(value);
  const encontrado = TIPOS_COMPRA_ITEM.find(item =>
    normalizarClaveTexto_(item) === normalizarClaveTexto_(tipo)
  );
  return encontrado || 'Insumo';
}

function leerProductosActivosCompraQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  return leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.productos))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      item: texto_(row.Producto_Estandar),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad_Default),
      tipoItem: 'Producto'
    }))
    .filter(row => row.item)
    .sort((a, b) => a.item.localeCompare(b.item));
}

function construirIndiceProductosCompraQTAS_(productos) {
  return (productos || []).reduce((acc, row) => {
    const key = normalizarClaveTexto_(row.item);
    if (key && !acc[key]) {
      acc[key] = row;
    }
    return acc;
  }, {});
}

function resolverProductoCanonicoCompraQTAS_(productosActivosIndex, item) {
  const key = normalizarClaveTexto_(item);
  if (!key || !productosActivosIndex) return null;
  return productosActivosIndex[key] || null;
}

function leerCostosHistoricosQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.costosReferencia);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.costosReferencia])) {
    return [];
  }

  return leerObjetos_(sheet)
    .map(row => ({
      costoId: texto_(row.Costo_ID),
      compraId: numero_(row.Compra_ID),
      item: texto_(row.Item),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      costoUnitario: redondear_(numero_(row.Costo_Unitario)),
      proveedor: texto_(row.Proveedor),
      fechaDesde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
      fechaHasta: row.Fecha_Hasta
        ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date())
        : null,
      activo: estaActivo_(row.Activo),
      fuenteTipo: texto_(row.Fuente_Tipo),
      fuenteId: texto_(row.Fuente_ID),
      nota: texto_(row.Nota)
    }))
    .filter(row => row.item && row.unidad && row.costoUnitario > 0 && row.activo);
}

function cargarCostosEnMemoria_() {
  return leerCostosHistoricosQTAS_()
    .filter(row => row.activo)
    .map(row => ({
      item: row.item,
      tipoItem: row.tipoItem,
      unidad: row.unidad,
      costoUnitario: row.costoUnitario,
      proveedor: row.proveedor,
      desde: fecha_(row.fechaDesde),
      hasta: row.fechaHasta ? fecha_(row.fechaHasta) : null
    }));
}

function obtenerCostoVigenteDesdeCache_(costosCache, item, unidad, fechaConsulta, tipoItem) {
  const fechaBase = resolverFechaOperacion_(fechaConsulta, new Date());
  const itemNombre = texto_(item);
  const itemKey = normalizarClaveTexto_(itemNombre);
  const itemUnidad = normalizarUnidadCanonicaQTAS_(unidad);
  const tipoItemKey = normalizarClaveTexto_(tipoItem);

  const matches = (costosCache || []).filter(row => {
    if (normalizarClaveTexto_(row.item) !== itemKey) return false;
    if (normalizarUnidadCanonicaQTAS_(row.unidad) !== itemUnidad) return false;
    return fechaBase >= row.desde && (!row.hasta || fechaBase <= row.hasta);
  });

  if (!matches.length) return 0;

  let scopedMatches = matches;
  if (tipoItemKey) {
    const exactTypeMatches = matches.filter(row =>
      normalizarClaveTexto_(row.tipoItem) === tipoItemKey
    );
    const blankTypeMatches = matches.filter(row =>
      !normalizarClaveTexto_(row.tipoItem)
    );
    scopedMatches = exactTypeMatches.length
      ? exactTypeMatches
      : (blankTypeMatches.length ? blankTypeMatches : matches);
  }

  scopedMatches.sort((a, b) => b.desde - a.desde);
  return redondear_(numero_(scopedMatches[0].costoUnitario));
}

function listarCostosVigentesQTAS_() {
  const hoy = resolverFechaOperacion_(new Date(), new Date());
  const vistos = {};

  return leerCostosHistoricosQTAS_()
    .filter(row =>
      row.activo &&
      row.fechaDesde &&
      row.fechaDesde <= hoy &&
      (!row.fechaHasta || row.fechaHasta >= hoy)
    )
    .sort((a, b) => {
      if (a.fechaDesde.getTime() !== b.fechaDesde.getTime()) return b.fechaDesde - a.fechaDesde;
      if (a.item !== b.item) return a.item.localeCompare(b.item);
      return a.unidad.localeCompare(b.unidad);
    })
    .filter(row => {
      const key = [
        normalizarClaveTexto_(row.tipoItem),
        normalizarClaveTexto_(row.item),
        normalizarClaveTexto_(row.unidad)
      ].join('|');
      if (vistos[key]) return false;
      vistos[key] = true;
      return true;
    })
    .map(row => ({
      costoId: row.costoId,
      compraId: row.compraId,
      item: row.item,
      tipoItem: row.tipoItem,
      unidad: row.unidad,
      costoUnitario: row.costoUnitario,
      proveedor: row.proveedor,
      fechaDesde: fechaInput_(row.fechaDesde),
      fechaHasta: row.fechaHasta ? fechaInput_(row.fechaHasta) : '',
      nota: row.nota
    }));
}

function listarComprasRecientesQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.compras);

  return leerObjetos_(sheet)
    .filter(row => numero_(row.Compra_ID) > 0 && !esRegistroAnulado_(row.Estado_Registro))
    .sort((a, b) => {
      const fechaA = valorFechaCompraCanonicaQTAS_(a, new Date());
      const fechaB = valorFechaCompraCanonicaQTAS_(b, new Date());
      if (fechaA.getTime() !== fechaB.getTime()) return fechaB - fechaA;
      return numero_(b.Compra_ID) - numero_(a.Compra_ID);
    })
    .slice(0, 12)
    .map(row => ({
      compraId: numero_(row.Compra_ID),
      fechaCompra: fechaInput_(row.Fecha_Compra),
      proveedor: texto_(row.Proveedor),
      itemsResumen: texto_(row.Items_Resumen),
      totalCompra: redondear_(numero_(row.Total_Compra)),
      medioPago: texto_(row.Medio_Pago)
    }));
}

function construirItemsSugeridosComprasQTAS_(productos, costosVigentes) {
  const mapa = {};

  (productos || []).forEach(row => {
    const key = [
      normalizarClaveTexto_(row.tipoItem || 'Producto'),
      normalizarClaveTexto_(row.item)
    ].join('|');
    if (!key) return;
    mapa[key] = {
      item: row.item,
      unidad: row.unidad,
      tipoItem: row.tipoItem,
      costoUnitario: 0,
      fechaDesde: ''
    };
  });

  (costosVigentes || []).forEach(row => {
    const key = [
      normalizarClaveTexto_(row.tipoItem || 'Insumo'),
      normalizarClaveTexto_(row.item)
    ].join('|');
    if (!key) return;

    if (!mapa[key]) {
      mapa[key] = {
        item: row.item,
        unidad: row.unidad,
        tipoItem: row.tipoItem || 'Insumo',
        costoUnitario: row.costoUnitario,
        fechaDesde: row.fechaDesde
      };
      return;
    }

    mapa[key].costoUnitario = row.costoUnitario;
    mapa[key].fechaDesde = row.fechaDesde;
    if (!mapa[key].unidad) mapa[key].unidad = row.unidad;
  });

  return Object.keys(mapa)
    .map(key => mapa[key])
    .sort((a, b) => a.item.localeCompare(b.item));
}

function actualizarCostosDesdeCompraQTAS_(context) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
  const headers = getHeaders_(sheet);
  const fechaCompra = resolverFechaOperacion_(context.fechaCompra, new Date());
  const rows = leerObjetosConMeta_(sheet);
  let actualizados = 0;

  (context.lineas || []).forEach(linea => {
    if (!linea.Impacta_Costo) return;
    if (!texto_(linea.Item) || !texto_(linea.Unidad)) return;
    if (numero_(linea.Costo_Unitario) <= 0) return;

    actualizados += upsertCostoReferenciaHistoricoQTAS_({
      sheet: sheet,
      headers: headers,
      rows: rows,
      fechaDesde: fechaCompra,
      proveedor: context.proveedor,
      compraId: context.compraId,
      fuenteTipo: 'Compra',
      fuenteId: texto_(linea.Compra_Detalle_ID),
      comentario: texto_(linea.Comentario_Linea),
      linea: linea
    });
  });

  return actualizados;
}

function upsertCostoReferenciaDesdeCompraQTAS_(context) {
  return upsertCostoReferenciaHistoricoQTAS_(Object.assign({}, context, {
    fechaDesde: context.fechaCompra,
    fuenteTipo: 'Compra',
    fuenteId: context && context.linea ? texto_(context.linea.Compra_Detalle_ID) : '',
    comentario: context && context.linea ? texto_(context.linea.Comentario_Linea) : ''
  }));
}

function upsertCostoReferenciaHistoricoQTAS_(context) {
  const sheet = context.sheet;
  const headers = context.headers;
  const rows = context.rows;
  const fechaCompra = resolverFechaOperacion_(context.fechaDesde || context.fechaCompra, new Date());
  const tipoItemKey = normalizarClaveTexto_(context.linea.Tipo_Item);
  const itemKey = normalizarClaveTexto_(context.linea.Item);
  const unitKey = normalizarClaveTexto_(context.linea.Unidad);
  const fuenteTipo = texto_(context.fuenteTipo) || 'Manual';
  const fuenteId = texto_(context.fuenteId);
  const comentario = texto_(context.comentario) || texto_(context.linea.Comentario_Linea);
  const nota = unirUnicos_([
    comentario,
    context.compraId ? `Compra ${context.compraId}` : '',
    texto_(context.nota)
  ]);

  const existentes = (rows || [])
    .filter(row =>
      normalizarClaveTexto_(row.Tipo_Item) === tipoItemKey &&
      normalizarClaveTexto_(row.Item) === itemKey &&
      normalizarClaveTexto_(row.Unidad) === unitKey &&
      estaActivo_(row.Activo)
    )
    .map(row => ({
      rowNumber: row.__rowNumber,
      desde: resolverFechaOperacion_(row.Fecha_Desde, fechaCompra),
      hasta: row.Fecha_Hasta ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || fechaCompra) : null,
      raw: row
    }))
    .sort((a, b) => a.desde - b.desde);

  const exacta = existentes.find(row => row.desde.getTime() === fechaCompra.getTime());

  if (exacta) {
    const actualizada = Object.assign({}, exacta.raw, {
      Compra_ID: numero_(context.compraId),
      Item: texto_(context.linea.Item),
      Tipo_Item: texto_(context.linea.Tipo_Item),
      Unidad: texto_(context.linea.Unidad),
      Costo_Unitario: redondear_(numero_(context.linea.Costo_Unitario)),
      Proveedor: texto_(context.proveedor),
      Activo: true,
      Fuente_Tipo: fuenteTipo,
      Fuente_ID: fuenteId,
      Nota: nota
    });
    actualizarFilaObjeto_(sheet, exacta.rowNumber, headers, actualizada);
    Object.assign(exacta.raw, actualizada);
    return 1;
  }

  const insertIndex = existentes.findIndex(row => row.desde.getTime() > fechaCompra.getTime());
  const prev = insertIndex > 0
    ? existentes[insertIndex - 1]
    : (insertIndex === -1 ? existentes[existentes.length - 1] : null);
  const next = insertIndex >= 0 ? existentes[insertIndex] : null;
  const caeDentroPrev = prev && (!prev.hasta || fechaCompra.getTime() <= prev.hasta.getTime());
  let fechaHastaNueva = '';

  if (caeDentroPrev) {
    const fechaHastaAnterior = prev.hasta;

    if (fechaCompra.getTime() > prev.desde.getTime()) {
      const filaCerrada = Object.assign({}, prev.raw, {
        Fecha_Hasta: diaAnterior_(fechaCompra)
      });
      actualizarFilaObjeto_(sheet, prev.rowNumber, headers, filaCerrada);
      Object.assign(prev.raw, filaCerrada);
    }

    fechaHastaNueva = fechaHastaAnterior || '';
  } else if (next) {
    fechaHastaNueva = diaAnterior_(next.desde);
  }

  const nueva = {
    Costo_ID: siguienteIdConPrefijo_(sheet, 'Costo_ID', 'COST-', 4),
    Compra_ID: numero_(context.compraId),
    Item: texto_(context.linea.Item),
    Tipo_Item: texto_(context.linea.Tipo_Item),
    Unidad: texto_(context.linea.Unidad),
    Costo_Unitario: redondear_(numero_(context.linea.Costo_Unitario)),
    Proveedor: texto_(context.proveedor),
    Fecha_Desde: fechaCompra,
    Fecha_Hasta: fechaHastaNueva,
    Activo: true,
    Fuente_Tipo: fuenteTipo,
    Fuente_ID: fuenteId,
    Nota: nota
  };

  escribirFilas_(sheet, [filaDesdeHeaders_(headers, nueva)]);
  rows.push(Object.assign({ __rowNumber: sheet.getLastRow() }, nueva));
  return 1;
}
