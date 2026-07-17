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
    origenesFondos: listarOrigenesFondosDisponiblesQTAS_(),
    productos: productos,
    itemsSugeridos: construirItemsSugeridosComprasQTAS_(productos, costosVigentes, { ss: ss }),
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
    const medioPago = normalizarMedioPagoQTAS_(payload.medioPago);

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
    const origenFondos = normalizarOrigenFondosQTAS_(payload.origenFondos);
    const productosActivos = leerProductosActivosCompraQTAS_(ss);
    const productosActivosIndex = construirIndiceProductosCompraQTAS_(productosActivos);
    const itemsCatalogoIndex = construirIndiceItemsSugeridosCompraQTAS_(
      construirItemsSugeridosComprasQTAS_(productosActivos, listarCostosVigentesQTAS_(), { ss: ss })
    );

    const lineasPreparadas = (payload.lineas || []).map((linea, index) =>
      prepararLineaCompraQTAS_({
        compraId: compraId,
        fechaCompra: fechaCompra,
        proveedor: proveedor,
        linea: linea,
        index: index,
        productosActivosIndex: productosActivosIndex,
        itemsCatalogoIndex: itemsCatalogoIndex
      })
    );

    const totalCompra = redondear_(sumar_(lineasPreparadas.map(item => item.Costo_Total_Linea)));
    const itemsResumen = resumenItemsCompraQTAS_(lineasPreparadas);
    const origenesFondosCompra = construirFilasCompraOrigenesFondosQTAS_({
      compraId: compraId,
      fechaCompra: fechaCompraBase,
      origenFondos: origenFondos,
      lineas: lineasPreparadas,
      comentarioCompra: comentarioCompra
    });

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

    registrarOrigenesFondosCompraQTAS_(origenesFondosCompra);

    const costosActualizados = actualizarCostosDesdeCompraQTAS_({
      compraId: compraId,
      proveedor: proveedor,
      fechaCompra: fechaCompraBase,
      lineas: lineasPreparadas
    });
    const inventario = sincronizarInventarioDesdeCompraQTAS_({
      ss: ss,
      compraId: compraId,
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
      inventario: inventario,
      origenFondos: origenesFondosCompra.origenFondos,
      origenesFondosAsignados: origenesFondosCompra.rows.length,
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
  const referenciaItem = resolverItemCanonicoCompraQTAS_(
    context && context.itemsCatalogoIndex,
    tipoItem,
    linea.item
  );
  const referenciaCanonica = referenciaProducto || referenciaItem;
  const item = referenciaCanonica ? referenciaCanonica.item : texto_(linea.item);
  const unidadBase = texto_(linea.unidad) || (referenciaCanonica ? texto_(referenciaCanonica.unidad) : 'und') || 'und';
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

function construirIndiceItemsSugeridosCompraQTAS_(items) {
  return (items || []).reduce((acc, row) => {
    const key = claveItemSugeridoCompraQTAS_(row && row.tipoItem, row && row.item);
    if (key && !acc[key]) {
      acc[key] = row;
    }
    return acc;
  }, {});
}

function resolverItemCanonicoCompraQTAS_(itemsCatalogoIndex, tipoItem, item) {
  const key = claveItemSugeridoCompraQTAS_(tipoItem, item);
  if (!key || !itemsCatalogoIndex) return null;
  return itemsCatalogoIndex[key] || null;
}

function leerCostosHistoricosQTAS_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActive();
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

function cargarCostosEnMemoria_(spreadsheet) {
  return leerCostosHistoricosQTAS_(spreadsheet)
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

  const cache = costosCache || [];
  const index = obtenerIndiceCostosHistoricosQTAS_(cache);
  const matches = (index[`${itemKey}|${itemUnidad}`] || []).filter(row =>
    fechaBase >= row.desde && (!row.hasta || fechaBase <= row.hasta)
  );

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

function obtenerIndiceCostosHistoricosQTAS_(costosCache) {
  const cache = costosCache || [];
  if (cache.__qtasCostosIndex) return cache.__qtasCostosIndex;

  const index = {};
  cache.forEach(row => {
    const key = [
      normalizarClaveTexto_(row.item),
      normalizarUnidadCanonicaQTAS_(row.unidad)
    ].join('|');
    if (!index[key]) index[key] = [];
    index[key].push(row);
  });

  Object.keys(index).forEach(key => {
    index[key].sort((a, b) => b.desde - a.desde);
  });
  Object.defineProperty(cache, '__qtasCostosIndex', {
    value: index,
    enumerable: false,
    configurable: true
  });
  return index;
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
  const origenesPorCompra = leerOrigenesFondosPorCompraQTAS_();

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
      medioPago: texto_(row.Medio_Pago),
      origenFondos: resumenOrigenesFondosCompraQTAS_(origenesPorCompra[numero_(row.Compra_ID)])
    }));
}

function listarOrigenesFondosDisponiblesQTAS_() {
  const vistos = {};

  return leerReglasOrigenesFondosQTAS_()
    .map(row => normalizarOrigenFondosQTAS_(row.origenFondos))
    .filter(Boolean)
    .filter(origen => {
      const key = normalizarClaveTexto_(origen);
      if (!key || vistos[key]) return false;
      vistos[key] = true;
      return true;
    })
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
}

function leerReglasOrigenesFondosQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.origenesFondosReglas);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.origenesFondosReglas])) {
    return [];
  }

  const agrupadas = {};

  leerObjetos_(sheet).forEach(row => {
    const reglaId = texto_(row.Regla_ID);
    const origenFondos = texto_(row.Origen_Fondos);
    if (!reglaId || !origenFondos) return;

    if (!agrupadas[reglaId]) {
      agrupadas[reglaId] = {
        reglaId: reglaId,
        origenFondos: origenFondos,
        desde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
        hasta: row.Fecha_Hasta
          ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date())
          : null,
        steve: 0,
        majo: 0,
        mush: 0,
        activo: true,
        nota: '',
        aportantes: []
      };
    }

    const aportante = normalizarAportanteOrigenFondosQTAS_(row.Aportante);
    const porcentaje = redondear_(numero_(row.Porcentaje));
    if (!aportante) return;

    if (aportante === 'Steve') agrupadas[reglaId].steve = porcentaje;
    if (aportante === 'Majo') agrupadas[reglaId].majo = porcentaje;
    if (aportante === 'Mush') agrupadas[reglaId].mush = porcentaje;
    agrupadas[reglaId].aportantes.push({
      aportante: aportante,
      porcentaje: porcentaje
    });
    agrupadas[reglaId].nota = unirUnicos_([
      agrupadas[reglaId].nota,
      texto_(row.Nota)
    ]);
  });

  const salida = Object.keys(agrupadas)
    .map(key => {
      const row = agrupadas[key];
      row.aportantes = row.aportantes
        .filter(item => numero_(item.porcentaje) > 0)
        .sort((a, b) => ordenAportanteOrigenFondosQTAS_(a.aportante) - ordenAportanteOrigenFondosQTAS_(b.aportante));
      return row;
    })
    .filter(row => row.origenFondos)
    .sort((a, b) => {
      const origen = a.origenFondos.localeCompare(b.origenFondos, undefined, { sensitivity: 'base' });
      if (origen !== 0) return origen;
      return a.desde - b.desde;
    });

  validarReglasOrigenesFondosQTAS_(salida);
  return salida;
}

function cargarReglasOrigenesFondosEnMemoriaQTAS_() {
  const namespace = 'origenes_fondos_reglas_memoria';

  return obtenerMemoEjecucionQTAS_(`cache:${namespace}`, () => {
    const cached = leerCacheDocumentoQTAS_(namespace);
    if (cached && Array.isArray(cached.rows)) {
      return cached.rows;
    }

    const rows = leerReglasOrigenesFondosQTAS_()
      .map(row => ({
        reglaId: texto_(row.reglaId),
        origenFondos: texto_(row.origenFondos),
        desde: fechaInput_(row.desde),
        hasta: row.hasta ? fechaInput_(row.hasta) : '',
        steve: redondear_(numero_(row.steve)),
        majo: redondear_(numero_(row.majo)),
        mush: redondear_(numero_(row.mush)),
        nota: texto_(row.nota),
        aportantes: construirAportantesOrigenFondosQTAS_(row)
      }))
      .sort((a, b) => {
        const origen = a.origenFondos.localeCompare(b.origenFondos, undefined, { sensitivity: 'base' });
        if (origen !== 0) return origen;
        return b.desde.localeCompare(a.desde);
      });

    guardarCacheDocumentoQTAS_(namespace, { rows: rows }, 300);
    return rows;
  });
}

function validarReglasOrigenesFondosQTAS_(reglas) {
  const porOrigen = {};

  (reglas || []).forEach(regla => {
    if (!regla || !texto_(regla.origenFondos)) return;

    const total = redondear_(
      numero_(regla.steve) +
      numero_(regla.majo) +
      numero_(regla.mush)
    );
    if (Math.abs(total - 100) > 0.01) {
      throw new Error(`La regla ${texto_(regla.reglaId)} de ${texto_(regla.origenFondos)} debe sumar 100%.`);
    }

    const key = normalizarClaveTexto_(regla.origenFondos);
    if (!porOrigen[key]) porOrigen[key] = [];
    porOrigen[key].push(regla);
  });

  Object.keys(porOrigen).forEach(key => {
    const rows = porOrigen[key]
      .slice()
      .sort((a, b) => a.desde - b.desde);

    for (let index = 0; index < rows.length - 1; index += 1) {
      const actual = rows[index];
      const siguiente = rows[index + 1];

      if (!actual.hasta) {
        throw new Error(`La regla ${texto_(actual.reglaId)} no puede quedar abierta si existe una regla posterior para ${texto_(actual.origenFondos)}.`);
      }

      if (resolverFechaOperacion_(actual.hasta, new Date()) >= resolverFechaOperacion_(siguiente.desde, new Date())) {
        throw new Error(`Las reglas ${texto_(actual.reglaId)} y ${texto_(siguiente.reglaId)} de ${texto_(actual.origenFondos)} se traslapan en fechas.`);
      }
    }
  });
}

function construirAportantesOrigenFondosQTAS_(row) {
  return [
    { aportante: 'Steve', porcentaje: redondear_(numero_(row && row.steve)) },
    { aportante: 'Majo', porcentaje: redondear_(numero_(row && row.majo)) },
    { aportante: 'Mush', porcentaje: redondear_(numero_(row && row.mush)) }
  ].filter(item => numero_(item.porcentaje) > 0);
}

function normalizarAportanteOrigenFondosQTAS_(value) {
  const key = normalizarClaveTexto_(value);
  if (!key) return '';
  if (key === 'steve') return 'Steve';
  if (key === 'majo') return 'Majo';
  if (key === 'mush') return 'Mush';
  return '';
}

function ordenAportanteOrigenFondosQTAS_(aportante) {
  const key = normalizarAportanteOrigenFondosQTAS_(aportante);
  if (key === 'Steve') return 1;
  if (key === 'Majo') return 2;
  if (key === 'Mush') return 3;
  return 99;
}

function obtenerReglaOrigenFondosVigenteDesdeCacheQTAS_(reglasCache, origenFondos, fechaBase) {
  const origenKey = normalizarClaveTexto_(normalizarOrigenFondosQTAS_(origenFondos));
  const fechaConsulta = fechaInput_(resolverFechaOperacion_(fechaBase, new Date()));
  const candidatas = (reglasCache || [])
    .filter(row => normalizarClaveTexto_(normalizarOrigenFondosQTAS_(row.origenFondos)) === origenKey);

  const matches = candidatas.filter(regla =>
    fechaConsulta >= regla.desde &&
    (!regla.hasta || fechaConsulta <= regla.hasta)
  );

  if (!matches.length) {
    const ordenadas = candidatas.slice().sort((a, b) => a.desde.localeCompare(b.desde));
    if (!ordenadas.length) return null;

    const masAntigua = ordenadas[0];
    if (fechaConsulta < masAntigua.desde) {
      return masAntigua;
    }

    return null;
  }

  matches.sort((a, b) => b.desde.localeCompare(a.desde));
  return matches[0];
}

function obtenerSnapshotOrigenFondosDesdeCacheQTAS_(reglasCache, origenFondos, fechaBase) {
  const regla = obtenerReglaOrigenFondosVigenteDesdeCacheQTAS_(reglasCache, origenFondos, fechaBase);
  if (!regla) {
    throw new Error(
      `No hay regla de origen de fondos vigente para ${texto_(origenFondos)} en la fecha ${fechaInput_(fechaBase)}.`
    );
  }

  return {
    reglaId: texto_(regla.reglaId),
    origenFondos: normalizarOrigenFondosQTAS_(regla.origenFondos),
    steve: redondear_(numero_(regla.steve)),
    majo: redondear_(numero_(regla.majo)),
    mush: redondear_(numero_(regla.mush)),
    aportantes: construirAportantesOrigenFondosQTAS_(regla)
  };
}

function construirFilasCompraOrigenesFondosQTAS_(context) {
  const origenFondos = normalizarOrigenFondosQTAS_(context && context.origenFondos);
  const lineas = context && context.lineas ? context.lineas : [];
  if (!origenFondos || !lineas.length) {
    return {
      origenFondos: origenFondos,
      rows: []
    };
  }

  const snapshot = obtenerSnapshotOrigenFondosDesdeCacheQTAS_(
    cargarReglasOrigenesFondosEnMemoriaQTAS_(),
    origenFondos,
    context.fechaCompra
  );
  const fechaCompra = resolverFechaOperacion_(context.fechaCompra, new Date());
  const rows = [];

  lineas.forEach(linea => {
    const montos = distribuirMontoOrigenFondosQTAS_(
      numero_(linea && linea.Costo_Total_Linea),
      snapshot.aportantes
    );

    montos.forEach(item => {
      rows.push({
        Compra_Origen_ID: construirCompraOrigenIdQTAS_(linea.Compra_Detalle_ID, item.aportante),
        Compra_ID: numero_(context.compraId),
        Compra_Detalle_ID: texto_(linea.Compra_Detalle_ID),
        Fecha_Compra: fechaCompra,
        Origen_Fondos: snapshot.origenFondos,
        Aportante: item.aportante,
        Porcentaje: item.porcentaje,
        Monto_Asignado: item.montoAsignado,
        Fuente_Registro: `Regla ${snapshot.reglaId}`,
        Nota: unirUnicos_([
          texto_(context.comentarioCompra),
          texto_(linea && linea.Comentario_Linea),
          `Asignacion automatica de ${snapshot.origenFondos}`
        ])
      });
    });
  });

  return {
    origenFondos: snapshot.origenFondos,
    reglaId: snapshot.reglaId,
    rows: rows
  };
}

function registrarOrigenesFondosCompraQTAS_(context) {
  const rows = context && context.rows ? context.rows : [];
  if (!rows.length) return 0;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.compraOrigenesFondos);
  const headers = getHeaders_(sheet);
  escribirFilas_(sheet, rows.map(row => filaDesdeHeaders_(headers, row)));
  return rows.length;
}

function distribuirMontoOrigenFondosQTAS_(montoBase, aportantes) {
  const activos = (aportantes || [])
    .map(item => ({
      aportante: normalizarAportanteOrigenFondosQTAS_(item && item.aportante),
      porcentaje: redondear_(numero_(item && item.porcentaje))
    }))
    .filter(item => item.aportante && item.porcentaje > 0)
    .sort((a, b) => ordenAportanteOrigenFondosQTAS_(a.aportante) - ordenAportanteOrigenFondosQTAS_(b.aportante));

  const monto = redondear_(Math.max(numero_(montoBase), 0));
  let restante = monto;

  return activos.map((item, index) => {
    const esUltimo = index === activos.length - 1;
    const montoAsignado = esUltimo
      ? restante
      : redondear_(monto * item.porcentaje / 100);
    restante = redondear_(restante - montoAsignado);

    return {
      aportante: item.aportante,
      porcentaje: item.porcentaje,
      montoAsignado: montoAsignado
    };
  });
}

function construirCompraOrigenIdQTAS_(compraDetalleId, aportante) {
  const detalle = texto_(compraDetalleId).replace(/[^A-Za-z0-9_-]+/g, '_');
  const persona = texto_(aportante).replace(/[^A-Za-z0-9_-]+/g, '_');
  return `COF-${detalle}-${persona}`.slice(0, 99);
}

function leerOrigenesFondosPorCompraQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.compraOrigenesFondos);

  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }

  if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.compraOrigenesFondos])) {
    return {};
  }

  return leerObjetos_(sheet).reduce((acc, row) => {
    const compraId = numero_(row.Compra_ID);
    if (compraId <= 0) return acc;
    if (!acc[compraId]) {
      acc[compraId] = {
        origenes: [],
        aportantes: []
      };
    }

    const origen = texto_(row.Origen_Fondos);
    if (origen && acc[compraId].origenes.indexOf(origen) < 0) {
      acc[compraId].origenes.push(origen);
    }

    const aportante = normalizarAportanteOrigenFondosQTAS_(row.Aportante);
    if (aportante && acc[compraId].aportantes.indexOf(aportante) < 0) {
      acc[compraId].aportantes.push(aportante);
    }

    return acc;
  }, {});
}

function resumenOrigenesFondosCompraQTAS_(row) {
  if (!row || !row.origenes || !row.origenes.length) return '';
  return row.origenes.join(' | ');
}

function construirItemsSugeridosComprasQTAS_(productos, costosVigentes, options) {
  const settings = options || {};
  const ss = settings.ss || SpreadsheetApp.getActive();
  const mapa = {};

  (productos || []).forEach(row => {
    registrarItemSugeridoCompraQTAS_(mapa, row, {
      prioridad: 400,
      origenCatalogo: 'Productos activos'
    });
  });

  leerItemsCosteoCompraQTAS_(ss).forEach(row => {
    registrarItemSugeridoCompraQTAS_(mapa, row, {
      prioridad: 300,
      origenCatalogo: 'Plantilla costeo'
    });
  });

  (costosVigentes || []).forEach(row => {
    registrarItemSugeridoCompraQTAS_(mapa, row, {
      prioridad: 350,
      origenCatalogo: 'Costos vigentes',
      proveedor: row.proveedor,
      fechaDesde: row.fechaDesde
    });
  });

  leerItemsHistoricosCompraQTAS_(ss).forEach(row => {
    registrarItemSugeridoCompraQTAS_(mapa, row, {
      prioridad: 200,
      origenCatalogo: 'Historial compras',
      proveedor: row.proveedor,
      fechaDesde: row.fechaDesde,
      sumarUso: true
    });
  });

  return Object.keys(mapa)
    .map(key => {
      const item = Object.assign({}, mapa[key]);
      delete item._prioridad;
      delete item._ultimaFechaComparable;
      return item;
    })
    .sort((a, b) => a.item.localeCompare(b.item));
}

function leerItemsHistoricosCompraQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(QTAS.sheets.compraDetalle);

  return leerObjetos_(sheet)
    .filter(row => !esRegistroAnulado_(row.Estado_Registro))
    .map(row => ({
      item: texto_(row.Item),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Item),
      costoUnitario: redondear_(numero_(row.Costo_Unitario)),
      proveedor: texto_(row.Proveedor),
      fechaDesde: row.Fecha_Compra ? fechaInput_(row.Fecha_Compra) : ''
    }))
    .filter(row => row.item && row.unidad);
}

function leerItemsCosteoCompraQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const componentes = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.productoComponentes))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      item: texto_(row.Item_Componente),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad_Componente),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Componente)
    }));
  const reglas = leerObjetos_(spreadsheet.getSheetByName(QTAS.sheets.productoReglasCosto))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      item: texto_(row.Item_Componente),
      unidad: normalizarUnidadCanonicaQTAS_(row.Unidad_Componente),
      tipoItem: normalizarTipoCompraItemQTAS_(row.Tipo_Componente)
    }));

  return componentes
    .concat(reglas)
    .filter(row => row.item && row.unidad);
}

function registrarItemSugeridoCompraQTAS_(mapa, row, options) {
  const settings = options || {};
  const tipoItem = normalizarTipoCompraItemQTAS_(row && row.tipoItem);
  const item = texto_(row && row.item);
  const unidad = normalizarUnidadCanonicaQTAS_(row && row.unidad);
  const key = claveItemSugeridoCompraQTAS_(tipoItem, item);

  if (!key || !unidad) return;

  const prioridad = Math.max(0, numero_(settings.prioridad));
  if (!mapa[key]) {
    mapa[key] = {
      item: item,
      unidad: unidad,
      tipoItem: tipoItem,
      costoUnitario: 0,
      fechaDesde: '',
      proveedor: '',
      origenCatalogo: texto_(settings.origenCatalogo),
      usos: 0,
      ultimaFecha: '',
      _prioridad: prioridad,
      _ultimaFechaComparable: ''
    };
  }

  const target = mapa[key];
  const fechaDesde = fechaTextoCatalogoCompraQTAS_(settings.fechaDesde || row.fechaDesde);
  const costoUnitario = redondear_(numero_(row && row.costoUnitario));

  if (prioridad > numero_(target._prioridad)) {
    target.item = item;
    target.unidad = unidad;
    target.tipoItem = tipoItem;
    target.origenCatalogo = texto_(settings.origenCatalogo) || target.origenCatalogo;
    target._prioridad = prioridad;
  } else if (!target.unidad) {
    target.unidad = unidad;
  }

  if (costoUnitario > 0 && (!target.costoUnitario || fechaDesde >= texto_(target.fechaDesde))) {
    target.costoUnitario = costoUnitario;
    target.fechaDesde = fechaDesde;
    target.proveedor = texto_(settings.proveedor || row.proveedor);
  }

  if (settings.sumarUso) {
    target.usos = numero_(target.usos) + 1;
  }

  if (fechaDesde && fechaDesde >= texto_(target._ultimaFechaComparable)) {
    target.ultimaFecha = fechaDesde;
    target._ultimaFechaComparable = fechaDesde;
  }
}

function claveItemSugeridoCompraQTAS_(tipoItem, item) {
  const tipoKey = normalizarClaveTexto_(tipoItem || 'Insumo');
  const itemKey = normalizarClaveTexto_(item);
  if (!tipoKey || !itemKey) return '';
  return [tipoKey, itemKey].join('|');
}

function fechaTextoCatalogoCompraQTAS_(value) {
  return value ? fechaInput_(value) : '';
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

function eliminarCompraRecienteQTAS(payload) {
  return withScriptLock_('eliminar compra reciente', () => {
    validarModeloSoloLecturaQTAS_({
      sheetNames: [
        QTAS.sheets.compras,
        QTAS.sheets.compraDetalle,
        QTAS.sheets.costosReferencia
      ],
      validarConfig: false
    });

    const compraId = numero_(payload && payload.compraId);
    if (compraId <= 0) {
      throw new Error('Falta la compra a eliminar.');
    }

    const recientes = listarComprasRecientesQTAS_();
    if (!recientes.some(row => numero_(row.compraId) === compraId)) {
      throw new Error('Solo se pueden eliminar compras recientes desde el ERP.');
    }

    const ss = SpreadsheetApp.getActive();
    const comprasSheet = ss.getSheetByName(QTAS.sheets.compras);
    const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
    const costosSheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
    const comprasHeaders = getHeaders_(comprasSheet);
    const detalleHeaders = getHeaders_(detalleSheet);
    const costosHeaders = getHeaders_(costosSheet);
    const origenesRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.compraOrigenesFondos);

    if (!origenesRef.ok && origenesRef.sheet) {
      throw new Error(`${origenesRef.reason} Repara el modelo antes de eliminar compras desde el ERP.`);
    }

    const comprasAntes = leerObjetos_(comprasSheet);
    const compra = comprasAntes.find(row => numero_(row.Compra_ID) === compraId);
    if (!compra || esRegistroAnulado_(compra.Estado_Registro)) {
      throw new Error('La compra ya no esta disponible para eliminar.');
    }

    const detalleAntes = leerObjetos_(detalleSheet);
    const costosAntes = leerObjetos_(costosSheet);
    const origenesAntes = origenesRef.ok ? leerObjetos_(origenesRef.sheet) : [];

    const comprasDespues = comprasAntes.filter(row => numero_(row.Compra_ID) !== compraId);
    const detalleDespues = detalleAntes.filter(row => numero_(row.Compra_ID) !== compraId);
    const origenesDespues = origenesAntes.filter(row => numero_(row.Compra_ID) !== compraId);

    sobrescribirObjetosHojaQTAS_(comprasSheet, comprasHeaders, comprasDespues);
    sobrescribirObjetosHojaQTAS_(detalleSheet, detalleHeaders, detalleDespues);
    if (origenesRef.ok) {
      sobrescribirObjetosHojaQTAS_(origenesRef.sheet, origenesRef.headers, origenesDespues);
    }

    const costosReconstruidos = reconstruirCostosReferenciaDesdeFuentesQTAS_({
      ss: ss
    });
    const costoProductoCalculado = reconstruirCostoProductoCalculadoInternoQTAS_({
      ss: ss,
      fechaBase: new Date(),
      ahora: new Date()
    });
    const inventario = reconstruirInventarioInternoQTAS_({
      ss: ss
    });

    limpiarCachesEjecucionQTAS_();

    return {
      ok: true,
      compraId: compraId,
      proveedor: texto_(compra.Proveedor),
      removed: {
        compras: comprasAntes.length - comprasDespues.length,
        compraDetalle: detalleAntes.length - detalleDespues.length,
        compraOrigenesFondos: origenesAntes.length - origenesDespues.length,
        costosPreviosCompra: costosAntes.filter(row => numero_(row.Compra_ID) === compraId).length
      },
      costosReconstruidos: costosReconstruidos,
      costoProductoCalculado: costoProductoCalculado,
      inventario: inventario,
      comprasRecientes: listarComprasRecientesQTAS_()
    };
  });
}

function reconstruirCostosReferenciaDesdeFuentesQTAS_(payload) {
  const settings = Object.assign({
    ss: SpreadsheetApp.getActive()
  }, payload || {});
  const ss = settings.ss || SpreadsheetApp.getActive();
  const costosSheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
  const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
  const headers = getHeaders_(costosSheet);
  const rowsActuales = leerObjetos_(costosSheet);
  const detalleCompra = leerObjetos_(detalleSheet)
    .filter(row =>
      numero_(row.Compra_ID) > 0 &&
      !esRegistroAnulado_(row.Estado_Registro) &&
      (row.Impacta_Costo === true || String(row.Impacta_Costo).toLowerCase() === 'true') &&
      numero_(row.Costo_Unitario) > 0 &&
      texto_(row.Item) &&
      texto_(row.Unidad)
    );
  const eventos = [];
  const idsExistentes = {};

  rowsActuales.forEach((row, index) => {
    const tipoItem = normalizarTipoCompraItemQTAS_(row.Tipo_Item);
    const item = texto_(row.Item);
    const unidad = normalizarUnidadCanonicaQTAS_(row.Unidad);
    const key = claveCostoHistoricoQTAS_(tipoItem, item, unidad);
    const fuenteTipo = texto_(row.Fuente_Tipo) || (numero_(row.Compra_ID) > 0 ? 'Compra' : 'Manual');
    const fechaDesde = row.Fecha_Desde ? resolverFechaOperacion_(row.Fecha_Desde, new Date()) : null;
    if (!key || !fechaDesde || !estaActivo_(row.Activo)) return;

    const identity = construirIdentidadEventoCostoHistoricoQTAS_({
      key: key,
      fuenteTipo: fuenteTipo,
      fuenteId: texto_(row.Fuente_ID) || texto_(row.Costo_ID),
      fechaDesde: fechaDesde
    });
    if (!idsExistentes[identity]) {
      idsExistentes[identity] = texto_(row.Costo_ID);
    }

    if (normalizarClaveTexto_(fuenteTipo) === normalizarClaveTexto_('Compra') || numero_(row.Compra_ID) > 0) {
      return;
    }

    eventos.push({
      key: key,
      tipoItem: tipoItem,
      item: item,
      unidad: unidad,
      costoUnitario: redondear_(numero_(row.Costo_Unitario)),
      proveedor: texto_(row.Proveedor),
      fechaDesde: fechaDesde,
      fuenteTipo: fuenteTipo,
      fuenteId: texto_(row.Fuente_ID) || texto_(row.Costo_ID),
      compraId: numero_(row.Compra_ID),
      nota: texto_(row.Nota),
      sourcePriority: prioridadFuenteCostoHistoricoQTAS_(fuenteTipo),
      sourceOrder: index + 1
    });
  });

  detalleCompra.forEach((row, index) => {
    const tipoItem = normalizarTipoCompraItemQTAS_(row.Tipo_Item);
    const item = texto_(row.Item);
    const unidad = normalizarUnidadCanonicaQTAS_(row.Unidad);
    const key = claveCostoHistoricoQTAS_(tipoItem, item, unidad);
    if (!key) return;

    eventos.push({
      key: key,
      tipoItem: tipoItem,
      item: item,
      unidad: unidad,
      costoUnitario: redondear_(numero_(row.Costo_Unitario)),
      proveedor: texto_(row.Proveedor),
      fechaDesde: resolverFechaOperacion_(row.Fecha_Compra, new Date()),
      fuenteTipo: 'Compra',
      fuenteId: texto_(row.Compra_Detalle_ID),
      compraId: numero_(row.Compra_ID),
      nota: unirUnicos_([
        texto_(row.Comentario_Linea),
        numero_(row.Compra_ID) > 0 ? `Compra ${numero_(row.Compra_ID)}` : ''
      ]),
      sourcePriority: prioridadFuenteCostoHistoricoQTAS_('Compra'),
      sourceOrder: index + 1
    });
  });

  const eventosPorClave = {};
  eventos.forEach(evento => {
    if (!eventosPorClave[evento.key]) {
      eventosPorClave[evento.key] = [];
    }
    eventosPorClave[evento.key].push(evento);
  });

  let nextCostoId = siguienteIdConPrefijo_(costosSheet, 'Costo_ID', 'COST-', 4);
  const reconstruidas = [];

  Object.keys(eventosPorClave)
    .sort((a, b) => a.localeCompare(b))
    .forEach(key => {
      const eventosUnicosPorFecha = {};

      eventosPorClave[key]
        .slice()
        .sort((a, b) => {
          if (a.fechaDesde.getTime() !== b.fechaDesde.getTime()) return a.fechaDesde - b.fechaDesde;
          if (a.sourcePriority !== b.sourcePriority) return a.sourcePriority - b.sourcePriority;
          return a.sourceOrder - b.sourceOrder;
        })
        .forEach(evento => {
          eventosUnicosPorFecha[fechaInput_(evento.fechaDesde)] = evento;
        });

      const eventosOrdenados = Object.keys(eventosUnicosPorFecha)
        .map(dateKey => eventosUnicosPorFecha[dateKey])
        .sort((a, b) => a.fechaDesde - b.fechaDesde);

      eventosOrdenados.forEach((evento, index) => {
        const siguiente = eventosOrdenados[index + 1];
        const identity = construirIdentidadEventoCostoHistoricoQTAS_(evento);
        const costoId = idsExistentes[identity] || nextCostoId;

        reconstruidas.push({
          Costo_ID: costoId,
          Compra_ID: normalizarClaveTexto_(evento.fuenteTipo) === normalizarClaveTexto_('Compra')
            ? numero_(evento.compraId)
            : 0,
          Item: evento.item,
          Tipo_Item: evento.tipoItem,
          Unidad: evento.unidad,
          Costo_Unitario: redondear_(numero_(evento.costoUnitario)),
          Proveedor: evento.proveedor,
          Fecha_Desde: evento.fechaDesde,
          Fecha_Hasta: siguiente ? diaAnterior_(siguiente.fechaDesde) : '',
          Activo: true,
          Fuente_Tipo: evento.fuenteTipo,
          Fuente_ID: evento.fuenteId,
          Nota: evento.nota
        });

        if (!idsExistentes[identity]) {
          nextCostoId = siguienteIdConPrefijoDesdeValorQTAS_(nextCostoId, 'COST-', 4);
        }
      });
    });

  sobrescribirObjetosHojaQTAS_(
    costosSheet,
    headers,
    reconstruidas.sort(compararFilasCostoHistoricoQTAS_)
  );
  limpiarCachesEjecucionQTAS_();

  return {
    ok: true,
    eventosFuente: eventos.length,
    rows: reconstruidas.length
  };
}

function normalizarListaIdsObjetivoQTAS_(value) {
  const raw = Array.isArray(value) ? value : [value];
  const seen = {};

  return raw
    .map(numero_)
    .filter(id => id > 0)
    .filter(id => {
      const key = String(id);
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

function compararCompraOrigenesFondosQTAS_(a, b) {
  if (numero_(a.Compra_ID) !== numero_(b.Compra_ID)) {
    return numero_(a.Compra_ID) - numero_(b.Compra_ID);
  }
  if (texto_(a.Compra_Detalle_ID) !== texto_(b.Compra_Detalle_ID)) {
    return texto_(a.Compra_Detalle_ID).localeCompare(texto_(b.Compra_Detalle_ID));
  }
  return ordenAportanteOrigenFondosQTAS_(a.Aportante) - ordenAportanteOrigenFondosQTAS_(b.Aportante);
}

function claveCostoHistoricoQTAS_(tipoItem, item, unidad) {
  const tipo = normalizarTipoCompraItemQTAS_(tipoItem);
  const nombre = texto_(item);
  const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);

  if (!tipo || !nombre || !unidadCanonica) return '';
  return [
    normalizarClaveTexto_(tipo),
    normalizarClaveTexto_(nombre),
    normalizarClaveTexto_(unidadCanonica)
  ].join('|');
}

function construirIdentidadEventoCostoHistoricoQTAS_(evento) {
  return [
    texto_(evento && evento.key),
    normalizarClaveTexto_(evento && evento.fuenteTipo),
    texto_(evento && evento.fuenteId),
    fechaTextoPlanoQTAS_(evento && evento.fechaDesde)
  ].join('|');
}

function prioridadFuenteCostoHistoricoQTAS_(fuenteTipo) {
  const key = normalizarClaveTexto_(fuenteTipo);
  if (key === normalizarClaveTexto_('Directo')) return 10;
  if (key === normalizarClaveTexto_('Compra')) return 20;
  if (key === normalizarClaveTexto_('Manual')) return 30;
  return 25;
}

function compararFilasCostoHistoricoQTAS_(a, b) {
  const keyA = claveCostoHistoricoQTAS_(a.Tipo_Item, a.Item, a.Unidad);
  const keyB = claveCostoHistoricoQTAS_(b.Tipo_Item, b.Item, b.Unidad);
  if (keyA !== keyB) return keyA.localeCompare(keyB);

  const fechaA = fechaTextoPlanoQTAS_(a.Fecha_Desde);
  const fechaB = fechaTextoPlanoQTAS_(b.Fecha_Desde);
  if (fechaA !== fechaB) return fechaA.localeCompare(fechaB);

  return texto_(a.Costo_ID).localeCompare(texto_(b.Costo_ID));
}
