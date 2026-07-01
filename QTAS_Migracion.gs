function migrarLegacyComprasQTAS(payload) {
  const settings = Object.assign({
    dryRun: true,
    incluirReglasOrigenes: true,
    incluirOrigenesFondos: true,
    permitirCoincidenciasHeuristicas: false
  }, payload || {});
  asegurarModeloOperativoQTAS_();

  const ss = SpreadsheetApp.getActive();
  const legacy = leerFuentesLegacyComprasQTAS_(ss);
  const productosActivosIndex = construirIndiceProductosCompraQTAS_(leerProductosActivosCompraQTAS_(ss));
  const plan = construirPlanMigracionLegacyComprasQTAS_(legacy, productosActivosIndex);
  const report = construirReporteMigracionLegacyComprasQTAS_(ss, plan, settings);

  if (settings.dryRun) {
    return report;
  }

  const comprasSheet = ss.getSheetByName(QTAS.sheets.compras);
  const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
  const costosSheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
  const reglasSheet = ss.getSheetByName(QTAS.sheets.origenesFondosReglas);
  const fondosSheet = ss.getSheetByName(QTAS.sheets.compraOrigenesFondos);
  const compraHeaders = getHeaders_(comprasSheet);
  const detalleHeaders = getHeaders_(detalleSheet);
  const costoHeaders = getHeaders_(costosSheet);
  const reglasHeaders = getHeaders_(reglasSheet);
  const fondosHeaders = getHeaders_(fondosSheet);
  const comprasExistentes = leerObjetosConMeta_(comprasSheet);
  const detallesExistentes = leerObjetosConMeta_(detalleSheet);
  const costosExistentes = leerObjetosConMeta_(costosSheet);
  const reglasExistentes = leerObjetosConMeta_(reglasSheet);
  const fondosExistentes = leerObjetosConMeta_(fondosSheet);
  let nextCompraId = siguienteIdNumerico_(comprasSheet, 'Compra_ID');
  const comprasAppend = [];
  const comprasStartRow = comprasSheet.getLastRow() + 1;

  plan.compras.forEach(compra => {
    const match = encontrarCompraCanonicaMigradaQTAS_(comprasExistentes, compra, settings);
    if (match) {
      compra.compraIdCanonico = numero_(match.Compra_ID);
      compra.__rowNumber = match.__rowNumber;
      const comentarioCompra = unirUnicos_([match.Comentario_Compra, compra.compra.Comentario_Compra]);
      if (texto_(comentarioCompra) !== texto_(match.Comentario_Compra)) {
        const actualizada = Object.assign({}, match, {
          Comentario_Compra: comentarioCompra
        });
        actualizarFilaObjeto_(comprasSheet, match.__rowNumber, compraHeaders, actualizada);
        Object.assign(match, actualizada);
      }
      return;
    }

    compra.compraIdCanonico = nextCompraId++;
    const row = Object.assign({}, compra.compra, { Compra_ID: compra.compraIdCanonico });
    compra.__rowNumber = comprasStartRow + comprasAppend.length;
    comprasAppend.push(row);
    comprasExistentes.push(Object.assign({ __rowNumber: compra.__rowNumber }, row));
  });
  if (comprasAppend.length) {
    escribirFilasDesdeFilaQTAS_(
      comprasSheet,
      comprasStartRow,
      comprasAppend.map(row => filaDesdeHeaders_(compraHeaders, row))
    );
  }

  const compraIdByLegacy = plan.compras.reduce((acc, compra) => {
    acc[compra.legacyCompraId] = numero_(compra.compraIdCanonico);
    return acc;
  }, {});
  plan.detalles.forEach(item => {
    item.compraIdCanonico = compraIdByLegacy[item.legacyCompraId] || 0;
  });
  const detalleAppend = [];
  const detalleStartRow = detalleSheet.getLastRow() + 1;

  plan.detalles
    .slice()
    .sort((a, b) => {
      const fechaA = valorFechaCompraCanonicaQTAS_(a.detalle, new Date());
      const fechaB = valorFechaCompraCanonicaQTAS_(b.detalle, new Date());
      if (fechaA.getTime() !== fechaB.getTime()) return fechaA - fechaB;
      if (a.compraIdCanonico !== b.compraIdCanonico) return a.compraIdCanonico - b.compraIdCanonico;
      return a.lineIndex - b.lineIndex;
    })
    .forEach(item => {
      item.detalle.Compra_ID = item.compraIdCanonico;
      item.detalle.Compra_Detalle_ID = compraDetalleIdQTAS_(item.compraIdCanonico, item.lineIndex);
      const match = encontrarDetalleCanonicoMigradoQTAS_(detallesExistentes, item, settings);
      if (match) {
        const comentarioLinea = unirUnicos_([match.Comentario_Linea, item.detalle.Comentario_Linea]);
        if (texto_(comentarioLinea) !== texto_(match.Comentario_Linea)) {
          const actualizado = Object.assign({}, match, {
            Comentario_Linea: comentarioLinea
          });
          actualizarFilaObjeto_(detalleSheet, match.__rowNumber, detalleHeaders, actualizado);
          Object.assign(match, actualizado);
        }
        item.detalleCanonicoId = texto_(match.Compra_Detalle_ID);
        item.__rowNumber = match.__rowNumber;
        return;
      }

      item.detalleCanonicoId = texto_(item.detalle.Compra_Detalle_ID);
      item.__rowNumber = detalleStartRow + detalleAppend.length;
      detalleAppend.push(Object.assign({}, item.detalle));
      detallesExistentes.push(Object.assign({ __rowNumber: item.__rowNumber }, item.detalle));
    });
  if (detalleAppend.length) {
    escribirFilasDesdeFilaQTAS_(
      detalleSheet,
      detalleStartRow,
      detalleAppend.map(row => filaDesdeHeaders_(detalleHeaders, row))
    );
  }

  const costosRowsMutable = costosExistentes.slice();
  const costosAppend = [];
  const costosUpdates = [];
  const costoIdState = {
    nextNumber: siguienteNumeroPrefijoEnRowsQTAS_(costosRowsMutable, 'Costo_ID', 'COST-')
  };
  plan.detalles
    .filter(item => Boolean(item.detalleCanonicoId))
    .filter(item => item.impactaCosto)
    .forEach(item => {
      const fuenteId = texto_(item.detalleCanonicoId);
      const existente = costosRowsMutable.find(row =>
        normalizarClaveTexto_(row.Fuente_Tipo) === normalizarClaveTexto_('Compra') &&
        texto_(row.Fuente_ID) === fuenteId
      );
      if (existente) return;

      upsertCostoReferenciaEnMemoriaMigracionQTAS_({
        rows: costosRowsMutable,
        fechaCompra: item.detalle.Fecha_Compra,
        proveedor: item.detalle.Proveedor,
        compraId: item.compraIdCanonico,
        linea: item.detalle,
        costoIdState: costoIdState,
        appendRows: costosAppend,
        updates: costosUpdates
      });
    });
  const costosUpdatesDedupe = deduplicarActualizacionesFilaQTAS_(costosUpdates);
  costosUpdatesDedupe.forEach(update => {
    actualizarFilaObjeto_(costosSheet, update.rowNumber, costoHeaders, update.row);
  });
  if (costosAppend.length) {
    escribirFilasDesdeFilaQTAS_(
      costosSheet,
      costosSheet.getLastRow() + 1,
      costosAppend.map(row => filaDesdeHeaders_(costoHeaders, row))
    );
  }

  if (settings.incluirReglasOrigenes) {
    const reglasAppend = [];
    const reglasStartRow = reglasSheet.getLastRow() + 1;
    plan.reglasOrigenes.forEach(regla => {
      const existing = reglasExistentes.find(row =>
        texto_(row.Regla_ID) === texto_(regla.Regla_ID) &&
        normalizarClaveTexto_(row.Origen_Fondos) === normalizarClaveTexto_(regla.Origen_Fondos) &&
        normalizarClaveTexto_(row.Aportante) === normalizarClaveTexto_(regla.Aportante)
      );
      if (existing) return;
      const rowNumber = reglasStartRow + reglasAppend.length;
      reglasAppend.push(Object.assign({}, regla));
      reglasExistentes.push(Object.assign({ __rowNumber: rowNumber }, regla));
    });
    if (reglasAppend.length) {
      escribirFilasDesdeFilaQTAS_(
        reglasSheet,
        reglasStartRow,
        reglasAppend.map(row => filaDesdeHeaders_(reglasHeaders, row))
      );
    }
  }

  if (settings.incluirOrigenesFondos) {
    const detalleIdByLegacy = plan.detalles.reduce((acc, item) => {
      acc[`${item.legacyCompraId}|${item.legacyLineaId}`] = texto_(item.detalleCanonicoId);
      return acc;
    }, {});
    const fondosAppend = [];
    const fondosStartRow = fondosSheet.getLastRow() + 1;
    plan.fondos.forEach(row => {
      row.Compra_ID = compraIdByLegacy[row.legacyCompraId] || 0;
      row.Compra_Detalle_ID = detalleIdByLegacy[`${row.legacyCompraId}|${row.legacyLineaId}`] || '';
      const existing = fondosExistentes.find(item =>
        texto_(item.Compra_Origen_ID) === texto_(row.Compra_Origen_ID)
      );
      if (existing) return;
      const rowNumber = fondosStartRow + fondosAppend.length;
      fondosAppend.push(Object.assign({}, row));
      fondosExistentes.push(Object.assign({ __rowNumber: rowNumber }, row));
    });
    if (fondosAppend.length) {
      escribirFilasDesdeFilaQTAS_(
        fondosSheet,
        fondosStartRow,
        fondosAppend.map(row => filaDesdeHeaders_(fondosHeaders, row))
      );
    }
  }

  limpiarCachesEjecucionQTAS_();
  return Object.assign({}, report, {
    dryRun: false,
    migrated: {
      compras: plan.compras.length,
      detalle: plan.detalles.length,
      reglasOrigenes: settings.incluirReglasOrigenes ? plan.reglasOrigenes.length : 0,
      fondos: settings.incluirOrigenesFondos ? plan.fondos.length : 0
    }
  });
}

function construirPlanMigracionLegacyComprasQTAS_(legacy, productosActivosIndex) {
  const productosIndex = productosActivosIndex || {};
  const resumenPorId = {};
  const detallePorCompra = {};
  const costosSugeridosPorFuente = {};
  const reglasOrigenes = normalizarReglasOrigenesLegacyQTAS_(legacy.reglasOrigenes);
  const aportesLegacyIndex = construirIndiceAportesLegacyQTAS_(legacy.gastosAportesDesglosado);

  (legacy.comprasResumen || []).forEach(row => {
    const legacyCompraId = texto_(row.compra_id || row.Compra_ID || row.compraId);
    if (!legacyCompraId) return;
    resumenPorId[legacyCompraId] = row;
  });

  (legacy.costosReferenciaSugeridos || []).forEach(row => {
    const key = construirClaveCostoSugeridoLegacyQTAS_(row);
    if (!key) return;
    costosSugeridosPorFuente[key] = row;
  });

  (legacy.gastosComprasLimpio || []).forEach(row => {
    const legacyCompraId = texto_(row.compra_id || row.Compra_ID || row.compraId);
    if (!legacyCompraId) return;
    if (!detallePorCompra[legacyCompraId]) detallePorCompra[legacyCompraId] = [];
    detallePorCompra[legacyCompraId].push(row);
  });

  const compraIds = [...new Set(
    Object.keys(resumenPorId).concat(Object.keys(detallePorCompra))
  )]
    .filter(Boolean)
    .sort((a, b) => {
      const lineasA = detallePorCompra[a] || [];
      const lineasB = detallePorCompra[b] || [];
      const fechaA = resolverFechaLegacyQTAS_(
        (resumenPorId[a] && (resumenPorId[a].periodo || resumenPorId[a].Periodo)) ||
        (lineasA[0] && (lineasA[0].periodo || lineasA[0].Periodo)),
        new Date(0)
      );
      const fechaB = resolverFechaLegacyQTAS_(
        (resumenPorId[b] && (resumenPorId[b].periodo || resumenPorId[b].Periodo)) ||
        (lineasB[0] && (lineasB[0].periodo || lineasB[0].Periodo)),
        new Date(0)
      );
      if (fechaA.getTime() !== fechaB.getTime()) return fechaA - fechaB;
      return texto_(a).localeCompare(texto_(b), undefined, { numeric: true, sensitivity: 'base' });
    });
  const compras = [];
  const detalles = [];

  compraIds.forEach(legacyCompraId => {
    const resumen = resumenPorId[legacyCompraId] || {};
    const lineas = (detallePorCompra[legacyCompraId] || [])
      .slice()
      .sort(compararOrdenLineaLegacyQTAS_);
    const fechaCompra = resolverFechaLegacyQTAS_(
      resumen.periodo || resumen.Periodo || (lineas[0] && (lineas[0].periodo || lineas[0].Periodo)),
      new Date()
    );
    const proveedor = texto_(
      resumen.proveedor || resumen.Proveedor || (lineas[0] && (lineas[0].proveedor || lineas[0].Proveedor))
    );
    const totalCompraResumen = redondear_(numero_(resumen.valor_total_compra || resumen.Total_Compra));
    const totalCompraLineas = redondear_(sumar_(lineas.map(row => row.valor_linea || row.Valor_Linea)));
    const totalCompra = totalCompraResumen > 0 ? totalCompraResumen : totalCompraLineas;
    const itemsResumen = lineas.length
      ? lineas.map(row => {
        const item = texto_(row.item_stock || row.Item_Stock || row.concepto_compra || row.Concepto_Compra);
        const medida = normalizarCantidadUnidadQTAS_(
          numero_(row.cantidad || row.Cantidad) || 1,
          texto_(row.unidad || row.Unidad) || 'und'
        );
        return `${item} ${formatearCantidad_(medida.cantidad || 1)}${medida.unidad || 'und'}`;
      }).join(' + ')
      : texto_(resumen.concepto || resumen.Concepto || legacyCompraId);
    const notaCompra = unirUnicos_([
      `Migrado desde ${legacyCompraId}`,
      texto_(resumen.tipo || resumen.Tipo),
      texto_(resumen.origenes_fondos || resumen.Origenes_Fondos)
        ? `Origenes ${texto_(resumen.origenes_fondos || resumen.Origenes_Fondos)}`
        : '',
      texto_(resumen.notas || resumen.Notas)
    ]);

    compras.push({
      legacyCompraId: legacyCompraId,
      compra: {
        Compra_ID: 0,
        Fecha_Compra: combinarFechaYHora_(fechaCompra, fechaCompra),
        Proveedor: proveedor || 'Proveedor legacy',
        Items_Resumen: itemsResumen,
        Total_Compra: totalCompra,
        Medio_Pago: 'Otro',
        Comentario_Compra: notaCompra,
        Estado_Registro: normalizarEstadoRegistroLegacyQTAS_(resumen.finalizado || resumen.Finalizado)
      }
    });

    lineas.forEach((linea, index) => {
      const lineaItem = normalizarLegacyLineaIdQTAS_(linea.linea_item || linea.Linea_Item || index + 1);
      const costoSugerido = costosSugeridosPorFuente[`${legacyCompraId}|${lineaItem}`] || null;
      const referenciaProducto = resolverProductoCanonicoLegacyQTAS_(
        productosIndex,
        linea.item_stock || linea.Item_Stock || linea.concepto_compra || linea.Concepto_Compra
      );
      const unidadBase = texto_(linea.unidad || linea.Unidad) || (referenciaProducto ? referenciaProducto.unidad : 'und');
      const medida = normalizarCantidadUnidadQTAS_(numero_(linea.cantidad || linea.Cantidad) || 1, unidadBase);
      const cantidad = medida.cantidad || 1;
      const unidad = medida.unidad || normalizarUnidadCanonicaQTAS_(unidadBase) || 'und';
      const valorLinea = redondear_(numero_(linea.valor_linea || linea.Valor_Linea) || 0);
      const costoUnitarioBase = redondear_(numero_(
        costoSugerido
          ? costoSugerido.costo_unitario || costoSugerido.Costo_Unitario
          : linea.costo_unitario || linea.Costo_Unitario
      ));
      const costoUnitario = costoUnitarioBase > 0
        ? costoUnitarioBase
        : (cantidad > 0 ? redondear_(valorLinea / cantidad) : 0);
      const impactaCosto = esValorAfirmativoLegacyQTAS_(
        linea.afecta_costo_producto || linea.Afecta_Costo_Producto
      );
      const tipoItem = clasificarTipoItemLegacyComprasQTAS_(linea, referenciaProducto, impactaCosto);
      const item = referenciaProducto
        ? referenciaProducto.item
        : texto_(
          (costoSugerido && (costoSugerido.item_stock || costoSugerido.Item_Stock)) ||
          linea.item_stock || linea.Item_Stock ||
          linea.concepto_compra || linea.Concepto_Compra ||
          `Linea ${index + 1}`
        );
      const notaLinea = unirUnicos_([
        `Migrado desde ${legacyCompraId} linea ${lineaItem}`,
        texto_(linea.categoria || linea.Categoria),
        texto_(linea.tipo || linea.Tipo),
        texto_(linea.origenes_fondos || linea.Origenes_Fondos)
          ? `Origenes ${texto_(linea.origenes_fondos || linea.Origenes_Fondos)}`
          : '',
        texto_(linea.notas || linea.Notas)
      ]);

      detalles.push({
        legacyCompraId: legacyCompraId,
        legacyLineaId: lineaItem,
        lineIndex: index + 1,
        impactaCosto: impactaCosto,
        origenesFondosRaw: texto_(linea.origenes_fondos || linea.Origenes_Fondos),
        detalle: {
          Compra_Detalle_ID: '',
          Compra_ID: 0,
          Fecha_Compra: combinarFechaYHora_(fechaCompra, fechaCompra),
          Proveedor: proveedor || 'Proveedor legacy',
          Tipo_Item: tipoItem,
          Item: item,
          Cantidad: cantidad,
          Unidad: unidad,
          Costo_Total_Linea: valorLinea,
          Costo_Unitario: costoUnitario,
          Impacta_Costo: impactaCosto,
          Comentario_Linea: notaLinea,
          Estado_Registro: normalizarEstadoRegistroLegacyQTAS_(linea.finalizado || linea.Finalizado)
        }
      });
    });
  });

  const fondos = construirFondosComprasLegacyQTAS_(compras, detalles, reglasOrigenes, aportesLegacyIndex);
  return {
    compras: compras,
    detalles: detalles,
    reglasOrigenes: reglasOrigenes,
    fondos: fondos,
    warnings: construirWarningsMigracionLegacyQTAS_(legacy, {
      aportesLegacyIndex: aportesLegacyIndex
    })
  };
}

function construirReporteMigracionLegacyComprasQTAS_(ss, plan, settings) {
  const comprasSheet = ss.getSheetByName(QTAS.sheets.compras);
  const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
  const costosSheet = ss.getSheetByName(QTAS.sheets.costosReferencia);
  const comprasExistentes = leerObjetosConMeta_(comprasSheet);
  const detallesExistentes = leerObjetosConMeta_(detalleSheet);
  const costosExistentes = leerObjetosConMeta_(costosSheet);
  const warnings = (plan.warnings || []).slice();
  const comprasNuevas = plan.compras.filter(item =>
    !encontrarCompraCanonicaMigradaQTAS_(comprasExistentes, item, settings)
  ).length;
  const detallesNuevos = plan.detalles.filter(item =>
    !encontrarDetalleCanonicoMigradoQTAS_(detallesExistentes, item, settings)
  ).length;
  const costosNuevos = contarCostosReferenciaNuevosMigracionQTAS_(plan.detalles, costosExistentes);
  const comprasCanonicasSinMarca = comprasExistentes.filter(row =>
    !leerLegacyIdDesdeComentarioQTAS_(row.Comentario_Compra)
  ).length;
  const detallesCanonicosSinMarca = detallesExistentes.filter(row =>
    !leerLegacyLineaIdDesdeComentarioQTAS_(row.Comentario_Linea)
  ).length;
  const costosCanonicosSinFuenteCompra = costosExistentes.filter(row =>
    normalizarClaveTexto_(row.Fuente_Tipo) !== normalizarClaveTexto_('Compra') || !texto_(row.Fuente_ID)
  ).length;
  const comprasCanonicasFusionadas = comprasExistentes.filter(row =>
    contarOcurrenciasTextoQTAS_(row.Comentario_Compra, 'Migrado desde ') > 1
  ).length;
  const detallesCanonicosFusionados = detallesExistentes.filter(row =>
    contarOcurrenciasTextoQTAS_(row.Comentario_Linea, 'Migrado desde ') > 1
  ).length;

  if (comprasCanonicasSinMarca || detallesCanonicosSinMarca || costosCanonicosSinFuenteCompra) {
    warnings.push(
      'Las hojas canonicas ya tienen datos previos sin marca de migracion ' +
      `(Compras=${comprasCanonicasSinMarca}, Compra_Detalle=${detallesCanonicosSinMarca}, ` +
      `Costos_Referencia=${costosCanonicosSinFuenteCompra}). ` +
      'Revisa antes de mezclar migracion con pruebas o capturas manuales.'
    );
  }
  if (comprasCanonicasFusionadas || detallesCanonicosFusionados) {
    warnings.push(
      'La hoja canonica actual parece venir de una migracion previa con filas fusionadas ' +
      `(Compras=${comprasCanonicasFusionadas}, Compra_Detalle=${detallesCanonicosFusionados}). ` +
      'Para validar esta version, usa una copia limpia y vuelve a migrar.'
    );
  }

  return {
    ok: true,
    dryRun: Boolean(settings.dryRun),
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    legacy: {
      comprasResumen: (plan.compras || []).length,
      gastoLineas: (plan.detalles || []).length,
      reglasOrigenes: (plan.reglasOrigenes || []).length,
      fondosDetalle: (plan.fondos || []).length
    },
    toImport: {
      compras: comprasNuevas,
      compraDetalle: detallesNuevos,
      costosReferencia: costosNuevos,
      reglasOrigenes: settings.incluirReglasOrigenes ? (plan.reglasOrigenes || []).length : 0,
      compraOrigenesFondos: settings.incluirOrigenesFondos ? (plan.fondos || []).length : 0
    },
    warnings: warnings
  };
}

function contarCostosReferenciaNuevosMigracionQTAS_(detallesPlan, costosExistentes) {
  const rows = (costosExistentes || [])
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      tipoItemKey: normalizarClaveTexto_(row.Tipo_Item),
      itemKey: normalizarClaveTexto_(row.Item),
      unidadKey: normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad)),
      desde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
      hasta: row.Fecha_Hasta
        ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date())
        : null
    }));
  let nuevos = 0;

  (detallesPlan || [])
    .filter(item => item && item.impactaCosto && item.detalle)
    .forEach(item => {
      const fechaCompra = resolverFechaOperacion_(item.detalle.Fecha_Compra, new Date());
      const tipoItemKey = normalizarClaveTexto_(item.detalle.Tipo_Item);
      const itemKey = normalizarClaveTexto_(item.detalle.Item);
      const unidadKey = normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(item.detalle.Unidad));
      const existentes = rows
        .filter(row =>
          row.tipoItemKey === tipoItemKey &&
          row.itemKey === itemKey &&
          row.unidadKey === unidadKey
        )
        .sort((a, b) => a.desde - b.desde);
      const exacta = existentes.find(row => row.desde.getTime() === fechaCompra.getTime());
      if (exacta) return;

      const insertIndex = existentes.findIndex(row => row.desde.getTime() > fechaCompra.getTime());
      const prev = insertIndex > 0
        ? existentes[insertIndex - 1]
        : (insertIndex === -1 ? existentes[existentes.length - 1] : null);
      const next = insertIndex >= 0 ? existentes[insertIndex] : null;
      const caeDentroPrev = prev && (!prev.hasta || fechaCompra.getTime() <= prev.hasta.getTime());
      let fechaHastaNueva = null;

      if (caeDentroPrev) {
        const fechaHastaAnterior = prev.hasta;
        if (fechaCompra.getTime() > prev.desde.getTime()) {
          prev.hasta = diaAnterior_(fechaCompra);
        }
        fechaHastaNueva = fechaHastaAnterior || null;
      } else if (next) {
        fechaHastaNueva = diaAnterior_(next.desde);
      }

      rows.push({
        tipoItemKey: tipoItemKey,
        itemKey: itemKey,
        unidadKey: unidadKey,
        desde: fechaCompra,
        hasta: fechaHastaNueva
      });
      nuevos += 1;
    });

  return nuevos;
}

function siguienteNumeroPrefijoEnRowsQTAS_(rows, fieldName, prefix) {
  return (rows || [])
    .map(row => {
      const value = texto_(row && row[fieldName]);
      const match = value.match(/(\d+)$/);
      return match ? Number(match[1]) : 0;
    })
    .reduce((max, value) => Math.max(max, value), 0) + 1;
}

function deduplicarActualizacionesFilaQTAS_(updates) {
  const map = {};
  (updates || []).forEach(item => {
    const rowNumber = numero_(item && item.rowNumber);
    if (rowNumber <= 0 || !item.row) return;
    map[rowNumber] = {
      rowNumber: rowNumber,
      row: item.row
    };
  });
  return Object.keys(map)
    .map(key => map[key])
    .sort((a, b) => a.rowNumber - b.rowNumber);
}

function upsertCostoReferenciaEnMemoriaMigracionQTAS_(context) {
  const rows = context.rows || [];
  const fechaCompra = resolverFechaOperacion_(context.fechaCompra, new Date());
  const tipoItemKey = normalizarClaveTexto_(context.linea.Tipo_Item);
  const itemKey = normalizarClaveTexto_(context.linea.Item);
  const unitKey = normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(context.linea.Unidad));
  const nota = unirUnicos_([
    texto_(context.linea.Comentario_Linea),
    `Compra ${context.compraId}`
  ]);
  const existentes = rows
    .filter(row =>
      normalizarClaveTexto_(row.Tipo_Item) === tipoItemKey &&
      normalizarClaveTexto_(row.Item) === itemKey &&
      normalizarClaveTexto_(normalizarUnidadCanonicaQTAS_(row.Unidad)) === unitKey &&
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
      Compra_ID: context.compraId,
      Item: texto_(context.linea.Item),
      Tipo_Item: texto_(context.linea.Tipo_Item),
      Unidad: texto_(context.linea.Unidad),
      Costo_Unitario: redondear_(numero_(context.linea.Costo_Unitario)),
      Proveedor: texto_(context.proveedor),
      Activo: true,
      Fuente_Tipo: 'Compra',
      Fuente_ID: texto_(context.linea.Compra_Detalle_ID),
      Nota: nota
    });
    Object.assign(exacta.raw, actualizada);
    if (numero_(exacta.rowNumber) > 0) {
      context.updates.push({
        rowNumber: exacta.rowNumber,
        row: actualizada
      });
    }
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
      prev.raw.Fecha_Hasta = diaAnterior_(fechaCompra);
      if (numero_(prev.rowNumber) > 0) {
        context.updates.push({
          rowNumber: prev.rowNumber,
          row: Object.assign({}, prev.raw)
        });
      }
    }
    fechaHastaNueva = fechaHastaAnterior || '';
  } else if (next) {
    fechaHastaNueva = diaAnterior_(next.desde);
  }

  const siguienteNumero = numero_(context.costoIdState && context.costoIdState.nextNumber) || 1;
  if (context.costoIdState) {
    context.costoIdState.nextNumber = siguienteNumero + 1;
  }
  const nueva = {
    Costo_ID: 'COST-' + String(siguienteNumero).padStart(4, '0'),
    Compra_ID: context.compraId,
    Item: texto_(context.linea.Item),
    Tipo_Item: texto_(context.linea.Tipo_Item),
    Unidad: texto_(context.linea.Unidad),
    Costo_Unitario: redondear_(numero_(context.linea.Costo_Unitario)),
    Proveedor: texto_(context.proveedor),
    Fecha_Desde: fechaCompra,
    Fecha_Hasta: fechaHastaNueva,
    Activo: true,
    Fuente_Tipo: 'Compra',
    Fuente_ID: texto_(context.linea.Compra_Detalle_ID),
    Nota: nota
  };
  rows.push(nueva);
  context.appendRows.push(nueva);
  return 1;
}

function leerFuentesLegacyComprasQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  return {
    comprasResumen: leerHojaLegacyOpcionalQTAS_(spreadsheet, 'compras_resumen'),
    gastosComprasLimpio: leerHojaLegacyOpcionalQTAS_(spreadsheet, 'gastos_compras_limpio'),
    gastosAportesDesglosado: leerHojaLegacyOpcionalQTAS_(spreadsheet, 'gastos_aportes_desglosado'),
    costosReferenciaSugeridos: leerHojaLegacyOpcionalQTAS_(spreadsheet, 'costos_referencia_sugeridos'),
    reglasOrigenes: leerHojaLegacyOpcionalQTAS_(spreadsheet, 'reglas_origen_fondos'),
    insumos: leerHojaLegacyOpcionalQTAS_(spreadsheet, 'Insumos')
  };
}

function leerHojaLegacyOpcionalQTAS_(ss, sheetName) {
  const sheet = (ss || SpreadsheetApp.getActive()).getSheetByName(sheetName);
  if (!sheet) return [];
  return leerObjetos_(sheet);
}

function normalizarReglasOrigenesLegacyQTAS_(rows) {
  return (rows || [])
    .map(row => ({
      Regla_ID: texto_(row.regla_id || row.Regla_ID),
      Origen_Fondos: texto_(row.origen_fondos || row.Origen_Fondos),
      Fecha_Desde: resolverFechaLegacyQTAS_(row.periodo_desde || row.Fecha_Desde, new Date()),
      Fecha_Hasta: valorLegacyFechaOpcionalQTAS_(row.periodo_hasta || row.Fecha_Hasta),
      Aportante: texto_(row.aportante || row.Aportante),
      Porcentaje: normalizarPorcentajeLegacyQTAS_(row.porcentaje || row.Porcentaje),
      Nota: texto_(row.nota || row.Nota)
    }))
    .filter(row => row.Regla_ID && row.Origen_Fondos && row.Aportante);
}

function construirFondosComprasLegacyQTAS_(compras, detalles, reglasOrigenes, aportesLegacyIndex) {
  const reglas = reglasOrigenes || [];
  const fondos = [];

  (detalles || []).forEach(item => {
    const aportesExactos = obtenerAportesLegacyPorLineaQTAS_(
      aportesLegacyIndex,
      item.legacyCompraId,
      item.legacyLineaId
    );
    if (aportesExactos.length) {
      const fondosExactos = construirFondosExactosDesdeAportesLegacyQTAS_(item, aportesExactos);
      if (fondosExactos.length) {
        fondos.push.apply(fondos, fondosExactos);
        return;
      }
    }

    const tokens = parsearOrigenesFondosLegacyQTAS_(item.origenesFondosRaw);
    const montoLinea = redondear_(numero_(item.detalle && item.detalle.Costo_Total_Linea));
    tokens.forEach((token, index) => {
      const montoOrigen = token.monto > 0
        ? redondear_(numero_(token.monto))
        : (tokens.length === 1 ? montoLinea : 0);
      const reglasActivas = reglas.filter(regla =>
        normalizarClaveTexto_(regla.Origen_Fondos) === normalizarClaveTexto_(token.origenFondos) &&
        fechaLegacyDentroDeRangoQTAS_(item.detalle.Fecha_Compra, regla.Fecha_Desde, regla.Fecha_Hasta)
      );
      const expansiones = reglasActivas.length
        ? reglasActivas.map(regla => ({
          aportante: texto_(regla.Aportante),
          porcentaje: redondear_(numero_(regla.Porcentaje)),
          reglaId: texto_(regla.Regla_ID)
        }))
        : [{
          aportante: token.origenFondos,
          porcentaje: 100,
          reglaId: ''
        }];

      expansiones.forEach(expansion => {
        fondos.push({
          Compra_Origen_ID: construirIdCompraOrigenQTAS_(
            item.legacyCompraId,
            item.legacyLineaId,
            token.origenFondos,
            expansion.aportante,
            index
          ),
          Compra_ID: 0,
          Compra_Detalle_ID: '',
          Fecha_Compra: item.detalle.Fecha_Compra,
          Origen_Fondos: token.origenFondos,
          Aportante: expansion.aportante,
          Porcentaje: expansion.porcentaje,
          Monto_Asignado: redondear_(montoOrigen * expansion.porcentaje / 100),
          Fuente_Registro: `Legacy:${item.legacyCompraId}#${item.legacyLineaId}`,
          Nota: unirUnicos_([
            expansion.reglaId ? `Regla ${expansion.reglaId}` : '',
            token.monto > 0 ? `Origen bruto ${token.origenFondos}:${token.monto}` : '',
            token.monto <= 0 && tokens.length === 1 && montoLinea > 0
              ? `Monto inferido desde costo total de linea ${montoLinea}`
              : ''
          ]),
          legacyCompraId: item.legacyCompraId,
          legacyLineaId: item.legacyLineaId
        });
      });
    });
  });
  return fondos;
}

function construirWarningsMigracionLegacyQTAS_(legacy, context) {
  const warnings = [];
  const aportesLegacyIndex = context && context.aportesLegacyIndex ? context.aportesLegacyIndex : {};
  if ((legacy.insumos || []).length) {
    warnings.push(
      `La hoja legacy "Insumos" existe con ${(legacy.insumos || []).length} fila(s), ` +
      'pero no se migra automaticamente a Producto_Componentes porque no trae cantidades de receta confiables.'
    );
  }
  if ((legacy.gastosAportesDesglosado || []).length) {
    const keys = Object.keys(aportesLegacyIndex);
    const multiOrigenKeys = keys.filter(key => {
      const rows = aportesLegacyIndex[key] || [];
      if (!rows.length) return false;
      return parsearOrigenesFondosLegacyQTAS_(rows[0].origenesFondosRaw).length > 1;
    });
    warnings.push(
      `La hoja legacy "gastos_aportes_desglosado" existe con ${(legacy.gastosAportesDesglosado || []).length} fila(s). ` +
      `Se usa como fuente exacta para lineas con un solo origen; ` +
      `${multiOrigenKeys.length} linea(s) con varios origenes siguen reconstruyendose desde texto bruto + reglas.`
    );
  }
  return warnings;
}

function construirIndiceAportesLegacyQTAS_(rows) {
  return (rows || []).reduce((acc, row) => {
    const legacyCompraId = texto_(row.compra_id || row.Compra_ID || row.compraId);
    const legacyLineaId = normalizarLegacyLineaIdQTAS_(row.linea_item || row.Linea_Item || row.lineaId);
    const aportante = texto_(row.aportante || row.Aportante);
    const montoAsignado = redondear_(numero_(row.valor_aporte_linea || row.Valor_Aporte_Linea));
    if (!legacyCompraId || !legacyLineaId || !aportante || montoAsignado <= 0) return acc;
    const key = construirLegacyLineaKeyQTAS_(legacyCompraId, legacyLineaId);
    if (!acc[key]) acc[key] = [];
    acc[key].push({
      legacyCompraId: legacyCompraId,
      legacyLineaId: legacyLineaId,
      aportante: aportante,
      montoAsignado: montoAsignado,
      origenesFondosRaw: texto_(row.origenes_fondos || row.Origenes_Fondos),
      nota: texto_(row.nota || row.Nota)
    });
    return acc;
  }, {});
}

function obtenerAportesLegacyPorLineaQTAS_(aportesLegacyIndex, legacyCompraId, legacyLineaId) {
  if (!aportesLegacyIndex) return [];
  const key = construirLegacyLineaKeyQTAS_(legacyCompraId, legacyLineaId);
  return (aportesLegacyIndex[key] || []).slice();
}

function construirFondosExactosDesdeAportesLegacyQTAS_(item, aportesExactos) {
  const tokens = parsearOrigenesFondosLegacyQTAS_(item.origenesFondosRaw);
  if (tokens.length !== 1) return [];

  const origen = tokens[0];
  const montoTotal = redondear_(sumar_(aportesExactos.map(row => row.montoAsignado)));
  if (montoTotal <= 0) return [];

  return aportesExactos
    .slice()
    .sort((a, b) => a.aportante.localeCompare(b.aportante))
    .map((aporte, index) => ({
      Compra_Origen_ID: construirIdCompraOrigenQTAS_(
        item.legacyCompraId,
        item.legacyLineaId,
        origen.origenFondos,
        aporte.aportante,
        index
      ),
      Compra_ID: 0,
      Compra_Detalle_ID: '',
      Fecha_Compra: item.detalle.Fecha_Compra,
      Origen_Fondos: origen.origenFondos,
      Aportante: aporte.aportante,
      Porcentaje: redondear_(aporte.montoAsignado * 100 / montoTotal),
      Monto_Asignado: aporte.montoAsignado,
      Fuente_Registro: `LegacyAportes:${item.legacyCompraId}#${item.legacyLineaId}`,
      Nota: unirUnicos_([
        'Migrado desde gastos_aportes_desglosado',
        origen.monto > 0 ? `Origen bruto ${origen.origenFondos}:${origen.monto}` : '',
        aporte.nota
      ]),
      legacyCompraId: item.legacyCompraId,
      legacyLineaId: item.legacyLineaId
    }));
}

function encontrarCompraCanonicaMigradaQTAS_(rows, compraPlan, options) {
  const settings = options || {};
  const legacyCompraId = texto_(compraPlan && compraPlan.legacyCompraId);
  const proveedor = normalizarClaveTexto_(compraPlan && compraPlan.compra && compraPlan.compra.Proveedor);
  const total = redondear_(numero_(compraPlan && compraPlan.compra && compraPlan.compra.Total_Compra));
  const items = normalizarClaveTexto_(compraPlan && compraPlan.compra && compraPlan.compra.Items_Resumen);
  const fecha = fechaInput_(compraPlan && compraPlan.compra && compraPlan.compra.Fecha_Compra);
  const bySource = (rows || []).find(row =>
    leerLegacyIdDesdeComentarioQTAS_(row.Comentario_Compra) === legacyCompraId
  );
  if (bySource) return bySource;

  if (settings.permitirCoincidenciasHeuristicas !== true) {
    return null;
  }

  return (rows || []).find(row =>
    !leerLegacyIdDesdeComentarioQTAS_(row.Comentario_Compra) &&
    normalizarClaveTexto_(row.Proveedor) === proveedor &&
    redondear_(numero_(row.Total_Compra)) === total &&
    normalizarClaveTexto_(row.Items_Resumen) === items &&
    fechaInput_(row.Fecha_Compra) === fecha
  ) || null;
}

function encontrarDetalleCanonicoMigradoQTAS_(rows, detallePlan, options) {
  const settings = options || {};
  const legacyKey = construirLegacyLineaKeyQTAS_(
    detallePlan && detallePlan.legacyCompraId,
    detallePlan && detallePlan.legacyLineaId
  );
  const bySource = (rows || []).find(row =>
    leerLegacyLineaIdDesdeComentarioQTAS_(row.Comentario_Linea) === legacyKey
  );
  if (bySource) return bySource;

  if (settings.permitirCoincidenciasHeuristicas !== true) {
    return null;
  }

  return (rows || []).find(row =>
    !leerLegacyLineaIdDesdeComentarioQTAS_(row.Comentario_Linea) &&
    numero_(row.Compra_ID) === numero_(detallePlan.compraIdCanonico) &&
    normalizarClaveTexto_(row.Item) === normalizarClaveTexto_(detallePlan.detalle.Item) &&
    normalizarClaveTexto_(row.Tipo_Item) === normalizarClaveTexto_(detallePlan.detalle.Tipo_Item) &&
    normalizarClaveTexto_(row.Unidad) === normalizarClaveTexto_(detallePlan.detalle.Unidad) &&
    redondear_(numero_(row.Cantidad)) === redondear_(numero_(detallePlan.detalle.Cantidad)) &&
    redondear_(numero_(row.Costo_Total_Linea)) === redondear_(numero_(detallePlan.detalle.Costo_Total_Linea)) &&
    fechaInput_(row.Fecha_Compra) === fechaInput_(detallePlan.detalle.Fecha_Compra)
  ) || null;
}

function resolverProductoCanonicoLegacyQTAS_(productosIndex, value) {
  const key = normalizarClaveTexto_(value);
  if (!key) return null;
  return productosIndex[key] || null;
}

function clasificarTipoItemLegacyComprasQTAS_(row, referenciaProducto, impactaCosto) {
  if (impactaCosto !== true) {
    return 'Gasto';
  }
  if (referenciaProducto) {
    return 'Producto';
  }
  return 'Insumo';
}

function construirClaveCostoSugeridoLegacyQTAS_(row) {
  const compraId = texto_(row.fuente_compra_id || row.Fuente_Compra_ID);
  const lineaItem = normalizarLegacyLineaIdQTAS_(row.fuente_linea_item || row.Fuente_Linea_Item);
  if (!compraId || !lineaItem) return '';
  return `${compraId}|${lineaItem}`;
}

function resolverFechaLegacyQTAS_(value, fallback) {
  if (value instanceof Date) {
    return resolverFechaOperacion_(value, fallback || new Date());
  }

  const numeric = numero_(value);
  if (numeric > 20000 && numeric < 80000) {
    const utcBase = Date.UTC(1899, 11, 30);
    return resolverFechaOperacion_(new Date(utcBase + Math.floor(numeric) * 86400000), fallback || new Date());
  }

  return resolverFechaOperacion_(value, fallback || new Date());
}

function valorLegacyFechaOpcionalQTAS_(value) {
  if (value === '' || value === null || value === undefined) return '';
  return resolverFechaLegacyQTAS_(value, new Date());
}

function esValorAfirmativoLegacyQTAS_(value) {
  const text = normalizarClaveTexto_(value);
  if (!text) return false;
  return ['si', 'sí', 's', 'true', '1', 'activo', 'finalizado', 'yes'].indexOf(text) >= 0;
}

function parsearOrigenesFondosLegacyQTAS_(value) {
  const text = texto_(value);
  if (!text) return [];

  const chunks = text
    .split(/\s*\|\s*|\s*;\s*/)
    .map(texto_)
    .filter(Boolean);

  return chunks
    .map(chunk => {
      const match = chunk.match(/^([^:]+):\s*(-?[\d.,]+)$/);
      if (!match) {
        return {
          origenFondos: chunk,
          monto: 0
        };
      }
      return {
        origenFondos: texto_(match[1]),
        monto: redondear_(numero_(match[2]))
      };
    })
    .filter(item => item.origenFondos);
}

function fechaLegacyDentroDeRangoQTAS_(fecha, desde, hasta) {
  const fechaBase = resolverFechaOperacion_(fecha, new Date());
  const fechaDesde = resolverFechaOperacion_(desde, fechaBase);
  const fechaHasta = hasta ? resolverFechaOperacion_(hasta, fechaBase) : null;
  return fechaBase >= fechaDesde && (!fechaHasta || fechaBase <= fechaHasta);
}

function normalizarLegacyLineaIdQTAS_(value) {
  const text = texto_(value);
  if (!text) return '';
  const numeric = numero_(text);
  if (numeric && Math.abs(numeric - Math.round(numeric)) < 0.000001) {
    return String(Math.round(numeric));
  }
  return text;
}

function compararOrdenLineaLegacyQTAS_(a, b) {
  const lineA = normalizarLegacyLineaIdQTAS_(a && (a.linea_item || a.Linea_Item));
  const lineB = normalizarLegacyLineaIdQTAS_(b && (b.linea_item || b.Linea_Item));
  const numA = numero_(lineA);
  const numB = numero_(lineB);
  const isNumA = lineA !== '' && !(numA === 0 && lineA !== '0');
  const isNumB = lineB !== '' && !(numB === 0 && lineB !== '0');
  if (isNumA && isNumB && numA !== numB) return numA - numB;
  return texto_(lineA).localeCompare(texto_(lineB), undefined, { numeric: true, sensitivity: 'base' });
}

function construirIdCompraOrigenQTAS_(legacyCompraId, legacyLineaId, origenFondos, aportante, index) {
  const base = [
    texto_(legacyCompraId),
    normalizarLegacyLineaIdQTAS_(legacyLineaId),
    texto_(origenFondos),
    texto_(aportante),
    String(index + 1)
  ]
    .join('-')
    .replace(/[^A-Za-z0-9_-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
  return `COF-${base}`.slice(0, 99);
}

function esValorNegativoLegacyQTAS_(value) {
  const text = normalizarClaveTexto_(value);
  if (!text) return false;
  return ['no', 'false', '0', 'anulado', 'cancelado'].indexOf(text) >= 0;
}

function normalizarEstadoRegistroLegacyQTAS_(value) {
  return esValorNegativoLegacyQTAS_(value)
    ? QTAS.status.registro.anulado
    : QTAS.status.registro.activo;
}

function normalizarPorcentajeLegacyQTAS_(value) {
  const parsed = redondear_(numero_(value));
  if (parsed <= 0) return 0;
  return parsed <= 1 ? redondear_(parsed * 100) : parsed;
}

function leerLegacyIdDesdeComentarioQTAS_(comentario) {
  const match = texto_(comentario).match(/Migrado desde ([^\s|]+)/i);
  return match ? texto_(match[1]) : '';
}

function construirLegacyLineaKeyQTAS_(legacyCompraId, legacyLineaId) {
  return `${texto_(legacyCompraId)}#${normalizarLegacyLineaIdQTAS_(legacyLineaId)}`;
}

function leerLegacyLineaIdDesdeComentarioQTAS_(comentario) {
  const match = texto_(comentario).match(/Migrado desde ([^\s|]+)\s+linea\s+([^\s|]+)/i);
  return match ? construirLegacyLineaKeyQTAS_(match[1], match[2]) : '';
}

function contarOcurrenciasTextoQTAS_(value, needle) {
  const text = texto_(value);
  const target = texto_(needle);
  if (!text || !target) return 0;
  return text.split(target).length - 1;
}
function migrarLegacyComprasQTAS_DryRun() {
  return migrarLegacyComprasQTAS({
    dryRun: true,
    incluirReglasOrigenes: true,
    incluirOrigenesFondos: true
  });
}

function migrarLegacyComprasQTAS_Ejecutar() {
  return migrarLegacyComprasQTAS({
    dryRun: false,
    incluirReglasOrigenes: true,
    incluirOrigenesFondos: true
  });
}

function migrarLegacyComprasQTAS_Ejecutar_Log() {
  const resultado = migrarLegacyComprasQTAS({
    dryRun: false,
    incluirReglasOrigenes: true,
    incluirOrigenesFondos: true
  });
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}

function migrarLegacyComprasQTAS_DryRun_Log() {
  const resultado = migrarLegacyComprasQTAS({
    dryRun: true,
    incluirReglasOrigenes: true,
    incluirOrigenesFondos: true
  });
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
