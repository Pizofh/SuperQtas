// Edicion controlada de operaciones recientes. Mantiene los IDs y conserva
// respaldo verificado antes de alterar datos comerciales.
function versionOperacionEdicionQTAS_(cabecera, detalle, pagos) {
  const plano = row => {
    const copy = Object.assign({}, row);
    delete copy.__rowNumber;
    return copy;
  };
  const contenido = JSON.stringify({
    cabecera: plano(cabecera),
    detalle: (detalle || []).map(plano),
    pagos: (pagos || []).map(plano)
  });
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, contenido)
  );
}

function reemplazarFilasOperacionEdicionQTAS_(sheet, existentes, nuevas) {
  const headers = getHeaders_(sheet);
  const compartir = Math.min(existentes.length, nuevas.length);
  for (let i = 0; i < compartir; i += 1) {
    actualizarFilaObjeto_(sheet, existentes[i].__rowNumber, headers, nuevas[i]);
  }
  if (nuevas.length > existentes.length) {
    escribirFilas_(sheet, nuevas.slice(existentes.length).map(row => filaDesdeHeaders_(headers, row)));
  }
  if (existentes.length > nuevas.length) {
    const plan = planificarEliminacionFilasQTAS_(sheet, headers, existentes.slice(nuevas.length));
    ejecutarEliminacionFilasQTAS_(plan);
  }
}

function obtenerCompraParaEdicionQTAS(payload) {
  const compraId = numero_(payload && payload.compraId);
  if (compraId <= 0) throw new Error('Falta el ID de la compra.');
  const ss = validarModeloSoloLecturaQTAS_({
    sheetNames: [QTAS.sheets.compras, QTAS.sheets.compraDetalle, QTAS.sheets.compraOrigenesFondos],
    validarConfig: false
  });
  if (!listarComprasRecientesQTAS_().some(row => row.compraId === compraId)) {
    throw new Error('Solo se permite editar compras recientes.');
  }
  const compra = leerObjetos_(ss.getSheetByName(QTAS.sheets.compras))
    .find(row => numero_(row.Compra_ID) === compraId && !esRegistroAnulado_(row.Estado_Registro));
  if (!compra) throw new Error('Compra inexistente o anulada.');
  const detalle = leerObjetos_(ss.getSheetByName(QTAS.sheets.compraDetalle))
    .filter(row => numero_(row.Compra_ID) === compraId && !esRegistroAnulado_(row.Estado_Registro))
    .sort((a, b) => texto_(a.Compra_Detalle_ID).localeCompare(texto_(b.Compra_Detalle_ID)));
  const origen = leerObjetos_(ss.getSheetByName(QTAS.sheets.compraOrigenesFondos))
    .find(row => numero_(row.Compra_ID) === compraId);
  return {
    ok: true, compraId: compraId,
    version: versionOperacionEdicionQTAS_(compra, detalle, []),
    fechaCompra: fechaInput_(compra.Fecha_Compra),
    proveedor: texto_(compra.Proveedor),
    medioPago: texto_(compra.Medio_Pago),
    comentarioCompra: texto_(compra.Comentario_Compra),
    origenFondos: texto_(origen && origen.Origen_Fondos),
    lineas: detalle.map(row => ({
      tipoItem: texto_(row.Tipo_Item), item: texto_(row.Item),
      cantidad: numero_(row.Cantidad), unidad: texto_(row.Unidad),
      costoTotalLinea: numero_(row.Costo_Total_Linea),
      impactaCosto: estaActivo_(row.Impacta_Costo),
      comentarioLinea: texto_(row.Comentario_Linea)
    }))
  };
}

function editarCompraQTAS(payload) {
  const compraId = numero_(payload && payload.compraId);
  const version = texto_(payload && payload.version);
  if (!version || compraId <= 0) throw new Error('Abre primero la compra desde Editar.');
  validarCompraQTAS_(payload);
  // Copia nativa verificada antes del primer cambio. Si falla, no se modifica nada.
  const backup = crearBackupEdicionQTAS_();
  if (!backup || !backup.ok || !backup.backupFileId) throw new Error('No se pudo verificar el respaldo de la edicion.');
  return withScriptLock_('editar compra reciente', () => {
    const ss = validarModeloSoloLecturaQTAS_({
      sheetNames: [QTAS.sheets.compras, QTAS.sheets.compraDetalle, QTAS.sheets.compraOrigenesFondos,
        QTAS.sheets.costosReferencia, QTAS.sheets.inventarioMovimientos, QTAS.sheets.inventarioSnapshot],
      validarConfig: false
    });
    if (!listarComprasRecientesQTAS_().some(row => row.compraId === compraId)) {
      throw new Error('La compra ya no esta entre las recientes.');
    }
    if (leerColaPostCompraQTAS_().some(row =>
        texto_(row.spreadsheetId) === ss.getId() && numero_(row.compraId) === compraId)) {
      throw new Error('Espera a que termine el postproceso de esta compra antes de editar.');
    }
    const comprasSheet = ss.getSheetByName(QTAS.sheets.compras);
    const detalleSheet = ss.getSheetByName(QTAS.sheets.compraDetalle);
    const origenSheet = ss.getSheetByName(QTAS.sheets.compraOrigenesFondos);
    const original = leerObjetosConMeta_(comprasSheet)
      .find(row => numero_(row.Compra_ID) === compraId && !esRegistroAnulado_(row.Estado_Registro));
    const anteriores = leerObjetosConMeta_(detalleSheet)
      .filter(row => numero_(row.Compra_ID) === compraId && !esRegistroAnulado_(row.Estado_Registro))
      .sort((a, b) => texto_(a.Compra_Detalle_ID).localeCompare(texto_(b.Compra_Detalle_ID)));
    if (!original || versionOperacionEdicionQTAS_(original, anteriores, []) !== version) {
      throw new Error('La compra cambio desde que abriste el editor. Recarga antes de continuar.');
    }
    const medioPago = normalizarMedioPagoQTAS_(payload.medioPago);
    if (!leerMediosPagoConfiguradosQTAS_().some(row =>
      row.activo && normalizarClaveTexto_(row.medioPago) === normalizarClaveTexto_(medioPago))) {
      throw new Error('Medio de pago inactivo.');
    }
    const ahora = new Date();
    const fechaBase = resolverFechaOperacion_(payload.fechaCompra, ahora);
    const fechaCompra = combinarFechaYHora_(fechaBase, original.Fecha_Compra || ahora);
    const productos = construirIndiceProductosCompraQTAS_(leerProductosActivosCompraQTAS_(ss));
    const catalogo = construirIndiceItemsSugeridosCompraQTAS_(
      construirItemsSugeridosComprasQTAS_(leerProductosActivosCompraQTAS_(ss),
        listarCostosVigentesQTAS_(), { ss: ss })
    );
    const nuevas = payload.lineas.map((linea, index) => prepararLineaCompraQTAS_({
      compraId: compraId, fechaCompra: fechaCompra, proveedor: texto_(payload.proveedor),
      linea: linea, index: index, productosActivosIndex: productos, itemsCatalogoIndex: catalogo
    }));
    const totalCompra = redondear_(sumar_(nuevas.map(row => row.Costo_Total_Linea)));
    const origenes = construirFilasCompraOrigenesFondosQTAS_({
      compraId: compraId, fechaCompra: fechaBase,
      origenFondos: normalizarOrigenFondosQTAS_(payload.origenFondos),
      lineas: nuevas, comentarioCompra: texto_(payload.comentarioCompra)
    });
    const prevOrigen = leerObjetosConMeta_(origenSheet).filter(row => numero_(row.Compra_ID) === compraId);
    const actualizada = Object.assign({}, original, {
      Fecha_Compra: fechaCompra, Proveedor: texto_(payload.proveedor),
      Items_Resumen: resumenItemsCompraQTAS_(nuevas),
      Total_Compra: totalCompra, Medio_Pago: medioPago,
      Comentario_Compra: texto_(payload.comentarioCompra)
    });
    // Se retiran solo los movimientos de ESTA compra, nunca ajustes fisicos.
    eliminarMovimientosInventarioPorFuentesQTAS_({ ss: ss, compraId: compraId });
    actualizarFilaObjeto_(comprasSheet, original.__rowNumber, getHeaders_(comprasSheet), actualizada);
    reemplazarFilasOperacionEdicionQTAS_(detalleSheet, anteriores, nuevas);
    reemplazarFilasOperacionEdicionQTAS_(origenSheet, prevOrigen, origenes.rows);
    reconstruirCostosReferenciaDesdeFuentesQTAS_({ ss: ss });
    const postProceso = encolarPostProcesoCompraQTAS_({
      spreadsheetId: ss.getId(), compraId: compraId, programar: true
    });
    limpiarCachesEjecucionQTAS_();
    return {
      ok: true, compraId: compraId, totalCompra: totalCompra, backupId: backup.backupFileId,
      postProceso: postProceso, compraReciente: resumenCompraRecienteQTAS_(actualizada,
        leerOrigenesFondosPorCompraQTAS_()[compraId])
    };
  });
}

function obtenerVentaParaEdicionQTAS(payload) {
  const ventaId = numero_(payload && payload.ventaId);
  if (ventaId <= 0) throw new Error('Falta el ID de la venta.');
  const ss = validarModeloSoloLecturaQTAS_({
    sheetNames: [QTAS.sheets.ventas, QTAS.sheets.detalle, QTAS.sheets.pagos],
    validarConfig: false
  });
  if (!listarVentasRecientesDirectasQTAS_(ss).some(row => row.ventaId === ventaId)) {
    throw new Error('Solo se permite editar ventas recientes.');
  }
  const venta = leerObjetos_(ss.getSheetByName(QTAS.sheets.ventas))
    .find(row => numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro));
  if (!venta) throw new Error('Venta inexistente o anulada.');
  const detalle = leerObjetos_(ss.getSheetByName(QTAS.sheets.detalle))
    .filter(row => numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro))
    .sort((a, b) => texto_(a.Detalle_ID).localeCompare(texto_(b.Detalle_ID)));
  const pagos = leerObjetos_(ss.getSheetByName(QTAS.sheets.pagos))
    .filter(row => numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro))
    .sort((a, b) => texto_(a.Pago_ID).localeCompare(texto_(b.Pago_ID)));
  return {
    ok: true, ventaId: ventaId, version: versionOperacionEdicionQTAS_(venta, detalle, pagos),
    fechaVenta: fechaInput_(venta.Fecha_Venta),
    cliente: { clienteId: texto_(venta.Cliente_ID), nombre: texto_(venta.Nombre) },
    comentarioVenta: texto_(venta.Comentario_Venta),
    lineas: detalle.map(row => ({
      producto: texto_(row.Producto_Estandar),
      cantidad: numero_(row.Cantidad), unidad: texto_(row.Unidad),
      precioVendidoUnitario: numero_(row.Precio_Vendido_Unitario),
      descuentoLinea: numero_(row.Descuento_Linea),
      comentarioLinea: texto_(row.Comentario_Linea)
    })),
    pagos: pagos.map(row => ({
      pagoId: texto_(row.Pago_ID),
      fechaPago: fechaInput_(row.Fecha_Pago),
      medio: texto_(row.Medio_Pago),
      monto: numero_(row.Monto_Pago),
      comentarioPago: texto_(row.Comentario_Pago)
    }))
  };
}

function editarVentaQTAS(payload) {
  const ventaId = numero_(payload && payload.ventaId);
  const version = texto_(payload && payload.version);
  if (!version || ventaId <= 0) throw new Error('Abre primero la venta desde Editar.');
  validarVenta_(payload);
  const backup = crearBackupEdicionQTAS_();
  if (!backup || !backup.ok || !backup.backupFileId) throw new Error('No se pudo verificar el respaldo de la edicion.');
  return withScriptLock_('editar venta reciente', () => {
    const ss = validarModeloSoloLecturaQTAS_({
      sheetNames: [QTAS.sheets.ventas, QTAS.sheets.detalle, QTAS.sheets.pagos,
        QTAS.sheets.distribucionIngresos, QTAS.sheets.inventarioMovimientos,
        QTAS.sheets.inventarioSnapshot], validarConfig: false
    });
    if (!listarVentasRecientesDirectasQTAS_(ss).some(row => row.ventaId === ventaId)) {
      throw new Error('La venta ya no esta entre las recientes.');
    }
    if (leerColaPostVentaQTAS_().some(row =>
        texto_(row.spreadsheetId) === ss.getId() && numero_(row.ventaId) === ventaId)) {
      throw new Error('Espera a que termine el postproceso de esta venta antes de editar.');
    }
    const ventasSheet = ss.getSheetByName(QTAS.sheets.ventas);
    const detalleSheet = ss.getSheetByName(QTAS.sheets.detalle);
    const pagosSheet = ss.getSheetByName(QTAS.sheets.pagos);
    const original = leerObjetosConMeta_(ventasSheet)
      .find(row => numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro));
    const anteriores = leerObjetosConMeta_(detalleSheet)
      .filter(row => numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro))
      .sort((a, b) => texto_(a.Detalle_ID).localeCompare(texto_(b.Detalle_ID)));
    const pagosAnteriores = leerObjetosConMeta_(pagosSheet)
      .filter(row => numero_(row.Venta_ID) === ventaId && !esRegistroAnulado_(row.Estado_Registro))
      .sort((a, b) => texto_(a.Pago_ID).localeCompare(texto_(b.Pago_ID)));
    if (!original || versionOperacionEdicionQTAS_(original, anteriores, pagosAnteriores) !== version) {
      throw new Error('La venta cambio desde que abriste el editor. Recarga antes de continuar.');
    }
    // Los pagos con su propia trazabilidad conservan IDs y fechas.
    if ((payload.pagos || []).length !== pagosAnteriores.length) {
      throw new Error('No agregues ni quites pagos en Editar; usa el modulo de pagos pendientes.');
    }
    const pagosPreparadosInput = (payload.pagos || []).map((pago, index) =>
      Object.assign({}, pago, { fechaPago: pagosAnteriores[index].Fecha_Pago })
    );
    const ahora = new Date();
    const fechaBase = resolverFechaOperacion_(payload.fechaVenta, ahora);
    const fechaVenta = combinarFechaYHora_(fechaBase, original.Fecha_Venta || ahora);
    const cliente = resolverClienteQTAS_(
      Object.assign({}, payload.cliente || {}, {
        confirmadoCatalogo: Boolean(payload.cliente && payload.cliente.clienteId)
      }),
      ahora, { ss: ss }
    );
    const priceCache = cargarPreciosReferenciaEnMemoria_();
    const distributionCache = cargarReglasDistribucionEnMemoria_();
    const regla = obtenerSnapshotDistribucionDesdeCache_(distributionCache, fechaVenta);
    const preparadas = payload.lineas.map((linea, index) => prepararLineaVentaQTAS_({
      linea: linea, index: index, ventaId: ventaId, fechaVenta: fechaVenta,
      clienteId: cliente.clienteId, clienteNombre: cliente.nombre,
      priceCache: priceCache, creadoEn: ahora
    }));
    const nuevos = preparadas.map(row => row.row);
    const pagados = prepararPagosVentaQTAS_({
      ventaId: ventaId, nombre: cliente.nombre, pagos: pagosPreparadosInput,
      fechaPagoDefault: fechaVenta, distributionCache: distributionCache, creadoEn: ahora
    });
    const pagosNuevos = pagados.rows.map((row, index) => Object.assign({}, row, {
      Pago_ID: texto_(pagosAnteriores[index].Pago_ID),
      Fecha_Pago: pagosAnteriores[index].Fecha_Pago
    }));
    const total = redondear_(sumar_(preparadas.map(row => row.subtotalNeto)));
    if (pagados.totalPagado > total + 0.009) throw new Error('Los pagos exceden el nuevo total.');
    const producciones = leerObjetos_(ss.getSheetByName(QTAS.sheets.producciones))
      .filter(row => texto_(row.Comentario_Produccion).indexOf('Produccion automatica por venta V' + ventaId) >= 0);
    const stockHaCambiado = anteriores.length !== nuevos.length || anteriores.some((row, index) =>
      texto_(row.Producto_Estandar) !== texto_(nuevos[index].Producto_Estandar) ||
      texto_(row.Unidad) !== texto_(nuevos[index].Unidad) ||
      Math.abs(numero_(row.Cantidad) - numero_(nuevos[index].Cantidad)) > 0.0001
    );
    if (producciones.length && stockHaCambiado) {
      throw new Error('Esta venta genero produccion automatica. No se puede cambiar productos o cantidades sin reconciliar producciones.');
    }
    const actualizada = Object.assign({}, original, {
      Fecha_Venta: fechaVenta, Cliente_ID: cliente.clienteId, Nombre: cliente.nombre,
      Comentario_Venta: texto_(payload.comentarioVenta),
      Productos_Resumen: resumenProductos_(payload.lineas),
      Total_Venta: total, Total_Pagado: pagados.totalPagado,
      Saldo: redondear_(Math.max(0, total - pagados.totalPagado)),
      Estado_Pago: obtenerEstadoPago_(total, pagados.totalPagado, QTAS.status.registro.activo),
      Regla_Distribucion_Venta_ID: regla.reglaId,
      Steve_Pct_Venta: regla.steve, Majo_Pct_Venta: regla.majo, Mush_Pct_Venta: regla.mush
    });
    eliminarMovimientosInventarioPorFuentesQTAS_({
      ss: ss, ventaId: ventaId,
      detalleIds: anteriores.map(row => texto_(row.Detalle_ID))
    });
    actualizarFilaObjeto_(ventasSheet, original.__rowNumber, getHeaders_(ventasSheet), actualizada);
    reemplazarFilasOperacionEdicionQTAS_(detalleSheet, anteriores, nuevos);
    reemplazarFilasOperacionEdicionQTAS_(pagosSheet, pagosAnteriores, pagosNuevos);
    const costoRef = resolverHojaCanonicaOperativaQTAS_(ss, QTAS.sheets.ventaDetalleCostosCalculado);
    if (costoRef.ok) {
      const costos = leerObjetosConMeta_(costoRef.sheet).filter(row => numero_(row.Venta_ID) === ventaId);
      if (costos.length) {
        ejecutarEliminacionFilasQTAS_(
          planificarEliminacionFilasQTAS_(costoRef.sheet, costoRef.headers, costos)
        );
      }
    }
    sincronizarDistribucionVentaQTAS_(ventaId);
    const postProceso = encolarPostProcesoVentaQTAS_({
      spreadsheetId: ss.getId(), ventaId: ventaId, programar: true
    });
    limpiarCachesEjecucionQTAS_();
    return {
      ok: true, ventaId: ventaId, clienteId: cliente.clienteId, nombre: cliente.nombre,
      totalVenta: total, totalPagado: pagados.totalPagado,
      saldo: actualizada.Saldo, backupId: backup.backupFileId,
      postProceso: postProceso
    };
  });
}
