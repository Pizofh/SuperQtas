function reconstruirDistribucionIngresosQTAS_(options) {
  assertOperacionDestructivaPermitidaQTAS_('reconstruirDistribucionIngresosQTAS_');
  const settings = options || {};
  const ss = SpreadsheetApp.getActive();
  const ventasSheet = ss.getSheetByName(QTAS.sheets.ventas);
  const pagosSheet = ss.getSheetByName(QTAS.sheets.pagos);
  const distribucionSheet = ss.getSheetByName(QTAS.sheets.distribucionIngresos);
  const ventaHeaders = getHeaders_(ventasSheet);
  const pagoHeaders = getHeaders_(pagosSheet);
  const distribucionHeaders = getHeaders_(distribucionSheet);
  const reglasCache = cargarReglasDistribucionEnMemoria_();
  const estadoVentas = construirEstadoVentasQTAS_();
  const estadoPorVenta = {};

  estadoVentas.ventasActualizadas.forEach(venta => {
    estadoPorVenta[texto_(venta.Venta_ID)] = venta;
  });

  const ventas = leerObjetosConMeta_(ventasSheet).map(row => {
    const merged = Object.assign({}, row, estadoPorVenta[texto_(row.Venta_ID)] || {});
    return completarSnapshotDistribucionVentaFilaQTAS_(merged, reglasCache);
  });

  ventas.forEach(venta => {
    if (venta.__needsSave) {
      actualizarFilaObjeto_(ventasSheet, venta.__rowNumber, ventaHeaders, venta);
    }
  });

  const ventasPorId = {};
  ventas.forEach(venta => {
    ventasPorId[texto_(venta.Venta_ID)] = venta;
  });

  const pagos = leerObjetosConMeta_(pagosSheet).map(row =>
    completarSnapshotDistribucionPagoFilaQTAS_(row, reglasCache)
  );

  pagos.forEach(pago => {
    if (pago.__needsSave) {
      actualizarFilaObjeto_(pagosSheet, pago.__rowNumber, pagoHeaders, pago);
    }
  });

  const filasDistribucion = [];

  ventas.forEach(venta => {
    filasDistribucion.push(construirFilaDistribucionVentaQTAS_(venta));
  });

  pagos.forEach(pago => {
    filasDistribucion.push(construirFilaDistribucionPagoQTAS_(pago, ventasPorId[texto_(pago.Venta_ID)]));
  });

  filasDistribucion.sort((a, b) => {
    const fechaA = fecha_(a.Fecha_Base);
    const fechaB = fecha_(b.Fecha_Base);
    if (fechaA.getTime() !== fechaB.getTime()) return fechaA - fechaB;
    return texto_(a.Distribucion_ID).localeCompare(texto_(b.Distribucion_ID));
  });

  reemplazarObjetos_(distribucionSheet, distribucionHeaders, filasDistribucion);

  if (!settings.silent) {
    maybeAlert_(
      `Distribucion reconstruida: ${filasDistribucion.length} fila(s), ${ventas.length} venta(s) y ${pagos.length} pago(s).`
    );
  }

  return {
    ok: true,
    filas: filasDistribucion.length,
    ventas: ventas.length,
    pagos: pagos.length
  };
}

function sincronizarDistribucionVentaQTAS_(ventaId) {
  const ventaIdNumero = numero_(ventaId);
  if (ventaIdNumero <= 0) {
    return { ok: false, filas: 0, pagos: 0 };
  }

  const ss = SpreadsheetApp.getActive();
  const ventasSheet = ss.getSheetByName(QTAS.sheets.ventas);
  const pagosSheet = ss.getSheetByName(QTAS.sheets.pagos);
  const distribucionSheet = ss.getSheetByName(QTAS.sheets.distribucionIngresos);
  const ventaHeaders = getHeaders_(ventasSheet);
  const pagoHeaders = getHeaders_(pagosSheet);
  const distribucionHeaders = getHeaders_(distribucionSheet);
  const reglasCache = cargarReglasDistribucionEnMemoria_();

  const ventaOriginal = leerObjetosConMeta_(ventasSheet)
    .find(row => numero_(row.Venta_ID) === ventaIdNumero);

  if (!ventaOriginal) {
    return { ok: false, filas: 0, pagos: 0 };
  }

  const venta = completarSnapshotDistribucionVentaFilaQTAS_(ventaOriginal, reglasCache);
  if (venta.__needsSave) {
    actualizarFilaObjeto_(ventasSheet, venta.__rowNumber, ventaHeaders, venta);
  }

  const pagos = leerObjetosConMeta_(pagosSheet)
    .filter(row => numero_(row.Venta_ID) === ventaIdNumero)
    .map(row => completarSnapshotDistribucionPagoFilaQTAS_(row, reglasCache));

  pagos.forEach(pago => {
    if (pago.__needsSave) {
      actualizarFilaObjeto_(pagosSheet, pago.__rowNumber, pagoHeaders, pago);
    }
  });

  const filasObjetivo = [construirFilaDistribucionVentaQTAS_(venta)]
    .concat(pagos.map(pago => construirFilaDistribucionPagoQTAS_(pago, venta)));

  const existentes = leerObjetosConMeta_(distribucionSheet)
    .filter(row => numero_(row.Venta_ID) === ventaIdNumero);
  const existentesPorId = {};
  const objetivosPorId = {};

  existentes.forEach(row => {
    existentesPorId[texto_(row.Distribucion_ID)] = row;
  });

  filasObjetivo.forEach(fila => {
    objetivosPorId[texto_(fila.Distribucion_ID)] = fila;
    const existente = existentesPorId[texto_(fila.Distribucion_ID)];

    if (existente) {
      actualizarFilaObjeto_(
        distribucionSheet,
        existente.__rowNumber,
        distribucionHeaders,
        Object.assign({}, existente, fila)
      );
      return;
    }

    escribirFilas_(distribucionSheet, [filaDesdeHeaders_(distribucionHeaders, fila)]);
  });

  existentes.forEach(row => {
    const distribucionId = texto_(row.Distribucion_ID);
    if (objetivosPorId[distribucionId]) return;

    const anulada = Object.assign({}, row, {
      Monto_Base: 0,
      Steve_Valor: 0,
      Majo_Valor: 0,
      Mush_Valor: 0,
      Estado_Registro: QTAS.status.registro.anulado
    });
    actualizarFilaObjeto_(distribucionSheet, row.__rowNumber, distribucionHeaders, anulada);
  });

  return {
    ok: true,
    filas: filasObjetivo.length,
    pagos: pagos.length
  };
}

function registrarDistribucionVentaNuevaQTAS_(venta, pagos, options) {
  const settings = options || {};
  const distribucionSheet = settings.distribucionSheet ||
    SpreadsheetApp.getActive().getSheetByName(QTAS.sheets.distribucionIngresos);
  const distribucionHeaders = settings.distribucionHeaders || getHeaders_(distribucionSheet);
  const startRow = numero_(settings.startRow) > 0
    ? numero_(settings.startRow)
    : distribucionSheet.getLastRow() + 1;
  const filas = [construirFilaDistribucionVentaQTAS_(venta)]
    .concat((pagos || []).map(pago => construirFilaDistribucionPagoQTAS_(pago, venta)));

  if (filas.length) {
    escribirFilasDesdeFilaQTAS_(
      distribucionSheet,
      startRow,
      filas.map(fila => filaDesdeHeaders_(distribucionHeaders, fila))
    );
  }

  return {
    ok: true,
    filas: filas.length,
    pagos: (pagos || []).length
  };
}

function leerReglasDistribucionQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.distribucionReglas);

  if (!sheet || sheet.getLastRow() < 2) {
    return construirReglasDistribucionBaseQTAS_();
  }

  if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.distribucionReglas])) {
    return construirReglasDistribucionBaseQTAS_();
  }

  const reglas = leerObjetos_(sheet)
    .map(row => ({
      reglaId: texto_(row.Regla_ID),
      desde: resolverFechaOperacion_(row.Fecha_Desde, new Date()),
      hasta: row.Fecha_Hasta
        ? resolverFechaOperacion_(row.Fecha_Hasta, row.Fecha_Desde || new Date())
        : null,
      steve: redondear_(numero_(row.Steve_Pct)),
      majo: redondear_(numero_(row.Majo_Pct)),
      mush: redondear_(numero_(row.Mush_Pct)),
      activo: estaActivo_(row.Activo),
      nota: texto_(row.Nota)
    }))
    .filter(regla => regla.reglaId && regla.activo);

  const salida = reglas.length ? reglas : construirReglasDistribucionBaseQTAS_();
  validarReglasDistribucionQTAS_(salida);
  return salida;
}

function construirReglasDistribucionBaseQTAS_() {
  return DISTRIBUCION_REGLAS_INICIALES.map((regla, index) => ({
    reglaId: 'DIST-' + String(index + 1).padStart(4, '0'),
    desde: resolverFechaOperacion_(regla.desde, new Date()),
    hasta: regla.hasta ? resolverFechaOperacion_(regla.hasta, regla.desde) : null,
    steve: redondear_(numero_(regla.steve)),
    majo: redondear_(numero_(regla.majo)),
    mush: redondear_(numero_(regla.mush)),
    activo: true,
    nota: texto_(regla.nota)
  }));
}

function cargarReglasDistribucionEnMemoria_() {
  const namespace = 'distribucion_reglas_memoria';

  return obtenerMemoEjecucionQTAS_(`cache:${namespace}`, () => {
    const cached = leerCacheDocumentoQTAS_(namespace);
    if (cached && Array.isArray(cached.rows)) {
      return cached.rows;
    }

    const rows = leerReglasDistribucionQTAS_()
      .filter(regla => regla.activo)
      .map(regla => ({
        reglaId: texto_(regla.reglaId),
        desde: fechaInput_(regla.desde),
        hasta: regla.hasta ? fechaInput_(regla.hasta) : '',
        steve: redondear_(numero_(regla.steve)),
        majo: redondear_(numero_(regla.majo)),
        mush: redondear_(numero_(regla.mush)),
        activo: regla.activo !== false,
        nota: texto_(regla.nota)
      }))
      .sort((a, b) => b.desde.localeCompare(a.desde));

    guardarCacheDocumentoQTAS_(namespace, { rows: rows }, 300);
    return rows;
  });
}

function validarReglasDistribucionQTAS_(reglas) {
  const activas = (reglas || [])
    .filter(regla => regla && regla.activo)
    .slice()
    .sort((a, b) => a.desde - b.desde);

  activas.forEach(regla => {
    const total = redondear_(numero_(regla.steve) + numero_(regla.majo) + numero_(regla.mush));
    if (Math.abs(total - 100) > 0.01) {
      throw new Error(`La regla ${regla.reglaId} debe sumar 100%.`);
    }
  });

  for (let index = 0; index < activas.length - 1; index += 1) {
    const actual = activas[index];
    const siguiente = activas[index + 1];

    if (!actual.hasta) {
      throw new Error(`La regla ${actual.reglaId} no puede quedar abierta si existe una regla posterior.`);
    }

    if (resolverFechaOperacion_(actual.hasta, new Date()) >= resolverFechaOperacion_(siguiente.desde, new Date())) {
      throw new Error(`Las reglas ${actual.reglaId} y ${siguiente.reglaId} se traslapan en fechas.`);
    }
  }
}

function obtenerReglaDistribucionVigenteDesdeCache_(reglasCache, fechaBase) {
  const fechaConsulta = fechaInput_(resolverFechaOperacion_(fechaBase, new Date()));

  const matches = reglasCache.filter(regla => {
    const desde = fechaInput_(regla.desde);
    const hasta = regla.hasta ? fechaInput_(regla.hasta) : '';
    return fechaConsulta >= desde && (!hasta || fechaConsulta <= hasta);
  });

  if (!matches.length) {
    const activasOrdenadas = (reglasCache || []).slice().sort((a, b) => a.desde.localeCompare(b.desde));
    if (!activasOrdenadas.length) return null;

    const masAntigua = activasOrdenadas[0];
    if (fechaConsulta < masAntigua.desde) {
      return masAntigua;
    }

    return null;
  }

  matches.sort((a, b) => b.desde.localeCompare(a.desde));
  return matches[0];
}

function obtenerSnapshotDistribucionDesdeCache_(reglasCache, fechaBase) {
  const regla = obtenerReglaDistribucionVigenteDesdeCache_(reglasCache, fechaBase);

  if (!regla) {
    throw new Error(
      `No hay regla de distribucion vigente para la fecha ${fechaInput_(fechaBase)}.`
    );
  }

  return {
    reglaId: regla.reglaId,
    steve: redondear_(numero_(regla.steve)),
    majo: redondear_(numero_(regla.majo)),
    mush: redondear_(numero_(regla.mush))
  };
}

function completarSnapshotDistribucionVentaFilaQTAS_(row, reglasCache) {
  const updated = Object.assign({}, row);

  if (snapshotDistribucionVentaCompletoQTAS_(row)) {
    return updated;
  }

  const snapshot = obtenerSnapshotDistribucionDesdeCache_(
    reglasCache,
    valorFechaVentaCanonicaQTAS_(row, new Date())
  );

  updated.Regla_Distribucion_Venta_ID = snapshot.reglaId;
  updated.Steve_Pct_Venta = snapshot.steve;
  updated.Majo_Pct_Venta = snapshot.majo;
  updated.Mush_Pct_Venta = snapshot.mush;
  updated.Actualizado_En = new Date();
  updated.__needsSave = true;

  return updated;
}

function completarSnapshotDistribucionPagoFilaQTAS_(row, reglasCache) {
  const updated = Object.assign({}, row);

  if (snapshotDistribucionPagoCompletoQTAS_(row)) {
    return updated;
  }

  const snapshot = obtenerSnapshotDistribucionDesdeCache_(
    reglasCache,
    valorFechaPagoCanonicaQTAS_(row, new Date())
  );

  updated.Regla_Distribucion_Pago_ID = snapshot.reglaId;
  updated.Steve_Pct_Pago = snapshot.steve;
  updated.Majo_Pct_Pago = snapshot.majo;
  updated.Mush_Pct_Pago = snapshot.mush;
  updated.Actualizado_En = new Date();
  updated.__needsSave = true;

  return updated;
}

function snapshotDistribucionVentaCompletoQTAS_(row) {
  return snapshotDistribucionValidoQTAS_(
    row.Regla_Distribucion_Venta_ID,
    row.Steve_Pct_Venta,
    row.Majo_Pct_Venta,
    row.Mush_Pct_Venta
  );
}

function snapshotDistribucionPagoCompletoQTAS_(row) {
  return snapshotDistribucionValidoQTAS_(
    row.Regla_Distribucion_Pago_ID,
    row.Steve_Pct_Pago,
    row.Majo_Pct_Pago,
    row.Mush_Pct_Pago
  );
}

function snapshotDistribucionValidoQTAS_(reglaId, steve, majo, mush) {
  if (!texto_(reglaId)) return false;
  if (!porcentajeDistribucionValidoQTAS_(steve)) return false;
  if (!porcentajeDistribucionValidoQTAS_(majo)) return false;
  if (!porcentajeDistribucionValidoQTAS_(mush)) return false;

  const total = redondear_(numero_(steve) + numero_(majo) + numero_(mush));
  return Math.abs(total - 100) <= 0.01;
}

function porcentajeDistribucionValidoQTAS_(value) {
  if (value instanceof Date) return false;
  if (value === '' || value === null || value === undefined) return false;

  if (typeof value === 'string') {
    const normalized = value.replace(/\s/g, '').replace(',', '.');
    if (!/^-?\d+(\.\d+)?$/.test(normalized)) {
      return false;
    }
  }

  const parsed = numero_(value);
  return isFinite(parsed) && parsed >= 0 && parsed <= 100;
}

function construirFilaDistribucionVentaQTAS_(venta) {
  const fechaVenta = valorFechaVentaCanonicaQTAS_(venta, new Date());
  const montoBase = esRegistroAnulado_(venta.Estado_Registro)
    ? 0
    : redondear_(numero_(venta.Total_Venta));
  const snapshot = {
    reglaId: texto_(venta.Regla_Distribucion_Venta_ID),
    steve: redondear_(numero_(venta.Steve_Pct_Venta)),
    majo: redondear_(numero_(venta.Majo_Pct_Venta)),
    mush: redondear_(numero_(venta.Mush_Pct_Venta))
  };
  const valores = distribuirMontoDistribucionQTAS_(montoBase, snapshot);

  return {
    Distribucion_ID: 'DV-' + String(numero_(venta.Venta_ID)).padStart(6, '0'),
    Fuente_Tipo: 'Venta',
    Fuente_ID: numero_(venta.Venta_ID),
    Venta_ID: numero_(venta.Venta_ID),
    Pago_ID: '',
    Fecha_Base: fechaVenta,
    Fecha_Venta: fechaVenta,
    Fecha_Pago: '',
    Nombre: texto_(venta.Nombre),
    Cliente_ID: texto_(venta.Cliente_ID),
    Productos_Resumen: texto_(venta.Productos_Resumen),
    Base_Distribucion: 'Venta_Neta',
    Monto_Base: montoBase,
    Regla_ID_Usada: snapshot.reglaId,
    Steve_Pct: snapshot.steve,
    Majo_Pct: snapshot.majo,
    Mush_Pct: snapshot.mush,
    Steve_Valor: valores.steve,
    Majo_Valor: valores.majo,
    Mush_Valor: valores.mush,
    Medio_Pago: '',
    Estado_Pago: texto_(venta.Estado_Pago),
    Estado_Registro: texto_(venta.Estado_Registro) || QTAS.status.registro.activo
  };
}

function construirFilaDistribucionPagoQTAS_(pago, venta) {
  const ventaBase = venta || {};
  const fechaPago = valorFechaPagoCanonicaQTAS_(pago, new Date());
  const fechaVenta = ventaBase.Venta_ID
    ? valorFechaVentaCanonicaQTAS_(ventaBase, fechaPago)
    : '';
  const estadoRegistroVenta = texto_(ventaBase.Estado_Registro);
  const estadoRegistroPago = texto_(pago.Estado_Registro) || QTAS.status.registro.activo;
  const montoBase = esRegistroAnulado_(estadoRegistroPago) || esRegistroAnulado_(estadoRegistroVenta)
    ? 0
    : redondear_(numero_(pago.Monto_Pago));
  const snapshot = {
    reglaId: texto_(pago.Regla_Distribucion_Pago_ID),
    steve: redondear_(numero_(pago.Steve_Pct_Pago)),
    majo: redondear_(numero_(pago.Majo_Pct_Pago)),
    mush: redondear_(numero_(pago.Mush_Pct_Pago))
  };
  const valores = distribuirMontoDistribucionQTAS_(montoBase, snapshot);

  return {
    Distribucion_ID: 'DP-' + texto_(pago.Pago_ID),
    Fuente_Tipo: 'Pago',
    Fuente_ID: texto_(pago.Pago_ID),
    Venta_ID: numero_(pago.Venta_ID),
    Pago_ID: texto_(pago.Pago_ID),
    Fecha_Base: fechaPago,
    Fecha_Venta: fechaVenta,
    Fecha_Pago: fechaPago,
    Nombre: texto_(pago.Nombre || ventaBase.Nombre),
    Cliente_ID: texto_(ventaBase.Cliente_ID),
    Productos_Resumen: texto_(ventaBase.Productos_Resumen),
    Base_Distribucion: 'Pago_Real',
    Monto_Base: montoBase,
    Regla_ID_Usada: snapshot.reglaId,
    Steve_Pct: snapshot.steve,
    Majo_Pct: snapshot.majo,
    Mush_Pct: snapshot.mush,
    Steve_Valor: valores.steve,
    Majo_Valor: valores.majo,
    Mush_Valor: valores.mush,
    Medio_Pago: texto_(pago.Medio_Pago),
    Estado_Pago: texto_(ventaBase.Estado_Pago),
    Estado_Registro: texto_(pago.Estado_Registro) || QTAS.status.registro.activo
  };
}

function distribuirMontoDistribucionQTAS_(montoBase, snapshot) {
  const monto = redondear_(Math.max(numero_(montoBase), 0));
  const steve = redondear_(monto * numero_(snapshot.steve) / 100);
  const majo = redondear_(monto * numero_(snapshot.majo) / 100);
  const mush = redondear_(monto - steve - majo);

  return {
    steve: steve,
    majo: majo,
    mush: mush
  };
}
