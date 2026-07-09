function registrarVentaQTAS(payload) {
  const performance = crearPerfilRendimientoQTAS_('registrarVentaQTAS');
  const debeValidarModelo = !payload || payload.validarModelo !== false;
  const usarHeadersEstaticos = !debeValidarModelo;
  const tienePagosIniciales = Boolean((payload && payload.pagos || []).some(pago => numero_(pago && pago.monto) > 0));

  try {
    return withScriptLock_('registrar venta', () => {
      if (debeValidarModelo) {
        const sheetNames = [
          QTAS.sheets.ventas,
          QTAS.sheets.detalle,
          QTAS.sheets.clientes,
          QTAS.sheets.distribucionIngresos
        ];
        if (tienePagosIniciales) {
          sheetNames.push(QTAS.sheets.pagos);
        }
        medirBloqueRendimientoQTAS_(performance, 'validarModelo', () => {
          validarModeloSoloLecturaQTAS_({
            sheetNames: sheetNames,
            validarConfig: false
          });
        });
      } else {
        registrarMedicionRendimientoQTAS_(performance, 'validarModelo', 0);
      }
      medirBloqueRendimientoQTAS_(performance, 'validarPayload', () => validarVenta_(payload));

      let ss;
      let ventasSheet;
      let detalleSheet;
      let pagosSheet;
      let distribucionSheet;
      let ventasNextRow;
      let detalleNextRow;
      let pagosNextRow;
      let distribucionNextRow;
      let ventaHeaders;
      let detalleHeaders;
      let pagoHeaders;
      let distribucionHeaders;

      medirBloqueRendimientoQTAS_(performance, 'obtenerHojasHeaders', () => {
        ss = SpreadsheetApp.getActive();
        const sheetNames = [
          QTAS.sheets.ventas,
          QTAS.sheets.detalle,
          QTAS.sheets.distribucionIngresos
        ];
        if (tienePagosIniciales) {
          sheetNames.push(QTAS.sheets.pagos);
        }
        const sheets = obtenerHojasPorNombreQTAS_(ss, sheetNames);
        ventasSheet = sheets[QTAS.sheets.ventas];
        detalleSheet = sheets[QTAS.sheets.detalle];
        distribucionSheet = sheets[QTAS.sheets.distribucionIngresos];
        ventasNextRow = ventasSheet.getLastRow() + 1;
        detalleNextRow = detalleSheet.getLastRow() + 1;
        distribucionNextRow = distribucionSheet.getLastRow() + 1;
        ventaHeaders = usarHeadersEstaticos ? QTAS.schemas[QTAS.sheets.ventas] : getHeaders_(ventasSheet);
        detalleHeaders = usarHeadersEstaticos ? QTAS.schemas[QTAS.sheets.detalle] : getHeaders_(detalleSheet);
        distribucionHeaders = usarHeadersEstaticos
          ? QTAS.schemas[QTAS.sheets.distribucionIngresos]
          : getHeaders_(distribucionSheet);

        if (tienePagosIniciales) {
          pagosSheet = sheets[QTAS.sheets.pagos];
          pagosNextRow = pagosSheet.getLastRow() + 1;
          pagoHeaders = usarHeadersEstaticos ? QTAS.schemas[QTAS.sheets.pagos] : getHeaders_(pagosSheet);
        }
      });

      let ventaId;
      let ahora;
      let fechaVentaBase;
      let fechaVenta;
      let cliente;
      let comentarioVenta;

      ventaId = medirBloqueRendimientoQTAS_(performance, 'ventaId', () =>
        siguienteIdNumericoPersistenteQTAS_('venta_id', ventasSheet, 'Venta_ID')
      );
      medirBloqueRendimientoQTAS_(performance, 'resolverVentaBase', () => {
        ahora = new Date();
        fechaVentaBase = resolverFechaOperacion_(payload.fechaVenta, ahora);
        fechaVenta = combinarFechaYHora_(fechaVentaBase, ahora);
        comentarioVenta = texto_(payload.comentarioVenta);
      });
      cliente = medirBloqueRendimientoQTAS_(performance, 'resolverCliente', () =>
        resolverClienteQTAS_(payload.cliente, ahora, { ss: ss })
      );

      let priceCache;
      let distributionCache;
      let snapshotDistribucionVenta;

      priceCache = medirBloqueRendimientoQTAS_(performance, 'cargarPreciosCache', () =>
        cargarPreciosReferenciaEnMemoria_()
      );
      distributionCache = medirBloqueRendimientoQTAS_(performance, 'cargarReglasCache', () =>
        cargarReglasDistribucionEnMemoria_()
      );
      snapshotDistribucionVenta = medirBloqueRendimientoQTAS_(performance, 'resolverReglaVenta', () =>
        obtenerSnapshotDistribucionDesdeCache_(
          distributionCache,
          fechaVenta
        )
      );

      const lineasPreparadas = medirBloqueRendimientoQTAS_(performance, 'prepararLineas', () =>
        payload.lineas.map((linea, index) =>
          prepararLineaVentaQTAS_({
            linea,
            index,
            ventaId,
            fechaVenta,
            clienteId: cliente.clienteId,
            clienteNombre: cliente.nombre,
            priceCache,
            creadoEn: ahora
          })
        )
      );

      const totalVenta = redondear_(sumar_(lineasPreparadas.map(item => item.subtotalNeto)));
      const pagosPreparados = medirBloqueRendimientoQTAS_(performance, 'prepararPagos', () =>
        prepararPagosVentaQTAS_({
          ventaId,
          nombre: cliente.nombre,
          pagos: payload.pagos || [],
          fechaPagoDefault: fechaVenta,
          distributionCache: distributionCache,
          creadoEn: ahora
        })
      );

      if (pagosPreparados.totalPagado > totalVenta + 0.009) {
        throw new Error('El pago inicial no puede ser mayor al total de la venta.');
      }

      const saldo = redondear_(Math.max(totalVenta - pagosPreparados.totalPagado, 0));
      const estadoPago = obtenerEstadoPago_(totalVenta, pagosPreparados.totalPagado, QTAS.status.registro.activo);
      const productosResumen = resumenProductos_(payload.lineas);
      const ventaRow = filaDesdeHeaders_(ventaHeaders, {
        Venta_ID: ventaId,
        Fecha_Venta: fechaVenta,
        Cliente_ID: cliente.clienteId,
        Nombre: cliente.nombre,
        Productos_Resumen: productosResumen,
        Comentario_Venta: comentarioVenta,
        Total_Venta: totalVenta,
        Total_Pagado: pagosPreparados.totalPagado,
        Saldo: saldo,
        Estado_Pago: estadoPago,
        Regla_Distribucion_Venta_ID: snapshotDistribucionVenta.reglaId,
        Steve_Pct_Venta: snapshotDistribucionVenta.steve,
        Majo_Pct_Venta: snapshotDistribucionVenta.majo,
        Mush_Pct_Venta: snapshotDistribucionVenta.mush,
        Estado_Registro: QTAS.status.registro.activo
      });
      const detalleRows = lineasPreparadas.map(item => filaDesdeHeaders_(detalleHeaders, item.row));
      const pagoRows = pagosPreparados.rows.map(row => filaDesdeHeaders_(pagoHeaders, row));

      const ventaObj = {
        Venta_ID: ventaId,
        Fecha_Venta: fechaVenta,
        Cliente_ID: cliente.clienteId,
        Nombre: cliente.nombre,
        Productos_Resumen: productosResumen,
        Comentario_Venta: comentarioVenta,
        Total_Venta: totalVenta,
        Total_Pagado: pagosPreparados.totalPagado,
        Saldo: saldo,
        Estado_Pago: estadoPago,
        Regla_Distribucion_Venta_ID: snapshotDistribucionVenta.reglaId,
        Steve_Pct_Venta: snapshotDistribucionVenta.steve,
        Majo_Pct_Venta: snapshotDistribucionVenta.majo,
        Mush_Pct_Venta: snapshotDistribucionVenta.mush,
        Estado_Registro: QTAS.status.registro.activo
      };

      medirBloqueRendimientoQTAS_(performance, 'escribirVentas', () => {
        escribirFilasDesdeFilaQTAS_(ventasSheet, ventasNextRow, [ventaRow]);
      });
      medirBloqueRendimientoQTAS_(performance, 'escribirDetalle', () => {
        escribirFilasDesdeFilaQTAS_(detalleSheet, detalleNextRow, detalleRows);
      });

      if (pagosPreparados.rows.length) {
        medirBloqueRendimientoQTAS_(performance, 'escribirPagos', () => {
          escribirFilasDesdeFilaQTAS_(pagosSheet, pagosNextRow, pagoRows);
        });
      }

      medirBloqueRendimientoQTAS_(performance, 'registrarDistribucion', () => {
        registrarDistribucionVentaNuevaQTAS_(ventaObj, pagosPreparados.rows, {
          distribucionSheet: distribucionSheet,
          distribucionHeaders: distribucionHeaders,
          startRow: distribucionNextRow
        });
      });

      let analiticaCostos = null;
      analiticaCostos = medirBloqueRendimientoQTAS_(performance, 'analiticaCostos', () => {
        try {
          return sincronizarVentaDetalleCostosLoteQTAS_(
            lineasPreparadas.map(item => item.row),
            {
              ss: ss,
              ahora: ahora
            }
          );
        } catch (error) {
          Logger.log(`No se pudo sincronizar Venta_Detalle_Costos_Calc para Venta ${ventaId}: ${error.message}`);
          return {
            ok: false,
            skipped: true,
            reason: error.message,
            rows: 0,
            inserted: 0,
            updated: 0,
            stale: 0
          };
        }
      });

      let estadoEnvio = '';
      if (payload && payload.pendienteEnvio === true) {
        estadoEnvio = medirBloqueRendimientoQTAS_(performance, 'registrarEnvioPendiente', () => {
          const envio = upsertEstadoEnvioVentaQTAS_({
            ss: ss,
            ventaId: ventaId,
            estadoEnvio: QTAS.status.envio.pendiente,
            comentarioEnvio: texto_(payload.comentarioEnvio),
            ahora: ahora
          });
          return texto_(envio.Estado_Envio);
        });
      }

      let dashboard = null;
      if (payload.devolverDashboard !== false) {
        dashboard = medirBloqueRendimientoQTAS_(performance, 'dashboard', () =>
          dashboardVentasConsistenteQTAS_()
        );
      }

      return {
        ok: true,
        ventaId,
        clienteId: cliente.clienteId,
        nombre: cliente.nombre,
        fechaVenta: fechaInput_(fechaVenta),
        totalVenta,
        totalPagado: pagosPreparados.totalPagado,
        saldo,
        estadoPago,
        estadoEnvio,
        productosResumen,
        analiticaCostos,
        dashboard,
        performance: finalizarPerfilRendimientoQTAS_(performance, { ventaId: ventaId })
      };
    }, performance);
  } catch (error) {
    finalizarPerfilRendimientoQTAS_(performance, { error: error.message });
    throw error;
  }
}

function registrarPagoPendienteQTAS(payload) {
  const performance = crearPerfilRendimientoQTAS_('registrarPagoPendienteQTAS');
  const debeValidarModelo = !payload || payload.validarModelo !== false;
  const usarHeadersEstaticos = !debeValidarModelo;

  try {
    return withScriptLock_('registrar pago pendiente', () => {
      if (debeValidarModelo) {
        medirBloqueRendimientoQTAS_(performance, 'validarModelo', () => {
          validarModeloSoloLecturaQTAS_({
            sheetNames: [
              QTAS.sheets.ventas,
              QTAS.sheets.pagos,
              QTAS.sheets.distribucionIngresos
            ],
            validarConfig: false
          });
        });
      } else {
        registrarMedicionRendimientoQTAS_(performance, 'validarModelo', 0);
      }
      medirBloqueRendimientoQTAS_(performance, 'validarPayload', () => validarPagoPendienteQTAS_(payload));

      let ss;
      let ventasSheet;
      let pagosSheet;
      let pagosNextRow;
      let ventaHeaders;
      let pagoHeaders;

      medirBloqueRendimientoQTAS_(performance, 'obtenerHojasHeaders', () => {
        ss = SpreadsheetApp.getActive();
        const sheets = obtenerHojasPorNombreQTAS_(ss, [
          QTAS.sheets.ventas,
          QTAS.sheets.pagos
        ]);
        ventasSheet = sheets[QTAS.sheets.ventas];
        pagosSheet = sheets[QTAS.sheets.pagos];
        pagosNextRow = pagosSheet.getLastRow() + 1;
        ventaHeaders = usarHeadersEstaticos ? QTAS.schemas[QTAS.sheets.ventas] : getHeaders_(ventasSheet);
        pagoHeaders = usarHeadersEstaticos ? QTAS.schemas[QTAS.sheets.pagos] : getHeaders_(pagosSheet);
      });

      const ventas = medirBloqueRendimientoQTAS_(performance, 'leerVentas', () =>
        leerObjetosConMeta_(ventasSheet)
      );
      const ventaId = numero_(payload.ventaId);
      const venta = ventas.find(row => numero_(row.Venta_ID) === ventaId);

      if (!venta) {
        throw new Error('No se encontro la venta pendiente.');
      }

      const ventaActual = normalizarVentaResumenQTAS_(venta);

      if (esRegistroAnulado_(ventaActual.Estado_Registro) || numero_(ventaActual.Saldo) <= 0) {
        throw new Error('La venta seleccionada ya no tiene saldo pendiente.');
      }

      const monto = redondear_(numero_(payload.monto));
      if (monto > numero_(ventaActual.Saldo) + 0.009) {
        throw new Error('El pago no puede ser mayor al saldo pendiente.');
      }

      let ahora;
      let fechaPago;
      let distributionCache;
      let snapshotDistribucionPago;

      medirBloqueRendimientoQTAS_(performance, 'resolverFechaPago', () => {
        ahora = new Date();
        fechaPago = resolverFechaOperacion_(payload.fechaPago, ahora);
      });
      distributionCache = medirBloqueRendimientoQTAS_(performance, 'cargarReglasCache', () =>
        cargarReglasDistribucionEnMemoria_()
      );
      snapshotDistribucionPago = medirBloqueRendimientoQTAS_(performance, 'resolverReglaPago', () =>
        obtenerSnapshotDistribucionDesdeCache_(
          distributionCache,
          fechaPago
        )
      );

      const pagoId = medirBloqueRendimientoQTAS_(performance, 'calcularPagoId', () =>
        siguientePagoIdVentaQTAS_(pagoSheetRowsQTAS_(pagosSheet), ventaId)
      );

      const fechaPagoMomento = combinarFechaYHora_(fechaPago, ahora);
      const pagoObj = {
        Pago_ID: pagoId,
        Venta_ID: ventaId,
        Fecha_Pago: fechaPagoMomento,
        Medio_Pago: texto_(payload.medio),
        Monto_Pago: monto,
        Comentario_Pago: texto_(payload.comentarioPago),
        Regla_Distribucion_Pago_ID: snapshotDistribucionPago.reglaId,
        Steve_Pct_Pago: snapshotDistribucionPago.steve,
        Majo_Pct_Pago: snapshotDistribucionPago.majo,
        Mush_Pct_Pago: snapshotDistribucionPago.mush,
        Estado_Registro: QTAS.status.registro.activo
      };
      const pagoRow = filaDesdeHeaders_(pagoHeaders, pagoObj);

      medirBloqueRendimientoQTAS_(performance, 'escribirPago', () => {
        escribirFilasDesdeFilaQTAS_(pagosSheet, pagosNextRow, [pagoRow]);
      });

      const totalPagadoActualizado = redondear_(numero_(ventaActual.Total_Pagado) + monto);
      const saldoActualizado = redondear_(Math.max(numero_(ventaActual.Total_Venta) - totalPagadoActualizado, 0));
      const ventaActualizada = Object.assign({}, ventaActual, {
        Total_Pagado: totalPagadoActualizado,
        Saldo: saldoActualizado,
        Estado_Pago: obtenerEstadoPago_(
          numero_(ventaActual.Total_Venta),
          totalPagadoActualizado,
          texto_(ventaActual.Estado_Registro) || QTAS.status.registro.activo
        )
      });

      medirBloqueRendimientoQTAS_(performance, 'actualizarVenta', () => {
        actualizarFilaObjeto_(ventasSheet, venta.__rowNumber, ventaHeaders, ventaActualizada);
      });
      medirBloqueRendimientoQTAS_(performance, 'sincronizarDistribucion', () => {
        sincronizarDistribucionVentaQTAS_(ventaId);
      });

      let dashboard = null;
      if (payload.devolverDashboard !== false) {
        dashboard = medirBloqueRendimientoQTAS_(performance, 'dashboard', () =>
          dashboardVentasConsistenteQTAS_()
        );
      }

      return {
        ok: true,
        ventaId,
        monto,
        saldoRestante: saldoActualizado,
        dashboard,
        performance: finalizarPerfilRendimientoQTAS_(performance, { ventaId: ventaId })
      };
    }, performance);
  } catch (error) {
    finalizarPerfilRendimientoQTAS_(performance, { error: error.message });
    throw error;
  }
}

function actualizarEstadoEnvioVentaQTAS(payload) {
  const performance = crearPerfilRendimientoQTAS_('actualizarEstadoEnvioVentaQTAS');
  const debeValidarModelo = !payload || payload.validarModelo !== false;

  try {
    return withScriptLock_('actualizar estado envio', () => {
      if (debeValidarModelo) {
        medirBloqueRendimientoQTAS_(performance, 'validarModelo', () => {
          validarModeloSoloLecturaQTAS_({
            sheetNames: [QTAS.sheets.ventas, QTAS.sheets.pagos, QTAS.sheets.ventasEnvio],
            validarConfig: false
          });
        });
      } else {
        registrarMedicionRendimientoQTAS_(performance, 'validarModelo', 0);
      }

      const ventaId = numero_(payload && payload.ventaId);
      const estadoEnvio = normalizarEstadoEnvioVentaQTAS_(payload && payload.estadoEnvio);
      if (ventaId <= 0) throw new Error('Falta la venta para actualizar el envio.');
      if (!estadoEnvio) throw new Error('Falta un estado de envio valido.');

      let ss;
      let ventasSheet;
      let ventaHeaders;

      medirBloqueRendimientoQTAS_(performance, 'obtenerHojasHeaders', () => {
        ss = SpreadsheetApp.getActive();
        ventasSheet = ss.getSheetByName(QTAS.sheets.ventas);
        ventaHeaders = getHeaders_(ventasSheet);
      });

      const ventas = medirBloqueRendimientoQTAS_(performance, 'leerVentas', () =>
        leerObjetosConMeta_(ventasSheet)
      );
      const venta = ventas.find(row => numero_(row.Venta_ID) === ventaId);
      if (!venta) {
        throw new Error('No se encontro la venta para actualizar el envio.');
      }

      const ventaActual = normalizarVentaResumenQTAS_(venta);
      if (texto_(ventaActual.Estado_Registro) === QTAS.status.registro.anulado) {
        throw new Error('No se puede actualizar el envio de una venta anulada.');
      }

      medirBloqueRendimientoQTAS_(performance, 'guardarEnvio', () => {
        upsertEstadoEnvioVentaQTAS_({
          ss: ss,
          ventaId: ventaId,
          estadoEnvio: estadoEnvio,
          comentarioEnvio: texto_(payload && payload.comentarioEnvio),
          ahora: new Date()
        });
      });

      let dashboard = null;
      if (!payload || payload.devolverDashboard !== false) {
        dashboard = medirBloqueRendimientoQTAS_(performance, 'dashboard', () =>
          dashboardVentasConsistenteQTAS_()
        );
      }

      return {
        ok: true,
        ventaId: ventaId,
        estadoEnvio: estadoEnvio,
        dashboard: dashboard,
        performance: finalizarPerfilRendimientoQTAS_(performance, { ventaId: ventaId })
      };
    }, performance);
  } catch (error) {
    finalizarPerfilRendimientoQTAS_(performance, { error: error.message });
    throw error;
  }
}

function dashboardVentasDesdeEstadoQTAS_(estado) {
  const source = estado || construirEstadoVentasQTAS_();
  return {
    deudores: construirDeudoresQTAS_(source),
    ventasPendientes: ventasPendientesDesdeEstadoQTAS_(source),
    enviosPendientes: construirEnviosPendientesQTAS_(source)
  };
}

function dashboardVentasConsistenteQTAS_() {
  return dashboardVentasDesdeEstadoQTAS_(construirEstadoVentasQTAS_());
}

function obtenerHojaVentasEnvioQTAS_(ss, options) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const settings = Object.assign({
    create: false,
    validate: true
  }, options || {});
  let sheet = spreadsheet.getSheetByName(QTAS.sheets.ventasEnvio);

  if (!sheet) {
    if (!settings.create) return null;
    sheet = asegurarHojaModelo_(spreadsheet, QTAS.sheets.ventasEnvio, QTAS.schemas[QTAS.sheets.ventasEnvio]);
  }

  if (settings.validate !== false && !headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.ventasEnvio])) {
    throw new Error(
      `La hoja ${QTAS.sheets.ventasEnvio} no coincide con la estructura esperada. ` +
      'Ejecuta "Crear / reparar modelo" manualmente.'
    );
  }

  return sheet;
}

function leerEstadosEnvioVentasQTAS_(ss) {
  const sheet = obtenerHojaVentasEnvioQTAS_(ss, { create: false, validate: false });
  if (!sheet) return [];
  if (!headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.ventasEnvio])) {
    Logger.log(
      `Ignorando ${QTAS.sheets.ventasEnvio} porque no coincide con la estructura esperada.`
    );
    return [];
  }
  return leerObjetosConMeta_(sheet);
}

function normalizarEstadoEnvioVentaQTAS_(value) {
  const normalizado = normalizarClaveTexto_(value);
  if (normalizado === normalizarClaveTexto_(QTAS.status.envio.pendiente)) {
    return QTAS.status.envio.pendiente;
  }
  if (normalizado === normalizarClaveTexto_(QTAS.status.envio.enviado)) {
    return QTAS.status.envio.enviado;
  }
  return '';
}

function upsertEstadoEnvioVentaQTAS_(context) {
  const ss = context && context.ss ? context.ss : SpreadsheetApp.getActive();
  const sheet = context && context.sheet
    ? context.sheet
    : obtenerHojaVentasEnvioQTAS_(ss, { create: true, validate: true });
  const headers = context && context.headers
    ? context.headers
    : QTAS.schemas[QTAS.sheets.ventasEnvio];
  const rows = context && context.rows ? context.rows : leerObjetosConMeta_(sheet);
  const ventaId = numero_(context && context.ventaId);
  const ahora = context && context.ahora ? context.ahora : new Date();
  const estadoEnvio = normalizarEstadoEnvioVentaQTAS_(context && context.estadoEnvio);

  if (ventaId <= 0) {
    throw new Error('Falta la venta para actualizar el envio.');
  }
  if (!estadoEnvio) {
    throw new Error('Estado de envio invalido.');
  }

  const existente = rows.find(row => numero_(row.Venta_ID) === ventaId) || null;
  const actualizado = {
    Venta_ID: ventaId,
    Estado_Envio: estadoEnvio,
    Fecha_Pendiente_Envio: existente ? existente.Fecha_Pendiente_Envio : '',
    Fecha_Envio: existente ? existente.Fecha_Envio : '',
    Comentario_Envio: texto_(context && context.comentarioEnvio) || texto_(existente && existente.Comentario_Envio),
    Creado_En: existente && existente.Creado_En ? existente.Creado_En : ahora,
    Actualizado_En: ahora
  };

  if (estadoEnvio === QTAS.status.envio.pendiente) {
    actualizado.Fecha_Pendiente_Envio = existente && existente.Fecha_Pendiente_Envio
      ? existente.Fecha_Pendiente_Envio
      : ahora;
    actualizado.Fecha_Envio = '';
  } else if (estadoEnvio === QTAS.status.envio.enviado) {
    actualizado.Fecha_Envio = ahora;
  }

  if (existente) {
    actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, actualizado);
  } else {
    escribirFilas_(sheet, [filaDesdeHeaders_(headers, actualizado)]);
  }

  return actualizado;
}

function construirEnviosPendientesQTAS_(estado) {
  const source = estado || construirEstadoVentasQTAS_();
  const mapaEnvios = {};

  leerEstadosEnvioVentasQTAS_().forEach(row => {
    const ventaId = numero_(row.Venta_ID);
    const estadoEnvio = normalizarEstadoEnvioVentaQTAS_(row.Estado_Envio);
    if (ventaId <= 0 || estadoEnvio !== QTAS.status.envio.pendiente) return;
    mapaEnvios[ventaId] = row;
  });

  return (source.ventasActualizadas || [])
    .filter(venta =>
      texto_(venta.Estado_Registro) !== QTAS.status.registro.anulado &&
      mapaEnvios[numero_(venta.Venta_ID)]
    )
    .map(venta => {
      const ventaId = numero_(venta.Venta_ID);
      const envio = mapaEnvios[ventaId];
      const fechaVenta = valorFechaVentaCanonicaQTAS_(venta, new Date());
      const fechaPendiente = envio && envio.Fecha_Pendiente_Envio
        ? envio.Fecha_Pendiente_Envio
        : fechaVenta;
      const saldo = redondear_(numero_(venta.Saldo));

      return {
        ventaId: ventaId,
        clienteId: texto_(venta.Cliente_ID),
        nombre: texto_(venta.Nombre),
        fechaVenta: fechaInput_(fechaVenta),
        fechaPendienteEnvio: fechaInput_(fechaPendiente),
        productos: texto_(venta.Productos_Resumen),
        estadoPago: texto_(venta.Estado_Pago),
        saldo: saldo,
        comentarioEnvio: texto_(envio && envio.Comentario_Envio),
        label: [
          `V${ventaId}`,
          texto_(venta.Nombre),
          texto_(venta.Estado_Pago) || QTAS.status.pago.pendiente
        ].filter(Boolean).join(' | ')
      };
    })
    .sort((a, b) => {
      const delta = fecha_(a.fechaPendienteEnvio) - fecha_(b.fechaPendienteEnvio);
      return delta !== 0 ? delta : numero_(a.ventaId) - numero_(b.ventaId);
    });
}

function ventasPendientesDesdeEstadoQTAS_(estado) {
  const source = estado || construirEstadoVentasQTAS_();

  return (source.pendientes || [])
    .slice()
    .sort((a, b) => {
      const delta = fecha_(a.fechaVenta) - fecha_(b.fechaVenta);
      return delta !== 0 ? delta : numero_(a.ventaId) - numero_(b.ventaId);
    })
    .map(item => ({
      ventaId: numero_(item.ventaId),
      nombre: item.nombre,
      clienteId: item.clienteId,
      fechaVenta: fechaInput_(item.fechaVenta),
      saldo: item.saldo,
      productos: item.productos,
      label: `V${item.ventaId} | ${item.nombre} | ${moneda_(item.saldo)} | ${fechaInput_(item.fechaVenta)}`
    }));
}

function construirEstadoVentasQTAS_() {
  const ss = SpreadsheetApp.getActive();
  const ventasSheet = ss.getSheetByName(QTAS.sheets.ventas);
  const pagosSheet = ss.getSheetByName(QTAS.sheets.pagos);
  const ventas = leerObjetos_(ventasSheet);
  const pagos = leerObjetos_(pagosSheet).filter(row => !esRegistroAnulado_(row.Estado_Registro));
  const pagosPorVenta = {};

  pagos.forEach(pago => {
    const ventaId = texto_(pago.Venta_ID);
    pagosPorVenta[ventaId] = redondear_((pagosPorVenta[ventaId] || 0) + numero_(pago.Monto_Pago));
  });

  const ventasActualizadas = [];
  const pendientes = [];

  ventas.forEach(venta => {
    const ventaId = texto_(venta.Venta_ID);
    if (!ventaId) return;

    const totalVenta = redondear_(numero_(venta.Total_Venta));
    const totalPagado = redondear_(pagoPorVentaSeguraQTAS_(pagosPorVenta, ventaId));
    const estadoRegistro = esRegistroAnulado_(venta.Estado_Registro) ||
      normalizarClaveTexto_(venta.Estado_Pago) === normalizarClaveTexto_(QTAS.status.pago.anulado)
      ? QTAS.status.registro.anulado
      : QTAS.status.registro.activo;
    const saldo = estadoRegistro === QTAS.status.registro.anulado
      ? 0
      : redondear_(Math.max(totalVenta - totalPagado, 0));
    const estadoPago = obtenerEstadoPago_(totalVenta, totalPagado, estadoRegistro);
    const fechaVenta = valorFechaVentaCanonicaQTAS_(venta, new Date());

    const ventaActualizada = Object.assign({}, venta, {
      Total_Pagado: totalPagado,
      Saldo: saldo,
      Estado_Pago: estadoPago,
      Fecha_Venta: fechaVenta,
      Estado_Registro: estadoRegistro
    });

    ventasActualizadas.push(ventaActualizada);
    if (estadoRegistro !== QTAS.status.registro.anulado && saldo > 0) {
      pendientes.push({
        ventaId: numero_(ventaId),
        clienteId: texto_(venta.Cliente_ID),
        nombre: texto_(venta.Nombre),
        fechaVenta: fechaVenta,
        saldo: saldo,
        productos: texto_(venta.Productos_Resumen)
      });
    }
  });

  return {
    ventasActualizadas,
    pendientes
  };
}

function construirDeudoresQTAS_(estado) {
  const source = estado || construirEstadoVentasQTAS_();
  const deudasPorCliente = agrupar_(
    source.pendientes,
    item => texto_(item.clienteId) || normalizarClaveTexto_(item.nombre)
  );

  return Object.keys(deudasPorCliente)
    .map(key => {
      const ventasCliente = deudasPorCliente[key];
      const primera = ventasCliente[0];
      const saldoTotal = redondear_(sumar_(ventasCliente.map(item => item.saldo)));
      const ultimaFecha = ventasCliente
        .map(item => fecha_(item.fechaVenta))
        .sort((a, b) => b - a)[0];
      const detalle = ventasCliente
        .sort((a, b) => fecha_(a.fechaVenta) - fecha_(b.fechaVenta))
        .map(item => `V${item.ventaId}: ${moneda_(item.saldo)} - ${item.productos}`)
        .join(' | ');

      return {
        Cliente_ID: primera.clienteId,
        Nombre: primera.nombre,
        Saldo_Total: saldoTotal,
        Ventas_Pendientes: ventasCliente.length,
        Ultima_Fecha: ultimaFecha,
        Detalle_Ventas: detalle
      };
    })
    .sort((a, b) => numero_(b.Saldo_Total) - numero_(a.Saldo_Total));
}

function prepararLineaVentaQTAS_(context) {
  const linea = context.linea;
  const producto = texto_(linea.producto);
  const cantidad = redondear_(numero_(linea.cantidad));
  const unidad = normalizarUnidadCanonicaQTAS_(linea.unidad);
  const precioLista = redondear_(
    obtenerPrecioVigenteDesdeCache_(
      context.priceCache,
      producto,
      unidad,
      context.fechaVenta
    )
  );
  const precioVendidoUnitario = linea.precioVendidoUnitario === '' || linea.precioVendidoUnitario === undefined
    ? precioLista
    : redondear_(numero_(linea.precioVendidoUnitario));
  const descuentoLinea = redondear_(numero_(linea.descuentoLinea));
  const subtotalBruto = redondear_(cantidad * precioVendidoUnitario);

  if (cantidad <= 0) {
    throw new Error(`Cantidad invalida en la linea ${context.index + 1}.`);
  }
  if (precioVendidoUnitario <= 0) {
    throw new Error(`Precio vendido invalido en la linea ${context.index + 1}.`);
  }
  if (descuentoLinea < 0) {
    throw new Error(`Descuento invalido en la linea ${context.index + 1}.`);
  }
  if (descuentoLinea > subtotalBruto + 0.009) {
    throw new Error(`El descuento no puede ser mayor al subtotal bruto en la linea ${context.index + 1}.`);
  }

  const subtotalNeto = redondear_(subtotalBruto - descuentoLinea);

  return {
    subtotalNeto,
    row: {
      Detalle_ID: detalleId_(context.ventaId, context.index + 1),
      Venta_ID: context.ventaId,
      Fecha_Venta: context.fechaVenta,
      Cliente_ID: context.clienteId,
      Nombre: context.clienteNombre,
      Producto_Estandar: producto,
      Cantidad: cantidad,
      Unidad: unidad,
      Precio_Lista: precioLista,
      Precio_Vendido_Unitario: precioVendidoUnitario,
      Descuento_Linea: descuentoLinea,
      Subtotal_Bruto: subtotalBruto,
      Subtotal_Neto: subtotalNeto,
      Comentario_Linea: texto_(linea.comentarioLinea),
      Estado_Registro: QTAS.status.registro.activo
    }
  };
}

function prepararPagosVentaQTAS_(context) {
  const rows = [];
  let totalPagado = 0;

  context.pagos.forEach((pago, index) => {
    const monto = redondear_(numero_(pago.monto));
    if (monto <= 0) return;

    const fechaPago = resolverFechaOperacion_(pago.fechaPago || context.fechaPagoDefault, context.creadoEn);
    const snapshotDistribucionPago = obtenerSnapshotDistribucionDesdeCache_(
      context.distributionCache,
      fechaPago
    );
    totalPagado = redondear_(totalPagado + monto);

    rows.push({
      Pago_ID: siguientePagoIdVentaSimpleQTAS_(context.ventaId, index + 1),
      Venta_ID: context.ventaId,
      Fecha_Pago: combinarFechaYHora_(fechaPago, context.creadoEn),
      Medio_Pago: texto_(pago.medio),
      Monto_Pago: monto,
      Comentario_Pago: texto_(pago.comentarioPago),
      Regla_Distribucion_Pago_ID: snapshotDistribucionPago.reglaId,
      Steve_Pct_Pago: snapshotDistribucionPago.steve,
      Majo_Pct_Pago: snapshotDistribucionPago.majo,
      Mush_Pct_Pago: snapshotDistribucionPago.mush,
      Estado_Registro: QTAS.status.registro.activo
    });
  });

  return {
    rows,
    totalPagado
  };
}

function resolverClienteQTAS_(input, ahora, options) {
  const settings = options || {};
  if (!input || !texto_(input.nombre)) {
    throw new Error('Falta el nombre del cliente.');
  }

  const nombre = texto_(input.nombre);
  const clienteIdBuscado = texto_(input.clienteId);
  if (input.confirmadoCatalogo === true && clienteIdBuscado) {
    return {
      clienteId: clienteIdBuscado,
      nombre: nombre
    };
  }

  const ss = settings.ss || SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(QTAS.sheets.clientes);
  const headers = QTAS.schemas[QTAS.sheets.clientes];
  const colClienteId = headers.indexOf('Cliente_ID') + 1;
  const colNombre = headers.indexOf('Nombre') + 1;

  if (!colClienteId || !colNombre) {
    throw new Error('La hoja Clientes no tiene la estructura esperada.');
  }

  let existente = null;

  // Con nombres historicos poco estandarizados, solo reutilizamos clientes
  // cuando llega un Cliente_ID concreto desde el catalogo.
  if (clienteIdBuscado) {
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, colClienteId, lastRow - 1, 1).getValues().flat().map(texto_);
      const indexById = ids.findIndex(value => value === clienteIdBuscado);
      if (indexById >= 0) {
        const values = sheet.getRange(indexById + 2, 1, 1, headers.length).getValues()[0];
        existente = { __rowNumber: indexById + 2 };
        headers.forEach((header, headerIndex) => {
          existente[header] = values[headerIndex];
        });
      }
    }
  }

  if (existente) {
    const nombreResuelto = texto_(existente.Nombre) || nombre;
    const actualizado = Object.assign({}, existente, {
      Nombre: nombreResuelto,
      Activo: true
    });

    const nombreActual = texto_(existente.Nombre);
    const necesitaActualizar = !estaActivo_(existente.Activo) ||
      nombreActual !== nombreResuelto;

    if (necesitaActualizar) {
      actualizado.Actualizado_En = ahora;
      actualizarFilaObjeto_(sheet, existente.__rowNumber, headers, actualizado);
    }

    return {
      clienteId: texto_(actualizado.Cliente_ID),
      nombre: texto_(actualizado.Nombre) || nombre
    };
  }

  const clienteId = siguienteIdConPrefijoPersistenteQTAS_('cliente_id', sheet, 'Cliente_ID', 'CLI-');
  const nuevoCliente = {
    Cliente_ID: clienteId,
    Nombre: nombre,
    Activo: true,
    Creado_En: ahora,
    Actualizado_En: ahora
  };

  escribirFilas_(sheet, [filaDesdeHeaders_(headers, nuevoCliente)]);

  return {
    clienteId: clienteId,
    nombre: nombre
  };
}

function validarPagoPendienteQTAS_(payload) {
  if (!payload) throw new Error('Pago vacio.');
  if (numero_(payload.ventaId) <= 0) throw new Error('Falta la venta pendiente.');
  if (!texto_(payload.medio)) throw new Error('Falta el medio de pago.');
  if (numero_(payload.monto) <= 0) throw new Error('El monto debe ser mayor a cero.');
}

function pagoPorVentaSeguraQTAS_(pagosPorVenta, ventaId) {
  return pagosPorVenta[texto_(ventaId)] || 0;
}

function pagoSheetRowsQTAS_(sheet) {
  return leerObjetos_(sheet);
}

function siguientePagoIdVentaQTAS_(pagos, ventaId) {
  const correlativo = pagos.filter(row => numero_(row.Venta_ID) === numero_(ventaId)).length + 1;
  return siguientePagoIdVentaSimpleQTAS_(ventaId, correlativo);
}

function siguientePagoIdVentaSimpleQTAS_(ventaId, correlativo) {
  return `P-${String(ventaId).padStart(6, '0')}-${String(correlativo).padStart(2, '0')}`;
}

function normalizarVentaResumenQTAS_(venta) {
  const totalVenta = redondear_(numero_(venta.Total_Venta));
  const totalPagado = redondear_(numero_(venta.Total_Pagado));
  const estadoRegistro = esRegistroAnulado_(venta.Estado_Registro) ||
    normalizarClaveTexto_(venta.Estado_Pago) === normalizarClaveTexto_(QTAS.status.pago.anulado)
    ? QTAS.status.registro.anulado
    : QTAS.status.registro.activo;
  const saldo = estadoRegistro === QTAS.status.registro.anulado
    ? 0
    : redondear_(Math.max(totalVenta - totalPagado, 0));

  return Object.assign({}, venta, {
    Fecha_Venta: valorFechaVentaCanonicaQTAS_(venta, new Date()),
    Total_Venta: totalVenta,
    Total_Pagado: totalPagado,
    Saldo: saldo,
    Estado_Pago: obtenerEstadoPago_(totalVenta, totalPagado, estadoRegistro),
    Estado_Registro: estadoRegistro
  });
}
