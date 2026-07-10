var QTAS_RUNTIME_CACHE = QTAS_RUNTIME_CACHE || {
  headers: {},
  memo: {}
};

function valorPropiedadScriptQTAS_(key) {
  try {
    return texto_(PropertiesService.getScriptProperties().getProperty(texto_(key)));
  } catch (error) {
    return '';
  }
}

function propiedadBooleanaScriptQTAS_(key, fallback) {
  const raw = valorPropiedadScriptQTAS_(key);
  if (!raw) return fallback === undefined ? false : Boolean(fallback);
  return ['true', '1', 'si', 'yes', 'on'].includes(normalizarClaveTexto_(raw));
}

function operacionesDestructivasPermitidasQTAS_() {
  return propiedadBooleanaScriptQTAS_('QTAS_ALLOW_DESTRUCTIVE', false);
}

function assertOperacionDestructivaPermitidaQTAS_(operation) {
  if (operacionesDestructivasPermitidasQTAS_()) return;

  throw new Error(
    `La operacion "${texto_(operation)}" esta bloqueada en modo seguro. ` +
    'En produccion deja QTAS_ALLOW_DESTRUCTIVE apagado y habilitalo solo temporalmente en QA o mantenimiento manual.'
  );
}

function getHeaders_(sheet) {
  if (!sheet) return [];

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return [];

  const cacheKey = `${sheet.getSheetId()}:${lastColumn}`;
  const cached = obtenerCacheEjecucionQTAS_('headers')[cacheKey];
  if (cached) return cached.slice();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(h => texto_(h));

  while (headers.length && !headers[headers.length - 1]) {
    headers.pop();
  }

  obtenerCacheEjecucionQTAS_('headers')[cacheKey] = headers.slice();
  return headers;
}

function leerObjetos_(sheet) {
  return leerObjetosConMeta_(sheet).map(row => {
    const clean = {};
    Object.keys(row).forEach(key => {
      if (key !== '__rowNumber') clean[key] = row[key];
    });
    return clean;
  });
}

function leerObjetosConMeta_(sheet) {
  const headers = getHeaders_(sheet);
  const lastRow = sheet ? sheet.getLastRow() : 0;
  if (!sheet || !headers.length || lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  return values
    .map((row, index) => {
      if (!row.some(cell => cell !== '' && cell !== null)) return null;

      const obj = { __rowNumber: index + 2 };
      headers.forEach((header, headerIndex) => {
        obj[header] = row[headerIndex];
      });
      return obj;
    })
    .filter(Boolean);
}

function filaDesdeHeaders_(headers, obj) {
  return headers.map(header => {
    if (!obj || obj[header] === undefined || obj[header] === null) return '';
    if (esHeaderFechaSoloQTAS_(header)) return fechaTextoPlanoQTAS_(obj[header]);
    return obj[header];
  });
}

function escribirFilas_(sheet, rows) {
  if (!sheet || !rows || !rows.length) return;
  const lastRow = sheet.getLastRow();
  escribirFilasDesdeFilaQTAS_(sheet, lastRow + 1, rows);
}

function escribirFilasDesdeFilaQTAS_(sheet, startRow, rows) {
  if (!sheet || !rows || !rows.length || numero_(startRow) <= 0) return;

  sheet
    .getRange(startRow, 1, rows.length, rows[0].length)
    .setValues(rows);
}

function sobrescribirObjetosHojaQTAS_(sheet, headers, objects) {
  if (!sheet || !headers || !headers.length) return;

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  if (lastRow > 1 && lastColumn > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
  }

  if (!objects || !objects.length) return;

  sheet
    .getRange(2, 1, objects.length, headers.length)
    .setValues(objects.map(obj => filaDesdeHeaders_(headers, obj)));
}

function reemplazarObjetos_(sheet, headers, objects) {
  if (!sheet || !headers || !headers.length) return;

  assertOperacionDestructivaPermitidaQTAS_(`reemplazar contenido completo de ${sheet.getName()}`);
  limpiarDatos_(sheet, { destructiveAuthorized: true });
  if (!objects || !objects.length) return;

  sheet
    .getRange(2, 1, objects.length, headers.length)
    .setValues(objects.map(obj => filaDesdeHeaders_(headers, obj)));
}

function limpiarDatos_(sheet, options) {
  const settings = options || {};
  const lastRow = sheet ? sheet.getLastRow() : 0;
  if (!sheet || lastRow <= 1) return;

  if (!settings.destructiveAuthorized) {
    assertOperacionDestructivaPermitidaQTAS_(`limpiar datos de ${sheet.getName()}`);
  }

  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .clearContent();
}

function limpiarCachesEjecucionQTAS_() {
  QTAS_RUNTIME_CACHE.headers = {};
  QTAS_RUNTIME_CACHE.memo = {};
}

function obtenerCacheEjecucionQTAS_(bucket) {
  if (!QTAS_RUNTIME_CACHE[bucket]) {
    QTAS_RUNTIME_CACHE[bucket] = {};
  }

  return QTAS_RUNTIME_CACHE[bucket];
}

function limpiarCacheHeadersHojaQTAS_(sheet) {
  if (!sheet) return;

  const prefix = `${sheet.getSheetId()}:`;
  const headersCache = obtenerCacheEjecucionQTAS_('headers');
  Object.keys(headersCache).forEach(key => {
    if (key.indexOf(prefix) === 0) {
      delete headersCache[key];
    }
  });
}

function obtenerMemoEjecucionQTAS_(key, builder) {
  const memo = obtenerCacheEjecucionQTAS_('memo');
  if (Object.prototype.hasOwnProperty.call(memo, key)) {
    return memo[key];
  }

  const value = builder();
  memo[key] = value;
  return value;
}

function construirClaveCacheDocumentoQTAS_(namespace) {
  const ss = SpreadsheetApp.getActive();
  return [
    'qtas',
    ss ? ss.getId() : 'sin_ss',
    texto_(namespace)
  ].join(':');
}

function limpiarMemoCacheDocumentoQTAS_(namespace) {
  const prefix = `cache:${texto_(namespace)}`;
  const memo = obtenerCacheEjecucionQTAS_('memo');
  Object.keys(memo).forEach(key => {
    if (key.indexOf(prefix) === 0) {
      delete memo[key];
    }
  });
}

function leerCacheDocumentoQTAS_(namespace) {
  try {
    const cache = CacheService.getDocumentCache();
    const raw = cache.get(construirClaveCacheDocumentoQTAS_(namespace));
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    Logger.log(`No se pudo leer cache ${namespace}: ${error.message}`);
    return null;
  }
}

function guardarCacheDocumentoQTAS_(namespace, value, ttlSeconds) {
  try {
    const raw = JSON.stringify(value);
    if (!raw || raw.length > 90000) return false;

    CacheService.getDocumentCache().put(
      construirClaveCacheDocumentoQTAS_(namespace),
      raw,
      Math.max(60, numero_(ttlSeconds) || 300)
    );
    return true;
  } catch (error) {
    Logger.log(`No se pudo guardar cache ${namespace}: ${error.message}`);
    return false;
  }
}

function invalidarCacheDocumentoQTAS_(namespace) {
  try {
    limpiarMemoCacheDocumentoQTAS_(namespace);
    CacheService.getDocumentCache().remove(construirClaveCacheDocumentoQTAS_(namespace));
  } catch (error) {
    Logger.log(`No se pudo invalidar cache ${namespace}: ${error.message}`);
  }
}

function obtenerHojasPorNombreQTAS_(ss, sheetNames) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const requested = (sheetNames || []).filter(Boolean);
  const wanted = {};
  requested.forEach(name => {
    wanted[name] = null;
  });

  spreadsheet.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (Object.prototype.hasOwnProperty.call(wanted, name)) {
      wanted[name] = sheet;
    }
  });

  return wanted;
}

function esHojaOpcionalQTAS_(sheetName) {
  const optionalSheets = [
    QTAS.sheets.ventasEnvio,
    QTAS.sheets.costoProductoCalculado,
    QTAS.sheets.ventaDetalleCostosCalculado,
    QTAS.sheets.compraOrigenesFondos
  ].map(texto_);
  return optionalSheets.indexOf(texto_(sheetName)) >= 0;
}

function asegurarHojaModelo_(ss, nombre, headers) {
  let sheet = ss.getSheetByName(nombre);
  let debeAutoAjustar = false;

  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    debeAutoAjustar = true;
  }

  const currentHeaders = getHeaders_(sheet);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    limpiarCacheHeadersHojaQTAS_(sheet);
    debeAutoAjustar = true;
  } else if (!currentHeaders.length) {
    assertOperacionDestructivaPermitidaQTAS_(`reconstruir hoja ${nombre} sin headers`);
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    limpiarCacheHeadersHojaQTAS_(sheet);
    debeAutoAjustar = true;
  } else if (!headersIguales_(currentHeaders, headers)) {
    assertOperacionDestructivaPermitidaQTAS_(`sincronizar headers de ${nombre}`);
    sincronizarHeaders_(sheet, currentHeaders, headers);
    debeAutoAjustar = true;
  }

  if (sheet.getFrozenRows() !== 1) {
    sheet.setFrozenRows(1);
  }

  sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).setFontWeight('bold');

  if (debeAutoAjustar) {
    sheet.autoResizeColumns(1, Math.max(sheet.getLastColumn(), headers.length));
  }

  return sheet;
}

function obtenerHojaConfigQTAS_(ss, options) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const settings = options || {};
  const canonical = spreadsheet.getSheetByName(QTAS.sheets.config);
  if (canonical) return canonical;

  if (!settings.create) return null;
  return asegurarHojaModelo_(spreadsheet, QTAS.sheets.config, QTAS.schemas[QTAS.sheets.config]);
}

function esConfigMediosPagoCanonicoQTAS_(sheet) {
  return headersIguales_(getHeaders_(sheet), QTAS.schemas[QTAS.sheets.config]);
}

function construirMediosPagoBaseQTAS_() {
  return MEDIOS_PAGO.map(item => ({
    medioPago: item,
    activo: true,
    nota: ''
  }));
}

function leerMediosPagoConfiguradosQTAS_() {
  const sheet = obtenerHojaConfigQTAS_(SpreadsheetApp.getActive());
  if (!sheet || sheet.getLastRow() < 2) {
    return construirMediosPagoBaseQTAS_();
  }

  if (!esConfigMediosPagoCanonicoQTAS_(sheet)) {
    throw new Error(
      `La hoja ${QTAS.sheets.config} no coincide con la estructura esperada. ` +
      'Corrige la configuracion antes de continuar.'
    );
  }

  return leerObjetos_(sheet)
    .map(row => ({
      medioPago: texto_(row.Medio_Pago),
      activo: estaActivo_(row.Activo),
      nota: texto_(row.Nota)
    }))
    .filter(row => row.medioPago);
}

function sincronizarHeaders_(sheet, currentHeaders, expectedHeaders) {
  if (headersIguales_(currentHeaders, expectedHeaders)) return;

  if (!currentHeaders.length) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    limpiarCacheHeadersHojaQTAS_(sheet);
    return;
  }

  const sheetName = sheet.getName();
  const existingRows = leerObjetos_(sheet).map(row =>
    normalizarFilaModeloQTAS_(sheetName, row)
  );
  const currentMaxColumns = sheet.getMaxColumns();

  assertOperacionDestructivaPermitidaQTAS_(`sincronizar estructura de ${sheetName}`);

  if (currentMaxColumns < expectedHeaders.length) {
    sheet.insertColumnsAfter(currentMaxColumns, expectedHeaders.length - currentMaxColumns);
  } else if (currentMaxColumns > expectedHeaders.length) {
    sheet.deleteColumns(expectedHeaders.length + 1, currentMaxColumns - expectedHeaders.length);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  limpiarCacheHeadersHojaQTAS_(sheet);

  if (existingRows.length) {
    sheet
      .getRange(2, 1, existingRows.length, expectedHeaders.length)
      .setValues(existingRows.map(row => filaDesdeHeaders_(expectedHeaders, row)));
  }
}

function normalizarFilaModeloQTAS_(sheetName, row) {
  const normalized = Object.assign({}, row);
  const fallback = row.Actualizado_En || row.Creado_En || new Date();

  if (sheetName === QTAS.sheets.ventas || sheetName === QTAS.sheets.detalle) {
    normalized.Fecha_Venta = valorFechaVentaCanonicaQTAS_(row, fallback);
  }

  if (sheetName === QTAS.sheets.compras || sheetName === QTAS.sheets.compraDetalle) {
    normalized.Fecha_Compra = valorFechaCompraCanonicaQTAS_(row, fallback);
  }

  if (sheetName === QTAS.sheets.pagos) {
    normalized.Fecha_Pago = valorFechaPagoCanonicaQTAS_(row, fallback);
  }

  if (sheetName === QTAS.sheets.ventasEnvio) {
    normalized.Fecha_Pendiente_Envio = resolverFechaMomentoOpcionalQTAS_(
      row.Fecha_Pendiente_Envio,
      row.Fecha_Pendiente_Envio,
      fallback
    );
    normalized.Fecha_Envio = resolverFechaMomentoOpcionalQTAS_(
      row.Fecha_Envio,
      row.Fecha_Envio,
      fallback
    );
  }

  if (sheetName === QTAS.sheets.distribucionIngresos) {
    normalized.Fecha_Base = valorFechaBaseCanonicaQTAS_(row, fallback);
    normalized.Fecha_Venta = valorFechaDistribucionVentaCanonicaQTAS_(row, fallback);
    normalized.Fecha_Pago = valorFechaDistribucionPagoCanonicaQTAS_(row, fallback);
  }

  return normalized;
}

function headersIguales_(actual, expected) {
  if (actual.length !== expected.length) return false;
  return actual.every((value, index) => value === expected[index]);
}

function actualizarFilaObjeto_(sheet, rowNumber, headers, obj) {
  if (!sheet || !rowNumber || !headers || !headers.length) return;
  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([filaDesdeHeaders_(headers, obj)]);
}

function siguienteIdNumerico_(sheet, headerName) {
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(headerName) + 1;
  const lastRow = sheet ? sheet.getLastRow() : 0;

  if (!col || lastRow < 2) return 1;

  const ultimoValor = numero_(sheet.getRange(lastRow, col).getValue());
  if (ultimoValor > 0) {
    return ultimoValor + 1;
  }

  const values = sheet
    .getRange(2, col, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(numero_)
    .filter(n => n > 0);

  return values.length ? Math.max.apply(null, values) + 1 : 1;
}

function siguienteIdConPrefijo_(sheet, headerName, prefix, padLength) {
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(headerName) + 1;
  const size = padLength || 6;
  const lastRow = sheet ? sheet.getLastRow() : 0;

  if (!col || lastRow < 2) {
    return prefix + String(1).padStart(size, '0');
  }

  const ultimoValor = texto_(sheet.getRange(lastRow, col).getValue());
  const ultimoMatch = ultimoValor.match(/(\d+)$/);
  if (ultimoMatch) {
    return prefix + String(Number(ultimoMatch[1]) + 1).padStart(size, '0');
  }

  const maxValue = sheet
    .getRange(2, col, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(texto_)
    .map(value => {
      const match = value.match(/(\d+)$/);
      return match ? Number(match[1]) : 0;
    })
    .reduce((max, value) => Math.max(max, value), 0);

  return prefix + String(maxValue + 1).padStart(size, '0');
}

function siguienteSecuenciaPersistenteQTAS_(sequenceKey, resolverUltimoAsignado) {
  const props = PropertiesService.getDocumentProperties();
  const propKey = `QTAS_SEQ_${texto_(sequenceKey)}`;
  const ultimoAsignado = numero_(props.getProperty(propKey));

  if (ultimoAsignado > 0) {
    const siguiente = ultimoAsignado + 1;
    props.setProperty(propKey, String(siguiente));
    return siguiente;
  }

  const base = Math.max(0, numero_(resolverUltimoAsignado ? resolverUltimoAsignado() : 0));
  const siguiente = base + 1;
  props.setProperty(propKey, String(siguiente));
  return siguiente;
}

function siguienteIdNumericoPersistenteQTAS_(sequenceKey, sheet, headerName) {
  return siguienteSecuenciaPersistenteQTAS_(sequenceKey, () =>
    siguienteIdNumerico_(sheet, headerName) - 1
  );
}

function siguienteIdConPrefijoPersistenteQTAS_(sequenceKey, sheet, headerName, prefix, padLength) {
  const size = padLength || 6;
  const siguienteNumero = siguienteSecuenciaPersistenteQTAS_(sequenceKey, () => {
    const siguienteId = siguienteIdConPrefijo_(sheet, headerName, prefix, size);
    const match = texto_(siguienteId).match(/(\d+)$/);
    return match ? Number(match[1]) - 1 : 0;
  });

  return prefix + String(siguienteNumero).padStart(size, '0');
}

function maximoIdNumericoHojaQTAS_(sheet, headerName) {
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(headerName) + 1;
  const lastRow = sheet ? sheet.getLastRow() : 0;

  if (!col || lastRow < 2) return 0;

  return sheet
    .getRange(2, col, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(numero_)
    .reduce((max, value) => Math.max(max, value > 0 ? value : 0), 0);
}

function fijarSecuenciaPersistenteQTAS_(sequenceKey, lastAssigned) {
  const props = PropertiesService.getDocumentProperties();
  const propKey = `QTAS_SEQ_${texto_(sequenceKey)}`;
  const value = Math.max(0, Math.floor(numero_(lastAssigned)));

  if (value > 0) {
    props.setProperty(propKey, String(value));
  } else {
    props.deleteProperty(propKey);
  }

  return value;
}

function resincronizarIdNumericoPersistenteQTAS_(sequenceKey, sheet, headerName) {
  return fijarSecuenciaPersistenteQTAS_(
    sequenceKey,
    maximoIdNumericoHojaQTAS_(sheet, headerName)
  );
}

function withScriptLock_(label, fn, performance) {
  const waitStartedAt = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  registrarMedicionRendimientoQTAS_(performance, `lock:${texto_(label)}`, Date.now() - waitStartedAt);

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function crearPerfilRendimientoQTAS_(label) {
  return {
    label: texto_(label) || 'operacion',
    startedAtMs: Date.now(),
    steps: []
  };
}

function registrarMedicionRendimientoQTAS_(profile, step, durationMs) {
  if (!profile) return null;

  const duration = Math.max(0, Math.round(Number(durationMs) || 0));
  profile.steps.push({
    step: texto_(step) || 'paso',
    ms: duration
  });

  return duration;
}

function medirBloqueRendimientoQTAS_(profile, step, fn) {
  const startedAt = Date.now();

  try {
    return fn();
  } finally {
    registrarMedicionRendimientoQTAS_(profile, step, Date.now() - startedAt);
  }
}

function finalizarPerfilRendimientoQTAS_(profile, meta) {
  if (!profile || profile.__finalized) {
    return profile && profile.result ? profile.result : null;
  }

  const totalMs = Math.max(0, Date.now() - numero_(profile.startedAtMs));
  const steps = (profile.steps || []).map(item => ({
    step: texto_(item.step),
    ms: Math.max(0, numero_(item.ms))
  }));
  const topSteps = steps
    .slice()
    .sort((a, b) => b.ms - a.ms)
    .slice(0, 5);
  const summary = [`Total ${totalMs} ms`]
    .concat(topSteps.map(item => `${item.step}: ${item.ms} ms`))
    .join(' | ');
  const result = Object.assign({
    label: texto_(profile.label) || 'operacion',
    totalMs: totalMs,
    steps: steps,
    topSteps: topSteps,
    summary: summary
  }, meta || {});

  profile.__finalized = true;
  profile.result = result;
  Logger.log(`[QTAS PERF] ${result.label} | ${summary}`);
  return result;
}

function nombreHojaUnico_(ss, baseName) {
  let name = baseName.slice(0, 99);
  let index = 1;

  while (ss.getSheetByName(name)) {
    const suffix = `_${index++}`;
    name = baseName.slice(0, 99 - suffix.length) + suffix;
  }

  return name;
}

function validarVenta_(payload) {
  if (!payload) throw new Error('Venta vacia.');
  if (!payload.cliente || !texto_(payload.cliente.nombre)) {
    throw new Error('Falta el nombre del cliente.');
  }
  if (!payload.lineas || !payload.lineas.length) {
    throw new Error('La venta debe tener al menos un producto.');
  }

  payload.lineas.forEach((linea, index) => {
    if (!texto_(linea.producto)) {
      throw new Error(`Falta producto en la linea ${index + 1}.`);
    }
    if (numero_(linea.cantidad) <= 0) {
      throw new Error(`Cantidad invalida en la linea ${index + 1}.`);
    }
    if (numero_(linea.precioVendidoUnitario) <= 0) {
      throw new Error(`Precio vendido invalido en la linea ${index + 1}.`);
    }
    if (numero_(linea.descuentoLinea) < 0) {
      throw new Error(`Descuento invalido en la linea ${index + 1}.`);
    }
  });
}

function resumenProductos_(lineas) {
  return lineas.map(linea => {
    const producto = texto_(linea.producto || linea.Producto_Estandar);
    const cantidad = formatearCantidad_(linea.cantidad || linea.Cantidad);
    const unidad = texto_(linea.unidad || linea.Unidad);
    return `${producto} ${cantidad}${unidad}`;
  }).join(' + ');
}

function formatearCantidad_(value) {
  const cantidad = redondear_(numero_(value));
  return Number.isInteger(cantidad) ? String(cantidad) : String(cantidad);
}

function agrupar_(rows, keyFn) {
  return rows.reduce((acc, row) => {
    const key = keyFn(row);
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});
}

function fecha_(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (!value) return new Date();

  const text = texto_(value);

  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    const parts = text.slice(0, 10).split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }

  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(text)) {
    const parts = text.split(' ')[0].split('/').map(Number);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? new Date() : parsed;
}

function resolverFechaOperacion_(value, fallback) {
  const source = value || fallback || new Date();

  if (source instanceof Date) {
    const iso = Utilities.formatDate(source, zonaHorariaQTAS_(), 'yyyy-MM-dd');
    const parts = iso.split('-').map(Number);
    return crearFechaDiaSeguroQTAS_(parts[0], parts[1] - 1, parts[2]);
  }

  const base = fecha_(source);
  return crearFechaDiaSeguroQTAS_(base.getFullYear(), base.getMonth(), base.getDate());
}

function resolverFechaMomentoQTAS_(fechaHora, fechaSolo, fallback) {
  const referencia = fallback instanceof Date ? fallback : new Date();
  const directa = fechaMomentoExactaQTAS_(fechaHora) || fechaMomentoExactaQTAS_(fechaSolo);

  if (directa) return directa;
  if (fechaHora) return combinarFechaYHora_(resolverFechaOperacion_(fechaHora, referencia), referencia);
  if (fechaSolo) return combinarFechaYHora_(resolverFechaOperacion_(fechaSolo, referencia), referencia);
  return combinarFechaYHora_(referencia, referencia);
}

function resolverFechaMomentoOpcionalQTAS_(fechaHora, fechaSolo, fallback) {
  if (!fechaHora && !fechaSolo) return '';
  return resolverFechaMomentoQTAS_(fechaHora, fechaSolo, fallback);
}

function fechaMomentoExactaQTAS_(value) {
  if (value instanceof Date) return new Date(value.getTime());

  const text = texto_(value);
  if (!text) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text) || /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(text)) {
    return null;
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function valorFechaVentaCanonicaQTAS_(row, fallback) {
  return resolverFechaMomentoQTAS_(
    row && (row.Fecha_Hora || row.Fecha_Venta),
    row && row.Fecha_Venta,
    fallback
  );
}

function valorFechaCompraCanonicaQTAS_(row, fallback) {
  return resolverFechaMomentoQTAS_(
    row && (row.Fecha_Hora || row.Fecha_Compra),
    row && row.Fecha_Compra,
    fallback
  );
}

function valorFechaPagoCanonicaQTAS_(row, fallback) {
  return resolverFechaMomentoQTAS_(
    row && (row.Fecha_Hora || row.Fecha_Pago),
    row && row.Fecha_Pago,
    fallback
  );
}

function valorFechaBaseCanonicaQTAS_(row, fallback) {
  return resolverFechaMomentoQTAS_(
    row && (row.Fecha_Hora_Base || row.Fecha_Base),
    row && row.Fecha_Base,
    fallback
  );
}

function valorFechaDistribucionVentaCanonicaQTAS_(row, fallback) {
  return resolverFechaMomentoOpcionalQTAS_(
    row && (row.Fecha_Hora_Venta || row.Fecha_Venta),
    row && row.Fecha_Venta,
    fallback
  );
}

function valorFechaDistribucionPagoCanonicaQTAS_(row, fallback) {
  return resolverFechaMomentoOpcionalQTAS_(
    row && (row.Fecha_Hora_Pago || row.Fecha_Pago),
    row && row.Fecha_Pago,
    fallback
  );
}

function aplicarFormatosModeloQTAS_(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  if (!spreadsheet) return;

  const formatosPorHoja = {
    Productos: {},
    Precios_Referencia: {
      Precio: '#,##0.00',
      Fecha_Desde: '@',
      Fecha_Hasta: '@'
    },
    Compras: {
      Fecha_Compra: 'yyyy-mm-dd hh:mm:ss',
      Total_Compra: '#,##0.00'
    },
    Compra_Detalle: {
      Fecha_Compra: 'yyyy-mm-dd hh:mm:ss',
      Cantidad: '#,##0.00',
      Costo_Total_Linea: '#,##0.00',
      Costo_Unitario: '#,##0.00'
    },
    Costos_Referencia: {
      Costo_Unitario: '#,##0.00',
      Fecha_Desde: '@',
      Fecha_Hasta: '@'
    },
    Producto_Componentes: {
      Orden: '0',
      Cantidad_Componente: '#,##0.0000',
      Merma_Pct: '0.00'
    },
    Producto_Reglas_Costo: {
      Fecha_Desde: '@',
      Fecha_Hasta: '@',
      Cantidad_Min: '#,##0.0000',
      Cantidad_Max: '#,##0.0000',
      Orden: '0',
      Cantidad_Componente: '#,##0.0000',
      Merma_Pct: '0.00'
    },
    Costo_Producto_Calc: {
      Fecha_Calculo: 'yyyy-mm-dd hh:mm:ss',
      Costo_Unitario_Total: '#,##0.0000',
      Costo_Unitario_Componentes: '#,##0.0000',
      Cobertura_Costo_Pct: '0.00'
    },
    Venta_Detalle_Costos_Calc: {
      Fecha_Venta: 'yyyy-mm-dd hh:mm:ss',
      Cantidad: '#,##0.0000',
      Subtotal_Neto: '#,##0.00',
      Costo_Unitario_Usado: '#,##0.0000',
      Costo_Total_Estimado: '#,##0.00',
      Margen_Bruto_Estimado: '#,##0.00',
      Margen_Porcentaje_Estimado: '0.00',
      Cobertura_Costo_Pct: '0.00',
      Actualizado_En: 'yyyy-mm-dd hh:mm:ss'
    },
    Origenes_Fondos_Reglas: {
      Fecha_Desde: '@',
      Fecha_Hasta: '@',
      Porcentaje: '0.00'
    },
    Compra_Origenes_Fondos: {
      Fecha_Compra: 'yyyy-mm-dd hh:mm:ss',
      Porcentaje: '0.00',
      Monto_Asignado: '#,##0.00'
    },
    Clientes: {
      Creado_En: 'yyyy-mm-dd hh:mm:ss',
      Actualizado_En: 'yyyy-mm-dd hh:mm:ss'
    },
    Ventas: {
      Fecha_Venta: 'yyyy-mm-dd hh:mm:ss',
      Total_Venta: '#,##0.00',
      Total_Pagado: '#,##0.00',
      Saldo: '#,##0.00',
      Steve_Pct_Venta: '0.00',
      Majo_Pct_Venta: '0.00',
      Mush_Pct_Venta: '0.00'
    },
    Venta_Detalle: {
      Fecha_Venta: 'yyyy-mm-dd hh:mm:ss',
      Cantidad: '#,##0.00',
      Precio_Lista: '#,##0.00',
      Precio_Vendido_Unitario: '#,##0.00',
      Descuento_Linea: '#,##0.00',
      Subtotal_Bruto: '#,##0.00',
      Subtotal_Neto: '#,##0.00'
    },
    Pagos: {
      Fecha_Pago: 'yyyy-mm-dd hh:mm:ss',
      Monto_Pago: '#,##0.00',
      Steve_Pct_Pago: '0.00',
      Majo_Pct_Pago: '0.00',
      Mush_Pct_Pago: '0.00'
    },
    Ventas_Envio: {
      Fecha_Pendiente_Envio: 'yyyy-mm-dd hh:mm:ss',
      Fecha_Envio: 'yyyy-mm-dd hh:mm:ss',
      Creado_En: 'yyyy-mm-dd hh:mm:ss',
      Actualizado_En: 'yyyy-mm-dd hh:mm:ss'
    },
    Distribucion_Reglas: {
      Fecha_Desde: '@',
      Fecha_Hasta: '@',
      Steve_Pct: '0.00',
      Majo_Pct: '0.00',
      Mush_Pct: '0.00'
    },
    Distribucion_Ingresos: {
      Fecha_Base: 'yyyy-mm-dd hh:mm:ss',
      Fecha_Venta: 'yyyy-mm-dd hh:mm:ss',
      Fecha_Pago: 'yyyy-mm-dd hh:mm:ss',
      Monto_Base: '#,##0.00',
      Steve_Pct: '0.00',
      Majo_Pct: '0.00',
      Mush_Pct: '0.00',
      Steve_Valor: '#,##0.00',
      Majo_Valor: '#,##0.00',
      Mush_Valor: '#,##0.00'
    },
    Config_MediosPago: {}
  };

  Object.keys(formatosPorHoja).forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;
    aplicarFormatosPorHeaderQTAS_(sheet, formatosPorHoja[sheetName]);
  });
}

function aplicarFormatosPorHeaderQTAS_(sheet, formatosPorHeader) {
  if (!sheet || !formatosPorHeader) return;

  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  const lastRow = sheet.getLastRow();
  const totalRows = Math.max(lastRow - 1, 1);

  headers.forEach((header, index) => {
    const formato = formatosPorHeader[header];
    if (!formato) return;

    sheet
      .getRange(2, index + 1, totalRows, 1)
      .setNumberFormat(formato);
  });

  const extraRows = Math.max(sheet.getMaxRows() - lastRow, 0);
  if (extraRows > 0) {
    sheet
      .getRange(lastRow + 1, 1, extraRows, headers.length)
      .clearFormat();
  }
}

function combinarFechaYHora_(fechaBase, horaBase) {
  const date = resolverFechaOperacion_(fechaBase, horaBase || new Date());
  const time = horaBase instanceof Date ? horaBase : new Date();
  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    time.getHours(),
    time.getMinutes(),
    time.getSeconds(),
    time.getMilliseconds()
  );
}

function fechaInput_(value) {
  return Utilities.formatDate(
    resolverFechaOperacion_(value, new Date()),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function diaAnterior_(value) {
  const date = resolverFechaOperacion_(value, new Date());
  const previous = new Date(date);
  previous.setDate(previous.getDate() - 1);
  return resolverFechaOperacion_(previous, previous);
}

function diaSiguiente_(value) {
  const date = resolverFechaOperacion_(value, new Date());
  const next = new Date(date);
  next.setDate(next.getDate() + 1);
  return resolverFechaOperacion_(next, next);
}

function crearFechaDiaSeguroQTAS_(year, month, day) {
  return new Date(year, month, day, 12, 0, 0, 0);
}

function esHeaderFechaSoloQTAS_(header) {
  return [
    'Fecha_Desde',
    'Fecha_Hasta',
    'Ultima_Fecha'
  ].includes(texto_(header));
}

function fechaTextoPlanoQTAS_(value) {
  if (value === '' || value === null || value === undefined) return '';
  return fechaInput_(value);
}

function zonaHorariaQTAS_() {
  try {
    return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  } catch (error) {
    return Session.getScriptTimeZone();
  }
}

function numero_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isFinite(value) ? value : 0;

  let text = String(value)
    .replace(/\s/g, '')
    .replace(/\$/g, '')
    .trim();

  if (!text) return 0;

  const hasDot = text.indexOf('.') >= 0;
  const hasComma = text.indexOf(',') >= 0;

  if (hasDot && hasComma) {
    if (text.lastIndexOf('.') > text.lastIndexOf(',')) {
      text = text.replace(/,/g, '');
    } else {
      text = text.replace(/\./g, '').replace(',', '.');
    }
  } else if (hasComma) {
    const parts = text.split(',');
    if (parts.length > 2) {
      text = text.replace(/,/g, '');
    } else if (parts[1] && parts[1].length === 3 && parts[0].length > 0) {
      text = text.replace(/,/g, '');
    } else {
      text = text.replace(',', '.');
    }
  } else if (hasDot) {
    const parts = text.split('.');
    if (parts.length > 2) {
      text = text.replace(/\./g, '');
    } else if (parts[1] && parts[1].length === 3 && parts[0].length > 0) {
      text = text.replace(/\./g, '');
    }
  }

  const parsed = Number(text);
  return isNaN(parsed) ? 0 : parsed;
}

function texto_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function normalizarClaveTexto_(value) {
  return texto_(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizarUnidadCanonicaQTAS_(value) {
  const original = texto_(value);
  const key = normalizarClaveTexto_(original);
  if (!key) return '';

  if (['und', 'unidad', 'unidades', 'unid', 'u', 'unit', 'units'].includes(key)) {
    return 'und';
  }
  if (['g', 'gr', 'grs', 'gramo', 'gramos'].includes(key)) {
    return 'g';
  }
  if (['lb', 'lbs', 'libra', 'libras', 'pound', 'pounds'].includes(key)) {
    return 'lb';
  }

  return original;
}

function normalizarCantidadUnidadQTAS_(cantidad, unidad) {
  const cantidadBase = numero_(cantidad);
  const unidadCanonica = normalizarUnidadCanonicaQTAS_(unidad);

  if (unidadCanonica === 'lb') {
    return {
      cantidad: redondear_(cantidadBase * 453.59237),
      unidad: 'g',
      unidadOriginal: texto_(unidad),
      convertido: true
    };
  }

  return {
    cantidad: redondear_(cantidadBase),
    unidad: unidadCanonica,
    unidadOriginal: texto_(unidad),
    convertido: false
  };
}

function sumar_(values) {
  return values.reduce((acc, value) => acc + numero_(value), 0);
}

function redondear_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function unirUnicos_(values) {
  return [...new Set((values || []).map(texto_).filter(Boolean))].join(' | ');
}

function moneda_(value) {
  return '$' + Math.round(Number(value) || 0).toLocaleString('es-CO');
}

function estaActivo_(value) {
  const text = texto_(value).toLowerCase();
  if (!text) return true;
  return !['false', '0', 'no', 'n', 'inactivo'].includes(text);
}

function esRegistroAnulado_(value) {
  return normalizarClaveTexto_(value) === normalizarClaveTexto_(QTAS.status.registro.anulado);
}

function obtenerEstadoPago_(totalVenta, totalPagado, estadoRegistro) {
  if (esRegistroAnulado_(estadoRegistro)) {
    return QTAS.status.pago.anulado;
  }

  const venta = redondear_(Math.max(numero_(totalVenta), 0));
  const pagado = redondear_(Math.max(numero_(totalPagado), 0));

  if (venta <= 0) return QTAS.status.pago.pendiente;
  if (pagado + 0.009 >= venta) return QTAS.status.pago.pagado;
  if (pagado > 0) return QTAS.status.pago.parcial;
  return QTAS.status.pago.pendiente;
}

function maybeAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log(message);
  }
}

function detalleId_(ventaId, lineNumber) {
  return `DET-${String(ventaId).padStart(6, '0')}-${String(lineNumber).padStart(2, '0')}`;
}
