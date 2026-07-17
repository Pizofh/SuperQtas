// TEMPORAL: retirar este archivo despues de ejecutar y validar la migracion de Compra 90.
function previsualizarMigracionCompra90CajaQTAS() {
  return migrarCompra90CajaATitularesTemporalQTAS_({ dryRun: true });
}

function aplicarMigracionCompra90CajaQTAS() {
  return migrarCompra90CajaATitularesTemporalQTAS_({ dryRun: false });
}

function migrarCompra90CajaATitularesTemporalQTAS_(options) {
  const settings = Object.assign({ dryRun: true }, options || {});

  return withScriptLock_('migrar compra 90 caja', () => {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(QTAS.sheets.compraOrigenesFondos);
    if (!sheet) {
      throw new Error(`No existe la hoja ${QTAS.sheets.compraOrigenesFondos}.`);
    }

    const headers = getHeaders_(sheet);
    const requiredHeaders = [
      'Compra_Origen_ID',
      'Compra_ID',
      'Compra_Detalle_ID',
      'Fecha_Compra',
      'Origen_Fondos',
      'Aportante',
      'Porcentaje',
      'Monto_Asignado',
      'Fuente_Registro',
      'Nota'
    ];
    const missingHeaders = requiredHeaders.filter(header => headers.indexOf(header) < 0);
    if (missingHeaders.length) {
      throw new Error(`Faltan encabezados requeridos: ${missingHeaders.join(', ')}.`);
    }

    const rows = leerObjetosConMeta_(sheet);
    const totalAntes = totalMontoCompraOrigenesFondosTemporalQTAS_(rows);
    const candidatas = rows.filter(row =>
      numero_(row.Compra_ID) === 90 &&
      texto_(row.Compra_Detalle_ID) === 'COMDET-000090-01' &&
      texto_(row.Origen_Fondos) === 'Caja' &&
      normalizarClaveTexto_(row.Aportante) === 'caja' &&
      Math.abs(numero_(row.Monto_Asignado) - 30000) < 0.01
    );

    if (!candidatas.length) {
      const alreadyApplied = validarMigracionCompra90CajaTemporalQTAS_(rows, totalAntes);
      if (alreadyApplied.ok) {
        return {
          ok: true,
          applied: false,
          alreadyApplied: true,
          dryRun: Boolean(settings.dryRun),
          validation: alreadyApplied
        };
      }
      throw new Error(
        'No se encontro exactamente la fila objetivo de Compra 90 y la migracion no parece aplicada.'
      );
    }

    if (candidatas.length !== 1) {
      throw new Error(
        `Se esperaban 1 fila objetivo para Compra 90, pero se encontraron ${candidatas.length}. No se modifico nada.`
      );
    }

    const original = candidatas[0];
    const replacements = construirReemplazosCompra90CajaTemporalQTAS_(original);
    const rowsProyectadas = rows
      .filter(row => row.__rowNumber !== original.__rowNumber)
      .concat(replacements);
    const validacionProyectada = validarMigracionCompra90CajaTemporalQTAS_(rowsProyectadas, totalAntes);

    if (!validacionProyectada.ok) {
      throw new Error(
        `La prevalidacion de Compra 90 fallo: ${validacionProyectada.errors.join(' | ')}. No se modifico nada.`
      );
    }

    const preview = {
      filaOriginal: resumenFilaCompraOrigenTemporalQTAS_(original),
      filasNuevas: replacements.map(resumenFilaCompraOrigenTemporalQTAS_),
      totalAntes: totalAntes,
      totalDespues: totalMontoCompraOrigenesFondosTemporalQTAS_(rowsProyectadas),
      validacion: validacionProyectada
    };

    if (settings.dryRun !== false) {
      return {
        ok: true,
        applied: false,
        dryRun: true,
        preview: preview
      };
    }

    actualizarFilaObjeto_(sheet, original.__rowNumber, headers, replacements[0]);
    sheet.insertRowsAfter(original.__rowNumber, replacements.length - 1);
    escribirFilasDesdeFilaQTAS_(
      sheet,
      original.__rowNumber + 1,
      replacements.slice(1).map(row => filaDesdeHeaders_(headers, row))
    );
    SpreadsheetApp.flush();

    const validation = validarMigracionCompra90CajaTemporalQTAS_(
      leerObjetosConMeta_(sheet),
      totalAntes
    );
    if (!validation.ok) {
      throw new Error(
        `La validacion posterior de Compra 90 fallo: ${validation.errors.join(' | ')}.`
      );
    }

    return {
      ok: true,
      applied: true,
      dryRun: false,
      preview: preview,
      validation: validation
    };
  });
}

function construirReemplazosCompra90CajaTemporalQTAS_(original) {
  const detalleId = texto_(original.Compra_Detalle_ID);
  const repartos = [
    { aportante: 'Steve', porcentaje: 40, monto: 12000 },
    { aportante: 'Majo', porcentaje: 40, monto: 12000 },
    { aportante: 'Mush', porcentaje: 20, monto: 6000 }
  ];

  return repartos.map(item => Object.assign({}, original, {
    Compra_Origen_ID: construirCompraOrigenIdQTAS_(detalleId, item.aportante),
    Aportante: item.aportante,
    Porcentaje: item.porcentaje,
    Monto_Asignado: item.monto
  }));
}

function validarMigracionCompra90CajaTemporalQTAS_(rows, totalAntes) {
  const compraRows = (rows || []).filter(row => numero_(row.Compra_ID) === 90);
  const expected = {
    Steve: { porcentaje: 40, monto: 12000 },
    Majo: { porcentaje: 40, monto: 12000 },
    Mush: { porcentaje: 20, monto: 6000 }
  };
  const errors = [];
  const aportantes = {};

  compraRows.forEach(row => {
    const aportante = normalizarAportanteOrigenFondosQTAS_(row.Aportante);
    if (aportante) aportantes[aportante] = row;
  });

  if (compraRows.some(row => normalizarClaveTexto_(row.Aportante) === 'caja')) {
    errors.push('Compra 90 aun conserva Caja como aportante.');
  }
  if (compraRows.length !== 3) {
    errors.push(`Compra 90 debe tener exactamente 3 filas y tiene ${compraRows.length}.`);
  }

  Object.keys(expected).forEach(aportante => {
    const row = aportantes[aportante];
    if (!row) {
      errors.push(`Falta la fila de ${aportante}.`);
      return;
    }
    if (texto_(row.Compra_Detalle_ID) !== 'COMDET-000090-01') {
      errors.push(`${aportante} no conserva Compra_Detalle_ID.`);
    }
    if (texto_(row.Origen_Fondos) !== 'Caja') {
      errors.push(`${aportante} no conserva Origen_Fondos Caja.`);
    }
    if (Math.abs(numero_(row.Porcentaje) - expected[aportante].porcentaje) > 0.01) {
      errors.push(`${aportante} tiene un porcentaje distinto al esperado.`);
    }
    if (Math.abs(numero_(row.Monto_Asignado) - expected[aportante].monto) > 0.01) {
      errors.push(`${aportante} tiene un monto distinto al esperado.`);
    }
  });

  const porcentajeTotal = redondear_(sumar_(compraRows.map(row => numero_(row.Porcentaje))));
  const montoCompra = redondear_(sumar_(compraRows.map(row => numero_(row.Monto_Asignado))));
  if (Math.abs(porcentajeTotal - 100) > 0.01) errors.push('Los porcentajes de Compra 90 no suman 100%.');
  if (Math.abs(montoCompra - 30000) > 0.01) errors.push('Los montos de Compra 90 no suman 30000.');

  const ids = {};
  const duplicateIds = [];
  (rows || []).forEach(row => {
    const id = texto_(row.Compra_Origen_ID);
    if (!id) return;
    if (ids[id]) duplicateIds.push(id);
    ids[id] = true;
  });
  if (duplicateIds.length) errors.push(`Hay identificadores duplicados: ${unirUnicos_(duplicateIds)}.`);

  const totalDespues = totalMontoCompraOrigenesFondosTemporalQTAS_(rows);
  if (Math.abs(totalDespues - numero_(totalAntes)) > 0.01) {
    errors.push('El total general de Monto_Asignado cambio.');
  }

  const totalesPorAportante = totalPorAportanteCompraOrigenesFondosTemporalQTAS_(rows);
  const esperadosFinales = {
    Steve: 13953575,
    Majo: 14418425,
    Mush: 34012500
  };
  Object.keys(esperadosFinales).forEach(aportante => {
    if (Math.abs(numero_(totalesPorAportante[aportante]) - esperadosFinales[aportante]) > 0.01) {
      errors.push(`El total final de ${aportante} no coincide con el esperado.`);
    }
  });

  return {
    ok: errors.length === 0,
    errors: errors,
    compraRows: compraRows.map(resumenFilaCompraOrigenTemporalQTAS_),
    porcentajes: porcentajeTotal,
    montoCompra: montoCompra,
    totalGeneral: totalDespues,
    totalesPorAportante: totalesPorAportante,
    duplicateIds: unirUnicos_(duplicateIds)
  };
}

function totalMontoCompraOrigenesFondosTemporalQTAS_(rows) {
  return redondear_(sumar_((rows || []).map(row => numero_(row.Monto_Asignado))));
}

function totalPorAportanteCompraOrigenesFondosTemporalQTAS_(rows) {
  return (rows || []).reduce((totales, row) => {
    const aportante = normalizarAportanteOrigenFondosQTAS_(row.Aportante);
    if (!aportante) return totales;
    totales[aportante] = redondear_((totales[aportante] || 0) + numero_(row.Monto_Asignado));
    return totales;
  }, {});
}

function resumenFilaCompraOrigenTemporalQTAS_(row) {
  return {
    compraOrigenId: texto_(row.Compra_Origen_ID),
    compraId: numero_(row.Compra_ID),
    compraDetalleId: texto_(row.Compra_Detalle_ID),
    fechaCompra: fechaTextoPlanoQTAS_(row.Fecha_Compra),
    origenFondos: texto_(row.Origen_Fondos),
    aportante: texto_(row.Aportante),
    porcentaje: redondear_(numero_(row.Porcentaje)),
    montoAsignado: redondear_(numero_(row.Monto_Asignado)),
    fuenteRegistro: texto_(row.Fuente_Registro),
    nota: texto_(row.Nota)
  };
}
