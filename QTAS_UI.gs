function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('QTAS ERP')
    .addItem('Abrir ERP', 'mostrarAppQTAS')
    .addToUi();
}

function crearSalidaAppQTAS_(vistaInicial) {
  const template = HtmlService.createTemplateFromFile('App');
  template.vistaInicial = vistaInicial || 'ventas';

  return template
    .evaluate()
    .setTitle('QTAS ERP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGet(e) {
  const vista = e && e.parameter ? e.parameter.vista : 'ventas';
  return crearSalidaAppQTAS_(vista);
}

function mostrarAppQTAS() {
  const html = crearSalidaAppQTAS_('ventas')
    .setWidth(1280)
    .setHeight(920);

  SpreadsheetApp.getUi().showModalDialog(html, 'QTAS ERP');
}

function getCatalogoQTAS(fechaVenta) {
  validarModeloSoloLecturaQTAS_({
    sheetNames: [
      QTAS.sheets.productos,
      QTAS.sheets.precios,
      QTAS.sheets.clientes
    ],
    validarConfig: true
  });

  const ss = SpreadsheetApp.getActive();
  const fechaConsulta = resolverFechaOperacion_(fechaVenta, new Date());
  const productos = leerObjetos_(ss.getSheetByName(QTAS.sheets.productos))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      producto: texto_(row.Producto_Estandar),
      unidad: texto_(row.Unidad_Default)
  }));

  const priceCache = cargarPreciosReferenciaEnMemoria_();
  cargarReglasDistribucionEnMemoria_();

  const productosConPrecio = productos.map(producto => ({
    producto: producto.producto,
    unidad: producto.unidad,
    precio: obtenerPrecioVigenteDesdeCache_(
      priceCache,
      producto.producto,
      producto.unidad,
      fechaConsulta
    )
  }));

  const mediosPago = leerMediosPagoConfiguradosQTAS_()
    .filter(row => row.activo)
    .map(row => row.medioPago);

  const clientes = leerObjetos_(ss.getSheetByName(QTAS.sheets.clientes))
    .filter(row => estaActivo_(row.Activo))
    .map(row => ({
      clienteId: texto_(row.Cliente_ID),
      nombre: texto_(row.Nombre)
    }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));

  return {
    productos: productosConPrecio,
    mediosPago,
    clientes,
    hoy: fechaInput_(new Date())
  };
}

function getDeudoresQTAS() {
  validarModeloSoloLecturaQTAS_({
    sheetNames: [QTAS.sheets.ventas, QTAS.sheets.pagos],
    validarConfig: false
  });
  return dashboardVentasConsistenteQTAS_().deudores;
}

function getDashboardQTAS() {
  validarModeloSoloLecturaQTAS_({
    sheetNames: [QTAS.sheets.ventas, QTAS.sheets.pagos],
    validarConfig: false
  });
  return dashboardVentasConsistenteQTAS_();
}
